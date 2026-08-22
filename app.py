# -*- coding: utf-8 -*-
import streamlit as st
import zipfile, re, io, os
from datetime import datetime
import openpyxl

st.set_page_config(page_title='Cert Report Generator', page_icon='📋', layout='wide')
st.title('认证报告生成器')

def sanitize(name):
    return re.sub(r'[\\/:*?<>|]', '_', str(name))

def format_date(val):
    if val is None: return ''
    if isinstance(val, datetime): return val.strftime('%Y-%m-%d')
    if isinstance(val, (int, float)):
        try: return (datetime(1899, 12, 30) + __import__('datetime').timedelta(days=int(val))).strftime('%Y-%m-%d')
        except: return ''
    s = str(val).strip()
    if not s or s == '#N/A': return ''
    m = re.search(r'(\d{4})[年\-\/](\d{1,2})[月\-\/](\d{1,2})', s)
    if m: return f'{m.group(1)}-{m.group(2).zfill(2)}-{m.group(3).zfill(2)}'
    m = re.search(r'(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})', s)
    if m:
        y = int(m.group(3))
        if y < 100: y += 2000
        return f'{y}-{m.group(1).zfill(2)}-{m.group(2).zfill(2)}'
    return s

def count_data_rows(ws, company_col):
    last_data = 0
    count = 0
    for ri, row in enumerate(ws.iter_rows(min_row=2), 2):
        vals = [c.value for c in row]
        if any(v is not None for v in vals):
            last_data = ri
        if vals[company_col] is not None:
            count += 1
    return count

def set_cb(cell_xml, target):
    """Set checkbox checked for target text."""
    cb_off = chr(9633)  # □ U+25A1
    cb_on = chr(9745)   # ☑ U+2611
    cell_xml = cell_xml.replace(cb_off + target, cb_on + target)
    for m in re.finditer(re.escape(target), cell_xml):
        pos = m.start()
        open_tag = cell_xml.rfind('<w:t', 0, pos)
        if open_tag < 0:
            open_tag = cell_xml.rfind('<w:t>', 0, pos)
        if open_tag >= 0:
            close_tag = cell_xml.find('</w:t>', pos)
            if close_tag > pos:
                tag_match = re.match(r'<w:t[^>]*>', cell_xml[open_tag:])
                tag_len = len(tag_match.group()) if tag_match else 5
                content = cell_xml[open_tag+tag_len:close_tag]
                if cb_off in content:
                    cell_xml = cell_xml[:open_tag+tag_len] + content.replace(cb_off, cb_on, 1) + cell_xml[close_tag:]
                    break
    return cell_xml

def unset_cb(cell_xml, target):
    """Unset checkbox for target text."""
    cb_off = chr(9633)  # □ U+25A1
    cb_on = chr(9745)   # ☑ U+2611
    cell_xml = cell_xml.replace(cb_on + target, cb_off + target)
    for m in re.finditer(re.escape(target), cell_xml):
        pos = m.start()
        open_tag = cell_xml.rfind('<w:t', 0, pos)
        if open_tag < 0:
            open_tag = cell_xml.rfind('<w:t>', 0, pos)
        if open_tag >= 0:
            close_tag = cell_xml.find('</w:t>', pos)
            if close_tag > pos:
                tag_match = re.match(r'<w:t[^>]*>', cell_xml[open_tag:])
                tag_len = len(tag_match.group()) if tag_match else 5
                content = cell_xml[open_tag+tag_len:close_tag]
                if cb_on in content:
                    cell_xml = cell_xml[:open_tag+tag_len] + content.replace(cb_on, cb_off, 1) + cell_xml[close_tag:]
                    break
    return cell_xml

def set_form_checkbox(xml_str, index):
    pattern = r'<w:fldChar[^>]*w:fldCharType="begin"[^>]*>.*?<w:fldChar[^>]*w:fldCharType="end"[^>]*>'
    matches = list(re.finditer(pattern, xml_str, re.DOTALL))
    if index < len(matches):
        cb = matches[index].group(0)
        new_cb = re.sub(r'w:checked w:val="0"', 'w:checked w:val="1"', cb)
        xml_str = xml_str[:matches[index].start()] + new_cb + xml_str[matches[index].end():]
    return xml_str

def unset_form_checkbox(xml_str, index):
    pattern = r'<w:fldChar[^>]*w:fldCharType="begin"[^>]*>.*?<w:fldChar[^>]*w:fldCharType="end"[^>]*>'
    matches = list(re.finditer(pattern, xml_str, re.DOTALL))
    if index < len(matches):
        cb = matches[index].group(0)
        new_cb = re.sub(r'w:checked w:val="1"', 'w:checked w:val="0"', cb)
        xml_str = xml_str[:matches[index].start()] + new_cb + xml_str[matches[index].end():]
    return xml_str

def extract_form_data(form_bytes):
    with zipfile.ZipFile(io.BytesIO(form_bytes)) as z:
        xml = z.read('word/document.xml').decode('utf-8')
    texts = re.findall(r'<w:t[^>]*>([^<]*)</w:t>', xml)
    data = {}
    i = 0
    while i < len(texts):
        t = texts[i].strip()
        if not data.get('company') and t in ('组织名称', '公司名称'):
            if i + 1 < len(texts):
                nxt = texts[i + 1].strip()
                if nxt: data['company'] = nxt
        elif not data.get('taskNo') and t in ('任务编号', '任务号'):
            if i + 1 < len(texts):
                nxt = texts[i + 1].strip()
                if nxt and not any(k in nxt for k in ['合同', '备注', '认证', '审核']):
                    data['taskNo'] = nxt
        elif not data.get('leader') and ('审核组长' in t or t == '组长') and i + 1 < len(texts):
            nxt = texts[i + 1].strip()
            if nxt and len(nxt) < 30:
                data['leader'] = nxt.split('+')[0].strip()
        elif not data.get('address') and '审核地址' in t and i + 1 < len(texts):
            nxt = texts[i + 1].strip()
            if nxt: data['address'] = nxt
        elif not data.get('scope') and '认证范围' in t and i + 1 < len(texts):
            nxt = texts[i + 1].strip()
            if nxt: data['scope'] = nxt
        elif not data.get('auditType') and '审核性质' in t and i + 1 < len(texts):
            nxt = texts[i + 1].strip()
            if nxt: data['auditType'] = nxt
        elif not data.get('date') and '现场审核日期' in t and i + 1 < len(texts):
            nxt = texts[i + 1].strip()
            if nxt:
                m = re.search(r'(\d{4})[\/-]\d{1,2}[\/-]\d{1,2}', nxt)
                if m: data['date'] = m.group(0)
        i += 1
    return data

def parse_template(template_bytes):
    with zipfile.ZipFile(io.BytesIO(template_bytes)) as zfin:
        contents = {n: zfin.read(n) for n in zfin.namelist()}
    xml_str = contents['word/document.xml'].decode('utf-8')
    tbl_start = xml_str.find('<w:tbl>')
    tbl_match = re.search(r'<w:tbl[^>]*>(?:(?!</w:tbl>).)*</w:tbl>', xml_str[tbl_start:], re.DOTALL)
    if tbl_match is None: raise Exception('No table found in template')
    tbl_xml = xml_str[tbl_start:tbl_start + tbl_match.end()]
    row_parts = re.split(r'(<w:tr[^>]*>.*?</w:tr>)', tbl_xml, flags=re.DOTALL)
    trs = [p for p in row_parts if p.startswith('<w:tr')]
    if len(trs) < 22: raise Exception(f'Expected 22+ rows, got {len(trs)}')
    other_files = {n: v for n, v in contents.items() if n != 'word/document.xml'}
    return other_files, xml_str, list(trs)

def fill_one_row(original_trs, data):
    company = str(data.get('company', '')).strip()
    task_no = str(data.get('taskNo', '')).strip()
    leader = str(data.get('leader', '')).strip()
    audit_type = str(data.get('auditType', '')).strip()
    scope = str(data.get('scope', '')).strip()
    conclusion = str(data.get('conclusion', '')).strip()
    address = str(data.get('address', '')).strip()
    cert_standards = str(data.get('certStandards', '')).strip()
    trs = list(original_trs)
    # Row 0: Company name and task number
    cells0 = re.findall(r'(<w:tc[^>]*>.*?</w:tc>)', trs[0])
    if len(cells0) >= 2:
        c0, c1 = cells0[0], cells0[1]
        m = re.search(r'<w:t[^>]*>公司名称[^<]*</w:t>', c0)
        if not m: m = re.search(r'<w:t[^>]*>组织名称[^<]*</w:t>', c0)
        if m:
            after = c0[m.end():]
            tm = re.search(r'<w:t[^>]*>([^<]*)</w:t>', after)
            if tm:
                ph = tm.group(1)
                if ph and ph.strip() and ph not in ('', ':', '\uff1a'):
                    c0 = c0[:tm.start()] + f'<w:t>{company}</w:t>' + c0[tm.end():]
                else:
                    c0 = c0[:m.end()] + f'<w:t>{company}</w:t>' + c0[m.end():]
            else:
                c0 = c0[:m.end()] + f'<w:t>{company}</w:t>' + c0[m.end():]
        m = re.search(r'<w:t[^>]*>任务号[^<]*</w:t>', c1)
        if not m: m = re.search(r'<w:t[^>]*>任务编号[^<]*</w:t>', c1)
        if m:
            after = c1[m.end():]
            tm = re.search(r'<w:t[^>]*>([^<]*)</w:t>', after)
            if tm:
                ph = tm.group(1)
                if ph and ph.strip() and ph not in ('', '\uff1a'):
                    c1 = c1[:tm.start()] + f'<w:t>{task_no}</w:t>' + c1[tm.end():]
                else:
                    c1 = c1[:m.end()] + f'<w:t>{task_no}</w:t>' + c1[m.end():]
            else:
                c1 = c1[:m.end()] + f'<w:t>{task_no}</w:t>' + c1[m.end():]
        trs[0] = re.sub(r'(<w:tc[^>]*>.*?</w:tc>){2}', lambda m: c0+c1, trs[0], count=1, flags=re.DOTALL)
    # Row 1: Audit team leader
    cells1 = re.findall(r'(<w:tc[^>]*>.*?</w:tc>)', trs[1])
    if len(cells1) >= 2:
        c0 = cells1[0]
        c1 = cells1[1]
        m = re.search(r'<w:t[^>]*>审核组长[^<]*</w:t>', c0)
        if not m: m = re.search(r'<w:t[^>]*>组长[^<]*</w:t>', c0)
        if m:
            tm = re.search(r'<w:t[^>]*>([^<]*)</w:t>', c1)
            if tm:
                ph = tm.group(1)
                if ph and ph.strip():
                    c1 = c1[:tm.start()] + f'<w:t>{leader}</w:t>' + c1[tm.end():]
                else:
                    lm = re.search(r'<w:t[^>]*>审核组长[^<]*</w:t>', c1)
                    if not lm: lm = re.search(r'<w:t[^>]*>组长[^<]*</w:t>', c1)
                    if lm:
                        c1 = c1[:lm.end()] + f'<w:t>{leader}</w:t>' + c1[lm.end():]
                    else:
                        c1 = c1.replace('</w:tc>', f'<w:t>{leader}</w:t></w:tc>', 1)
            else:
                c1 = c1.replace('</w:tc>', f'<w:t>{leader}</w:t></w:tc>', 1)
        trs[1] = re.sub(r'(<w:tc[^>]*>.*?</w:tc>){2}', lambda m: cells1[0]+c1, trs[1], count=1, flags=re.DOTALL)
    # Row 2: Cert standards (unicode checkboxes)
    cells2 = re.findall(r'(<w:tc[^>]*>.*?</w:tc>)', trs[2])
    if len(cells2) >= 2:
        has_iatf = 'IATF' in cert_standards or 'IATF' in scope
        has_iso = 'ISO' in cert_standards or ('ISO' in scope and 'IATF' not in scope)
        c2 = cells2[1]
        if has_iatf: c2 = set_cb(c2, 'IATF16949:2016')
        else: c2 = unset_cb(c2, 'IATF16949:2016')
        if has_iso: c2 = set_cb(c2, 'ISO9001:2015')
        else: c2 = unset_cb(c2, 'ISO9001:2015')
        cells2[1] = c2
        trs[2] = re.sub(r'(<w:tc[^>]*>.*?</w:tc>){2}', lambda m: cells2[0]+cells2[1], trs[2], count=1, flags=re.DOTALL)
    # Row 3: Audit type (unicode checkboxes)
    cells3 = re.findall(r'(<w:tc[^>]*>.*?</w:tc>)', trs[3])
    if len(cells3) >= 2:
        c3 = cells3[1]
        is_initial = '二阶段' in audit_type or '一阶段' in audit_type
        is_surv = '监' in audit_type
        is_recert = '再认证' in audit_type
        is_transfer = '转移' in audit_type
        is_special = '特殊' in audit_type
        c3 = unset_cb(c3, '初审')
        c3 = unset_cb(c3, '监')
        c3 = unset_cb(c3, '再认证/转移')
        c3 = unset_cb(c3, '特殊审核')
        c3 = unset_cb(c3, '其它')
        if is_initial:
            c3 = set_cb(c3, '初审')
        elif is_surv:
            c3 = set_cb(c3, '监')
        if is_recert or is_transfer:
            c3 = set_cb(c3, '再认证/转移')
        elif is_special:
            c3 = set_cb(c3, '特殊审核')
        cells3[1] = c3
        trs[3] = re.sub(r'(<w:tc[^>]*>.*?</w:tc>){2}', lambda m: cells3[0]+cells3[1], trs[3], count=1, flags=re.DOTALL)
    # Row 4: Audit address
    cells4 = re.findall(r'(<w:tc[^>]*>.*?</w:tc>)', trs[4])
    if len(cells4) >= 2:
        c4 = cells4[1]
        m = re.search(r'<w:t[^>]*>审核地址[^<]*</w:t>', c4)
        if m:
            after = c4[m.end():]
            tm = re.search(r'<w:t[^>]*>([^<]*)</w:t>', after)
            if tm:
                ph = tm.group(1)
                if ph and ph.strip():
                    c4 = c4[:tm.start()] + f'<w:t>{address}</w:t>' + c4[tm.end():]
                else:
                    c4 = c4[:m.end()] + f'<w:t>{address}</w:t>' + c4[m.end():]
            else:
                c4 = c4[:m.end()] + f'<w:t>{address}</w:t>' + c4[m.end():]
        trs[4] = re.sub(r'(<w:tc[^>]*>.*?</w:tc>){2}', lambda m: cells4[0]+c4, trs[4], count=1, flags=re.DOTALL)
    # Row 5: Certification scope
    cells5 = re.findall(r'(<w:tc[^>]*>.*?</w:tc>)', trs[5])
    if len(cells5) >= 2:
        c5 = cells5[1]
        m = re.search(r'<w:t[^>]*>认证范围[^<]*</w:t>', c5)
        if m:
            after = c5[m.end():]
            tm = re.search(r'<w:t[^>]*>([^<]*)</w:t>', after)
            if tm:
                ph = tm.group(1)
                if ph and ph.strip():
                    c5 = c5[:tm.start()] + f'<w:t>{scope}</w:t>' + c5[tm.end():]
                else:
                    c5 = c5[:m.end()] + f'<w:t>{scope}</w:t>' + c5[m.end():]
            else:
                c5 = c5[:m.end()] + f'<w:t>{scope}</w:t>' + c5[m.end():]
        trs[5] = re.sub(r'(<w:tc[^>]*>.*?</w:tc>){2}', lambda m: cells5[0]+c5, trs[5], count=1, flags=re.DOTALL)
    # Row 21: Conclusion checkboxes (form fields)
    cells21 = re.findall(r'(<w:tc[^>]*>.*?</w:tc>)', trs[21])
    if len(cells21) >= 2:
        c21 = cells21[1]
        for idx in range(7):
            c21 = unset_form_checkbox(c21, idx)
        is_initial = '二阶段' in audit_type or '一阶段' in audit_type or '再认证' in audit_type
        is_surv = '监' in audit_type
        is_transfer = '转移' in audit_type
        if is_initial:
            c21 = set_form_checkbox(c21, 0)
        elif is_transfer:
            c21 = set_form_checkbox(c21, 2)
        elif is_surv:
            if '不换证' in conclusion:
                c21 = set_form_checkbox(c21, 4)
            elif '换发' in conclusion:
                c21 = set_form_checkbox(c21, 5)
            else:
                c21 = set_form_checkbox(c21, 0)
        cells21[1] = c21
        trs[21] = re.sub(r'(<w:tc[^>]*>.*?</w:tc>){2}', lambda m: cells21[0]+c21, trs[21], count=1, flags=re.DOTALL)
    return trs

def build_docx(other_files, original_doc_xml, trs):
    tbl_start = original_doc_xml.find('<w:tbl>')
    tbl_match = re.search(r'<w:tbl[^>]*>(?:(?!</w:tbl>).)*</w:tbl>', original_doc_xml[tbl_start:], re.DOTALL)
    tbl_end = tbl_start + tbl_match.end()
    new_tbl = '<w:tbl>' + trs[0] + ''.join(trs[1:]) + '</w:tbl>'
    new_doc_xml = original_doc_xml[:tbl_start] + new_tbl + original_doc_xml[tbl_end:]
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
        zout.writestr('word/document.xml', new_doc_xml.encode('utf-8'))
        for name, content in other_files.items():
            zout.writestr(name, content)
    return buf.getvalue()

def fill_template(template_bytes, data):
    other_files, original_doc_xml, original_trs = parse_template(template_bytes)
    trs = fill_one_row(original_trs, data)
    return build_docx(other_files, original_doc_xml, trs)

mode = st.radio('选择模式', ['Single Report', 'Batch Generation'], horizontal=True)

if mode == 'Single Report':
    st.header('Single Report Generation')
    col1, col2 = st.columns(2)
    with col1:
        form_file = st.file_uploader('FORM6101 (.docx)', type=['docx'], key='s1')
    with col2:
        tpl_file = st.file_uploader('Report Template (.docx)', type=['docx'], key='s2')
    data = None
    if form_file and tpl_file:
        try:
            form_bytes = form_file.getvalue()
            data = extract_form_data(form_bytes)
        except Exception as e:
            st.error('Parse failed: ' + str(e))
            data = None
    if data:
        st.markdown('### Extracted Fields')
        for k, v in data.items():
            st.text('**' + k + '**: ' + str(v or '(not found)'))
        if st.button('Generate Report', type='primary'):
            try:
                r = fill_template(tpl_file.getvalue(), data)
                ts = datetime.now().strftime('%Y%m%d_%H%M')
                fn = sanitize(data.get('company', 'Report')) + '_' + ts + '.docx'
                st.download_button('Download', data=r, file_name=fn,
                    mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                st.success('Generated!')
            except Exception as e:
                st.error('Failed: ' + str(e))
    elif form_file or tpl_file:
        st.info('Please upload both files')

else:
    st.header('Batch Generation')
    col1, col2 = st.columns(2)
    with col1:
        excel_file = st.file_uploader('Cert Excel (.xlsx)', type=['xlsx'], key='b1')
    with col2:
        tpl_file2 = st.file_uploader('Report Template (.docx)', type=['docx'], key='b2')
    BATCH_SIZE = 20
    for k in ['batch_index','batch_total','batch_parsed','batch_done','_excel_key']:
        if k not in st.session_state: st.session_state[k] = {'batch_index':0,'batch_total':0,'batch_parsed':False,'batch_done':False,'_excel_key':''}[k]
    if excel_file and tpl_file2:
        try:
            if not st.session_state.get('batch_parsed', False):
                with st.spinner('Parsing template...'):
                    of,od,ot = parse_template(tpl_file2.getvalue())
                st.session_state['batch_other_files'] = of
                st.session_state['batch_original_doc_xml'] = od
                st.session_state['batch_original_trs'] = ot
                st.session_state['batch_parsed'] = True
            if st.session_state.get('_excel_key') != excel_file.name:
                excel_bytes = excel_file.getvalue()
                wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True)
                ws = wb.active
                headers = [c.value for c in ws[1]]
                ncols = ws.max_column
                if ncols >= 15:
                    cc,lc,tc = 3,4,5
                    ac,sc,tkc = 12,13,14
                    conc,dtc = None,None
                    total = count_data_rows(ws, cc)
                    fmt = 'A'
                else:
                    cc,lc,tc = 2,3,4
                    ac,sc,tkc = 6,7,8
                    conc,dtc = 10,11
                    total = count_data_rows(ws, cc)
                    fmt = 'B'
                st.write('Total data rows: ' + str(total))
                st.session_state['batch_total'] = total
                st.session_state['_excel_key'] = excel_file.name
                st.session_state['batch_index'] = 0
                st.session_state['batch_done'] = False
                st.session_state['batch_excel_bytes'] = excel_bytes
                st.session_state.pop('batch_zip_buf', None)
                st.session_state.pop('batch_data_rows', None)
                st.session_state['batch_col_fmt'] = fmt
            total = st.session_state.get('batch_total', 0)
            index = st.session_state.get('batch_index', 0)
            st.write('Progress: ' + str(index) + ' / ' + str(total))
            if total > 0: st.progress(index / total)
            col_a,col_b,col_c = st.columns(3)
            with col_a:
                gen_btn = st.button('Generate Next '+str(BATCH_SIZE), type='primary', disabled=(index>=total))
            with col_b:
                zb = st.session_state.get('batch_zip_buf')
                if zb is not None and index > 0:
                    zb.seek(0)
                    st.download_button('Download ZIP', data=zb, file_name='reports_'+str(index)+'.zip', mime='application/zip')
            with col_c:
                if st.button('Clear & Start Over'):
                    for k in ['batch_index','batch_total','batch_parsed','batch_done','_excel_key','batch_other_files','batch_original_doc_xml','batch_original_trs','batch_excel_bytes','batch_zip_buf','batch_data_rows']:
                        st.session_state.pop(k, None)
                    st.rerun()
            if gen_btn and index < total:
                with st.spinner('Generating '+str(BATCH_SIZE)+' reports...'):
                    of=st.session_state['batch_other_files']
                    od=st.session_state['batch_original_doc_xml']
                    ot=st.session_state['batch_original_trs']
                    eb=st.session_state['batch_excel_bytes']
                    end_idx = min(index+BATCH_SIZE, total)
                    fmt=st.session_state.get('batch_col_fmt','B')
                    cc=3 if fmt=='A' else 2
                    lc,tc = (4,5) if fmt=='A' else (3,4)
                    ac,sc,tkc = (12,13,14) if fmt=='A' else (6,7,8)
                    conc,dtc = (None,None) if fmt=='A' else (10,11)
                    wb=openpyxl.load_workbook(io.BytesIO(eb),data_only=True)
                    ws=wb.active
                    data_rows=[]
                    for row in ws.iter_rows(min_row=2):
                        vals=[c.value for c in row]
                        if vals[cc]:
                            data_rows.append(vals)
                    st.session_state['batch_data_rows']=data_rows
                    buf=io.BytesIO()
                    with zipfile.ZipFile(buf,'w',compression=zipfile.ZIP_DEFLATED) as z:
                        for di in range(index, min(end_idx,len(data_rows))):
                            row=data_rows[di]
                            company=row[cc];audit_team=row[lc];audit_type=row[tc]
                            audit_address=row[ac];cert_scope=row[sc];task_no=row[tkc]
                            conclusion=row[conc] if conc is not None else None
                            date_val=row[dtc] if dtc is not None else None
                            ds=format_date(date_val)
                            leader=str(audit_team).split('+')[0].strip() if audit_team else ''
                            d={'company':str(company) if company else '',
                               'taskNo':str(task_no) if task_no else '',
                               'leader':leader,
                               'auditType':str(audit_type) if audit_type else '',
                               'address':str(audit_address) if audit_address else '',
                               'scope':str(cert_scope) if cert_scope else '',
                               'date':ds,
                               'conclusion':str(conclusion) if conclusion else ''}
                            try:
                                rt=fill_one_row(ot,d)
                                rb=build_docx(of,od,rt)
                                z.writestr(sanitize(str(company))+'.docx',rb)
                            except Exception as e:
                                st.text('Error: '+str(company)+' - '+str(e))
                    buf.seek(0)
                    st.session_state['batch_zip_buf']=buf
                    st.session_state['batch_index']=end_idx
                    if end_idx>=len(data_rows): st.session_state['batch_done']=True
                    st.rerun()
            if st.session_state.get('batch_done'):
                st.success('All '+str(total)+' reports generated!')
                zb=st.session_state.get('batch_zip_buf')
                if zb is not None:
                    zb.seek(0)
                    st.download_button('Download All (ZIP)',data=zb,file_name='reports_all.zip',mime='application/zip')
        except Exception as e:
            st.error('Error: '+str(e))
    elif excel_file or tpl_file2:
        st.info('Please upload both files')

