# -*- coding: utf-8 -*-
import streamlit as st
import zipfile, re, io, os, gc
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

def set_form_checkbox(xml_str, index):
    pattern = r'<w:fldChar[^>]*w:fldCharType="begin"[^>]*>.*?</w:fldChar>'
    matches = list(re.finditer(pattern, xml_str, re.DOTALL))
    if index < len(matches):
        cb = matches[index].group(0)
        new_cb = re.sub(r'w:checked w:val="0"', 'w:checked w:val="1"', cb)
        xml_str = xml_str[:matches[index].start()] + new_cb + xml_str[matches[index].end():]
    return xml_str

def replace_unicode_checkbox(cell_xml, target_char, new_char):
    return cell_xml.replace(f'<w:t>{target_char}</w:t>', f'<w:t>{new_char}</w:t>')

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
    original_doc_xml = xml_str
    original_trs = list(trs)
    other_files = {n: v for n, v in contents.items() if n != 'word/document.xml'}
    del contents, xml_str, tbl_xml, row_parts, trs
    gc.collect()
    return other_files, original_doc_xml, original_trs

def fill_one_row(original_trs, data):
    company = str(data.get('company', '')).strip()
    task_no = str(data.get('taskNo', '')).strip()
    leader = str(data.get('leader', '')).strip()
    audit_type = str(data.get('auditType', '')).strip()
    scope = str(data.get('scope', '')).strip()
    conclusion = str(data.get('conclusion', '')).strip()
    trs = list(original_trs)
    cells0 = re.findall(r'(<w:tc[^>]*>.*?</w:tc>)', trs[0])
    if len(cells0) >= 2:
        c0, c1 = cells0[0], cells0[1]
        m = re.search(r'<w:t[^>]*>公司名称[^<]*</w:t>', c0)
        if m:
            after = c0[m.end():]
            tm = re.search(r'<w:t[^>]*>([^<]*)</w:t>', after)
            if tm:
                ph = tm.group(1)
                if ph and ph.strip() and ph not in ('', ':'):
                    c0 = c0[:tm.start()] + f'<w:t>{company}</w:t>' + c0[tm.end():]
                else:
                    c0 = c0[:m.end()] + f'<w:t>{company}</w:t>' + c0[m.end():]
            else:
                c0 = c0[:m.end()] + f'<w:t>{company}</w:t>' + c0[m.end():]
        m = re.search(r'<w:t[^>]*>任务号[^<]*</w:t>', c1)
        if m:
            after = c1[m.end():]
            tm = re.search(r'<w:t[^>]*>([^<]*)</w:t>', after)
            if tm:
                ph = tm.group(1)
                if ph and ph.strip() and ph not in ('', '：'):
                    c1 = c1[:tm.start()] + f'<w:t>{task_no}</w:t>' + c1[tm.end():]
                else:
                    c1 = c1[:m.end()] + f'<w:t>{task_no}</w:t>' + c1[m.end():]
            else:
                c1 = c1[:m.end()] + f'<w:t>{task_no}</w:t>' + c1[m.end():]
        trs[0] = re.sub(r'(<w:tc[^>]*>.*?</w:tc>){2}', lambda m: c0+c1, trs[0], count=1, flags=re.DOTALL)
    cells1 = re.findall(r'(<w:tc[^>]*>.*?</w:tc>)', trs[1])
    if len(cells1) >= 2:
        c1 = cells1[1]
        m = re.search(r'<w:t[^>]*>组长[^<]*</w:t>', c1)
        if m:
            after = c1[m.end():]
            tm = re.search(r'<w:t[^>]*>([^<]*)</w:t>', after)
            if tm:
                ph = tm.group(1)
                if ph and ph.strip():
                    c1 = c1[:tm.start()] + f'<w:t>{leader}</w:t>' + c1[tm.end():]
                else:
                    c1 = c1[:m.end()] + f'<w:t>{leader}</w:t>' + c1[m.end():]
            else:
                c1 = c1[:m.end()] + f'<w:t>{leader}</w:t>' + c1[m.end():]
        trs[1] = re.sub(r'(<w:tc[^>]*>.*?</w:tc>){2}', lambda m: cells1[0]+c1, trs[1], count=1, flags=re.DOTALL)
    cells2 = re.findall(r'(<w:tc[^>]*>.*?</w:tc>)', trs[2])
    if len(cells2) >= 2:
        is_iatf = 'IATF' in scope
        is_iso = 'ISO' in scope
        if is_iatf: cells2[0] = replace_unicode_checkbox(cells2[0], '□', '☑')
        else: cells2[0] = replace_unicode_checkbox(cells2[0], '☑', '□')
        if is_iso: cells2[1] = replace_unicode_checkbox(cells2[1], '□', '☑')
        else: cells2[1] = replace_unicode_checkbox(cells2[1], '☑', '□')
        trs[2] = re.sub(r'(<w:tc[^>]*>.*?</w:tc>){2}', lambda m: cells2[0]+cells2[1], trs[2], count=1, flags=re.DOTALL)
    cells3 = re.findall(r'(<w:tc[^>]*>.*?</w:tc>)', trs[3])
    if len(cells3) >= 2:
        is_initial = '二阶段' in audit_type or '一阶段' in audit_type
        is_surv = '监' in audit_type
        is_recert = '再认证' in audit_type
        is_transfer = '转移' in audit_type
        is_special = '特殊' in audit_type
        if is_initial:
            cells3[0] = replace_unicode_checkbox(cells3[0], '□', '☑')
            cells3[0] = replace_unicode_checkbox(cells3[0], '☑', '□')
        elif is_surv:
            cells3[0] = replace_unicode_checkbox(cells3[0], '□', '☑')
            cells3[0] = replace_unicode_checkbox(cells3[0], '☑', '□')
        else:
            cells3[0] = replace_unicode_checkbox(cells3[0], '☑', '□')
        if is_recert or is_transfer:
            cells3[1] = replace_unicode_checkbox(cells3[1], '□', '☑')
        elif is_special:
            cells3[1] = replace_unicode_checkbox(cells3[1], '□', '☑')
        trs[3] = re.sub(r'(<w:tc[^>]*>.*?</w:tc>){2}', lambda m: cells3[0]+cells3[1], trs[3], count=1, flags=re.DOTALL)
    cells21 = re.findall(r'(<w:tc[^>]*>.*?</w:tc>)', trs[21])
    if len(cells21) >= 2:
        is_initial = '二阶段' in audit_type or '一阶段' in audit_type or '再认证' in audit_type
        is_surv = '监' in audit_type
        is_transfer = '转移' in audit_type
        c21 = cells21[1]
        if is_initial: c21 = set_form_checkbox(c21, 0)
        if is_transfer: c21 = set_form_checkbox(c21, 2)
        if is_surv:
            if '不换证' in conclusion: c21 = set_form_checkbox(c21, 4)
            elif '换发' in conclusion: c21 = set_form_checkbox(c21, 5)
        cells21[1] = c21
        trs[21] = re.sub(r'(<w:tc[^>]*>.*?</w:tc>){2}', lambda m: cells21[0]+c21, trs[21], count=1, flags=re.DOTALL)
    return trs

def build_docx(other_files, original_doc_xml, trs):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
        tbl_start = original_doc_xml.find('<w:tbl>')
        tbl_match = re.search(r'<w:tbl[^>]*>(?:(?!</w:tbl>).)*</w:tbl>', original_doc_xml[tbl_start:], re.DOTALL)
        tbl_end = tbl_start + tbl_match.end()
        new_tbl = trs[0] + ''.join(trs[1:])
        new_doc_xml = original_doc_xml[:tbl_start] + new_tbl + original_doc_xml[tbl_end:]
        zout.writestr('word/document.xml', new_doc_xml.encode('utf-8'))
        for name, content in other_files.items():
            zout.writestr(name, content)
    result = buf.getvalue()
    buf.close()
    return result

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
    if excel_file and tpl_file2:
        try:
            wb = openpyxl.load_workbook(io.BytesIO(excel_file.getvalue()), data_only=True)
            ws = wb.active
            headers = [c.value for c in ws[1]]
            st.write('Headers: ' + str(headers))
            rows_data = []
            for row in ws.iter_rows(min_row=2):
                vals = [c.value for c in row]
                if vals[3]:
                    rows_data.append(vals)
            total = len(rows_data)
            st.success('Read ' + str(total) + ' rows')
            if st.button('Generate All (ZIP)', type='primary'):
                if total == 0:
                    st.warning('No valid rows found')
                else:
                    # Show status updates
                    status_placeholder = st.empty()
                    status_placeholder.text('Parsing template...')
                    try:
                        other_files, original_doc_xml, original_trs = parse_template(tpl_file2.getvalue())
                        gc.collect()
                        status_placeholder.text('Generating reports... 0/' + str(total))
                        results = []
                        errors = []
                        for ri, rv in enumerate(rows_data):
                            try:
                                company = rv[3]
                                audit_team = rv[4]
                                audit_type = rv[5]
                                audit_address = rv[12]
                                cert_scope = rv[13]
                                task_no = rv[14]
                                conclusion = rv[16]
                                date_val = rv[17]
                                ds = format_date(date_val)
                                leader = str(audit_team).split('+')[0].strip() if audit_team else ''
                                d = {
                                    'company': str(company) if company else '',
                                    'taskNo': str(task_no) if task_no else '',
                                    'leader': leader,
                                    'auditType': str(audit_type) if audit_type else '',
                                    'address': str(audit_address) if audit_address else '',
                                    'scope': str(cert_scope) if cert_scope else '',
                                    'date': ds,
                                    'conclusion': str(conclusion) if conclusion else '',
                                }
                                row_trs = fill_one_row(original_trs, d)
                                rb = build_docx(other_files, original_doc_xml, row_trs)
                                ts = datetime.now().strftime('%Y%m%d_%H%M%S')
                                fname = sanitize(str(company)) + '_' + ts + '.docx'
                                results.append((fname, rb))
                            except Exception as e:
                                errors.append(str(rv[3]) + ': ' + str(e))
                            status_placeholder.text('Generating reports... ' + str(ri + 1) + '/' + str(total))
                        gc.collect()
                        if results:
                            buf = io.BytesIO()
                            with zipfile.ZipFile(buf, 'w', compression=zipfile.ZIP_DEFLATED) as zfout:
                                for fn, fd in results:
                                    zfout.writestr(fn, fd)
                            zipped = buf.getvalue()
                            buf.close()
                            st.download_button('Download All (ZIP)', data=zipped,
                                file_name='reports_' + datetime.now().strftime('%Y%m%d_%H%M') + '.zip',
                                mime='application/zip')
                            st.success('Generated ' + str(len(results)) + ' reports')
                        if errors:
                            st.warning(str(len(errors)) + ' failed')
                            for e in errors[:5]:
                                st.text('  - ' + e)
                    except Exception as e:
                        st.error('Generation error: ' + str(e))
        except ImportError:
            st.error('Need openpyxl')
        except Exception as e:
            st.error('Error: ' + str(e))
    elif excel_file or tpl_file2:
        st.info('Please upload both files')
