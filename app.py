# -*- coding: utf-8 -*-

import streamlit as st

import zipfile, re, io, os

from datetime import datetime



st.set_page_config(page_title='Cert Report Generator', page_icon='📋', layout='wide')

st.title('认证报告生成器')



def sanitize(name):

    return re.sub(r'[\\/:*?<>|]', '_', str(name))



def format_date(val):

    if val is None:

        return ''

    if isinstance(val, datetime):

        return val.strftime('%Y-%m-%d')

    if isinstance(val, (int, float)):

        # Excel serial date (days since 1899-12-30)

        try:

            return (datetime(1899, 12, 30) + __import__('datetime').timedelta(days=int(val))).strftime('%Y-%m-%d')

        except Exception:

            return ''

    s = str(val).strip()

    if not s or s == '#N/A':

        return ''

    m = re.search(r'(\d{4})[年\-\/](\d{1,2})[月\-\/](\d{1,2})', s)

    if m:

        return f'{m.group(1)}-{m.group(2).zfill(2)}-{m.group(3).zfill(2)}'

    m = re.search(r'(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})', s)

    if m:

        y = int(m.group(3))

        if y < 100:

            y += 2000

        return f'{y}-{m.group(1).zfill(2)}-{m.group(2).zfill(2)}'

    return s



def set_checkbox_in_xml(xml_str, index):

    pattern = r'<w:fldChar[^>]*w:fldCharType="begin"[^>]*>.*?</w:fldChar>'

    matches = list(re.finditer(pattern, xml_str, re.DOTALL))

    if index < len(matches):

        cb = matches[index].group(0)

        new_cb = re.sub(r'w:checked w:val="0"', 'w:checked w:val="1"', cb)

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

                if nxt:

                    data['company'] = nxt

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

            if nxt:

                data['address'] = nxt

        elif not data.get('scope') and '认证范围' in t and i + 1 < len(texts):

            nxt = texts[i + 1].strip()

            if nxt:

                data['scope'] = nxt

        elif not data.get('auditType') and '审核性质' in t and i + 1 < len(texts):

            nxt = texts[i + 1].strip()

            if nxt:

                data['auditType'] = nxt

        elif not data.get('date') and '现场审核日期' in t and i + 1 < len(texts):

            nxt = texts[i + 1].strip()

            if nxt:

                m = re.search(r'(\d{4})[\/-]\d{1,2}[\/-]\d{1,2}', nxt)

                if m:

                    data['date'] = m.group(0)

        i += 1

    return data



def fill_template(template_bytes, data):

    company = str(data.get('company', '')).strip()

    task_no = str(data.get('taskNo', '')).strip()

    leader = str(data.get('leader', '')).strip()

    audit_type = str(data.get('auditType', '')).strip()

    address = str(data.get('address', '')).strip()

    scope = str(data.get('scope', '')).strip()

    date = str(data.get('date', '')).strip()

    conclusion = str(data.get('conclusion', '')).strip()



    with zipfile.ZipFile(io.BytesIO(template_bytes)) as zfin:

        contents = {n: zfin.read(n) for n in zfin.namelist()}

    xml_str = contents['word/document.xml'].decode('utf-8')



    tbl_start = xml_str.find('<w:tbl>')

    tbl_match = re.search(r'<w:tbl[^>]*>(?:(?!</w:tbl>).)*</w:tbl>', xml_str[tbl_start:], re.DOTALL)

    if tbl_match is None:

        raise Exception('No table found in template')

    tbl_end = tbl_start + tbl_match.end()

    tbl_xml = xml_str[tbl_start:tbl_end]



    row_parts = re.split(r'(<w:tr[^>]*>.*?</w:tr>)', tbl_xml, flags=re.DOTALL)

    tr_indices = [i for i, p in enumerate(row_parts) if p.startswith('<w:tr')]

    trs = [row_parts[i] for i in tr_indices]



    if len(trs) < 22:

        raise Exception(f'Expected at least 22 rows, got {len(trs)}')



    # Row 0: 公司名称 + 任务号

    cells0 = re.findall(r'(<w:tc[^>]*>.*?</w:tc>)', trs[0])

    if len(cells0) >= 2:

        c0 = cells0[0]

        m = re.search(r'<w:t[^>]*>公司名称</w:t>', c0)

        if m:

            after_label = c0[m.end():]

            cm = re.match(r'<w:t[^>]*>[:：\s]*</w:t>', after_label)

            if cm:

                insert_pos = m.end() + cm.end()

                c0 = c0[:insert_pos] + f'<w:t>{company}</w:t>' + c0[insert_pos:]

            else:

                c0 = c0[:m.end()] + f'<w:t>{company}</w:t>' + c0[m.end():]

        c1 = cells0[1]

        m = re.search(r'<w:t[^>]*>任务号[：:]</w:t>', c1)

        if m:

            after_label = c1[m.end():]

            cm = re.match(r'<w:t[^>]*>[\s]*</w:t>', after_label)

            if cm:

                insert_pos = m.end() + cm.end()

                c1 = c1[:insert_pos] + f'<w:t>{task_no}</w:t>' + c1[insert_pos:]

            else:

                c1 = c1[:m.end()] + f'<w:t>{task_no}</w:t>' + c1[m.end():]

        trs[0] = trs[0].replace(cells0[0], c0, 1).replace(cells0[1], c1, 1)



    # Row 1: 审核组长

    cells1 = re.findall(r'(<w:tc[^>]*>.*?</w:tc>)', trs[1])

    if len(cells1) >= 2 and leader:

        c1 = cells1[1]

        m = re.search(r'</w:t>\s*$', c1)

        if m:

            c1 = c1[:m.end()] + f'<w:t>{leader}</w:t>' + c1[m.end():]

        else:

            c1 = c1.replace('</w:p>', f'<w:t>{leader}</w:t></w:p>', 1)

        trs[1] = trs[1].replace(cells1[1], c1, 1)



    # Row 2: 认证标准 (IATF/ISO) - replace checkbox characters

    new_row2 = trs[2]

    new_row2 = new_row2.replace('☑', '□')

    has_ts = 'TS' in task_no

    has_er = 'ER' in task_no

    cb2_matches = list(re.finditer(r'<w:t[^>]*>□</w:t>', new_row2))

    if len(cb2_matches) >= 5:

        if has_ts and not has_er:

            pos = cb2_matches[0].start()

            new_row2 = new_row2[:pos] + '<w:t>☑</w:t>' + new_row2[cb2_matches[0].end():]

        if has_er and not has_ts:

            pos = cb2_matches[1].start()

            new_row2 = new_row2[:pos] + '<w:t>☑</w:t>' + new_row2[cb2_matches[1].end():]

    trs[2] = new_row2



    # Row 3: 审核类型 - detect template style first, then apply

    new_row3 = trs[3]

    new_row3 = new_row3.replace('☑', '□')

    is_initial = '二阶段' in audit_type or '一阶段' in audit_type or '再认证' in audit_type

    is_surv = '监' in audit_type

    is_transfer = '转移' in audit_type

    is_special = '特殊' in audit_type



    # Detect template style

    is_znl_style = bool(re.search(r'<w:t[^>]*>□</w:t>.*?<w:t>监</w:t>', new_row3, re.DOTALL))

    is_02_style = '□初审' in new_row3 and '□' in new_row3.split('初审')[1][:20]



    # Apply checks - for 02 style, apply 监 BEFORE 初审 to avoid pattern consumption

    if is_02_style:

        if is_surv:

            new_row3 = re.sub(r'(□初审[^□]*)(□)', r'\1☑', new_row3, count=1)

        if is_initial:

            new_row3 = new_row3.replace('□初审', '☑初审', 1)

    elif is_znl_style:

        if is_surv:

            new_row3 = re.sub(r'(<w:t[^>]*>)□(</w:t>.*?<w:t>)监', r'\1☑\2监', new_row3, count=1)

        if is_initial:

            new_row3 = new_row3.replace('□初审', '☑初审', 1)

    else:

        # Default: apply all

        if is_initial:

            new_row3 = new_row3.replace('□初审', '☑初审', 1)

        if is_surv:

            new_row3 = new_row3.replace('□监', '☑监', 1)



    if is_transfer:

        new_row3 = new_row3.replace('□再认证/转移', '☑再认证/转移', 1)

    if is_special:

        new_row3 = new_row3.replace('□特殊审核', '☑特殊审核', 1)

    trs[3] = new_row3



    # Row 4: 审核地址

    cells4 = re.findall(r'(<w:tc[^>]*>.*?</w:tc>)', trs[4])

    if len(cells4) >= 2 and address:

        c = cells4[1]

        m = re.search(r'审核地址[：:]', c)

        if m:

            after = c[m.end():]

            cm = re.match(r'<w:t[^>]*>[\s]*</w:t>', after)

            if cm:

                insert_pos = m.end() + cm.end()

                c = c[:insert_pos] + f'<w:t>{address}</w:t>' + c[insert_pos:]

            else:

                c = c[:m.end()] + f'<w:t>{address}</w:t>' + c[m.end():]

        trs[4] = trs[4].replace(cells4[1], c, 1)



    # Row 5: 认证范围

    cells5 = re.findall(r'(<w:tc[^>]*>.*?</w:tc>)', trs[5])

    if len(cells5) >= 2 and scope:

        c = cells5[1]

        m = re.search(r'认证范围[：:]', c)

        if m:

            after = c[m.end():]

            cm = re.match(r'<w:t[^>]*>[\s]*</w:t>', after)

            if cm:

                insert_pos = m.end() + cm.end()

                c = c[:insert_pos] + f'<w:t>{scope}</w:t>' + c[insert_pos:]

            else:

                c = c[:m.end()] + f'<w:t>{scope}</w:t>' + c[m.end():]

        trs[5] = trs[5].replace(cells5[1], c, 1)



    # Row 21: 认证决定结论 (form field checkboxes)

    cells21 = re.findall(r'(<w:tc[^>]*>.*?</w:tc>)', trs[21])

    if len(cells21) >= 2:

        c21 = cells21[1]

        c21 = re.sub(r'w:checked w:val="1"', 'w:checked w:val="0"', c21)

        if '二阶段' in audit_type or '一阶段' in audit_type or '再认证' in audit_type:

            c21 = set_checkbox_in_xml(c21, 0)

        if '转移' in audit_type:

            c21 = set_checkbox_in_xml(c21, 2)

        if '监' in audit_type:

            if '不换证' in conclusion:

                c21 = set_checkbox_in_xml(c21, 4)

            elif '换发' in conclusion:

                c21 = set_checkbox_in_xml(c21, 5)

        if '特殊' in audit_type and '换发' in conclusion:

            c21 = set_checkbox_in_xml(c21, 3)

        trs[21] = trs[21].replace(cells21[1], c21, 1)



    # Date

    if date:

        for ri in range(len(trs)):

            if '日期' in trs[ri] and '日期：' in trs[ri]:

                m = re.search(r'日期[：:]', trs[ri])

                if m:

                    after = trs[ri][m.end():]

                    cm = re.match(r'<w:t[^>]*>[\s]*</w:t>', after)

                    if cm:

                        insert_pos = m.end() + cm.end()

                        trs[ri] = trs[ri][:insert_pos] + f'<w:t>{date}</w:t>' + trs[ri][insert_pos:]

                    else:

                        trs[ri] = trs[ri][:m.end()] + f'<w:t>{date}</w:t>' + trs[ri][m.end():]

                break



    for i, idx in enumerate(tr_indices):

        row_parts[idx] = trs[i]

    new_tbl_xml = ''.join(row_parts)

    out_xml = xml_str[:tbl_start] + new_tbl_xml + xml_str[tbl_end:]



    buf = io.BytesIO()

    with zipfile.ZipFile(buf, 'w', compression=zipfile.ZIP_DEFLATED) as zout:

        for name, content in contents.items():

            if name == 'word/document.xml':

                zout.writestr(name, out_xml.encode('utf-8'))

            else:

                zout.writestr(name, content)

    return buf.getvalue()



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

            import openpyxl

            wb = openpyxl.load_workbook(io.BytesIO(excel_file.getvalue()), data_only=True)

            ws = wb.active

            headers = [c.value for c in ws[1]]

            st.write('Headers: ' + str(headers))

            rows_data = []

            for row in ws.iter_rows(min_row=2):

                vals = [c.value for c in row]

                if vals[3]:

                    rows_data.append(vals)

            st.success('Read ' + str(len(rows_data)) + ' rows')

            if st.button('Generate All (ZIP)', type='primary'):

                results, errors = [], []

                for ri, rv in enumerate(rows_data):

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

                    try:

                        rb = fill_template(tpl_file2.getvalue(), d)

                        ts = datetime.now().strftime('%Y%m%d_%H%M%S')

                        fname = sanitize(str(company)) + '_' + ts + '.docx'

                        results.append((fname, rb))

                    except Exception as e:

                        errors.append(str(company) + ': ' + str(e))

                if results:

                    buf = io.BytesIO()

                    with zipfile.ZipFile(buf, 'w', compression=zipfile.ZIP_DEFLATED) as zfout:

                        for fn, fd in results:

                            zfout.writestr(fn, fd)

                    buf.seek(0)

                    st.download_button('Download All (ZIP)', data=buf,

                        file_name='reports_' + datetime.now().strftime('%Y%m%d_%H%M') + '.zip',

                        mime='application/zip')

                    st.success('Generated ' + str(len(results)) + ' reports')

                if errors:

                    st.warning(str(len(errors)) + ' failed')

                    for e in errors[:5]:

                        st.text('  - ' + e)

        except ImportError:

            st.error('Need openpyxl')

        except Exception as e:

            st.error('Error: ' + str(e))

    elif excel_file or tpl_file2:

        st.info('Please upload both files')



