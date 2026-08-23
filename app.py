# -*- coding: utf-8 -*-
import io, re, zipfile
from datetime import datetime
import streamlit as st
import openpyxl
from docx import Document

CHK_EMPTY = chr(0x25A1)
CHK_FILLED = chr(0x25A0)

def set_cb(run, text):
    if run.text is None:
        run.text = text
    else:
        run.text = run.text.replace(CHK_EMPTY, text).replace(CHK_FILLED, text)

def fill_cb(doc, target, checked=True):
    fill = CHK_FILLED if checked else CHK_EMPTY
    for para in doc.paragraphs:
        for run in para.runs:
            if run.text and target in run.text:
                set_cb(run, fill)
                return True
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        if run.text and target in run.text:
                            set_cb(run, fill)
                            return True
    return False

def format_date(val):
    if val is None:
        return ''
    if isinstance(val, datetime):
        return val.strftime('%Y-%m-%d')
    s = str(val).strip()
    m = re.search(r'(\d{4})[年\-/](\d{1,2})[月\-/](\d{1,2})', s)
    if m:
        return m.group(1) + '-' + m.group(2).zfill(2) + '-' + m.group(3).zfill(2)
    return s[:10] if len(s) >= 10 else s

def get_conclusion(atype):
    atype = str(atype).strip() if atype else ''
    is_surv = '监' in atype
    if '一阶段' in atype or '二阶段' in atype or '再认证' in atype:
        return {'checked': [True, False, False, False, False, False],
                'fields': ['通过，可发证', '不通过', '通过，可换发证书', '不符合发证条件', '通过，不换证', '通过，可换发新的认证证书']}
    elif '转移' in atype:
        return {'checked': [False, False, True, False, False, False],
                'fields': ['通过，可发证', '不通过', '通过，可换发证书', '不符合发证条件', '通过，不换证', '通过，可换发新的认证证书']}
    elif is_surv:
        if '换发' in atype:
            return {'checked': [False, False, False, False, False, True],
                    'fields': ['通过，可发证', '不通过', '通过，可换发证书', '不符合发证条件', '通过，不换证', '通过，可换发新的认证证书']}
        else:
            return {'checked': [False, False, False, False, True, False],
                    'fields': ['通过，可发证', '不通过', '通过，可换发证书', '不符合发证条件', '通过，不换证', '通过，可换发新的认证证书']}
    else:
        return {'checked': [True, False, False, False, False, False],
                'fields': ['通过，可发证', '不通过', '通过，可换发证书', '不符合发证条件', '通过，不换证', '通过，可换发新的认证证书']}

def fill_report(doc, fields):
    company = fields.get('company', '')
    taskNo = fields.get('taskNo', '')
    leader = fields.get('leader', '')
    auditType = fields.get('auditType', '')
    address = fields.get('address', '')
    scope = fields.get('scope', '')
    
    # 填充文本
    filled = {'company': False, 'taskNo': False, 'leader': False, 'address': False, 'scope': False}
    
    for para in doc.paragraphs:
        runs = list(para.runs)
        full_text = ''.join(r.text or '' for r in runs)
        
        # 公司名称
        if company and not filled['company'] and '公司名称' in full_text:
            for i, run in enumerate(runs):
                if run.text and '公司名称' in run.text:
                    if i + 1 < len(runs):
                        runs[i + 1].text = (runs[i + 1].text or '') + company
                    else:
                        run.text = run.text + company
                    filled['company'] = True
                    break
        
        # 任务号
        if taskNo and not filled['taskNo'] and ('任务号' in full_text or '任务编号' in full_text):
            for i, run in enumerate(runs):
                if run.text and ('任务号' in run.text or '任务编号' in run.text):
                    if i + 1 < len(runs):
                        runs[i + 1].text = (runs[i + 1].text or '') + taskNo
                    else:
                        run.text = run.text + taskNo
                    filled['taskNo'] = True
                    break
        
        # 审核组长
        if leader and not filled['leader'] and '审核组长' in full_text:
            for i, run in enumerate(runs):
                if run.text and '审核组长' in run.text:
                    if i + 1 < len(runs):
                        runs[i + 1].text = (runs[i + 1].text or '') + leader
                    else:
                        run.text = run.text + leader
                    filled['leader'] = True
                    break
        
        # 审核地址
        if address and not filled['address'] and '审核地址' in full_text:
            for i, run in enumerate(runs):
                if run.text and '审核地址' in run.text:
                    if i + 1 < len(runs):
                        runs[i + 1].text = (runs[i + 1].text or '') + address
                    else:
                        run.text = run.text + address
                    filled['address'] = True
                    break
        
        # 认证范围
        if scope and not filled['scope'] and '认证范围' in full_text:
            for i, run in enumerate(runs):
                if run.text and '认证范围' in run.text:
                    if i + 1 < len(runs):
                        runs[i + 1].text = (runs[i + 1].text or '') + scope
                    else:
                        run.text = run.text + scope
                    filled['scope'] = True
                    break
    
    # 审核类型勾选
    if auditType:
        for para in doc.paragraphs:
            for run in para.runs:
                if run.text:
                    # 清除所有勾选
                    run.text = run.text.replace(CHK_FILLED, CHK_EMPTY)
        
        # 根据审核类型勾选
        if '一阶段' in auditType or '二阶段' in auditType or '初审' in auditType:
            for para in doc.paragraphs:
                for run in para.runs:
                    if run.text and '初审' in run.text:
                        run.text = run.text.replace(CHK_EMPTY, CHK_FILLED)
                        break
        if '监' in auditType:
            for para in doc.paragraphs:
                for run in para.runs:
                    if run.text and '监' in run.text:
                        run.text = run.text.replace(CHK_EMPTY, CHK_FILLED)
                        break
        if '再认证' in auditType:
            for para in doc.paragraphs:
                for run in para.runs:
                    if run.text and '再认证' in run.text:
                        run.text = run.text.replace(CHK_EMPTY, CHK_FILLED)
                        break
        if '转移' in auditType:
            for para in doc.paragraphs:
                for run in para.runs:
                    if run.text and '转移' in run.text:
                        run.text = run.text.replace(CHK_EMPTY, CHK_FILLED)
                        break
        if '特殊' in auditType:
            for para in doc.paragraphs:
                for run in para.runs:
                    if run.text and '特殊' in run.text:
                        run.text = run.text.replace(CHK_EMPTY, CHK_FILLED)
                        break
    
    con = get_conclusion(auditType)
    for name, chk in zip(con['fields'], con['checked']):
        fill_cb(doc, name, chk)
    return doc

def extract_form_fields(doc):
    fields = {}
    for table in doc.tables:
        for ri, row in enumerate(table.rows):
            cells = [c.text.strip() for c in row.cells]
            if ri == 1 and len(cells) > 1 and cells[1]:
                fields['taskNo'] = cells[1]
            if ri == 2 and len(cells) > 1 and cells[1]:
                fields['company'] = cells[1]
            if ri == 3 and len(cells) > 1 and cells[1]:
                fields['address'] = cells[1]
            if ri == 11 and len(cells) > 1 and cells[1]:
                s = cells[1]
                if s.startswith('IATF:'):
                    s = s[5:]
                fields['scope'] = s
            if ri == 14 and len(cells) > 1 and cells[1]:
                fields['auditType'] = cells[1]
            if ri == 17 and len(cells) > 1 and cells[1]:
                fields['leader'] = cells[1]
    for para in doc.paragraphs:
        t = para.text
        if '任务编号' in t or '任务号' in t:
            parts = re.split(r'[:：]', t)
            if len(parts) > 1:
                fields['taskNo'] = parts[-1].strip()
        if '组织名称' in t or '公司名称' in t:
            parts = re.split(r'[:：]', t)
            if len(parts) > 1:
                fields['company'] = parts[-1].strip()
        if '审核组长' in t:
            parts = re.split(r'[:：]', t)
            if len(parts) > 1:
                fields['leader'] = parts[-1].strip()
        if '审核地址' in t:
            parts = re.split(r'[:：]', t)
            if len(parts) > 1:
                fields['address'] = parts[-1].strip()
        if '认证范围' in t or 'IATF:' in t:
            parts = re.split(r'[:：]', t)
            if len(parts) > 1:
                s = parts[-1].strip()
                if s.startswith('IATF:'):
                    s = s[5:]
                fields['scope'] = s
        if '审核性质' in t or '审核类型' in t:
            parts = re.split(r'[:：]', t)
            if len(parts) > 1:
                fields['auditType'] = parts[-1].strip()
        m = re.search(r'(\d{4})[年\-/](\d{1,2})[月\-/](\d{1,2})', t)
        if m and ('审核日期' in t or '现场审核日期' in t):
            fields['date'] = m.group(1) + '-' + m.group(2).zfill(2) + '-' + m.group(3).zfill(2)
    fields.setdefault('company', '')
    fields.setdefault('taskNo', '')
    fields.setdefault('leader', '')
    fields.setdefault('auditType', '')
    fields.setdefault('address', '')
    fields.setdefault('scope', '')
    fields.setdefault('date', datetime.now().strftime('%Y-%m-%d'))
    return fields

def detect_format(wb):
    ws = wb.active
    if ws.max_column >= 15:
        return 'A'
    elif ws.max_column >= 12:
        return 'B'
    return 'unknown'

def read_row(ws, row, fmt):
    vals = [c.value for c in row]
    if fmt == 'A':
        return {
            'company': str(vals[3]).strip() if vals[3] else '',
            'leader': str(vals[4]).split('+')[0].strip() if vals[4] else '',
            'auditType': str(vals[5]).strip() if vals[5] else '',
            'address': str(vals[12]).strip() if vals[12] else '',
            'scope': str(vals[13]).strip() if vals[13] else '',
            'taskNo': str(vals[14]).strip() if vals[14] else '',
        }
    else:
        return {
            'company': str(vals[2]).strip() if vals[2] else '',
            'leader': str(vals[3]).split('+')[0].strip() if vals[3] else '',
            'auditType': str(vals[4]).strip() if vals[4] else '',
            'address': str(vals[6]).strip() if vals[6] else '',
            'scope': str(vals[7]).strip() if vals[7] else '',
            'taskNo': str(vals[8]).strip() if vals[8] else '',
            'date': format_date(vals[11]) if vals[11] else '',
        }

def count_rows(ws):
    c = 0
    for row in ws.iter_rows(min_row=2):
        vals = [cv.value for cv in row]
        if vals and vals[3]:
            c += 1
        elif vals and not any(v for v in vals):
            if c > 0:
                break
    return c

st.title('Cert Report Generator')

if 'mode' not in st.session_state:
    st.session_state.mode = 'single'
if 'batch_step' not in st.session_state:
    st.session_state.batch_step = 0
if 'batch_total' not in st.session_state:
    st.session_state.batch_total = 0
if 'batch_files' not in st.session_state:
    st.session_state.batch_files = []
if 'batch_processed' not in st.session_state:
    st.session_state.batch_processed = 0
if 'single_fields' not in st.session_state:
    st.session_state.single_fields = {}
if 'form_doc' not in st.session_state:
    st.session_state.form_doc = None
if 'tpl_bytes' not in st.session_state:
    st.session_state.tpl_bytes = None
if 'expl' not in st.session_state:
    st.session_state.expl = None
if 'ws' not in st.session_state:
    st.session_state.ws = None
if 'excel_fmt' not in st.session_state:
    st.session_state.excel_fmt = None

mode = st.radio('Select Mode', ['Single Report', 'Batch Generation'], key='mode')

if mode == 'Single Report':
    st.header('Single Report')
    c1, c2 = st.columns(2)
    with c1:
        st.subheader('Step 1: Upload FORM6101')
        ff = st.file_uploader('Upload FORM6101', type=['docx'], key='form_up')
    with c2:
        st.subheader('Step 2: Upload Template')
        tf = st.file_uploader('Upload Template', type=['docx'], key='tpl_up')
    if ff and tf:
        if st.session_state.form_doc is None or st.session_state.get('form_name') != ff.name:
            st.session_state.form_doc = ff.read()
            st.session_state.form_name = ff.name
            st.session_state.single_fields = extract_form_fields(Document(io.BytesIO(st.session_state.form_doc)))
        if st.session_state.tpl_bytes is None or st.session_state.get('tpl_name') != tf.name:
            st.session_state.tpl_bytes = tf.read()
            st.session_state.tpl_name = tf.name
        f = st.session_state.single_fields
        st.info('Extracted: company=' + str(f.get('company', '')) + ', taskNo=' + str(f.get('taskNo', '')) + ', leader=' + str(f.get('leader', '')))
        if st.button('Generate Report', type='primary', key='gen_s'):
            with st.spinner('Generating...'):
                try:
                    doc = Document(io.BytesIO(st.session_state.tpl_bytes))
                    doc = fill_report(doc, f)
                    out = io.BytesIO()
                    doc.save(out)
                    out.seek(0)
                    st.success('Report generated!')
                    fname = str(f.get('company', 'report')) + '.docx'
                    st.download_button(label='Download Report', data=out.getvalue(),
                                       file_name=fname,
                                       mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                except Exception as e:
                    st.error('Generation failed: ' + str(e))
    elif ff or tf:
        st.warning('Please upload both FORM6101 and template')
else:
    st.header('Batch Generation')
    c1, c2 = st.columns(2)
    with c1:
        ef = st.file_uploader('Upload Excel', type=['xlsx'], key='exc_up')
    with c2:
        tf = st.file_uploader('Upload Template', type=['docx'], key='b_tpl_up')
    if ef and tf:
        if st.session_state.expl is None or st.session_state.get('exc_name') != ef.name or st.session_state.get('b_tpl_name') != tf.name:
            try:
                wb = openpyxl.load_workbook(ef, data_only=True)
                ws = wb.active
                fmt = detect_format(wb)
                total = count_rows(ws)
                st.session_state.expl = wb
                st.session_state.ws = ws
                st.session_state.excel_fmt = fmt
                st.session_state.batch_total = total
                st.session_state.batch_processed = 0
                st.session_state.batch_step = 0
                st.session_state.batch_files = []
                st.session_state.expl_name = ef.name
                st.session_state.b_tpl_name = tf.name
                st.session_state.b_tpl_bytes = tf.read()
            except Exception as e:
                st.error('Failed to load Excel: ' + str(e))
                st.stop()
        fmt = st.session_state.excel_fmt
        total = st.session_state.batch_total
        processed = st.session_state.batch_processed
        fmt_label = 'Format A (15 cols)' if fmt == 'A' else 'Format B (12 cols)'
        st.info('Excel format: ' + fmt_label + ', total rows: ' + str(total))
        prog = min((processed + 20) / max(total, 1), 1.0) if total > 0 else 0
        st.progress(prog)
        st.write('Progress: ' + str(processed) + '/' + str(total))
        if st.button('Start Generation (20 at a time)', key='b_start'):
            with st.spinner('Generating...'):
                ws = st.session_state.ws
                fmt = st.session_state.excel_fmt
                step = st.session_state.batch_step
                start = 2 + step * 20
                end = min(start + 20, total + 1)
                new = []
                for ri in range(start, end):
                    row = list(ws.iter_rows(min_row=ri, max_row=ri))[0]
                    f = read_row(ws, row, fmt)
                    if not f.get('company'):
                        continue
                    try:
                        doc = Document(io.BytesIO(st.session_state.b_tpl_bytes))
                        doc = fill_report(doc, f)
                        out = io.BytesIO()
                        doc.save(out)
                        new.append((f['company'], out.getvalue()))
                    except Exception as e:
                        st.error(str(f['company']) + ' failed: ' + str(e))
                if new:
                    st.session_state.batch_files.extend(new)
                    st.session_state.batch_step += 1
                    st.session_state.batch_processed = min(step * 20 + len(new), total)
                    zb = io.BytesIO()
                    with zipfile.ZipFile(zb, 'w', zipfile.ZIP_DEFLATED) as zf:
                        for n, d in new:
                            zf.writestr(n + '.docx', d)
                    zb.seek(0)
                    st.session_state.curr_zip = zb.getvalue()
                    st.success('Generated ' + str(len(new)) + ' reports (total: ' + str(st.session_state.batch_processed) + '/' + str(total) + ')')
                    fname = 'reports_batch_' + str(step + 1) + '.zip'
                    st.download_button(label='Download Batch (ZIP)', data=st.session_state.curr_zip,
                                       file_name=fname,
                                       mime='application/zip')
                    if st.session_state.batch_processed >= total:
                        st.success('All reports generated!')
                        if len(st.session_state.batch_files) > 1:
                            azb = io.BytesIO()
                            with zipfile.ZipFile(azb, 'w', zipfile.ZIP_DEFLATED) as zf:
                                for n, d in st.session_state.batch_files:
                                    zf.writestr(n + '.docx', d)
                            azb.seek(0)
                            st.download_button(label='Download All (ZIP)', data=azb.getvalue(),
                                               file_name='all_reports.zip',
                                               mime='application/zip')
        if st.button('Clear and Restart', key='b_clear'):
            st.session_state.batch_step = 0
            st.session_state.batch_processed = 0
            st.session_state.batch_files = []
            st.session_state.expl = None
            st.rerun()
    elif ef or tf:
        st.warning('Please upload both Excel and template')

st.markdown('---')
st.caption('Cert Report Generator v2.0')
