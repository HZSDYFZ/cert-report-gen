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
    reps = {
        '{{company}}': fields.get('company', ''),
        '{{taskNo}}': fields.get('taskNo', ''),
        '{{leader}}': fields.get('leader', ''),
        '{{auditType}}': fields.get('auditType', ''),
        '{{address}}': fields.get('address', ''),
        '{{scope}}': fields.get('scope', ''),
        '{{date}}': fields.get('date', datetime.now().strftime('%Y-%m-%d')),
    }
    for para in doc.paragraphs:
        for run in para.runs:
            for k, v in reps.items():
                if k in run.text:
                    run.text = run.text.replace(k, v)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        for k, v in reps.items():
                            if k in run.text:
                                run.text = run.text.replace(k, v)
    con = get_conclusion(fields.get('auditType', ''))
    for name, chk in zip(con['fields'], con['checked']):
        fill_cb(doc, name, chk)
    return doc

def gen_docx(tpl_bytes, fields):
    buf = io.BytesIO(tpl_bytes)
    with zipfile.ZipFile(buf, 'r') as zf:
        xml = zf.read('word/document.xml').decode('utf-8')
        reps = {
            '{{company}}': fields.get('company', ''),
            '{{taskNo}}': fields.get('taskNo', ''),
            '{{leader}}': fields.get('leader', ''),
            '{{auditType}}': fields.get('auditType', ''),
            '{{address}}': fields.get('address', ''),
            '{{scope}}': fields.get('scope', ''),
            '{{date}}': fields.get('date', datetime.now().strftime('%Y-%m-%d')),
        }
        for k, v in reps.items():
            xml = xml.replace(k, v)
        out = io.BytesIO()
        with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as oz:
            for name in zf.namelist():
                if name == 'word/document.xml':
                    oz.writestr(name, xml.encode('utf-8'))
                else:
                    oz.writestr(name, zf.read(name))
        return out.getvalue()

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
        if any(cv.value for cv in row):
            c += 1
    return c

st.set_page_config(page_title='Cert Report Generator', layout='wide')
st.title('Cert Report Generator')

# 模式选择
mode = st.radio('Select Mode', ['Single Report', 'Batch Generation'], key='mode_choice')

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
        try:
            form_bytes = ff.getvalue()
            tpl_bytes = tf.getvalue()
            
            single_fields = extract_form_fields(Document(io.BytesIO(form_bytes)))
            
            st.info(f"Extracted: company={single_fields.get('company', '')}, taskNo={single_fields.get('taskNo', '')}, leader={single_fields.get('leader', '')}")
            
            if st.button('Generate Report', type='primary', key='gen_s'):
                with st.spinner('Generating...'):
                    doc = Document(io.BytesIO(tpl_bytes))
                    doc = fill_report(doc, single_fields)
                    out = io.BytesIO()
                    doc.save(out)
                    out.seek(0)
                    st.success('Report generated!')
                    st.download_button(
                        label='Download Report',
                        data=out.getvalue(),
                        file_name=f"{single_fields.get('company', 'report')}.docx",
                        mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                    )
        except Exception as e:
            st.error(f"Processing failed: {str(e)}")
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
        try:
            excel_bytes = ef.getvalue()
            b_tpl_bytes = tf.getvalue()
            
            wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True)
            ws = wb.active
            fmt = detect_format(wb)
            total = count_rows(ws)
            fmt_label = 'Format A (15 cols)' if fmt == 'A' else 'Format B (12 cols)'
            
            st.info(f"Excel format: {fmt_label}, total rows: {total}")
            
            if st.button('Start Generation (All Reports)', type='primary', key='b_start'):
                with st.spinner('Generating all reports...'):
                    progress_bar = st.progress(0)
                    azb = io.BytesIO()
                    
                    with zipfile.ZipFile(azb, 'w', zipfile.ZIP_DEFLATED) as zf:
                        processed_count = 0
                        for ri in range(2, total + 2):
                            row = list(ws.iter_rows(min_row=ri, max_row=ri))[0]
                            f = read_row(ws, row, fmt)
                            if not f.get('company'):
                                continue
                            
                            doc_bytes = gen_docx(b_tpl_bytes, f)
                            zf.writestr(f"{f['company']}.docx", doc_bytes)
                            
                            processed_count += 1
                            progress_bar.progress(min(processed_count / max(total, 1), 1.0))
                    
                    azb.seek(0)
                    st.success(f"Successfully generated {processed_count} reports!")
                    st.download_button(
                        label='Download All Reports (ZIP)',
                        data=azb.getvalue(),
                        file_name='all_reports.zip',
                        mime='application/zip'
                    )
        except Exception as e:
            st.error(f"Batch generation failed: {str(e)}")
    elif ef or tf:
        st.warning('Please upload both Excel and template')

st.markdown('---')
st.caption('Cert Report Generator v2.0')
