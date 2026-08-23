import sys, io, re, zipfile
from datetime import datetime
import streamlit as st
import openpyxl
from docx import Document

CHK_EMPTY = chr(0x25A1)
CHK_FILLED = chr(0x25A0)

def format_date(val):
    if val is None: return ''
    if isinstance(val, datetime): return val.strftime('%Y-%m-%d')
    s = str(val).strip()
    m = re.search(r'(\d{4})[年\-/](\d{1,2})[月\-/](\d{1,2})', s)
    if m: return m.group(1)+'-'+m.group(2).zfill(2)+'-'+m.group(3).zfill(2)
    m = re.search(r'(\d{1,2})/(\d{1,2})/(\d{4})', s)
    if m: return m.group(3)+'-'+m.group(1).zfill(2)+'-'+m.group(2).zfill(2)
    return s[:10] if len(s) >= 10 else s

def get_conclusion_indices(atype):
    atype = str(atype).strip() if atype else ''
    is_surv = '监' in atype and '再认证' not in atype and '二阶段' not in atype and '一阶段' not in atype
    if '一阶段' in atype or '二阶段' in atype or '再认证' in atype: return [0]
    elif '转移' in atype: return [2]
    elif is_surv:
        if '换发' in atype: return [5]
        else: return [4]
    else: return [0]

def is_audit_surv(atype):
    return '监' in atype and '再认证' not in atype and '二阶段' not in atype and '一阶段' not in atype

def fill_cert_standard_cell(cell_text, atype):
    at = atype.strip()
    result = cell_text
    if 'IATF16949' in cell_text and 'IATF' in at:
        result = result.replace(CHK_EMPTY+' IATF16949', CHK_FILLED+' IATF16949')
    if 'ISO9001' in cell_text and '9001' in at:
        result = result.replace(CHK_EMPTY+' ISO9001', CHK_FILLED+' ISO9001')
    if 'ISO14001' in cell_text and ('EMS' in at or '14001' in at):
        result = result.replace(CHK_EMPTY+' ISO14001', CHK_FILLED+' ISO14001')
    if 'ISO45001' in cell_text:
        if 'OHS' in at or '45001' in at:
            result = result.replace(CHK_EMPTY+'ISO 45001', CHK_FILLED+'ISO 45001')
            if result == cell_text:
                result = result.replace(CHK_EMPTY+'ISO45001', CHK_FILLED+'ISO45001')
    return result

def fill_audit_type_cell(cell_text, atype):
    at = atype.strip()
    is_surv = is_audit_surv(at)
    result = cell_text
    if '一阶段' in at or '二阶段' in at or '再认证' in at:
        result = result.replace(CHK_EMPTY+'初审', CHK_FILLED+'初审')
    if is_surv:
        result = result.replace(CHK_EMPTY+'监审', CHK_FILLED+'监审')
    if '再认证' in at or '转移' in at:
        result = result.replace(CHK_EMPTY+'再认证/转移', CHK_FILLED+'再认证/转移')
    if '特殊' in at:
        result = result.replace(CHK_EMPTY+'特殊审核', CHK_FILLED+'特殊审核')
    return result

def fill_report(doc, fields):
    company = fields.get('company','')
    taskNo = fields.get('taskNo','')
    leader = fields.get('leader','')
    auditType = fields.get('auditType','')
    address = fields.get('address','')
    scope = fields.get('scope','')
    filled = {'company':False,'taskNo':False,'leader':False,'address':False,'scope':False}
    for para in doc.paragraphs:
        runs = list(para.runs)
        full_text = ''.join(r.text or '' for r in runs)
        if company and not filled['company'] and '公司名称' in full_text:
            for i, run in enumerate(runs):
                if run.text and '公司名称' in run.text:
                    if i+1 < len(runs): runs[i+1].text = (runs[i+1].text or '') + company
                    else: run.text = run.text + company
                    filled['company'] = True; break
        if taskNo and not filled['taskNo'] and ('任务号' in full_text or '任务编号' in full_text):
            for i, run in enumerate(runs):
                if run.text and ('任务号' in run.text or '任务编号' in run.text):
                    if i+1 < len(runs): runs[i+1].text = (runs[i+1].text or '') + taskNo
                    else: run.text = run.text + taskNo
                    filled['taskNo'] = True; break
        if leader and not filled['leader'] and '审核组长' in full_text:
            for i, run in enumerate(runs):
                if run.text and '审核组长' in run.text:
                    if i+1 < len(runs): runs[i+1].text = (runs[i+1].text or '') + leader
                    else: run.text = run.text + leader
                    filled['leader'] = True; break
        if address and not filled['address'] and '审核地址' in full_text:
            for i, run in enumerate(runs):
                if run.text and '审核地址' in run.text:
                    if i+1 < len(runs): runs[i+1].text = (runs[i+1].text or '') + address
                    else: run.text = run.text + address
                    filled['address'] = True; break
        if scope and not filled['scope'] and '认证范围' in full_text:
            for i, run in enumerate(runs):
                if run.text and '认证范围' in run.text:
                    if i+1 < len(runs): runs[i+1].text = (runs[i+1].text or '') + scope
                    else: run.text = run.text + scope
                    filled['scope'] = True; break
    for table in doc.tables:
        for ri, row in enumerate(table.rows):
            cells = row.cells
            if ri == 0:
                if company and not filled['company']:
                    for ci in range(min(3, len(cells))):
                        for para in cells[ci].paragraphs:
                            runs = list(para.runs)
                            for i, run in enumerate(runs):
                                if run.text and '公司名称' in run.text:
                                    if i+1 < len(runs): runs[i+1].text = (runs[i+1].text or '') + company
                                    else: run.text = run.text + company
                                    filled['company'] = True; break
                            if filled['company']: break
                        if filled['company']: break
                if taskNo and not filled['taskNo'] and len(cells) > 3:
                    for para in cells[3].paragraphs:
                        runs = list(para.runs)
                        for i, run in enumerate(runs):
                            if run.text and ('任务号' in run.text or '任务编号' in run.text):
                                if i+1 < len(runs): runs[i+1].text = (runs[i+1].text or '') + taskNo
                                else: run.text = run.text + taskNo
                                filled['taskNo'] = True; break
                        if filled['taskNo']: break
            if ri == 1 and leader and not filled['leader'] and len(cells) > 2:
                cell = cells[2]
                has_content = any(r.text and r.text.strip() for para in cell.paragraphs for r in para.runs)
                if not has_content:
                    para = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
                    para.add_run(leader); filled['leader'] = True
            if ri == 2 and auditType:
                at = auditType.strip()
                for ci in range(2, min(5, len(cells))):
                    for para in cells[ci].paragraphs:
                        new_text = fill_cert_standard_cell(para.text, at)
                        if new_text != para.text:
                            for run in para.runs: run.text = ''
                            if para.runs: para.runs[0].text = new_text
                            else: para.add_run(new_text)
            if ri == 3 and auditType:
                at = auditType.strip()
                for ci in range(2, min(5, len(cells))):
                    for para in cells[ci].paragraphs:
                        new_text = fill_audit_type_cell(para.text, at)
                        if new_text != para.text:
                            for run in para.runs: run.text = ''
                            if para.runs: para.runs[0].text = new_text
                            else: para.add_run(new_text)
            if ri == 4 and address and not filled['address']:
                for ci in range(2, min(5, len(cells))):
                    for para in cells[ci].paragraphs:
                        for run in para.runs:
                            if run.text and '审核地址' in run.text:
                                if run.text.strip() == '审核地址：': run.text = run.text + address
                                else:
                                    runs = list(para.runs)
                                    for i, r in enumerate(runs):
                                        if r.text and '审核地址' in r.text:
                                            if i+1 < len(runs): runs[i+1].text = (runs[i+1].text or '') + address
                                            else: r.text = r.text + address
                                            filled['address'] = True; break
                                    if filled['address']: break
                                if filled['address']: break
                        if filled['address']: break
            if ri == 5 and scope and not filled['scope']:
                for ci in range(2, min(5, len(cells))):
                    for para in cells[ci].paragraphs:
                        for run in para.runs:
                            if run.text and '认证范围' in run.text:
                                if run.text.strip() == '认证范围：': run.text = run.text + scope
                                else:
                                    runs = list(para.runs)
                                    for i, r in enumerate(runs):
                                        if r.text and '认证范围' in r.text:
                                            if i+1 < len(runs): runs[i+1].text = (runs[i+1].text or '') + scope
                                            else: r.text = r.text + scope
                                            filled['scope'] = True; break
                                    if filled['scope']: break
                                if filled['scope']: break
                        if filled['scope']: break
    if auditType:
        con_indices = get_conclusion_indices(auditType)
        buf = io.BytesIO(); doc.save(buf); buf.seek(0)
        with zipfile.ZipFile(buf,'r') as z:
            content = {name: z.read(name) for name in z.namelist()}
        doc_xml = content['word/document.xml'].decode('utf-8')
        fcb_positions = []
        for m in re.finditer(r'FORMCHECKBOX', doc_xml):
            pos = m.start()
            chunk = doc_xml[max(0,pos-300):pos+100]
            if re.search(r'w:checked w:val=\"0\"', chunk): val='0'
            elif re.search(r'w:checked w:val=\"1\"', chunk): val='1'
            elif re.search(r'<w:checked/>', chunk): val='1'
            else: val='?'
            fcb_positions.append((pos, val))
        for idx in con_indices:
            if idx+66 < len(fcb_positions):
                abs_pos, old_val = fcb_positions[idx+66]
                chunk_start = max(0, abs_pos-300)
                chunk = doc_xml[chunk_start:abs_pos+100]
                if old_val == '0':
                    new_chunk = re.sub(r'w:checked w:val=\"0\"', 'w:checked', chunk)
                    doc_xml = doc_xml[:chunk_start] + new_chunk + doc_xml[abs_pos+100:]
                elif old_val == '?':
                    new_chunk = chunk.replace('FORMCHECKBOX', 'w:checked/>FORMCHECKBOX')
                    doc_xml = doc_xml[:chunk_start] + new_chunk + doc_xml[abs_pos+100:]
        for idx in range(7):
            if idx not in con_indices and idx+66 < len(fcb_positions):
                abs_pos, old_val = fcb_positions[idx+66]
                chunk_start = max(0, abs_pos-300)
                chunk = doc_xml[chunk_start:abs_pos+100]
                if old_val == '1':
                    if re.search(r'<w:checked/>', chunk):
                        new_chunk = chunk.replace('<w:checked/>', 'w:checked w:val=\"0\"/>')
                        doc_xml = doc_xml[:chunk_start] + new_chunk + doc_xml[abs_pos+100:]
        content['word/document.xml'] = doc_xml.encode('utf-8')
        out = io.BytesIO()
        with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as zout:
            for name, data in content.items(): zout.writestr(name, data)
        out.seek(0); doc = Document(out)
    return doc

def extract_form_fields(doc):
    fields = {}
    for table in doc.tables:
        for ri, row in enumerate(table.rows):
            cells = [c.text.strip() for c in row.cells]
            if ri == 1 and len(cells) > 1 and cells[1]: fields['taskNo'] = cells[1]
            if ri == 2 and len(cells) > 1 and cells[1]: fields['company'] = cells[1]
            if ri == 3 and len(cells) > 1 and cells[1]: fields['leader'] = cells[1]
            if ri == 4 and len(cells) > 1 and cells[1]: fields['auditType'] = cells[1]
            if ri == 5 and len(cells) > 1 and cells[1]: fields['address'] = cells[1]
            if ri == 6 and len(cells) > 1 and cells[1]: fields['scope'] = cells[1]
    for k in ['company','taskNo','leader','auditType','address','scope']:
        fields.setdefault(k, '')
    return fields

def detect_format(wb):
    ws = wb.active
    if ws.max_column >= 15: return 'A'
    elif ws.max_column >= 12: return 'B'
    return 'unknown'

def read_row(ws, row, fmt):
    vals = [c.value for c in row]
    if fmt == 'A':
        return {'company':str(vals[2]).strip() if vals[2] else '',
                'leader':str(vals[3]).split('+')[0].strip() if vals[3] else '',
                'auditType':str(vals[4]).strip() if vals[4] else '',
                'address':str(vals[6]).strip() if vals[6] else '',
                'scope':str(vals[7]).strip() if vals[7] else '',
                'taskNo':str(vals[8]).strip() if vals[8] else ''}
    else:
        return {'company':str(vals[2]).strip() if vals[2] else '',
                'leader':str(vals[3]).split('+')[0].strip() if vals[3] else '',
                'auditType':str(vals[4]).strip() if vals[4] else '',
                'address':str(vals[6]).strip() if vals[6] else '',
                'scope':str(vals[7]).strip() if vals[7] else '',
                'taskNo':str(vals[8]).strip() if vals[8] else '',
                'date':format_date(vals[11]) if len(vals)>11 and vals[11] else ''}

def count_rows(ws):
    '''Count all rows with valid company name (col 2), skip blanks'''
    c = 0
    for ri in range(2, ws.max_row + 1):
        row = list(ws.iter_rows(min_row=ri, max_row=ri))[0]
        vals = [cv.value for cv in row]
        if vals and len(vals) > 2 and vals[2]:
            c += 1
    return c

st.set_page_config(page_title='Cert Report Generator', page_icon='\\U0001F4C4')
if 'mode' not in st.session_state: st.session_state.mode = 'single'
if 'batch_step' not in st.session_state: st.session_state.batch_step = 0
if 'batch_total' not in st.session_state: st.session_state.batch_total = 0
if 'batch_files' not in st.session_state: st.session_state.batch_files = []
if 'batch_processed' not in st.session_state: st.session_state.batch_processed = 0
if 'single_fields' not in st.session_state: st.session_state.single_fields = {}
if 'form_doc' not in st.session_state: st.session_state.form_doc = None
if 'tpl_bytes' not in st.session_state: st.session_state.tpl_bytes = None
if 'expl' not in st.session_state: st.session_state.expl = None
if 'ws' not in st.session_state: st.session_state.ws = None
if 'excel_fmt' not in st.session_state: st.session_state.excel_fmt = None
if 'curr_zip' not in st.session_state: st.session_state.curr_zip = None
if 'all_rows' not in st.session_state: st.session_state.all_rows = None
if 'all_row_data' not in st.session_state: st.session_state.all_row_data = None

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
        st.info('Extracted: company=' + str(f.get('company','')) + ', taskNo=' + str(f.get('taskNo','')) + ', leader=' + str(f.get('leader','')))
        if st.button('Generate Report', type='primary', key='gen_s'):
            with st.spinner('Generating...'):
                try:
                    doc = Document(io.BytesIO(st.session_state.tpl_bytes))
                    doc = fill_report(doc, f)
                    out = io.BytesIO()
                    doc.save(out)
                    out.seek(0)
                    st.success('Report generated!')
                    fname = str(f.get('company','report')) + '.docx'
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
                # Pre-read ALL valid rows into memory
                all_row_data = []
                for ri in range(2, ws.max_row + 1):
                    row = list(ws.iter_rows(min_row=ri, max_row=ri))[0]
                    vals = [cv.value for cv in row]
                    if vals and len(vals) > 2 and vals[2]:
                        f = read_row(ws, row, fmt)
                        if f.get('company'):
                            all_row_data.append(f)
                st.session_state.expl = wb
                st.session_state.ws = ws
                st.session_state.excel_fmt = fmt
                st.session_state.batch_total = total
                st.session_state.all_row_data = all_row_data
                st.session_state.batch_processed = 0
                st.session_state.batch_step = 0
                st.session_state.batch_files = []
                st.session_state.expl_name = ef.name
                st.session_state.b_tpl_name = tf.name
                st.session_state.b_tpl_bytes = tf.read()
                st.info('Found ' + str(total) + ' valid rows in Excel')
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
                step = st.session_state.batch_step
                all_data = st.session_state.all_row_data
                start_idx = step * 20
                end_idx = min(start_idx + 20, len(all_data))
                batch = all_data[start_idx:end_idx]
                new = []
                for f in batch:
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
                    st.session_state.batch_processed = min(start_idx + len(new), total)
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
            st.session_state.curr_zip = None
            st.session_state.all_row_data = None
            st.rerun()
    elif ef or tf:
        st.warning('Please upload both Excel and template')

st.markdown('---')
st.caption('Cert Report Generator v2.7')
