import sys, io, re, zipfile
from datetime import datetime
import streamlit as st
import openpyxl
from docx import Document

# 页面基础配置及标题设置
st.set_page_config(
    page_title="认证报告自动生成系统",
    page_icon="📄",
    layout="wide"
)

# 界面主标题与简要说明
st.title("📄 认证报告自动化生成系统")
st.caption("支持单份 FORM6101 报告匹配生成与 Excel 数据低 CPU 占用分批导出")
st.markdown("---")

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

def get_conclusion_idx(atype):
    atype = str(atype).strip() if atype else ''
    is_surv = '监' in atype and '再认证' not in atype and '二阶段' not in atype and '一阶段' not in atype
    if '一阶段' in atype or '二阶段' in atype or '再认证' in atype:
        return 0   # FCB[66]: 通过，可发证
    elif '转移' in atype:
        return 2   # FCB[68]: 通过，可换发证书
    elif is_surv:
        if '换发' in atype:
            return 5  # FCB[71]: 通过，可换发新的认证证书
        else:
            return 4  # FCB[70]: 通过，不换证
    else:
        return 0   # FCB[66]: 通过，可发证

def is_audit_surv(atype):
    return '监' in atype and '再认证' not in atype and '二阶段' not in atype and '一阶段' not in atype

def set_fcb(doc_xml, pos, checked):
    cs = max(0, pos - 300)
    chunk = doc_xml[cs:pos + 100]
    if checked:
        if 'w:checked w:val=\"0\"' in chunk:
            new_chunk = chunk.replace('w:checked w:val=\"0\"/>', 'w:checked/>')
            return doc_xml[:cs] + new_chunk + doc_xml[pos + 100:]
        if 'w:checked' not in chunk:
            new_chunk = chunk.replace('FORMCHECKBOX', 'w:checked/>FORMCHECKBOX')
            return doc_xml[:cs] + new_chunk + doc_xml[pos + 100:]
    else:
        if 'w:checked/>' in chunk:
            new_chunk = chunk.replace('w:checked/>', 'w:checked w:val=\"0\"/>')
            return doc_xml[:cs] + new_chunk + doc_xml[pos + 100:]
    return doc_xml

def fill_cert_standard(cell_text, atype):
    at = atype.strip()
    result = cell_text
    if 'IATF16949' in cell_text and 'IATF' in at:
        result = result.replace(CHK_EMPTY + ' IATF16949', CHK_FILLED + ' IATF16949')
    if 'ISO9001' in cell_text and '9001' in at:
        result = result.replace(CHK_EMPTY + ' ISO9001', CHK_FILLED + ' ISO9001')
    if 'ISO14001' in cell_text and ('EMS' in at or '14001' in at):
        result = result.replace(CHK_EMPTY + ' ISO14001', CHK_FILLED + ' ISO14001')
    if 'ISO45001' in cell_text:
        if 'OHS' in at or '45001' in at:
            result = result.replace(CHK_EMPTY + 'ISO 45001', CHK_FILLED + 'ISO 45001')
            if result == cell_text:
                result = result.replace(CHK_EMPTY + 'ISO45001', CHK_FILLED + 'ISO45001')
    return result

def fill_audit_type(cell_text, atype):
    at = atype.strip()
    is_surv = is_audit_surv(at)
    result = cell_text
    if '一阶段' in at or '二阶段' in at or '再认证' in at:
        result = result.replace(CHK_EMPTY + '初审', CHK_FILLED + '初审')
    if is_surv:
        result = result.replace(CHK_EMPTY + '监审', CHK_FILLED + '监审')
    if '再认证' in at or '转移' in at:
        result = result.replace(CHK_EMPTY + '再认证/转移', CHK_FILLED + '再认证/转移')
    if '特殊' in at:
        result = result.replace(CHK_EMPTY + '特殊审核', CHK_FILLED + '特殊审核')
    return result

# Single Report 专属填充逻辑
def fill_report(doc, fields):
    company = fields.get('company', '')
    taskNo = fields.get('taskNo', '')
    leader = fields.get('leader', '')
    auditType = fields.get('auditType', '')
    address = fields.get('address', '')
    scope = fields.get('scope', '')
    filled = {'company': False, 'taskNo': False, 'leader': False, 'address': False, 'scope': False}

    for para in doc.paragraphs:
        runs = list(para.runs)
        full_text = ''.join(r.text or '' for r in runs)
        if company and not filled['company'] and '公司名称' in full_text:
            for i, run in enumerate(runs):
                if run.text and '公司名称' in run.text:
                    if i + 1 < len(runs): runs[i + 1].text = (runs[i + 1].text or '') + company
                    else: run.text = run.text + company
                    filled['company'] = True; break
        if taskNo and not filled['taskNo'] and ('任务号' in full_text or '任务编号' in full_text):
            for i, run in enumerate(runs):
                if run.text and ('任务号' in run.text or '任务编号' in run.text):
                    if i + 1 < len(runs): runs[i + 1].text = (runs[i + 1].text or '') + taskNo
                    else: run.text = run.text + taskNo
                    filled['taskNo'] = True; break
        if leader and not filled['leader'] and '审核组长' in full_text:
            for i, run in enumerate(runs):
                if run.text and '审核组长' in run.text:
                    if i + 1 < len(runs): runs[i + 1].text = (runs[i + 1].text or '') + leader
                    else: run.text = run.text + leader
                    filled['leader'] = True; break
        if address and not filled['address'] and '审核地址' in full_text:
            for i, run in enumerate(runs):
                if run.text and '审核地址' in run.text:
                    if i + 1 < len(runs): runs[i + 1].text = (runs[i + 1].text or '') + address
                    else: run.text = run.text + address
                    filled['address'] = True; break
        if scope and not filled['scope'] and '认证范围' in full_text:
            for i, run in enumerate(runs):
                if run.text and '认证范围' in run.text:
                    if i + 1 < len(runs): runs[i + 1].text = (runs[i + 1].text or '') + scope
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
                                    if i + 1 < len(runs): runs[i + 1].text = (runs[i + 1].text or '') + company
                                    else: run.text = run.text + company
                                    filled['company'] = True; break
                            if filled['company']: break
                        if filled['company']: break
                if taskNo and not filled['taskNo'] and len(cells) > 3:
                    for para in cells[3].paragraphs:
                        runs = list(para.runs)
                        for i, run in enumerate(runs):
                            if run.text and ('任务号' in run.text or '任务编号' in run.text):
                                if i + 1 < len(runs): runs[i + 1].text = (runs[i + 1].text or '') + taskNo
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
                        new_text = fill_cert_standard(para.text, at)
                        if new_text != para.text:
                            for run in para.runs: run.text = ''
                            if para.runs: para.runs[0].text = new_text
                            else: para.add_run(new_text)
            if ri == 3 and auditType:
                at = auditType.strip()
                for ci in range(2, min(5, len(cells))):
                    for para in cells[ci].paragraphs:
                        new_text = fill_audit_type(para.text, at)
                        if new_text != para.text:
                            for run in para.runs: run.text = ''
                            if para.runs: para.runs[0].text = new_text
                            else: para.add_run(new_text)
            if ri == 4 and address and not filled['address']:
                for ci in range(2, min(5, len(cells))):
                    for para in cells[ci].paragraphs:
                        for run in para.runs:
                            if run.text and '审核地址' in run.text:
                                if run.text.strip() == '审核地址：':
                                    run.text = run.text + address; filled['address'] = True
                                else:
                                    runs = list(para.runs)
                                    for i, r in enumerate(runs):
                                        if r.text and '审核地址' in r.text:
                                            if i + 1 < len(runs): runs[i + 1].text = (runs[i + 1].text or '') + address
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
                                if run.text.strip() == '认证范围：':
                                    run.text = run.text + scope; filled['scope'] = True
                                else:
                                    runs = list(para.runs)
                                    for i, r in enumerate(runs):
                                        if r.text and '认证范围' in r.text:
                                            if i + 1 < len(runs): runs[i + 1].text = (runs[i + 1].text or '') + scope
                                            else: r.text = r.text + scope
                                            filled['scope'] = True; break
                                    if filled['scope']: break
                                if filled['scope']: break
                        if filled['scope']: break

    if auditType:
        con_idx = get_conclusion_idx(auditType)
        buf = io.BytesIO(); doc.save(buf); buf.seek(0)
        with zipfile.ZipFile(buf, 'r') as z:
            content = {name: z.read(name) for name in z.namelist()}
        doc_xml = content['word/document.xml'].decode('utf-8')
        fcb_positions = []
        for m in re.finditer(r'FORMCHECKBOX', doc_xml):
            pos = m.start()
            cs = max(0, pos - 300)
            chunk = doc_xml[cs:pos + 100]
            if 'w:checked w:val=\"0\"' in chunk: val = '0'
            elif 'w:checked/>' in chunk: val = '1'
            elif 'w:checked' in chunk: val = '1'
            else: val = '?'
            fcb_positions.append((pos, val))
        for idx in range(7):
            abs_idx = idx + 66
            if abs_idx < len(fcb_positions):
                abs_pos, old_val = fcb_positions[abs_idx]
                is_target = (idx == con_idx)
                doc_xml = set_fcb(doc_xml, abs_pos, is_target)
        content['word/document.xml'] = doc_xml.encode('utf-8')
        out = io.BytesIO()
        with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
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

# BATCH GENERATION 专属解析与填充逻辑
def batch_read_row(row):
    vals = [c.value for c in row]
    leader_raw = str(vals[3]).strip() if len(vals) > 3 and vals[3] else ''
    leader = leader_raw.split('+')[0].strip()
    
    return {
        'company': str(vals[2]).strip() if len(vals) > 2 and vals[2] else '',
        'leader': leader,
        'auditType': str(vals[4]).strip() if len(vals) > 4 and vals[4] else '',
        'address': str(vals[6]).strip() if len(vals) > 6 and vals[6] else '',
        'scope': str(vals[7]).strip() if len(vals) > 7 and vals[7] else '',
        'taskNo': str(vals[8]).strip() if len(vals) > 8 and vals[8] else '',
        'decision': str(vals[10]).strip() if len(vals) > 10 and vals[10] else '',
        'date': format_date(vals[11]) if len(vals) > 11 and vals[11] else ''
    }

def batch_get_conclusion_idx(atype, decision):
    at = str(atype).strip()
    dec = str(decision).strip()
    is_surv = ('监一' in at or '监二' in at)
    
    if '二阶段' in at or '再认证' in at:
        return 0
    elif '转移' in at:
        return 2
    elif is_surv and '不换证' in dec:
        return 4
    elif is_surv and '换发' in dec:
        return 5
    elif '特殊' in at and '换发' in dec:
        return 3
    return 0

def batch_fill_report(doc, fields):
    company = fields.get('company', '')
    taskNo = fields.get('taskNo', '')
    leader = fields.get('leader', '')
    auditType = fields.get('auditType', '')
    address = fields.get('address', '')
    scope = fields.get('scope', '')
    date_val = fields.get('date', '')
    decision = fields.get('decision', '')

    tn_upper = str(taskNo).upper()
    has_ts = 'TS' in tn_upper
    has_er = 'ER' in tn_upper

    if date_val:
        for para in doc.paragraphs:
            if '日期' in para.text and date_val not in para.text:
                for run in para.runs:
                    if '日期' in run.text:
                        run.text = run.text + ' ' + str(date_val)

    for table in doc.tables:
        for ri, row in enumerate(table.rows):
            cells = row.cells
            for cell in cells:
                for para in cell.paragraphs:
                    full_p = para.text
                    
                    if taskNo and ('任务号' in full_p or '任务编号' in full_p) and taskNo not in full_p:
                        for run in para.runs:
                            if '任务号' in run.text or '任务编号' in run.text:
                                run.text = run.text + ' ' + taskNo
                    if company and '公司名称' in full_p and company not in full_p:
                        for run in para.runs:
                            if '公司名称' in run.text:
                                run.text = run.text + ' ' + company
                    if leader and '审核组长' in full_p and leader not in full_p:
                        for run in para.runs:
                            if '审核组长' in run.text:
                                run.text = run.text + ' ' + leader
                    if address and '审核地址' in full_p and address not in full_p:
                        for run in para.runs:
                            if '审核地址' in run.text:
                                run.text = run.text + ' ' + address
                    if scope and '认证范围' in full_p and scope not in full_p:
                        for run in para.runs:
                            if '认证范围' in run.text:
                                run.text = run.text + ' ' + scope
                    if date_val and '日期' in full_p and date_val not in full_p:
                        for run in para.runs:
                            if '日期' in run.text:
                                run.text = run.text + ' ' + date_val

                    if 'IATF16949' in full_p or 'ISO9001' in full_p:
                        new_text = para.text
                        if has_ts:
                            new_text = new_text.replace(CHK_EMPTY + ' IATF16949', CHK_FILLED + ' IATF16949').replace(CHK_EMPTY + 'IATF16949', CHK_FILLED + 'IATF16949')
                        if has_er:
                            new_text = new_text.replace(CHK_EMPTY + ' ISO9001', CHK_FILLED + ' ISO9001').replace(CHK_EMPTY + 'ISO9001', CHK_FILLED + 'ISO9001')
                        if new_text != para.text:
                            for r in para.runs: r.text = ''
                            para.runs[0].text = new_text

                    if auditType and ('初审' in full_p or '监审' in full_p or '再认证' in full_p or '特殊审核' in full_p):
                        new_text = para.text
                        if '二阶段' in auditType:
                            new_text = new_text.replace(CHK_EMPTY + '初审', CHK_FILLED + '初审')
                        if '监一' in auditType or '监二' in auditType:
                            new_text = new_text.replace(CHK_EMPTY + '监审', CHK_FILLED + '监审')
                        if '再认证' in auditType or '转移' in auditType:
                            new_text = new_text.replace(CHK_EMPTY + '再认证/转移', CHK_FILLED + '再认证/转移')
                        if '特殊' in auditType:
                            new_text = new_text.replace(CHK_EMPTY + '特殊审核', CHK_FILLED + '特殊审核')
                        if new_text != para.text:
                            for r in para.runs: r.text = ''
                            para.runs[0].text = new_text

    if auditType:
        con_idx = batch_get_conclusion_idx(auditType, decision)
        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        with zipfile.ZipFile(buf, 'r') as z:
            content = {name: z.read(name) for name in z.namelist()}
        doc_xml = content['word/document.xml'].decode('utf-8')
        
        fcb_positions = []
        for m in re.finditer(r'FORMCHECKBOX', doc_xml):
            pos = m.start()
            cs = max(0, pos - 300)
            chunk = doc_xml[cs:pos + 100]
            val = '1' if ('w:checked/>' in chunk or 'w:checked ' in chunk) else '0'
            fcb_positions.append((pos, val))
        
        for idx in range(7):
            abs_idx = idx + 66
            if abs_idx < len(fcb_positions):
                abs_pos, old_val = fcb_positions[abs_idx]
                is_target = (idx == con_idx)
                doc_xml = set_fcb(doc_xml, abs_pos, is_target)
                
        content['word/document.xml'] = doc_xml.encode('utf-8')
        out = io.BytesIO()
        with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
            for name, data in content.items(): zout.writestr(name, data)
        out.seek(0)
        doc = Document(out)
        
    return doc

# ==================== STREAMLIT UI 页面交互 ====================

mode = st.radio('选择操作模式', ['Single Report (单份生成)', 'Batch Generation (防降频分批导出)'], key='app_mode')

if 'Single Report' in mode:
    st.header('单份报告生成模式')
    c1, c2 = st.columns(2)
    with c1:
        ff = st.file_uploader('1. 上传 FORM6101 文件 (.docx)', type=['docx'], key='s_form')
    with c2:
        tf = st.file_uploader('2. 上传报告模板 (.docx)', type=['docx'], key='s_tpl')

    if ff and tf:
        if st.button('🚀 立即生成单份报告', type='primary'):
            try:
                fields = extract_form_fields(Document(io.BytesIO(ff.getvalue())))
                doc = fill_report(Document(io.BytesIO(tf.getvalue())), fields)
                out = io.BytesIO()
                doc.save(out)
                st.success(f"✅ 报告生成成功！包含公司：{fields.get('company', '未知')}")
                st.download_button('📥 下载生成的 Word 报告', out.getvalue(), file_name=f"{fields.get('company','report')}.docx")
            except Exception as e:
                st.error(f"生成失败: {str(e)}")

else:
    st.header('批量报告导出模式（低 CPU 占用分批版）')
    c1, c2 = st.columns(2)
    with c1:
        ef = st.file_uploader('1. 上传 Excel 数据源 (.xlsx)', type=['xlsx'], key='b_excel')
    with c2:
        tf = st.file_uploader('2. 上传 Word 模板 (.docx)', type=['docx'], key='b_tpl')

    if ef and tf:
        try:
            wb = openpyxl.load_workbook(io.BytesIO(ef.getvalue()), data_only=True)
            ws = wb.active
            rows_data = []
            for ri in range(2, ws.max_row + 1):
                row = list(ws.iter_rows(min_row=ri, max_row=ri))[0]
                if row and len(row) > 2 and row[2].value:
                    item = batch_read_row(row)
                    if item['company']:
                        rows_data.append(item)

            total_count = len(rows_data)
            st.info(f"📊 成功解析 Excel，共找到 **{total_count}** 条公司记录。")

            if total_count > 0:
                st.markdown("### ⚙️ 防 Throttling 极速分批生成控制")
                col_batch_size, col_batch_num = st.columns(2)
                
                with col_batch_size:
                    batch_size = st.number_input('每批次处理数量（推荐 10-20 条）：', min_value=1, max_value=50, value=15, step=5)
                
                total_batches = (total_count + batch_size - 1) // batch_size
                
                with col_batch_num:
                    selected_batch = st.selectbox(
                        f'选择要生成的批次（共 {total_batches} 批）：',
                        options=list(range(1, total_batches + 1)),
                        format_func=lambda x: f"第 {x} 批 (涵盖第 {(x-1)*batch_size + 1} ~ {min(x*batch_size, total_count)} 条数据)"
                    )
                
                start_idx = (selected_batch - 1) * batch_size
                end_idx = min(selected_batch * batch_size, total_count)
                current_batch_data = rows_data[start_idx:end_idx]

                st.write(f"📋 **当前批次预览**（包含 {len(current_batch_data)} 家公司）：", [d['company'] for d in current_batch_data])

                if st.button(f'🚀 生成第 {selected_batch} 批报告并打包 ZIP', type='primary'):
                    with st.spinner(f'正在处理第 {selected_batch} 批（共 {len(current_batch_data)} 份），保持 CPU 健康运行...'):
                        zb = io.BytesIO()
                        tpl_bytes = tf.getvalue()
                        with zipfile.ZipFile(zb, 'w', zipfile.ZIP_DEFLATED) as zf:
                            for item in current_batch_data:
                                try:
                                    doc = batch_fill_report(Document(io.BytesIO(tpl_bytes)), item)
                                    out = io.BytesIO()
                                    doc.save(out)
                                    zf.writestr(f"{item['company']}.docx", out.getvalue())
                                except Exception as err:
                                    st.warning(f"跳过 {item['company']}：{err}")

                        zb.seek(0)
                        st.success(f"🎉 第 {selected_batch} 批报告导出完成！")
                        st.download_button(
                            label=f'📥 下载第 {selected_batch} 批压缩包 (batch_{selected_batch}.zip)',
                            data=zb.getvalue(),
                            file_name=f"认证报告_第{selected_batch}批.zip",
                            mime="application/zip"
                        )
        except Exception as e:
            st.error(f"处理 Excel/Word 发生错误: {str(e)}")
