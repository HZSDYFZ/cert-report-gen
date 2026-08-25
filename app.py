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

st.title("📄 认证报告自动化生成系统")
st.caption("支持单份 FORM6101 报告匹配生成与 Excel 数据低 CPU 占用分批导出")
st.markdown("---")

CHK_EMPTY = chr(0x25A1)
CHK_FILLED = chr(0x25A0)

def clean_text(val):
    """清洗文本：去除 None、换行符及两端空格"""
    if val is None:
        return ''
    s = str(val).replace('\r', '').replace('\n', ' ').strip()
    return re.sub(r'\s+', ' ', s)

def format_date(val):
    if val is None: return ''
    if isinstance(val, datetime): return val.strftime('%Y-%m-%d')
    s = clean_text(val)
    if not s: return ''
    m = re.search(r'(\d{4})[年\-/](\d{1,2})[月\-/](\d{1,2})', s)
    if m: 
        return f"{m.group(1)}-{m.group(2).zfill(2)}-{m.group(3).zfill(2)}"
    m = re.search(r'(\d{1,2})/(\d{1,2})/(\d{4})', s)
    if m: 
        return f"{m.group(3)}-{m.group(1).zfill(2)}-{m.group(2).zfill(2)}"
    clean_d = re.search(r'\d{4}-\d{2}-\d{2}', s)
    if clean_d:
        return clean_d.group(0)
    return s[:10] if len(s) >= 10 else s

def set_fcb(doc_xml, pos, checked):
    cs = max(0, pos - 300)
    chunk = doc_xml[cs:pos + 100]
    if checked:
        if 'w:checked w:val=\"0\"' in chunk:
            return doc_xml[:cs] + chunk.replace('w:checked w:val=\"0\"/>', 'w:checked/>') + doc_xml[pos + 100:]
        if 'w:checked' not in chunk:
            return doc_xml[:cs] + chunk.replace('FORMCHECKBOX', 'w:checked/>FORMCHECKBOX') + doc_xml[pos + 100:]
    else:
        if 'w:checked/>' in chunk:
            return doc_xml[:cs] + chunk.replace('w:checked/>', 'w:checked w:val=\"0\"/>') + doc_xml[pos + 100:]
    return doc_xml

def fill_cert_standard(cell_text, atype):
    at = clean_text(atype)
    result = cell_text
    if 'IATF16949' in cell_text and ('IATF' in at or '16949' in at):
        result = result.replace(CHK_EMPTY + ' IATF16949', CHK_FILLED + ' IATF16949').replace(CHK_EMPTY + 'IATF16949', CHK_FILLED + 'IATF16949')
    if 'ISO9001' in cell_text and ('9001' in at or 'QMS' in at):
        result = result.replace(CHK_EMPTY + ' ISO9001', CHK_FILLED + ' ISO9001').replace(CHK_EMPTY + 'ISO9001', CHK_FILLED + 'ISO9001')
    if 'ISO14001' in cell_text and ('EMS' in at or '14001' in at):
        result = result.replace(CHK_EMPTY + ' ISO14001', CHK_FILLED + ' ISO14001').replace(CHK_EMPTY + 'ISO14001', CHK_FILLED + 'ISO14001')
    if 'ISO45001' in cell_text and ('OHS' in at or '45001' in at):
        result = result.replace(CHK_EMPTY + 'ISO 45001', CHK_FILLED + 'ISO 45001').replace(CHK_EMPTY + 'ISO45001', CHK_FILLED + 'ISO45001')
    return result

def fill_audit_type(cell_text, atype):
    at = clean_text(atype)
    result = cell_text
    if '一阶段' in at or '二阶段' in at or '初审' in at:
        result = result.replace(CHK_EMPTY + '初审', CHK_FILLED + '初审')
    if '监' in at:
        result = result.replace(CHK_EMPTY + '监审', CHK_FILLED + '监审')
    if '再认证' in at or '转移' in at:
        result = result.replace(CHK_EMPTY + '再认证/转移', CHK_FILLED + '再认证/转移')
    if '特殊' in at:
        result = result.replace(CHK_EMPTY + '特殊审核', CHK_FILLED + '特殊审核')
    return result

# ==================== SINGLE REPORT 模式 ====================

def extract_form_fields(doc):
    fields = {}
    for table in doc.tables:
        for ri, row in enumerate(table.rows):
            cells = [clean_text(c.text) for c in row.cells]
            if ri == 1 and len(cells) > 1: fields['taskNo'] = cells[1]
            if ri == 2 and len(cells) > 1: fields['company'] = cells[1]
            if ri == 3 and len(cells) > 1: fields['leader'] = cells[1]
            if ri == 4 and len(cells) > 1: fields['auditType'] = cells[1]
            if ri == 5 and len(cells) > 1: fields['address'] = cells[1]
            if ri == 6 and len(cells) > 1: fields['scope'] = cells[1]
    return fields

def fill_report(doc, fields):
    company = clean_text(fields.get('company', ''))
    taskNo = clean_text(fields.get('taskNo', ''))
    leader = clean_text(fields.get('leader', ''))
    auditType = clean_text(fields.get('auditType', ''))
    address = clean_text(fields.get('address', ''))
    scope = clean_text(fields.get('scope', ''))
    
    filled = {'company': False, 'taskNo': False, 'leader': False, 'address': False, 'scope': False}

    for para in doc.paragraphs:
        runs = list(para.runs)
        full_text = ''.join(r.text or '' for r in runs)
        for k, name in [('company', '公司名称'), ('taskNo', '任务号'), ('leader', '审核组长'), ('address', '审核地址'), ('scope', '认证范围')]:
            val = fields.get(k, '')
            if val and not filled[k] and name in full_text:
                for i, run in enumerate(runs):
                    if run.text and name in run.text:
                        if i + 1 < len(runs): runs[i + 1].text = (runs[i + 1].text or '') + val
                        else: run.text = run.text + val
                        filled[k] = True; break

    for table in doc.tables:
        for ri, row in enumerate(table.rows):
            cells = row.cells
            if ri == 0 and company and not filled['company']:
                for ci in range(min(3, len(cells))):
                    for para in cells[ci].paragraphs:
                        for i, run in enumerate(para.runs):
                            if run.text and '公司名称' in run.text:
                                run.text = run.text + company; filled['company'] = True; break
                        if filled['company']: break
                    if filled['company']: break
            if ri == 2 and auditType:
                for ci in range(2, min(5, len(cells))):
                    for para in cells[ci].paragraphs:
                        new_t = fill_cert_standard(para.text, auditType)
                        if new_t != para.text:
                            for r in para.runs: r.text = ''
                            para.runs[0].text = new_t if para.runs else para.add_run(new_t)
            if ri == 3 and auditType:
                for ci in range(2, min(5, len(cells))):
                    for para in cells[ci].paragraphs:
                        new_t = fill_audit_type(para.text, auditType)
                        if new_t != para.text:
                            for r in para.runs: r.text = ''
                            para.runs[0].text = new_t if para.runs else para.add_run(new_t)

    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    if auditType:
        con_idx = 0
        if '转移' in auditType: con_idx = 2
        elif '监' in auditType: con_idx = 5 if '换发' in auditType else 4
        
        with zipfile.ZipFile(buf, 'r') as z: content = {name: z.read(name) for name in z.namelist()}
        doc_xml = content['word/document.xml'].decode('utf-8')
        fcb_positions = [m.start() for m in re.finditer(r'FORMCHECKBOX', doc_xml)]
        for idx in range(7):
            if idx + 66 < len(fcb_positions):
                doc_xml = set_fcb(doc_xml, fcb_positions[idx + 66], idx == con_idx)
        content['word/document.xml'] = doc_xml.encode('utf-8')
        out = io.BytesIO()
        with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
            for name, data in content.items(): zout.writestr(name, data)
        return out.getvalue()
    return buf.getvalue()

# ==================== BATCH GENERATION 模式 ====================

def batch_read_row(row_tuple):
    """安全地解析单元格数据，补齐不足的列"""
    vals = list(row_tuple) + [''] * 20  # 自动补齐列表长度，防止 IndexError
    
    company = clean_text(vals[2])
    leader_raw = clean_text(vals[3])
    leader = leader_raw.split('+')[0].strip() if leader_raw else ''
    
    return {
        'company': company,
        'leader': leader,
        'auditType': clean_text(vals[4]),
        'address': clean_text(vals[6]),
        'scope': clean_text(vals[7]),
        'taskNo': clean_text(vals[8]),
        'decision': clean_text(vals[10]),
        'date': format_date(vals[11])
    }

def batch_get_conclusion_idx(atype, decision):
    """精准判断认证决定勾选索引 (FCB 66-72)"""
    at, dec = clean_text(atype), clean_text(decision)
    is_surv = ('监' in at or '监督' in at)
    
    if '转移' in at:
        return 2  # 通过，可换发证书
    elif is_surv:
        if '不换证' in dec or '保持' in dec:
            return 4  # 通过，不换证
        elif '换发' in dec or '换证' in dec:
            return 5  # 通过，可换发新的认证证书
        return 4  # 监督审核默认不换证
    elif '特殊' in at:
        return 3 if ('换发' in dec or '换证' in dec) else 0
    return 0  # 默认 (初审/二阶段/再认证): 通过，可发证

def replace_field_in_run(para, keyword, new_value):
    """把段落中的占位符替换为新值，清理多余占位字符"""
    if not new_value or keyword not in para.text:
        return
    for run in para.runs:
        if keyword in run.text:
            # 使用正则精准替换冒号后面的内容
            if re.search(f'{keyword}[：:]', run.text):
                run.text = re.sub(f'{keyword}[：:]\s*.*', f'{keyword}：{new_value}', run.text)
            else:
                run.text = f"{keyword}：{new_value}"

def batch_fill_report_fast(tpl_doc_bytes, fields):
    doc = Document(io.BytesIO(tpl_doc_bytes))
    company = fields.get('company', '')
    taskNo = fields.get('taskNo', '')
    leader = fields.get('leader', '')
    auditType = fields.get('auditType', '')
    address = fields.get('address', '')
    scope = fields.get('scope', '')
    date_val = fields.get('date', '')
    decision = fields.get('decision', '')

    tn_upper = str(taskNo).upper()
    has_ts, has_er = 'TS' in tn_upper, 'ER' in tn_upper

    # 替换正文段落
    for para in doc.paragraphs:
        replace_field_in_run(para, '公司名称', company)
        replace_field_in_run(para, '任务号', taskNo)
        replace_field_in_run(para, '任务编号', taskNo)
        replace_field_in_run(para, '审核组长', leader)
        replace_field_in_run(para, '审核地址', address)
        replace_field_in_run(para, '认证范围', scope)
        replace_field_in_run(para, '日期', date_val)

    # 替换表格中的单元格
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    full_p = para.text
                    if not full_p.strip(): continue
                    
                    replace_field_in_run(para, '公司名称', company)
                    replace_field_in_run(para, '任务号', taskNo)
                    replace_field_in_run(para, '任务编号', taskNo)
                    replace_field_in_run(para, '审核组长', leader)
                    replace_field_in_run(para, '审核地址', address)
                    replace_field_in_run(para, '认证范围', scope)
                    replace_field_in_run(para, '日期', date_val)

                    # 体系标准替换
                    if 'IATF16949' in full_p or 'ISO9001' in full_p:
                        new_t = para.text
                        if has_ts: new_t = new_t.replace(CHK_EMPTY + ' IATF16949', CHK_FILLED + ' IATF16949').replace(CHK_EMPTY + 'IATF16949', CHK_FILLED + 'IATF16949')
                        if has_er: new_t = new_t.replace(CHK_EMPTY + ' ISO9001', CHK_FILLED + ' ISO9001').replace(CHK_EMPTY + 'ISO9001', CHK_FILLED + 'ISO9001')
                        if new_t != para.text:
                            for r in para.runs: r.text = ''
                            para.runs[0].text = new_t

                    # 审核类型替换
                    if auditType and ('初审' in full_p or '监审' in full_p or '再认证' in full_p or '特殊审核' in full_p):
                        new_t = para.text
                        if '二阶段' in auditType or '初审' in auditType or '一阶段' in auditType: 
                            new_t = new_t.replace(CHK_EMPTY + '初审', CHK_FILLED + '初审')
                        if '监' in auditType: 
                            new_t = new_t.replace(CHK_EMPTY + '监审', CHK_FILLED + '监审')
                        if '再认证' in auditType or '转移' in auditType: 
                            new_t = new_t.replace(CHK_EMPTY + '再认证/转移', CHK_FILLED + '再认证/转移')
                        if '特殊' in auditType: 
                            new_t = new_t.replace(CHK_EMPTY + '特殊审核', CHK_FILLED + '特殊审核')
                        if new_t != para.text:
                            for r in para.runs: r.text = ''
                            para.runs[0].text = new_t

    out_buf = io.BytesIO()
    doc.save(out_buf)
    out_buf.seek(0)

    # 修改 XML 勾选认证决定 FORMCHECKBOX
    con_idx = batch_get_conclusion_idx(auditType, decision)
    with zipfile.ZipFile(out_buf, 'r') as z:
        content = {name: z.read(name) for name in z.namelist()}
    
    doc_xml = content['word/document.xml'].decode('utf-8')
    fcb_positions = [m.start() for m in re.finditer(r'FORMCHECKBOX', doc_xml)]
    
    for idx in range(7):
        abs_idx = idx + 66
        if abs_idx < len(fcb_positions):
            doc_xml = set_fcb(doc_xml, fcb_positions[abs_idx], idx == con_idx)
            
    content['word/document.xml'] = doc_xml.encode('utf-8')
    
    res_buf = io.BytesIO()
    with zipfile.ZipFile(res_buf, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in content.items(): zout.writestr(name, data)
    return res_buf.getvalue()

# ==================== STREAMLIT UI 主页面 ====================

mode = st.radio('选择操作模式', ['Single Report (单份生成)', 'Batch Generation (防降频极速版)'], key='app_mode')

if 'Single Report' in mode:
    st.header('单份报告生成模式')
    c1, c2 = st.columns(2)
    with c1: ff = st.file_uploader('1. 上传 FORM6101 文件 (.docx)', type=['docx'], key='s_form')
    with c2: tf = st.file_uploader('2. 上传报告模板 (.docx)', type=['docx'], key='s_tpl')

    if ff and tf:
        if st.button('🚀 立即生成单份报告', type='primary'):
            try:
                fields = extract_form_fields(Document(io.BytesIO(ff.getvalue())))
                doc_bytes = fill_report(Document(io.BytesIO(tf.getvalue())), fields)
                st.success(f"✅ 报告生成成功！公司：{fields.get('company', '未知')}")
                st.download_button('📥 下载生成的 Word 报告', doc_bytes, file_name=f"{fields.get('company','report')}.docx")
            except Exception as e:
                st.error(f"生成失败: {str(e)}")

else:
    st.header('批量报告导出模式（防降频极速版）')
    c1, c2 = st.columns(2)
    with c1: ef = st.file_uploader('1. 上传 Excel 数据源 (.xlsx)', type=['xlsx'], key='b_excel')
    with c2: tf = st.file_uploader('2. 上传 Word 模板 (.docx)', type=['docx'], key='b_tpl')

    if ef and tf:
        try:
            # 彻底去掉 read_only=True，防止 openpyxl 遗漏数据行或对非完整行做截断
            wb = openpyxl.load_workbook(io.BytesIO(ef.getvalue()), data_only=True)
            ws = wb.active
            rows_data = []
            
            # 从第 2 行开始逐行读取
            for row in ws.iter_rows(min_row=2, values_only=True):
                # 只要该行不为空，且能提取到有效公司名称
                if row:
                    item = batch_read_row(row)
                    if item['company']:
                        rows_data.append(item)
            wb.close()

            total_count = len(rows_data)
            st.info(f"📊 成功解析 Excel，共找到 **{total_count}** 条公司记录。")

            if total_count > 0:
                st.markdown("### ⚙️ 极速分批生成控制")
                col_batch_size, col_batch_num = st.columns(2)
                
                with col_batch_size:
                    batch_size = st.number_input('每批次处理数量：', min_value=1, max_value=50, value=15, step=5)
                
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

                st.write(f"📋 **当前批次预览**（包含 {len(current_batch_data)} 家公司）：")
                for idx, item in enumerate(current_batch_data, 1):
                    st.text(f"  [{idx}] 公司: {item['company']} | 任务号: {item['taskNo']} | 类型: {item['auditType']}")

                if st.button(f'🚀 生成第 {selected_batch} 批报告并打包 ZIP', type='primary'):
                    with st.spinner(f'正在极速生成第 {selected_batch} 批（共 {len(current_batch_data)} 份）...'):
                        zb = io.BytesIO()
                        tpl_bytes = tf.getvalue()
                        
                        with zipfile.ZipFile(zb, 'w', zipfile.ZIP_DEFLATED) as zf:
                            for item in current_batch_data:
                                try:
                                    doc_bytes = batch_fill_report_fast(tpl_bytes, item)
                                    zf.writestr(f"{item['company']}.docx", doc_bytes)
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
