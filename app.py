import sys, io, re, zipfile
from datetime import datetime
import streamlit as st
import openpyxl
from docx import Document

st.set_page_config(
    page_title="认证报告自动生成系统",
    page_icon="📄",
    layout="wide"
)

st.title("📄 认证报告自动化生成系统")
st.caption("支持单份 FORM6101 报告匹配生成与 Excel 数据批量报告导出")
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
    if '一阶段' in atype or '二阶段' in atype or '再认证' in atype: return 0   
    elif '转移' in atype: return 2   
    elif is_surv: return 5 if '换发' in atype else 4
    else: return 0   

def is_audit_surv(atype):
    return '监' in atype and '再认证' not in atype and '二阶段' not in atype and '一阶段' not in atype

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
    at = atype.strip()
    result = cell_text
    if 'IATF16949' in cell_text and 'IATF' in at: result = result.replace(CHK_EMPTY + ' IATF16949', CHK_FILLED + ' IATF16949')
    if 'ISO9001' in cell_text and '9001' in at: result = result.replace(CHK_EMPTY + ' ISO9001', CHK_FILLED + ' ISO9001')
    if 'ISO14001' in cell_text and ('EMS' in at or '14001' in at): result = result.replace(CHK_EMPTY + ' ISO14001', CHK_FILLED + ' ISO14001')
    if 'ISO45001' in cell_text and ('OHS' in at or '45001' in at):
        result = result.replace(CHK_EMPTY + 'ISO 45001', CHK_FILLED + 'ISO 45001').replace(CHK_EMPTY + 'ISO45001', CHK_FILLED + 'ISO45001')
    return result

def fill_audit_type(cell_text, atype):
    at = atype.strip()
    result = cell_text
    if '一阶段' in at or '二阶段' in at or '再认证' in at: result = result.replace(CHK_EMPTY + '初审', CHK_FILLED + '初审')
    if is_audit_surv(at): result = result.replace(CHK_EMPTY + '监审', CHK_FILLED + '监审')
    if '再认证' in at or '转移' in at: result = result.replace(CHK_EMPTY + '再认证/转移', CHK_FILLED + '再认证/转移')
    if '特殊' in at: result = result.replace(CHK_EMPTY + '特殊审核', CHK_FILLED + '特殊审核')
    return result

def fill_report(doc, fields):
    company, taskNo, leader = fields.get('company', ''), fields.get('taskNo', ''), fields.get('leader', '')
    auditType, address, scope = fields.get('auditType', ''), fields.get('address', ''), fields.get('scope', '')
    filled = {'company': False, 'taskNo': False, 'leader': False, 'address': False, 'scope': False}

    for para in doc.paragraphs:
        runs = list(para.runs)
        full_text = ''.join(r.text or '' for r in runs)
        for k, name in [('company', '公司名称'), ('taskNo', '任务号'), ('leader', '审核组长'), ('address', '审核地址'), ('scope', '认证范围')]:
            if fields.get(k) and not filled[k] and name in full_text:
                for i, run in enumerate(runs):
                    if run.text and name in run.text:
                        if i + 1 < len(runs): runs[i + 1].text = (runs[i + 1].text or '') + fields[k]
                        else: run.text = run.text + fields[k]
                        filled[k] = True; break

    for table in doc.tables:
        for ri, row in enumerate(table.rows):
            cells = row.cells
            if ri == 0:
                if company and not filled['company']:
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

    if auditType:
        con_idx = get_conclusion_idx(auditType)
        buf = io.BytesIO(); doc.save(buf); buf.seek(0)
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
        out.seek(0); doc = Document(out)
    return doc

def extract_form_fields(doc):
    fields = {}
    for table in doc.tables:
        for ri, row in enumerate(table.rows):
            cells = [c.text.strip() for c in row.cells]
            if ri == 1 and len(cells) > 1: fields['taskNo'] = cells[1]
            if ri == 2 and len(cells) > 1: fields['company'] = cells[1]
            if ri == 3 and len(cells) > 1: fields['leader'] = cells[1]
            if ri == 4 and len(cells) > 1: fields['auditType'] = cells[1]
            if ri == 5 and len(cells) > 1: fields['address'] = cells[1]
            if ri == 6 and len(cells) > 1: fields['scope'] = cells[1]
    return fields

def batch_read_row(row):
    vals = [c.value for c in row]
    leader_raw = str(vals[3]).strip() if len(vals) > 3 and vals[3] else ''
    return {
        'company': str(vals[2]).strip() if len(vals) > 2 and vals[2] else '',
        'leader': leader_raw.split('+')[0].strip(),
        'auditType': str(vals[4]).strip() if len(vals) > 4 and vals[4] else '',
        'address': str(vals[6]).strip() if len(vals) > 6 and vals[6] else '',
        'scope': str(vals[7]).strip() if len(vals) > 7 and vals[7] else '',
        'taskNo': str(vals[8]).strip() if len(vals) > 8 and vals[8] else '',
        'decision': str(vals[10]).strip() if len(vals) > 10 and vals[10] else '',
        'date': format_date(vals[11]) if len(vals) > 11 and vals[11] else ''
    }

# UI 页面部分
mode = st.radio('选择操作模式', ['Single Report (单份生成)', 'Batch Generation (批量导出)'])

if 'Single Report' in mode:
    st.header('单份报告生成模式')
    ff = st.file_uploader('1. 上传 FORM6101 文件 (.docx)', type=['docx'], key='s_form')
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
                st.error(f"生成失败，错误原因：{str(e)}")

else:
    st.header('批量报告导出模式')
    ef = st.file_uploader('1. 上传 Excel 数据源 (.xlsx)', type=['xlsx'], key='b_excel')
    tf = st.file_uploader('2. 上传 Word 模板 (.docx)', type=['docx'], key='b_tpl')

    if ef and tf:
        # 直接现场解析，不走复杂 Session 缓存，避免挂起
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
            
            st.info(f"📊 成功读取 Excel 数据，共解析出 {len(rows_data)} 条公司记录。")
            if len(rows_data) > 0:
                st.write("数据预览（前3条）：", rows_data[:3])

            if st.button('🚀 开始批量生成并打包下载', type='primary'):
                with st.spinner('正在批量填报报告中，请稍候...'):
                    zip_buffer = io.BytesIO()
                    tpl_bytes = tf.getvalue()
                    
                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                        success_count = 0
                        for f in rows_data:
                            try:
                                doc = fill_report(Document(io.BytesIO(tpl_bytes)), f)
                                out = io.BytesIO()
                                doc.save(out)
                                zf.writestr(f"{f['company']}.docx", out.getvalue())
                                success_count += 1
                            except Exception as err:
                                st.warning(f"公司 {f['company']} 跳过（出错: {err}）")
                    
                    zip_buffer.seek(0)
                    st.success(f"🎉 批量生成完毕！成功处理 {success_count} / {len(rows_data)} 份文档。")
                    st.download_button(
                        label="📥 一键下载所有报告压缩包 (.zip)",
                        data=zip_buffer.getvalue(),
                        file_name="批量认证报告结果.zip",
                        mime="application/zip"
                    )
        except Exception as e:
            st.error(f"读取或生成过程出错：{str(e)}")
