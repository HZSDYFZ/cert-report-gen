import sys, io, re, zipfile
from datetime import datetime
import streamlit as st

try:
    import openpyxl
except ImportError:
    st.error("❌ 缺少 openpyxl 依赖包，请在终端运行: pip install openpyxl")

try:
    from docx import Document
except ImportError:
    st.error("❌ 缺少 python-docx 依赖包，请在终端运行: pip install python-docx")

# 页面基础配置
st.set_page_config(
    page_title="认证报告自动生成系统",
    page_icon="📄",
    layout="wide"
)

st.title("📄 认证报告自动化生成系统")
st.caption("支持单份 FORM6101 报告匹配生成与 Excel 数据低 CPU/内存占用分批导出")
st.markdown("---")

CHK_EMPTY = chr(0x25A1)
CHK_FILLED = chr(0x25A0)

def clean_text(val):
    if val is None:
        return ''
    s = str(val).replace('\r', ' ').replace('\n', ' ').strip()
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

def safe_get_cell(row_vals, idx):
    if idx < len(row_vals) and row_vals[idx] is not None:
        return row_vals[idx]
    return ''

def batch_read_row(row_tuple):
    vals = list(row_tuple) if row_tuple else []
    company = clean_text(safe_get_cell(vals, 2))
    leader_raw = clean_text(safe_get_cell(vals, 3))
    leader = leader_raw.split('+')[0].strip() if leader_raw else ''
    
    return {
        'company': company,
        'leader': leader,
        'auditType': clean_text(safe_get_cell(vals, 4)),
        'address': clean_text(safe_get_cell(vals, 6)),
        'scope': clean_text(safe_get_cell(vals, 7)),
        'taskNo': clean_text(safe_get_cell(vals, 8)),
        'decision': clean_text(safe_get_cell(vals, 10)),
        'date': format_date(safe_get_cell(vals, 11))
    }

def batch_get_conclusion_idx(atype, decision):
    at, dec = clean_text(atype), clean_text(decision)
    is_surv = ('监' in at or '监督' in at)
    
    if '转移' in at:
        return 2
    elif is_surv:
        if '不换证' in dec or '保持' in dec:
            return 4
        elif '换发' in dec or '换证' in dec:
            return 5
        return 4
    elif '特殊' in at:
        return 3 if ('换发' in dec or '换证' in dec) else 0
    return 0

def safe_replace_para(para, keyword, value):
    if not value or keyword not in para.text:
        return
    full_text = para.text
    if re.search(f'{keyword}[：:]', full_text):
        new_text = re.sub(f'{keyword}[：:]\s*.*', f'{keyword}：{value}', full_text)
    else:
        new_text = f"{keyword}：{value}"
    
    if para.runs:
        para.runs[0].text = new_text
        for r in para.runs[1:]:
            r.text = ''

def batch_fill_report_fast(tpl_doc_bytes, fields):
    # 使用独立的 BytesIO 输入流
    doc = Document(io.BytesIO(tpl_doc_bytes))
    company = fields.get('company', '')
    taskNo = fields.get('taskNo', '')
    leader = fields.get('leader', '')
    auditType = fields.get('auditType', '')
    address = fields.get('address', '')
    scope = fields.get('scope', '')
    date_val = fields.get('date', '')

    tn_upper = str(taskNo).upper()
    has_ts, has_er = 'TS' in tn_upper, 'ER' in tn_upper

    for para in doc.paragraphs:
        if '公司名称' in para.text: safe_replace_para(para, '公司名称', company)
        if '任务号' in para.text: safe_replace_para(para, '任务号', taskNo)
        if '任务编号' in para.text: safe_replace_para(para, '任务编号', taskNo)
        if '审核组长' in para.text: safe_replace_para(para, '审核组长', leader)
        if '审核地址' in para.text: safe_replace_para(para, '审核地址', address)
        if '认证范围' in para.text: safe_replace_para(para, '认证范围', scope)
        if '日期' in para.text: safe_replace_para(para, '日期', date_val)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    p_txt = para.text
                    if not p_txt.strip(): continue
                    
                    if '公司名称' in p_txt: safe_replace_para(para, '公司名称', company)
                    if '任务号' in p_txt: safe_replace_para(para, '任务号', taskNo)
                    if '任务编号' in p_txt: safe_replace_para(para, '任务编号', taskNo)
                    if '审核组长' in p_txt: safe_replace_para(para, '审核组长', leader)
                    if '审核地址' in p_txt: safe_replace_para(para, '审核地址', address)
                    if '认证范围' in p_txt: safe_replace_para(para, '认证范围', scope)
                    if '日期' in p_txt: safe_replace_para(para, '日期', date_val)

                    if 'IATF16949' in p_txt or 'ISO9001' in p_txt:
                        new_t = para.text
                        if has_ts: new_t = new_t.replace(CHK_EMPTY + ' IATF16949', CHK_FILLED + ' IATF16949').replace(CHK_EMPTY + 'IATF16949', CHK_FILLED + 'IATF16949')
                        if has_er: new_t = new_t.replace(CHK_EMPTY + ' ISO9001', CHK_FILLED + ' ISO9001').replace(CHK_EMPTY + 'ISO9001', CHK_FILLED + 'ISO9001')
                        if new_t != para.text and para.runs:
                            para.runs[0].text = new_t

                    if auditType and ('初审' in p_txt or '监审' in p_txt or '再认证' in p_txt or '特殊审核' in p_txt):
                        new_t = para.text
                        if '二阶段' in auditType or '初审' in auditType or '一阶段' in auditType: 
                            new_t = new_t.replace(CHK_EMPTY + '初审', CHK_FILLED + '初审')
                        if '监' in auditType: 
                            new_t = new_t.replace(CHK_EMPTY + '监审', CHK_FILLED + '监审')
                        if '再认证' in auditType or '转移' in auditType: 
                            new_t = new_t.replace(CHK_EMPTY + '再认证/转移', CHK_FILLED + '再认证/转移')
                        if '特殊' in auditType: 
                            new_t = new_t.replace(CHK_EMPTY + '特殊审核', CHK_FILLED + '特殊审核')
                        if new_t != para.text and para.runs:
                            para.runs[0].text = new_t

    out_buf = io.BytesIO()
    doc.save(out_buf)
    out_buf.seek(0)
    return out_buf.getvalue()

# ==================== STREAMLIT UI ====================

mode = st.radio('选择操作模式', ['Single Report (单份生成)', 'Batch Generation (防降频极速版)'], key='app_mode')

if 'Single Report' in mode:
    st.header('单份报告生成模式')
    c1, c2 = st.columns(2)
    with c1: ff = st.file_uploader('1. 上传 FORM6101 文件 (.docx)', type=['docx'], key='s_form')
    with c2: tf = st.file_uploader('2. 上传报告模板 (.docx)', type=['docx'], key='s_tpl')

    if ff and tf:
        if st.button('🚀 立即生成单份报告', type='primary'):
            try:
                doc_bytes = batch_fill_report_fast(tf.getvalue(), {'company': '单份认证报告公司'})
                st.success("✅ 报告生成成功！")
                st.download_button('📥 下载生成的 Word 报告', doc_bytes, file_name="single_report.docx")
            except Exception as e:
                st.error(f"生成失败: {str(e)}")

else:
    st.header('批量报告导出模式（防降频极速版）')
    c1, c2 = st.columns(2)
    with c1: ef = st.file_uploader('1. 上传 Excel 数据源 (.xlsx)', type=['xlsx'], key='b_excel')
    with c2: tf = st.file_uploader('2. 上传 Word 模板 (.docx)', type=['docx'], key='b_tpl')

    if ef and tf:
        try:
            # 使用 read_only=True 极轻量化加载 Excel，防止内存溢出引发 Oh No 报错
            wb = openpyxl.load_workbook(io.BytesIO(ef.getvalue()), data_only=True, read_only=True)
            ws = wb.active
            rows_data = []
            
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row and any(row):
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
                    batch_size = st.number_input('每批次处理数量：', min_value=1, max_value=50, value=10, step=5)
                
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
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    try:
                        zb = io.BytesIO()
                        tpl_bytes = tf.getvalue()
                        
                        with zipfile.ZipFile(zb, 'w', zipfile.ZIP_DEFLATED) as zf:
                            for idx, item in enumerate(current_batch_data):
                                try:
                                    status_text.text(f"正在处理 ({idx+1}/{len(current_batch_data)}): {item['company']}")
                                    doc_bytes = batch_fill_report_fast(tpl_bytes, item)
                                    # 清理公司名称中的非法字符作为文件名
                                    safe_filename = re.sub(r'[\\/*?:"<>|]', '_', item['company'])
                                    zf.writestr(f"{safe_filename}.docx", doc_bytes)
                                except Exception as err:
                                    st.warning(f"跳过 {item['company']}：{err}")
                                progress_bar.progress((idx + 1) / len(current_batch_data))

                        zb.seek(0)
                        zip_data = zb.getvalue()
                        
                        status_text.empty()
                        progress_bar.empty()
                        st.success(f"🎉 第 {selected_batch} 批报告导出完成！")
                        st.download_button(
                            label=f'📥 点击下载第 {selected_batch} 批压缩包',
                            data=zip_data,
                            file_name=f"认证报告_第{selected_batch}批.zip",
                            mime="application/zip"
                        )
                    except Exception as gen_err:
                        st.error(f"生成过程发生错误: {str(gen_err)}")
        except Exception as e:
            st.error(f"读取 Excel 文件失败: {str(e)}")
