import streamlit as st
import zipfile
import re
import os
import io
from datetime import datetime

st.set_page_config(page_title="认证报告生成器", page_icon="📋", layout="centered")

st.title("📋 认证报告生成器")
st.markdown("上传 FORM6101 文件和报告模板，自动生成认证报告")


INITIAL_TYPES = {"初审二阶段","LOC初审二阶段","初审二阶段(LOC升级)","初审二阶段(QMS）",
                 "初审二阶段（LOC）","初审二阶段（loc升级）","初审二阶段（严重）",
                 "初审二阶段（主场所）","初审二阶段（免一阶段）","初审二阶段（搬迁）",
                 "初审二阶段（本机构搬迁）","特殊审核（二阶段审核后扩范围）",
                 "特殊审核(变更)","特殊审核（变更）","特殊审核（扩范围）"}
SURVEILLANCE_TYPES = {"监一","监一（ISO9001）","监一（严重）","监一（主场所）",
                      "监二","监二（IATF)","监二（ISO）","监二（ISO9001:2015）",
                      "监二（Q）","监二（严重）","监一(严重)","监一（IATF）",
                      "监一（ISO）","监一（Q）","监二(IATF)","监二(ISO 9001)",
                      "监二(ISO9001)","监二（IATF16949)","监二（IATF）","监二（QMS)",
                      "监二（主场所）","监二（主场所，严重）"}
RECERT_TYPES = {"再认证","再认证（严重）","再认证（主场所）"}
TRANSFER_TYPES = {"转移","转移审核","转移（严重）"}

def sanitize_filename(name):
    return re.sub(r"[\\\\/:*?<>|]", "_", str(name))

def extract_form_data(file_bytes):
    """从 FORM6101 Word 文件中提取字段"""
    try:
        with zipfile.ZipFile(io.BytesIO(file_bytes)) as zf:
            xml_content = zf.read("word/document.xml").decode("utf-8")
    except Exception as e:
        return None, f"解析失败: {e}"
    
    # 提取所有文本
    texts = re.findall(r"<w:t[^>]*>([^<]*)</w:t>", xml_content)
    full_text = "".join(texts)
    lines = [l.strip() for l in full_text.split("\n") if l.strip()]
    
    # 按行提取字段（基于表格结构）
    # 找到表格内容
    table_match = re.search(r"<w:tbl>(.*?)</w:tbl>", xml_content, re.DOTALL)
    if not table_match:
        return None, "未找到表格"
    
    table_xml = table_match.group(1)
    rows = re.findall(r"<w:tr[^>]*>(.*?)</w:tr>", table_xml, re.DOTALL)
    
    data = {}
    for i, row in enumerate(rows):
        cells = re.findall(r"<w:tc[^>]*>(.*?)</w:tc>", row, re.DOTALL)
        cell_texts = []
        for cell in cells:
            t_matches = re.findall(r"<w:t[^>]*>([^<]*)</w:t>", cell)
            cell_texts.append("".join(t_matches).strip())
        
        if i == 1 and len(cell_texts) > 1:
            data["taskNo"] = cell_texts[1]
        elif i == 2 and len(cell_texts) > 1:
            data["company"] = cell_texts[1]
        elif i == 3 and len(cell_texts) > 1:
            data["address"] = cell_texts[1]
        elif i == 11 and len(cell_texts) > 1:
            scope = cell_texts[1]
            if scope.startswith("IATF:"):
                scope = scope[5:]
            data["scope"] = scope
        elif i == 13:
            if len(cell_texts) > 2:
                date_str = cell_texts[2]
                match = re.search(r"(\\d{4})[年\\-](\\d{1,2})[月\\-](\\d{1,2})", date_str)
                if match:
                    data["date"] = f"{match.group(1)}-{match.group(2).zfill(2)}-{match.group(3).zfill(2)}"
        elif i == 14 and len(cell_texts) > 1:
            data["auditType"] = cell_texts[1]
        elif i == 17 and len(cell_texts) > 1:
            data["leader"] = cell_texts[1]
    
    # 回退：用全文匹配
    if not data.get("company"):
        for line in lines:
            if "公司名称" in line:
                data["company"] = re.sub(r".*[:：]", "", line).strip()
                break
    if not data.get("taskNo"):
        for line in lines:
            if "任务编号" in line or "任务号" in line:
                data["taskNo"] = re.sub(r".*[:：]", "", line).strip()
                break
    if not data.get("leader"):
        for line in lines:
            if "审核组长" in line:
                data["leader"] = re.sub(r".*[:：]", "", line).strip()
                break
    if not data.get("address"):
        for line in lines:
            if "审核地址" in line:
                data["address"] = re.sub(r".*[:：]", "", line).strip()
                break
    if not data.get("scope"):
        for line in lines:
            if "认证范围" in line or "IATF:" in line:
                s = re.sub(r".*[:：]", "", line).strip()
                if s.startswith("IATF:"):
                    s = s[5:]
                data["scope"] = s
                break
    if not data.get("date"):
        for line in lines:
            if "审核日期" in line or re.search(r"\\d{4}年\\d{1,2}月\\d{1,2}日", line):
                match = re.search(r"(\\d{4})[年\\-](\\d{1,2})[月\\-](\\d{1,2})", line)
                if match:
                    data["date"] = f"{match.group(1)}-{match.group(2).zfill(2)}-{match.group(3).zfill(2)}"
                break
    if not data.get("auditType"):
        for line in lines:
            if "审核性质" in line or "审核类型" in line:
                data["auditType"] = re.sub(r".*[:：]", "", line).strip()
                break
    
    return data, None

def fill_template(template_bytes, form_data):
    """填充报告模板"""
    company = form_data.get("company", "")
    task_no = form_data.get("taskNo", "")
    leader = form_data.get("leader", "")
    audit_type = str(form_data.get("auditType", ""))
    address = form_data.get("address", "")
    scope = form_data.get("scope", "")
    date = form_data.get("date", "")
    
    with zipfile.ZipFile(io.BytesIO(template_bytes)) as zf:
        contents = {name: zf.read(name) for name in zf.namelist()}
    
    xml_str = contents["word/document.xml"].decode("utf-8")
    tbl_start = xml_str.find("<w:tbl>")
    tbl_end = xml_str.find("</w:tbl>", tbl_start) + 8
    tbl_xml = xml_str[tbl_start:tbl_end]
    
    row_parts = re.split(r"(<w:tr[^>]*>.*?</w:tr>)", tbl_xml, flags=re.DOTALL)
    trs = [p for p in row_parts if p.startswith("<w:tr")]
    
    if len(trs) < 24:
        raise Exception(f"期望24行表格，实际{len(trs)}行")
    
    has_ts = "TS" in str(task_no)
    has_er = "ER" in str(task_no)
    
    # Row 0: company + task number
    row0 = trs[0]
    cells0 = re.findall(r"(<w:tc[^>]*>.*?</w:tc>)", row0, re.DOTALL)
    if len(cells0) >= 2:
        c0, c1 = cells0[0], cells0[1]
        texts_c0 = re.findall(r"<w:t([^>]*?)>([^<]*)</w:t>", c0)
        for attrs, text in texts_c0:
            if text.strip() in (":", ""):
                new_text = text + company
                old = f"<w:t{attrs}>{text}</w:t>" if attrs else f"<w:t>{text}</w:t>"
                new = f"<w:t{attrs}>{new_text}</w:t>" if attrs else f"<w:t>{new_text}</w:t>"
                c0 = c0.replace(old, new, 1)
                break
        texts_c1 = re.findall(r"<w:t([^>]*?)>([^<]*)</w:t>", c1)
        for attrs, text in texts_c1:
            if text == "任务号：":
                new_text = text + str(task_no)
                old = f"<w:t{attrs}>{text}</w:t>" if attrs else f"<w:t>{text}</w:t>"
                new = f"<w:t{attrs}>{new_text}</w:t>" if attrs else f"<w:t>{new_text}</w:t>"
                c1 = c1.replace(old, new, 1)
                break
        trs[0] = row0.replace(cells0[0], c0, 1).replace(cells0[1], c1, 1)
    
    # Row 1: audit leader
    row1 = trs[1]
    cells1 = re.findall(r"(<w:tc[^>]*>.*?</w:tc>)", row1, re.DOTALL)
    if len(cells1) >= 2:
        c1 = cells1[1]
        texts1 = re.findall(r"<w:t[^>]*>([^<]+)</w:t>", c1)
        if not texts1 or not "".join(texts1).strip():
            c1 = c1.replace("</w:tc>",
                "<w:p><w:r><w:rPr><w:rFonts w:ascii=\"宋体\" w:hAnsi=\"宋体\" w:eastAsia=\"宋体\"/><w:sz w:val=\"20\"/></w:rPr><w:t>" + str(leader) + "</w:t></w:r></w:p></w:tc>", 1)
        trs[1] = row1.replace(cells1[1], c1, 1)
    
    # Row 3: audit type checkbox
    row3 = trs[3]
    cells3 = re.findall(r"(<w:tc[^>]*>.*?</w:tc>)", row3, re.DOTALL)
    if len(cells3) >= 2:
        cell_xml = cells3[1]
        def replace_cb(m):
            text = m.group(2)
            attrs = m.group(1)
            if audit_type in INITIAL_TYPES:
                text = text.replace("□初审", "☑初审", 1)
            elif audit_type in SURVEILLANCE_TYPES:
                text = text.replace("□初审      □", "☐初审      ☑", 1)
            elif audit_type in RECERT_TYPES | TRANSFER_TYPES:
                text = text.replace("□再认证/转移", "☑再认证/转移", 1)
            return f"<w:t{attrs}>{text}</w:t>" if attrs else f"<w:t>{text}</w:t>"
        cell_xml = re.sub(r"<w:t([^>]*?)>([^<]*)</w:t>", replace_cb, cell_xml)
        trs[3] = row3.replace(cells3[1], cell_xml, 1)
    
    # Row 4: audit address
    row4 = trs[4]
    cells4 = re.findall(r"(<w:tc[^>]*>.*?</w:tc>)", row4, re.DOTALL)
    if len(cells4) >= 2:
        cell_xml = cells4[1].replace("审核地址：", "审核地址：" + str(address), 1)
        trs[4] = row4.replace(cells4[1], cell_xml, 1)
    
    # Row 5: certification scope
    row5 = trs[5]
    cells5 = re.findall(r"(<w:tc[^>]*>.*?</w:tc>)", row5, re.DOTALL)
    if len(cells5) >= 2:
        cell_xml = cells5[1].replace("认证范围：", "认证范围：" + str(scope), 1)
        trs[5] = row5.replace(cells5[1], cell_xml, 1)
    
    # Row 21: certification decision
    row21 = trs[21]
    cells21 = re.findall(r"(<w:tc[^>]*>.*?</w:tc>)", row21, re.DOTALL)
    if len(cells21) >= 2:
        cell_xml = cells21[1]
        paras = re.findall(r"<w:p[^>]*>.*?</w:p>", cell_xml, re.DOTALL)
        should_check = -1
        if audit_type in INITIAL_TYPES | RECERT_TYPES:
            should_check = 0
        elif audit_type in TRANSFER_TYPES:
            should_check = 2
        elif audit_type in SURVEILLANCE_TYPES:
            should_check = 4
        new_paras = list(paras)
        for pi, para in enumerate(paras):
            if pi <= 6:
                if pi == should_check:
                    new_paras[pi] = para.replace('<w:checked w:val="0"/>', '<w:checked w:val="1"/>', 1)
                else:
                    new_paras[pi] = para.replace('<w:checked w:val="1"/>', '<w:checked w:val="0"/>', 1)
        para_positions, remaining, offset = [], cell_xml, 0
        for p in paras:
            idx = remaining.find(p)
            if idx < 0: break
            para_positions.append(offset + idx)
            offset = offset + idx + len(p)
            remaining = remaining[idx + len(p):]
        if len(new_paras) > 9 and "日期：" in new_paras[9]:
            new_paras[9] = new_paras[9].replace("日期：", "日期：" + str(date), 1)
        new_cell_parts = []
        prev_end = 0
        for i, pos in enumerate(para_positions):
            new_cell_parts.append(cell_xml[prev_end:pos])
            new_cell_parts.append(new_paras[i] if i < len(new_paras) else paras[i])
            prev_end = pos + len(paras[i])
        new_cell_parts.append(cell_xml[prev_end:])
        trs[21] = row21.replace(cells21[1], "".join(new_cell_parts), 1)
    
    # Rebuild table
    new_tbl = row_parts[0]
    for tr in trs:
        new_tbl += tr
    last_tr_idx = 0
    for i in range(len(row_parts)):
        if row_parts[i].startswith("<w:tr"):
            last_tr_idx = i
    new_tbl += "".join(row_parts[last_tr_idx + 1:])
    new_xml = xml_str[:tbl_start] + new_tbl + xml_str[tbl_end:]
    contents["word/document.xml"] = new_xml.encode("utf-8")
    
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", compression=zipfile.ZIP_DEFLATED) as zo:
        for name, data in contents.items():
            zo.writestr(name, data)
    return out.getvalue()

# Sidebar
with st.sidebar:
    st.markdown("### 📁 上传文件")
    form_file = st.file_uploader("上传 FORM6101 (.docx)", type=["docx"], key="form")
    template_file = st.file_uploader("上传报告模板 (.docx)", type=["docx"], key="template")
    
    if st.button("🔄 清空", use_container_width=True):
        st.rerun()

# Main area
col1, col2 = st.columns(2)
with col1:
    if form_file:
        st.success(f"✅ 已上传: {form_file.name}")
    else:
        st.info("📄 待上传 FORM6101")
with col2:
    if template_file:
        st.success(f"✅ 已上传: {template_file.name}")
    else:
        st.info("📝 待上传报告模板")

if form_file and template_file:
    with st.spinner("正在解析 FORM6101..."):
        form_data, err = extract_form_data(form_file.getvalue())
    
    if err:
        st.error(err)
    else:
        st.markdown("### 📋 已提取字段")
        for k, v in form_data.items():
            st.text(f"**{k}**: {v or '（未找到）'}")
        
        if st.button("📄 生成报告", type="primary", use_container_width=True):
            with st.spinner("正在生成报告..."):
                try:
                    result_bytes = fill_template(template_file.getvalue(), form_data)
                    ts = datetime.now()
                    filename = f"{sanitize_filename(form_data.get('company','报告'))}_{ts.strftime('%Y%m%d_%H%M')}.docx"
                    st.download_button(
                        label="⬇️ 下载生成的报告",
                        data=result_bytes,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
                    st.success("✅ 报告生成成功！")
                except Exception as e:
                    st.error(f"生成失败: {e}")
                    st.code(str(e))
elif form_file:
    st.info("请同时上传报告模板")
elif template_file:
    st.info("请同时上传 FORM6101 文件")
else:
    st.markdown("""
    ### 📋 使用说明
    1. 上传 **FORM6101** Word 文件（认证评定记录）
    2. 上传 **报告模板** Word 文件
    3. 系统自动提取公司名、任务号、审核组长等信息
    4. 点击「生成报告」下载填充好的报告
    """)