# -*- coding: utf-8 -*-
import io
import re
import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor

st.set_page_config(
    page_title="认证评定记录全量解析与模版生成系统", layout="wide"
)


# ==========================================
# 1. 核心解析引擎 (四表联动与异常修复)
# ==========================================
def parse_and_fix_excel(file_buffer):
    """四表联动解析与数据清洗核心引擎 (Sheet1 ~ Sheet4)"""
    xls = pd.ExcelFile(file_buffer)

    df1 = pd.read_excel(xls, sheet_name="Sheet1")
    df2 = (
        pd.read_excel(xls, sheet_name="Sheet2")
        if "Sheet2" in xls.sheet_names
        else pd.DataFrame()
    )
    df3 = (
        pd.read_excel(xls, sheet_name="Sheet3")
        if "Sheet3" in xls.sheet_names
        else pd.DataFrame()
    )
    df4 = (
        pd.read_excel(xls, sheet_name="Sheet4")
        if "Sheet4" in xls.sheet_names
        else pd.DataFrame()
    )

    if not df3.empty:
        df3 = df3.rename(
            columns={
                df3.columns[0]: "任务号",
                "Observations": "Sheet3_结论",
                "Date": "Sheet3_日期",
            }
        )
        df3 = df3.drop_duplicates(subset=["任务号"], keep="first")

    if not df4.empty:
        df4.columns = df4.iloc[0]
        df4 = df4[1:].reset_index(drop=True)
        df4 = df4.rename(columns={"File number(s)": "任务号"})
        df4 = df4.drop_duplicates(subset=["任务号"], keep="first")

    if not df2.empty and "任务号" in df2.columns:
        df2 = df2.drop_duplicates(subset=["任务号"], keep="first")

    master_list = []
    anomaly_log = []

    for idx, row in df1.iterrows():
        task_no = str(row.get("任务号", "")).strip()

        row2 = (
            df2[df2["任务号"] == task_no].iloc[0]
            if (not df2.empty and task_no in df2["任务号"].values)
            else pd.Series()
        )
        row3 = (
            df3[df3["任务号"] == task_no].iloc[0]
            if (not df3.empty and task_no in df3["任务号"].values)
            else pd.Series()
        )
        row4 = (
            df4[df4["任务号"] == task_no].iloc[0]
            if (not df4.empty and task_no in df4["任务号"].values)
            else pd.Series()
        )

        # 公司名称提取与邮箱污染修正
        s1_company = str(row.get("客户名称 Client Name", "")).strip()
        s2_company = (
            str(row2.get("企业中文名字", row2.get("企业名称", ""))).strip()
            if not row2.empty
            else ""
        )
        s4_company = (
            str(row4.get("Company name", "")).strip() if not row4.empty else ""
        )

        is_email_polluted = "@" in s1_company
        if (
            is_email_polluted
            or not s1_company
            or s1_company.lower() in ["nan", "none", "null"]
        ):
            company_name = (
                s2_company
                if s2_company and s2_company.lower() != "nan"
                else (
                    s4_company
                    if s4_company and s4_company.lower() != "nan"
                    else "未知企业"
                )
            )
            if is_email_polluted:
                anomaly_log.append(
                    {
                        "任务号": task_no,
                        "异常类型": "邮箱污染公司名",
                        "原污染值": s1_company,
                        "修复后值": company_name,
                    }
                )
        else:
            company_name = s1_company

        # 英文名称
        company_en = (
            str(
                row2.get(
                    "企业英文名字",
                    s4_company if s4_company != company_name else "",
                )
            ).strip()
            if not row2.empty
            else s4_company
        )
        if company_en.lower() in ["nan", "none", "null"]:
            company_en = ""

        # 审核团队组合
        lead = str(
            row.get("审核组长", row2.get("组长", "") if not row2.empty else "")
        ).strip()
        members = (
            str(row2.get("组员", "")).strip()
            if (not row2.empty and pd.notna(row2.get("组员")))
            else ""
        )
        team_str = (
            f"{lead} (成员: {members})"
            if (members and members.lower() != "nan")
            else lead
        )

        # 审核地址清洗
        address = str(row.get("审核地址", "")).strip()
        is_address_date = re.match(r"^\d{4}[-/]\d{2}[-/]\d{2}", address)
        if is_address_date or address.lower() in ["nan", "none", "null", ""]:
            real_address = (
                str(row2.get("审核地址", "")).strip() if not row2.empty else ""
            )
            if is_address_date:
                anomaly_log.append(
                    {
                        "任务号": task_no,
                        "异常类型": "地址错入日期时间戳",
                        "原污染值": address,
                        "修复后值": real_address,
                    }
                )
            address = (
                real_address if real_address.lower() != "nan" else "未填写"
            )

        # 认证范围与标准
        scope = str(row.get("认证范围", "")).strip()
        if not scope or scope.lower() in ["nan", "none", "null"]:
            scope = (
                str(row2.get("审核范围", "")).strip() if not row2.empty else ""
            )
        if scope.lower() in ["nan", "none", "null"]:
            scope = ""

        standard = (
            str(row2.get("标准", "")).strip() if not row2.empty else "ISO/IATF"
        )
        if standard.lower() in ["nan", "none", "null"]:
            standard = ""

        # 认证结论与日期
        decision = str(row.get("认证决定结论", "")).strip()
        if not decision or decision.lower() in ["nan", "none", "null"]:
            decision = (
                str(
                    row3.get("Sheet3_结论", row4.get("Observations", ""))
                ).strip()
                if not row3.empty
                else ""
            )

        date_val = str(row.get("日期", "")).strip()
        if not date_val or date_val.lower() in ["nan", "0", "none"]:
            date_val = (
                str(
                    row3.get("Sheet3_日期", row4.get("VP pass date", ""))
                ).strip()
                if not row3.empty
                else ""
            )

        # 其它列
        code = (
            str(row2.get("专业代码", "")).strip() if not row2.empty else ""
        )
        fin_status = (
            str(row2.get("财务收费", "")).strip() if not row2.empty else ""
        )
        vp_pass = (
            str(row4.get("VP pass date", row2.get("VP审批", ""))).strip()
            if not row4.empty
            else ""
        )

        master_list.append(
            {
                "序号": row.get("项目序号 No.", idx + 1),
                "合同号": (
                    str(row.get("合同号 Contract No.", "")).strip()
                    if str(row.get("合同号 Contract No.", "")).lower() != "nan"
                    else ""
                ),
                "任务号": task_no,
                "公司中文名": company_name,
                "公司英文名": company_en,
                "审核类型": (
                    str(
                        row.get(
                            "审核类型Audit Type",
                            row2.get("审核类型", "")
                            if not row2.empty
                            else "",
                        )
                    ).strip()
                ),
                "认证标准": standard,
                "审核团队": team_str,
                "评定人员": (
                    str(
                        row.get(
                            "评定人员",
                            row2.get("评定人员", "") if not row2.empty else "",
                        )
                    ).strip()
                ),
                "审核地址": address,
                "认证范围": scope,
                "认证结论": decision,
                "结论日期": (
                    date_val
                    if date_val.lower() not in ["nan", "none", "null", "0"]
                    else ""
                ),
                "专业代码": code if code.lower() not in ["nan", "0"] else "",
                "财务状态": (
                    fin_status if fin_status.lower() not in ["nan", "0"] else ""
                ),
                "VP审批时间": (
                    vp_pass if vp_pass.lower() not in ["nan", "0"] else ""
                ),
            }
        )

    return pd.DataFrame(master_list), pd.DataFrame(anomaly_log)


# ==========================================
# 2. Word 模板填充引擎 (支持用户自定义模板)
# ==========================================
def fill_word_template_single(data_dict, template_bytes=None):
    """为单条记录填充 Word 模板（若无模板则生成标准卡片）"""
    if template_bytes:
        # 使用用户上传的自定义 Word 模板填充
        doc = Document(io.BytesIO(template_bytes))

        def replace_in_paragraphs(paragraphs, data):
            for p in paragraphs:
                for k, v in data.items():
                    tag = f"{{{{{k}}}}}"  # {{字段名}}
                    if tag in p.text:
                        p.text = p.text.replace(
                            tag, str(v) if pd.notna(v) and v != "" else ""
                        )

        replace_in_paragraphs(doc.paragraphs, data_dict)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    replace_in_paragraphs(cell.paragraphs, data_dict)

    else:
        # 未上传模板时，使用系统内置默认文档布局
        doc = Document()
        doc.add_heading(f"认证评定单项报告 - {data_dict['公司中文名']}", level=1)

        p = doc.add_paragraph()
        p.add_run("• 公司中文名：").bold = True
        p.add_run(f"{data_dict['公司中文名']}\n")

        p.add_run("• 公司英文名：").bold = True
        p.add_run(f"{data_dict['公司英文名']}\n")

        p.add_run("• 任务号：").bold = True
        p.add_run(f"{data_dict['任务号']}   |   ")
        p.add_run("合同号：").bold = True
        p.add_run(f"{data_dict['合同号']}\n")

        p.add_run("• 认证标准：").bold = True
        p.add_run(f"{data_dict['认证标准']}   |   ")
        p.add_run("审核类型：").bold = True
        p.add_run(f"{data_dict['审核类型']}\n")

        p.add_run("• 审核团队：").bold = True
        p.add_run(f"{data_dict['审核团队']}   |   ")
        p.add_run("评定人员：").bold = True
        p.add_run(f"{data_dict['评定人员']}\n")

        p.add_run("• 认证结论：").bold = True
        r_dec = p_body = p.add_run(f"{data_dict['认证结论']}")
        r_dec.bold = True
        p.add_run(f"   |   结论日期：{data_dict['结论日期']}\n")

        p.add_run("• 审核地址：").bold = True
        p.add_run(f"{data_dict['审核地址']}\n")

        p.add_run("• 认证范围：").bold = True
        p.add_run(f"{data_dict['认证范围']}")

    target_stream = io.BytesIO()
    doc.save(target_stream)
    target_stream.seek(0)
    return target_stream.getvalue()


def generate_word_bytes_batch(df, template_bytes=None):
    """批量导出所有记录的汇总 Word 报告"""
    doc = Document()
    title = doc.add_heading("认证评定记录全量汇总报告", level=1)
    title.runs[0].font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)

    for idx, row in df.iterrows():
        p_head = doc.add_paragraph()
        r_head = p_head.add_run(f"【{idx + 1}】 {row['公司中文名']} ")
        r_head.bold = True
        r_head.font.size = Pt(12)

        if row["公司英文名"]:
            r_en = p_head.add_run(f"({row['公司英文名']})")
            r_en.italic = True

        p_body = doc.add_paragraph()
        p_body.paragraph_format.left_indent = Inches(0.2)
        p_body.add_run("• 任务号：").bold = True
        p_body.add_run(f"{row['任务号']}   |   合同号：{row['合同号']}\n")
        p_body.add_run("• 认证标准：").bold = True
        p_body.add_run(f"{row['认证标准']}   |   审核类型：{row['审核类型']}\n")
        p_body.add_run("• 审核团队：").bold = True
        p_body.add_run(f"{row['审核团队']}   |   评定人员：{row['评定人员']}\n")
        p_body.add_run("• 认证结论：").bold = True
        p_body.add_run(f"{row['认证结论']}   |   结论日期：{row['结论日期']}\n")
        p_body.add_run("• 审核地址：").bold = True
        p_body.add_run(f"{row['审核地址']}\n")
        p_body.add_run("• 认证范围：").bold = True
        p_body.add_run(f"{row['认证范围']}")

        doc.add_paragraph("-" * 65)

    target_stream = io.BytesIO()
    doc.save(target_stream)
    target_stream.seek(0)
    return target_stream.getvalue()


def generate_excel_bytes(df):
    """导出 Excel 二进制流"""
    target_stream = io.BytesIO()
    with pd.ExcelWriter(target_stream, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="认证评定解析全量表")
    target_stream.seek(0)
    return target_stream.getvalue()


# ==========================================
# 3. Streamlit 主界面
# ==========================================
st.title("🛡️ 认证评定自动化解析与模版生成系统")

# 顶部双列上传组件
up_col1, up_col2 = st.columns(2)
with up_col1:
    excel_file = st.file_uploader(
        "1. 上传认证评定 Excel 数据文件 (.xlsx / .xls)",
        type=["xlsx", "xls"],
    )
with up_col2:
    template_file = st.file_uploader(
        "2. (可选) 上传自定义 Word 模板 (.docx)", type=["docx"]
    )
    with st.expander("💡 提示：自定义模板占位符写法"):
        st.markdown(
            """
        在 Word 模板中可直接输入以下标签，系统自动替换对应文本：
        - `{{公司中文名}}` 、 `{{公司英文名}}`
        - `{{任务号}}` 、 `{{合同号}}`
        - `{{审核团队}}` 、 `{{评定人员}}`
        - `{{认证标准}}` 、 `{{审核类型}}`
        - `{{审核地址}}` 、 `{{认证范围}}`
        - `{{认证结论}}` 、 `{{结论日期}}`
        """
        )

# 读取模板文件流
template_bytes = template_file.getvalue() if template_file else None

if excel_file is not None:
    try:
        with st.spinner("正在执行多表深度匹配、数据脱敏与自动修复..."):
            df_master, df_anomalies = parse_and_fix_excel(excel_file)

        # 核心指标
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("总解析记录数", f"{len(df_master)} 条")
        col2.metric("邮箱污染修复数", f"{len(df_anomalies)} 处")
        col3.metric(
            "合格发证数",
            f"{len(df_master[df_master['认证结论'].str.contains('可发证', na=False)])} 家",
        )
        col4.metric(
            "审核标准覆盖",
            f"{df_master['认证标准'].nunique()} 种",
        )

        st.markdown("---")

        # 侧边栏过滤
        st.sidebar.header("🔍 数据筛选与检索")
        search_kw = st.sidebar.text_input("搜索企业名称或任务号:")
        selected_decision = st.sidebar.multiselect(
            "按认证结论筛选:",
            options=df_master["认证结论"].unique().tolist(),
            default=[],
        )
        selected_standard = st.sidebar.multiselect(
            "按认证标准筛选:",
            options=df_master["认证标准"].unique().tolist(),
            default=[],
        )

        filtered_df = df_master.copy()
        if search_kw:
            filtered_df = filtered_df[
                filtered_df["公司中文名"].str.contains(search_kw, na=False)
                | filtered_df["任务号"].str.contains(search_kw, na=False)
                | filtered_df["公司英文名"].str.contains(search_kw, na=False)
            ]
        if selected_decision:
            filtered_df = filtered_df[
                filtered_df["认证结论"].isin(selected_decision)
            ]
        if selected_standard:
            filtered_df = filtered_df[
                filtered_df["认证标准"].isin(selected_standard)
            ]

        # 核心五大功能选项卡
        tab_single, tab_batch, tab_data, tab_log, tab_chart = st.tabs(
            [
                "🎯 单条记录生成 (Single Generate)",
                "📦 批量数据导出 (Batch Export)",
                "📋 完整解析全量表",
                "⚠️ 异常数据修复日志",
                "📊 统计图表分析",
            ]
        )

        # ----------------------------------------------------
        # Tab 1: 单条记录生成 (Single Generate)
        # ----------------------------------------------------
        with tab_single:
            st.subheader("🎯 选定单家企业生成/下载独立文档")
            company_list = filtered_df["公司中文名"].tolist()

            if company_list:
                selected_company = st.selectbox(
                    "请选择目标企业:", options=company_list
                )
                single_row = filtered_df[
                    filtered_df["公司中文名"] == selected_company
                ].iloc[0]
                single_dict = single_row.to_dict()

                # 单条记录预览卡片
                st.info(
                    f"**已选中记录**：{single_dict['公司中文名']} (任务号: {single_dict['任务号']})"
                )

                card_col1, card_col2 = st.columns(2)
                with card_col1:
                    st.write(f"**公司英文名**: {single_dict['公司英文名']}")
                    st.write(f"**合同号**: {single_dict['合同号']}")
                    st.write(f"**审核团队**: {single_dict['审核团队']}")
                    st.write(f"**评定人员**: {single_dict['评定人员']}")
                    st.write(f"**认证标准**: {single_dict['认证标准']}")
                with card_col2:
                    st.write(f"**审核类型**: {single_dict['审核类型']}")
                    st.write(f"**认证结论**: {single_dict['认证结论']}")
                    st.write(f"**结论日期**: {single_dict['结论日期']}")
                    st.write(f"**审核地址**: {single_dict['审核地址']}")
                    st.write(f"**认证范围**: {single_dict['认证范围']}")

                st.markdown("---")

                # 生成单条 Word / Excel 按钮
                s_btn1, s_btn2 = st.columns(2)

                single_word_bytes = fill_word_template_single(
                    single_dict, template_bytes
                )
                template_status_label = (
                    "已使用套用自定义 Word 模版"
                    if template_bytes
                    else "已使用系统内置规范样式"
                )

                s_btn1.download_button(
                    label=f"📄 下载该单条记录 Word 报告 ({template_status_label})",
                    data=single_word_bytes,
                    file_name=f"{single_dict['公司中文名']}_评定记录.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )

                single_df = pd.DataFrame([single_dict])
                single_excel_bytes = generate_excel_bytes(single_df)
                s_btn2.download_button(
                    label="📊 下载该单条记录 Excel 表格",
                    data=single_excel_bytes,
                    file_name=f"{single_dict['公司中文名']}_评定记录.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.warning("当前筛选条件下未找到任何记录。")

        # ----------------------------------------------------
        # Tab 2: 批量导出 (Batch Export)
        # ----------------------------------------------------
        with tab_batch:
            st.subheader("📦 批量导出汇总报告与全量表格")
            st.write(f"当前选定导出记录条数: **{len(filtered_df)}** 条")

            b_btn1, b_btn2 = st.columns(2)
            batch_excel_data = generate_excel_bytes(filtered_df)
            b_btn1.download_button(
                label="📥 导出批量 Excel 汇总表 (.xlsx)",
                data=batch_excel_data,
                file_name="认证评定记录_批量汇总.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            batch_word_data = generate_word_bytes_batch(
                filtered_df, template_bytes
            )
            b_btn2.download_button(
                label="📥 导出批量 Word 汇总报告 (.docx)",
                data=batch_word_data,
                file_name="认证评定汇总报告.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

        # ----------------------------------------------------
        # Tab 3: 数据表格视图
        # ----------------------------------------------------
        with tab_data:
            st.dataframe(filtered_df, use_container_width=True)

        # ----------------------------------------------------
        # Tab 4: 修复日志视图
        # ----------------------------------------------------
        with tab_log:
            if not df_anomalies.empty:
                st.warning(
                    f"系统共自动识别并修复了 {len(df_anomalies)} 项错位及邮箱污染数据："
                )
                st.dataframe(df_anomalies, use_container_width=True)
            else:
                st.success("数据质量良好，未检测到邮箱污染或明显错位字段！")

        # ----------------------------------------------------
        # Tab 5: 图表视图
        # ----------------------------------------------------
        with tab_chart:
            c1, c2 = st.columns(2)
            with c1:
                st.subheader("认证结论分布")
                st.bar_chart(filtered_df["认证结论"].value_counts())
            with c2:
                st.subheader("认证标准分布")
                st.bar_chart(filtered_df["认证标准"].value_counts())

    except Exception as e:
        st.error(f"处理文件时发生错误，请检查上传的文件格式: {str(e)}")
else:
    st.info("👈 请在左上方上传您的 Excel 文件，可同时选择上传 Word 模版。")
