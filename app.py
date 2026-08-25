# -*- coding: utf-8 -*-
import io
import re
import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor

st.set_page_config(page_title="认证评定记录全量解析系统", layout="wide")


def parse_and_fix_excel(file_buffer):
    """四表联动解析与数据清洗核心引擎 (Sheet1 ~ Sheet4)"""
    xls = pd.ExcelFile(file_buffer)

    # 1. 加载所有工作表
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

    # 2. 预处理副表表头与索引
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

    # 3. 逐行联合解析
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

        # 公司名称提取与邮箱修复
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

        # 审核团队组合 (组长 + 组员)
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

        # 认证范围
        scope = str(row.get("认证范围", "")).strip()
        if not scope or scope.lower() in ["nan", "none", "null"]:
            scope = (
                str(row2.get("审核范围", "")).strip() if not row2.empty else ""
            )
        if scope.lower() in ["nan", "none", "null"]:
            scope = ""

        # 标准类型
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

        # 附加属性
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


def generate_word_bytes(df):
    """导出排版精美的 Word 报告文件"""
    doc = Document()

    # 标题
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
            r_en.font.color.rgb = RGBColor(0x59, 0x59, 0x59)

        p_body = doc.add_paragraph()
        p_body.paragraph_format.left_indent = Inches(0.2)
        p_body.paragraph_format.line_spacing = 1.15

        p_body.add_run("• 任务号：").bold = True
        p_body.add_run(f"{row['任务号']}   |   ")
        p_body.add_run("合同号：").bold = True
        p_body.add_run(f"{row['合同号'] if row['合同号'] else '无'}\n")

        p_body.add_run("• 认证标准：").bold = True
        p_body.add_run(f"{row['认证标准']}   |   ")
        p_body.add_run("审核类型：").bold = True
        p_body.add_run(f"{row['审核类型']}\n")

        p_body.add_run("• 审核团队：").bold = True
        p_body.add_run(f"{row['审核团队']}   |   ")
        p_body.add_run("评定人员：").bold = True
        p_body.add_run(f"{row['评定人员']}\n")

        p_body.add_run("• 认证结论：").bold = True
        r_dec = p_body.add_run(f"{row['认证结论']}")
        r_dec.bold = True
        r_dec.font.color.rgb = (
            RGBColor(0x00, 0x80, 0x00)
            if "通过" in row["认证结论"]
            else RGBColor(0xC0, 0x00, 0x00)
        )
        p_body.add_run(
            f"   |   结论日期：{row['结论日期'] if row['结论日期'] else '待定'}\n"
        )

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
    """导出包含完整 18 维度数据的 Excel 文件"""
    target_stream = io.BytesIO()
    with pd.ExcelWriter(target_stream, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="认证评定解析全量表")
    target_stream.seek(0)
    return target_stream.getvalue()


# ---------------- Streamlit UI 界面 ----------------
st.title("🛡️ 认证评定记录自动化全量解析与修复系统")

uploaded_file = st.file_uploader(
    "请上传认证评定记录文件 (支持 .xlsx, .xls)", type=["xlsx", "xls"]
)

if uploaded_file is not None:
    try:
        with st.spinner("正在执行多表深度匹配、数据脱敏与自动修复..."):
            df_master, df_anomalies = parse_and_fix_excel(uploaded_file)

        # 核心指标显示
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

        # 搜索与筛选控制栏
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

        # 应用筛选逻辑
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

        # 选项卡视图
        tab1, tab2, tab3 = st.tabs(
            [
                "📋 完整数据解析结果",
                "⚠️ 异常数据修复日志",
                "📊 统计图表分析",
            ]
        )

        with tab1:
            st.dataframe(filtered_df, use_container_width=True)

        with tab2:
            if not df_anomalies.empty:
                st.warning(
                    f"系统共自动识别并修复了 {len(df_anomalies)} 项错位及邮箱污染数据："
                )
                st.dataframe(df_anomalies, use_container_width=True)
            else:
                st.success("数据质量良好，未检测到邮箱污染或明显错位字段！")

        with tab3:
            col_chart1, col_chart2 = st.columns(2)
            with col_chart1:
                st.subheader("认证结论分布")
                st.bar_chart(filtered_df["认证结论"].value_counts())
            with col_chart2:
                st.subheader("认证标准分布")
                st.bar_chart(filtered_df["认证标准"].value_counts())

        # 下载按钮区域
        st.markdown("---")
        st.subheader("📥 导出导出报告与数据")
        btn_col1, btn_col2 = st.columns(2)

        excel_data = generate_excel_bytes(filtered_df)
        btn_col1.download_button(
            label="下载全量修复后的 Excel 表格 (.xlsx)",
            data=excel_data,
            file_name="认证评定记录_全量修复版.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        word_data = generate_word_bytes(filtered_df)
        btn_col2.download_button(
            label="导出规范排版的 Word 总结报告 (.docx)",
            data=word_data,
            file_name="认证评定汇总报告.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    except Exception as e:
        st.error(f"处理文件时发生错误，请检查上传的文件格式: {str(e)}")
else:
    st.info("👈 请在左上方选择并上传您的 Excel 评定文件以启动解析。")
