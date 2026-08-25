# -*- coding: utf-8 -*-
import pandas as pd
import re
try:
    from docx import Document
except ImportError:
    Document = None

def parse_and_fix_excel(excel_path):
    """
    读取评定记录 Excel 文件，修复字段错位与邮箱污染问题
    """
    xls = pd.ExcelFile(excel_path)
    df_sheet1 = pd.read_excel(excel_path, sheet_name='Sheet1')
    
    # 备用表 Sheet2，用于在主表丢失/污染公司名称时进行反查
    df_sheet2 = pd.read_excel(excel_path, sheet_name='Sheet2') if 'Sheet2' in xls.sheet_names else pd.DataFrame()

    cleaned_records = []

    for idx, row in df_sheet1.iterrows():
        # 1. 提取并清洗公司名称（修复邮箱错位问题）
        raw_company = str(row.get('客户名称 Client Name', '')).strip()
        
        # 过滤邮箱污染或空值，若误填邮箱则通过任务号从 Sheet2 进行反查补全
        if '@' in raw_company or not raw_company or raw_company.lower() in ['nan', 'none', 'null']:
            task_no = str(row.get('任务号', '')).strip()
            matched_s2 = df_sheet2[df_sheet2['任务号'] == task_no] if not df_sheet2.empty and '任务号' in df_sheet2.columns else pd.DataFrame()
            
            if not matched_s2.empty and '企业中文名字' in matched_s2.columns and pd.notna(matched_s2.iloc[0]['企业中文名字']):
                raw_company = str(matched_s2.iloc[0]['企业中文名字']).strip()
            elif not matched_s2.empty and '企业名称' in matched_s2.columns and pd.notna(matched_s2.iloc[0]['企业名称']):
                raw_company = str(matched_s2.iloc[0]['企业名称']).strip()
            else:
                raw_company = "未知企业"

        # 2. 提取并清理其他关联字段
        team_lead = str(row.get('审核组长', '')).strip()
        task_no = str(row.get('任务号', '')).strip()
        audit_type = str(row.get('审核类型Audit Type', '')).strip()
        audit_address = str(row.get('审核地址', '')).strip()
        scope = str(row.get('认证范围', '')).strip()
        evaluator = str(row.get('评定人员', '')).strip()
        decision = str(row.get('认证决定结论', '')).strip()
        decision_date = str(row.get('日期', '')).strip()

        # 过滤错位的时间戳或空值填入地址字段的情况
        if re.match(r'^\d{4}[-/]\d{2}[-/]\d{2}', audit_address) or audit_address.lower() == 'nan':
            audit_address = ""

        cleaned_records.append({
            "序号": row.get('项目序号 No.', idx + 1),
            "合同号": str(row.get('合同号 Contract No.', '')).strip(),
            "公司名称": raw_company,
            "审核组长": team_lead,
            "审核类型": audit_type,
            "评定人员": evaluator,
            "审核地址": audit_address,
            "认证范围": scope,
            "任务号": task_no,
            "认证决定结论": decision,
            "日期": decision_date if decision_date.lower() not in ['nan', '0'] else ""
        })

    return pd.DataFrame(cleaned_records)


def export_to_word(df, output_docx_path):
    """
    将清洗后的数据导出为结构化的 Word 文档 (.docx)
    """
    if Document is None:
        print("未检测到 python-docx 库，跳过 Word 导出。可使用 pip install python-docx 安装。")
        return

    doc = Document()
    doc.add_heading('认证评定记录汇总表', level=1)

    for idx, row in df.iterrows():
        doc.add_heading(f"{idx + 1}. {row['公司名称']}", level=2)
        
        p = doc.add_paragraph()
        p.add_run("任务号：").bold = True
        p.add_run(f"{row['任务号']}\n")
        
        p.add_run("审核组长：").bold = True
        p.add_run(f"{row['审核组长']}\n")
        
        p.add_run("审核类型：").bold = True
        p.add_run(f"{row['审核类型']}\n")
        
        p.add_run("评定人员：").bold = True
        p.add_run(f"{row['评定人员']}\n")
        
        p.add_run("认证决定结论：").bold = True
        p.add_run(f"{row['认证决定结论']}\n")
        
        p.add_run("审核地址：").bold = True
        p.add_run(f"{row['审核地址']}\n")
        
        p.add_run("认证范围：").bold = True
        p.add_run(f"{row['认证范围']}")

    doc.save(output_docx_path)
    print(f"Word 文档已成功生成：{output_docx_path}")


if __name__ == '__main__':
    excel_input = '认证评定记录-郑NEW.xlsx'
    excel_output = '认证评定记录_已修复.xlsx'
    docx_output = '认证评定记录汇总.docx'

    # 1. 执行清洗解析
    df_cleaned = parse_and_fix_excel(excel_input)

    # 2. 导出修复后的 Excel
    df_cleaned.to_excel(excel_output, index=False)
    print(f"Excel 文件已成功修复并保存至：{excel_output}")

    # 3. 导出 Word 文档
    export_to_word(df_cleaned, docx_output)
