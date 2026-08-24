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

def set_fcb(doc_xml, pos, checked):
    '''Set FORMCHECKBOX at absolute position to checked/unchecked'''
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

def fill_cert_standard_by_task(cell_text, task_no):
    """勾选规则：
    - 包含 TS 且无 ER -> 仅勾选 IATF16949
    - 包含 ER 且无 TS -> 仅勾选 ISO9001
    - 同时包含 TS 和 ER -> 同时勾选 IATF16949 和 ISO9001
    """
    tn = str(task_no).upper()
    has_ts = 'TS' in tn
    has_er = 'ER' in tn
    
    result = cell_text
    if has_ts and 'IATF16949' in cell_text:
        result = result.replace(CHK_EMPTY + ' IATF16949', CHK_FILLED + ' IATF16949')
        result = result.replace(CHK_EMPTY + 'IATF16949', CHK_FILLED + 'IATF16949')
    if has_er and 'ISO9001' in cell_text:
        result = result.replace(CHK_EMPTY + ' ISO9001', CHK_FILLED + ' ISO9001')
        result = result.replace(CHK_EMPTY + 'ISO9001', CHK_FILLED + 'ISO9001')
    return result

def fill_audit_type_new(cell_text, atype):
    """审核类型勾选逻辑:
    - 包含 二阶段 -> 初审
    - 包含 监一 或 监二 -> 监审
    - 包含 再认证 或 转移 -> 再认证/转移
    - 包含 特殊 -> 特殊审核
    """
    at = str(atype).strip()
    result = cell_text
    if '二阶段' in at:
        result = result.replace(CHK_EMPTY + '初审', CHK_FILLED + '初审')
    if '监一' in at or '监二' in at:
        result = result.replace(CHK_EMPTY + '监审', CHK_FILLED + '监审')
    if '再认证' in at or '转移' in at:
        result = result.replace(CHK_EMPTY + '再认证/转移', CHK_FILLED + '再认证/转移')
    if '特殊' in at:
        result = result.replace(CHK_EMPTY + '特殊审核', CHK_FILLED + '特殊审核')
    return result

def get_conclusion_idx_new(atype, decision):
    """认证决定结论选项索引 (0 - 6):
    0: 通过，可发证（适用于：初审、再认证）
    1: 通过，暂停恢复审核，可发证
    2: 通过，可换发证书（转机构）
    3: 通过，同意扩大认证范围，可发证
    4: 通过，不换证（适用于：至上次认证决定后，企业证书信息无变更的监督项目）
    5: 通过，可换发新的认证证书（适用于：至上次认证决定后，企业证书信息有变更的监督项目）
    6: 不予通过
    """
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
        return 3  # 对应“同意扩大认证范围，可发证”/换发
    return 0

def fill_report(doc, fields):
    company = fields.get('company', '')
    taskNo = fields.get('taskNo', '')
    leader = fields.get('leader', '')
    auditType = fields.get('auditType', '')
    address = fields.get('address', '')
    scope = fields.get('scope', '')
    date_val = fields.get('date', '')
    decision = fields.get('decision', '')

    # 1. 替换段落文本 (包含日期等)
    if date_val:
        for para in doc.paragraphs:
            if '日期' in para.text:
                for run in para.runs:
                    if '日期' in run.text and date_val not in para.text:
                        run.text = run.text + ' ' + str(date_val)

    # 2. 表格数据填入与选择框修改
    for table in doc.tables:
        for ri, row in enumerate(table.rows):
            cells = row.cells
            
            # 填入单元格基础信息
            for cell in cells:
                for para in cell.paragraphs:
                    full_p = para.text
                    # 任务号
                    if taskNo and '任务号' in full_p and taskNo not in full_p:
                        for run in para.runs:
                            if '任务号' in run.text:
                                run.text = run.text + ' ' + taskNo
                    # 公司名称
                    if company and '公司名称' in full_p and company not in full_p:
                        for run in para.runs:
                            if '公司名称' in run.text:
                                run.text = run.text + ' ' + company
                    # 审核组长
                    if leader and '审核组长' in full_p and leader not in full_p:
                        for run in para.runs:
                            if '审核组长' in run.text:
                                run.text = run.text + ' ' + leader
                    # 审核地址
                    if address and '审核地址' in full_p and address not in full_p:
                        for run in para.runs:
                            if '审核地址' in run.text:
                                run.text = run.text + ' ' + address
                    # 认证范围
                    if scope and '认证范围' in full_p and scope not in full_p:
                        for run in para.runs:
                            if '认证范围' in run.text:
                                run.text = run.text + ' ' + scope
                    # 日期
                    if date_val and '日期' in full_p and date_val not in full_p:
                        for run in para.runs:
                            if '日期' in run.text:
                                run.text = run.text + ' ' + date_val

                    # 认证标准（依据任务号 TS / ER 勾选）
                    if taskNo and ('IATF16949' in full_p or 'ISO9001' in full_p):
                        new_text = fill_cert_standard_by_task(para.text, taskNo)
                        if new_text != para.text:
                            for r in para.runs: r.text = ''
                            para.runs[0].text = new_text

                    # 审核类型勾选
                    if auditType and ('初审' in full_p or '监审' in full_p or '再认证' in full_p):
                        new_text = fill_audit_type_new(para.text, auditType)
                        if new_text != para.text:
                            for r in para.runs: r.text = ''
                            para.runs[0].text = new_text

    # 3. 认证决定结论 FORMCHECKBOX 勾选
    if auditType:
        con_idx = get_conclusion_idx_new(auditType, decision)
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
        
        # 结论复选框默认位于最后7个 FORMCHECKBOX 位置（按原程序 FCB 66-72）
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

def read_row(ws, row, fmt):
    vals = [c.value for c in row]
    
    # 提取第一个组长姓名
    leader_val = str(vals[3]).split('+')[0].strip() if len(vals) > 3 and vals[3] else ''
    
    return {
        'company': str(vals[2]).strip() if len(vals) > 2 and vals[2] else '',
        'leader': leader_val,
        'auditType': str(vals[4]).strip() if len(vals) > 4 and vals[4] else '',
        'address': str(vals[6]).strip() if len(vals) > 6 and vals[6] else '',
        'scope': str(vals[7]).strip() if len(vals) > 7 and vals[7] else '',
        'taskNo': str(vals[8]).strip() if len(vals) > 8 and vals[8] else '',
        'date': format_date(vals[11]) if len(vals) > 11 and vals[11] else '',
        'decision': str(vals[10]).strip() if len(vals) > 10 and vals[10] else ''  # 获取认证决定结论文本
    }

st.markdown('---')
st.caption('Cert Report Generator v2.8')
