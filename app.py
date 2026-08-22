# -*- coding: utf-8 -*-
import io, os, re, zipfile, sys
from datetime import datetime
import streamlit as st
import openpyxl
from docx import Document
sys.stdout.reconfigure(encoding="utf-8")
CHK_EMPTY = chr(0x25A1)
CHK_FILLED = chr(0x25A0)
def set_cb(run, text):
    if run.text is None: run.text = text
    else: run.text = run.text.replace(CHK_EMPTY, text).replace(CHK_FILLED, text)
def fill_cb(doc, target, checked=True):
    fill = CHK_FILLED if checked else CHK_EMPTY
    for para in doc.paragraphs:
        for run in para.runs:
            if run.text and target in run.text:
                set_cb(run, fill)
                return True
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        if run.text and target in run.text:
                            set_cb(run, fill)
                            return True
    return False
def format_date(val):
    if val is None: return ""
    if isinstance(val, datetime): return val.strftime("%Y-%m-%d")
    s = str(val).strip()
    m = re.search(r"(\d{4})[年\-/](\d{1,2})[月\-/](\d{1,2})", s)
    if m: return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    return s[:10] if len(s)>=10 else sdef get_conclusion(atype):
    atype=str(atype).strip() if atype else ""
    is_surv="监" in atype
    if "一阶段" in atype or "二阶段" in atype or "再认证" in atype:
        return {"checked":[True,False,False,False,False,False],"fields":["通过，可发证","不通过","通过，可换发证书","不符合发证条件","通过，不换证","通过，可换发新的认证证书"]}
    elif "转移" in atype:
        return {"checked":[False,False,True,False,False,False],"fields":["通过，可发证","不通过","通过，可换发证书","不符合发证条件","通过，不换证","通过，可换发新的认证证书"]}
    elif is_surv:
        if "换发" in atype:
            return {"checked":[False,False,False,False,False,True],"fields":["通过，可发证","不通过","通过，可换发证书","不符合发证条件","通过，不换证","通过，可换发新的认证证书"]}
        else:
            return {"checked":[False,False,False,False,True,False],"fields":["通过，可发证","不通过","通过，可换发证书","不符合发证条件","通过，不换证","通过，可换发新的认证证书"]}
    else:
        return {"checked":[True,False,False,False,False,False],"fields":["通过，可发证","不通过","通过，可换发证书","不符合发证条件","通过，不换证","通过，可换发新的认证证书"]}
def fill_report(doc, fields):
    reps={"{{company}}":fields.get("company",""),"{{taskNo}}":fields.get("taskNo",""),
          "{{leader}}":fields.get("leader",""),"{{auditType}}":fields.get("auditType",""),
          "{{address}}":fields.get("address",""),"{{scope}}":fields.get("scope",""),
          "{{date}}":fields.get("date",datetime.now().strftime("%Y-%m-%d"))}
    for para in doc.paragraphs:
        for run in para.runs:
            for k,v in reps.items():
                if k in run.text: run.text=run.text.replace(k,v)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        for k,v in reps.items():
                            if k in run.text: run.text=run.text.replace(k,v)
    con=get_conclusion(fields.get("auditType",""))
    for name,chk in zip(con["fields"],con["checked"]):
        fill_cb(doc,name,chk)
    return doc
def gen_docx(tpl_bytes, fields):
    buf=io.BytesIO(tpl_bytes)
    with zipfile.ZipFile(buf,"r") as zf:
        xml=zf.read("word/document.xml").decode("utf-8")
        reps={"{{company}}":fields.get("company",""),"{{taskNo}}":fields.get("taskNo",""),
              "{{leader}}":fields.get("leader",""),"{{auditType}}":fields.get("auditType",""),
              "{{address}}":fields.get("address",""),"{{scope}}":fields.get("scope",""),
              "{{date}}":fields.get("date",datetime.now().strftime("%Y-%m-%d"))}
        for k,v in reps.items(): xml=xml.replace(k,v)
        out=io.BytesIO()
        with zipfile.ZipFile(out,"w",zipfile.ZIP_DEFLATED) as oz:
            for name in zf.namelist():
                if name=="word/document.xml": oz.writestr(name,xml.encode("utf-8"))
                else: oz.writestr(name,zf.read(name))
        return out.getvalue()
def extract_form_fields(doc):
    fields = {}
    for table in doc.tables:
        for ri, row in enumerate(table.rows):
            cells = [c.text.strip() for c in row.cells]
            if ri==1 and len(cells)>1 and cells[1]: fields["taskNo"]=cells[1]
            if ri==2 and len(cells)>1 and cells[1]: fields["company"]=cells[1]
            if ri==3 and len(cells)>1 and cells[1]: fields["address"]=cells[1]
            if ri==11 and len(cells)>1 and cells[1]:
                s=cells[1]
                if s.startswith("IATF:"): s=s[5:]
                fields["scope"]=s
            if ri==14 and len(cells)>1 and cells[1]: fields["auditType"]=cells[1]
            if ri==17 and len(cells)>1 and cells[1]: fields["leader"]=cells[1]
    for para in doc.paragraphs:
        t = para.text
        if "任务编号" in t or "任务号" in t: fields["taskNo"]=re.split(r"[:：]",t)[-1].strip()
        if "组织名称" in t or "公司名称" in t: fields["company"]=re.split(r"[:：]",t)[-1].strip()
        if "审核组长" in t: fields["leader"]=re.split(r"[:：]",t)[-1].strip()
        if "审核地址" in t: fields["address"]=re.split(r"[:：]",t)[-1].strip()
        if "认证范围" in t or "IATF:" in t:
            s=re.split(r"[:：]",t)[-1].strip()
            if s.startswith("IATF:"): s=s[5:]
            fields["scope"]=s
        if "审核性质" in t or "审核类型" in t: fields["auditType"]=re.split(r"[:：]",t)[-1].strip()
        m = re.search(r"(\d{4})[年\-/](\d{1,2})[月\-/](\d{1,2})", t)
        if m and ("审核日期" in t or "现场审核日期" in t):
            fields["date"]=f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    fields.setdefault("company",""); fields.setdefault("taskNo","")
    fields.setdefault("leader",""); fields.setdefault("auditType","")
    fields.setdefault("address",""); fields.setdefault("scope","")
    fields.setdefault("date",datetime.now().strftime("%Y-%m-%d"))
    return fields
def detect_format(wb):
    ws=wb.active
    if ws.max_column>=15: return "A"
    elif ws.max_column>=12: return "B"
    return "unknown"
def read_row(ws, row, fmt):
    vals=[c.value for c in row]
    if fmt=="A":
        return {"company":str(vals[3]).strip() if vals[3] else "",
                "leader":str(vals[4]).split("+")[0].strip() if vals[4] else "",
                "auditType":str(vals[5]).strip() if vals[5] else "",
                "address":str(vals[12]).strip() if vals[12] else "",
                "scope":str(vals[13]).strip() if vals[13] else "",
                "taskNo":str(vals[14]).strip() if vals[14] else ""}
    else:
        return {"company":str(vals[2]).strip() if vals[2] else "",
                "leader":str(vals[3]).split("+")[0].strip() if vals[3] else "",
                "auditType":str(vals[4]).strip() if vals[4] else "",
                "address":str(vals[6]).strip() if vals[6] else "",
                "scope":str(vals[7]).strip() if vals[7] else "",
                "taskNo":str(vals[8]).strip() if vals[8] else "",
                "date":format_date(vals[11]) if vals[11] else ""}
def count_rows(ws):
    c=0
    for row in ws.iter_rows(min_row=2):
        if any(cv.value for cv in row): c+=1
    return c
st.set_page_config(page_title="认证报告生成器",page_icon="\U0001f4cb",layout="wide")
st.title("\U0001f4cb 认证报告生成器")
if "mode" not in st.session_state: st.session_state.mode="single"
if "batch_step" not in st.session_state: st.session_state.batch_step=0
if "batch_total" not in st.session_state: st.session_state.batch_total=0
if "batch_files" not in st.session_state: st.session_state.batch_files=[]
if "batch_processed" not in st.session_state: st.session_state.batch_processed=0
if "single_fields" not in st.session_state: st.session_state.single_fields={}
if "form_doc" not in st.session_state: st.session_state.form_doc=None
if "tpl_bytes" not in st.session_state: st.session_state.tpl_bytes=None
if "expl" not in st.session_state: st.session_state.expl=None
if "ws" not in st.session_state: st.session_state.ws=None
if "excel_fmt" not in st.session_state: st.session_state.excel_fmt=None
mode=st.radio("选择生成模式",["单个报告生成","批量生成"],key="mode")
if mode=="单个报告生成":
    st.header("\U0001f4c4 单个报告生成")
    c1,c2=st.columns(2)
    with c1:
        st.subheader("第一步：上传 FORM6101")
        ff=st.file_uploader("上传 FORM6101",type=["docx"],key="form_up")
    with c2:
        st.subheader("第二步：上传报告模板")
        tf=st.file_uploader("上传报告模板",type=["docx"],key="tpl_up")
    if ff and tf:
        if st.session_state.form_doc is None or st.session_state.get("form_name")!=ff.name:
            st.session_state.form_doc=ff.read()
            st.session_state.form_name=ff.name
            st.session_state.single_fields=extract_form_fields(Document(io.BytesIO(st.session_state.form_doc)))
        if st.session_state.tpl_bytes is None or st.session_state.get("tpl_name")!=tf.name:
            st.session_state.tpl_bytes=tf.read()
            st.session_state.tpl_name=tf.name
        f=st.session_state.single_fields
        st.info(f"已提取：公司={f.get(chr(39)+'company'+chr(39),'')}, 任务号={f.get(chr(39)+'taskNo'+chr(39),'')}, 审核组长={f.get(chr(39)+'leader'+chr(39),'')}")
        if st.button("\U0001f4c4 生成报告",type="primary",key="gen_s"):
            with st.spinner("生成中..."):
                try:
                    doc=Document(io.BytesIO(st.session_state.tpl_bytes))
                    doc=fill_report(doc,f)
                    out=io.BytesIO()
                    doc.save(out)
                    out.seek(0)
                    st.success("\u2705 报告生成成功！")
                    st.download_button(label="\u2b07\ufe0f 下载报告",data=out.getvalue(),file_name=f"{f.get('company','报告')}.docx",mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                except Exception as e:
                    st.error(f"生成失败：{e}")
    elif ff or tf:
        st.warning("请同时上传 FORM6101 和报告模板")
else:
    st.header("\U0001f4e6 批量生成")
    c1,c2=st.columns(2)
    with c1:
        ef=st.file_uploader("上传 Excel",type=["xlsx"],key="exc_up")
    with c2:
        tf=st.file_uploader("上传报告模板",type=["docx"],key="b_tpl_up")
    if ef and tf:
        if st.session_state.expl is None or st.session_state.get("exc_name")!=ef.name or st.session_state.get("b_tpl_name")!=tf.name:
            wb=openpyxl.load_workbook(ef,data_only=True)
            ws=wb.active
            fmt=detect_format(wb)
            total=count_rows(ws)
            st.session_state.expl=wb
            st.session_state.ws=ws
            st.session_state.excel_fmt=fmt
            st.session_state.batch_total=total
            st.session_state.batch_processed=0
            st.session_state.batch_step=0
            st.session_state.batch_files=[]
            st.session_state.expl_name=ef.name
            st.session_state.b_tpl_name=tf.name
            st.session_state.b_tpl_bytes=tf.read()
        fmt=st.session_state.excel_fmt
        total=st.session_state.batch_total
        processed=st.session_state.batch_processed
        fmt_label="A型(15列)" if fmt=="A" else "B型(12列)"
        st.info(f"Excel格式：{fmt_label}，共 {total} 行数据")
        prog=min((processed+20)/max(total,1),1.0) if total>0 else 0
        st.progress(prog)
        st.write(f"进度：{processed}/{total}")
        if st.button("\u25b6\ufe0f 开始生成（每次20个）",key="b_start"):
            with st.spinner("生成中..."):
                ws=st.session_state.ws
                fmt=st.session_state.excel_fmt
                step=st.session_state.batch_step
                start=2+step*20
                end=min(start+20,total+1)
                new=[]
                for ri in range(start,end):
                    row=list(ws.iter_rows(min_row=ri,max_row=ri))[0]
                    f=read_row(ws,row,fmt)
                    if not f.get("company"): continue
                    try:
                        b=gen_docx(st.session_state.b_tpl_bytes,f)
                        new.append((f["company"],b))
                    except Exception as e:
                        st.error(f"{f['company']} 失败：{e}")
                if new:
                    st.session_state.batch_files.extend(new)
                    st.session_state.batch_step+=1
                    st.session_state.batch_processed=min(step*20+len(new),total)
                    zb=io.BytesIO()
                    with zipfile.ZipFile(zb,"w",zipfile.ZIP_DEFLATED) as zf:
                        for n,d in new: zf.writestr(f"{n}.docx",d)
                    zb.seek(0)
                    st.session_state.curr_zip=zb.getvalue()
                    st.success(f"\u2705 已生成 {len(new)} 个（累计 {st.session_state.batch_processed}/{total}）")
                    st.download_button(label="\u2b07\ufe0f 下载本批 (ZIP)",data=st.session_state.curr_zip,file_name=f"reports_batch_{step+1}.zip",mime="application/zip")
                    if st.session_state.batch_processed>=total:
                        st.success("\U0001f389 全部生成完成！")
                        if len(st.session_state.batch_files)>1:
                            azb=io.BytesIO()
                            with zipfile.ZipFile(azb,"w",zipfile.ZIP_DEFLATED) as zf:
                                for n,d in st.session_state.batch_files: zf.writestr(f"{n}.docx",d)
                            azb.seek(0)
                            st.download_button(label="\u2b07\ufe0f 下载全部 (ZIP)",data=azb.getvalue(),file_name="all_reports.zip",mime="application/zip")
        if st.button("\U0001f504 清除并重新开始",key="b_clear"):
            st.session_state.batch_step=0; st.session_state.batch_processed=0
            st.session_state.batch_files=[]; st.session_state.expl=None; st.rerun()
    elif ef or tf:
        st.warning("请同时上传 Excel 和报告模板")
st.markdown("---")
st.caption("认证报告生成器 v2.0")
