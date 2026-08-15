import streamlit as st
import zipfile, re, io, os
from datetime import datetime

st.set_page_config(page_title='Cert Report', page_icon='K', layout='wide')
st.title('Cert Report Generator')

INITIAL_TYPES = {'初审二阶段','LOC初审二阶段','初审二阶段(LOC升级)','初审二阶段(QMS）',
                 '初审二阶段（LOC）','初审二阶段（loc升级）','初审二阶段（严重）',
                 '初审二阶段（主场所）','初审二阶段（免一阶段）','初审二阶段（搬迁）',
                 '初审二阶段（本机构搬迁）','特殊审核（二阶段审核后扩范围）',
                 '特殊审核(变更)','特殊审核（变更）','特殊审核（扩范围）'}
SURVEILLANCE_TYPES = {'监一','监一（ISO9001）','监一（严重）','监一（主场所）',
                      '监二','监二（IATF)','监二（ISO）','监二（ISO9001:2015）',
                      '监二（Q）','监二（严重）','监一(严重)','监一（IATF）',
                      '监一（ISO）','监一（Q）','监二(IATF)','监二(ISO 9001)',
                      '监二(ISO9001)','监二（IATF16949）','监二（IATF）','监二（QMS）',
                      '监二（主场所）','监二（主场所，严重）'}
RECERT_TYPES = {'再认证','再认证（严重）','再认证（主场所）'}
TRANSFER_TYPES = {'转移','转移审核','转移（严重）'}

def sanitize(name):
    return re.sub(r'[\\/:*?<>|]', '_', str(name))

def fill_template(template_bytes, data):
    company = data.get('company', '')
    task_no = data.get('taskNo', '')
    leader = data.get('leader', '')
    audit_type = str(data.get('auditType', ''))
    address = data.get('address', '')
    scope = data.get('scope', '')
    date = data.get('date', '')
    with zipfile.ZipFile(io.BytesIO(template_bytes)) as zfin:
        contents = {n: zfin.read(n) for n in zfin.namelist()}
    xml_str = contents['word/document.xml'].decode('utf-8')
    tbl_start = xml_str.find('<w:tbl>')
    tbl_end = xml_str.find('</w:tbl>', tbl_start) + 8
    tbl_xml = xml_str[tbl_start:tbl_end]
    row_parts = re.split(r'(<w:tr[^>]*>.*?</w:tr>)', tbl_xml, flags=re.DOTALL)
    trs = [p for p in row_parts if p.startswith('<w:tr')]
    if len(trs) < 24:
        raise Exception('Expected 24 rows, got ' + str(len(trs)))
    c0, c1 = re.findall(r'(<w:tc[^>]*>.*?</w:tc>)', trs[0])[:2]
    for attrs, text in re.findall(r'<w:t([^>]*?)>([^<]*)</w:t>', c0):
        if text.strip() in (':', ''):
            new = ('<w:t'+attrs+'>'+text+company+'</w:t>') if attrs else ('<w:t>'+text+company+'</w:t>')
            c0 = c0.replace(('<w:t'+attrs+'>'+text+'</w:t>') if attrs else ('<w:t>'+text+'</w:t>'), new, 1)
            break
    for attrs, text in re.findall(r'<w:t([^>]*?)>([^<]*)</w:t>', c1):
        if text == '任务号：':
            new = ('<w:t'+attrs+'>'+text+task_no+'</w:t>') if attrs else ('<w:t>'+text+task_no+'</w:t>')
            c1 = c1.replace(('<w:t'+attrs+'>'+text+'</w:t>') if attrs else ('<w:t>'+text+'</w:t>'), new, 1)
            break
    trs[0] = trs[0].replace(c0, c0, 1).replace(c1, c1, 1)
    cells1 = re.findall(r'(<w:tc[^>]*>.*?</w:tc>)', trs[1])
    if len(cells1) >= 2:
        c = cells1[1]
        if not re.findall(r'<w:t[^>]*>([^<]+)</w:t>', c):
            c = c.replace('</w:tc>', '<w:p><w:r><w:rPr><w:rFonts w:ascii="宋体" w:hAnsi="宋体" w:eastAsia="宋体"/><w:sz w:val="20"/></w:rPr><w:t>'+str(leader)+'</w:t></w:r></w:p></w:tc>', 1)
        trs[1] = trs[1].replace(cells1[1], c, 1)
    cells3 = re.findall(r'(<w:tc[^>]*>.*?</w:tc>)', trs[3])
    if len(cells3) >= 2:
        def rcb(m):
            t, a = m.group(2), m.group(1)
            if audit_type in INITIAL_TYPES: t = t.replace('□初审', '☑初审', 1)
            elif audit_type in SURVEILLANCE_TYPES: t = t.replace('□初审      □', '☐初审      ☑', 1)
            elif audit_type in RECERT_TYPES or audit_type in TRANSFER_TYPES: t = t.replace('□再认证/转移', '☑再认证/转移', 1)
            return ('<w:t'+a+'>'+t+'</w:t>') if a else ('<w:t>'+t+'</w:t>')
        cx = re.sub(r'<w:t([^>]*?)>([^<]*)</w:t>', rcb, cells3[1])
        trs[3] = trs[3].replace(cells3[1], cx, 1)
    cells4 = re.findall(r'(<w:tc[^>]*>.*?</w:tc>)', trs[4])
    if len(cells4) >= 2:
        trs[4] = trs[4].replace(cells4[1], cells4[1].replace('审核地址：', '审核地址：'+str(address), 1), 1)
    cells5 = re.findall(r'(<w:tc[^>]*>.*?</w:tc>)', trs[5])
    if len(cells5) >= 2:
        trs[5] = trs[5].replace(cells5[1], cells5[1].replace('认证范围：', '认证范围：'+str(scope), 1), 1)
    cells21 = re.findall(r'(<w:tc[^>]*>.*?</w:tc>)', trs[21])
    if len(cells21) >= 2:
        cx = cells21[1]
        paras = re.findall(r'<w:p[^>]*>.*?</w:p>', cx, re.DOTALL)
        sc = -1
        if audit_type in INITIAL_TYPES or audit_type in RECERT_TYPES: sc = 0
        elif audit_type in TRANSFER_TYPES: sc = 2
        elif audit_type in SURVEILLANCE_TYPES: sc = 4
        np = list(paras)
        for pi, p in enumerate(paras):
            if pi <= 6:
                np[pi] = p.replace('<w:checked w:val="0"/>', '<w:checked w:val="1"/>', 1) if pi == sc else p.replace('<w:checked w:val="1"/>', '<w:checked w:val="0"/>', 1)
        pp, rem, off = [], cx, 0
        for p in paras:
            idx = rem.find(p)
            if idx < 0: break
            pp.append(off + idx); off = off + idx + len(p); rem = rem[idx + len(p):]
        if len(np) > 9 and '日期：' in np[9]: np[9] = np[9].replace('日期：', '日期：'+str(date), 1)
        parts = []
        pe = 0
        for i, pos in enumerate(pp):
            parts.append(cx[pe:pos]); parts.append(np[i] if i < len(np) else paras[i]); pe = pos + len(paras[i])
        parts.append(cx[pe:])
        trs[21] = trs[21].replace(cells21[1], ''.join(parts), 1)
    new_tbl = row_parts[0]
    for tr in trs: new_tbl += tr
    last = max(i for i, p in enumerate(row_parts) if p.startswith('<w:tr'))
    new_tbl += ''.join(row_parts[last+1:])
    new_xml = xml_str[:tbl_start] + new_tbl + xml_str[tbl_end:]
    contents['word/document.xml'] = new_xml.encode('utf-8')
    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w', compression=zipfile.ZIP_DEFLATED) as zo:
        for n, d in contents.items(): zo.writestr(n, d)
    return out.getvalue()

tab1, tab2 = st.tabs(['Single Mode', 'Batch Mode'])

with tab1:
    st.header('Single Generation')
    col1, col2 = st.columns(2)
    with col1:
        form_file = st.file_uploader('FORM6101 (.docx)', type=['docx'], key='s1')
    with col2:
        tpl_file = st.file_uploader('Report Template (.docx)', type=['docx'], key='s2')
    if form_file and tpl_file:
        with st.spinner('Parsing...'):
            try:
                with zipfile.ZipFile(io.BytesIO(form_file.getvalue())) as zfin:
                    xml = zfin.read('word/document.xml').decode('utf-8')
                texts = re.findall(r'<w:t[^>]*>([^<]*)</w:t>', xml)
                lines = [t.strip() for t in texts if t.strip()]
                data = {}
                tm = re.search(r'<w:tbl>(.*?)</w:tbl>', xml, re.DOTALL)
                if tm:
                    for i, row in enumerate(re.findall(r'<w:tr[^>]*>(.*?)</w:tr>', tm.group(1), re.DOTALL)):
                        ct = [' '.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', c)) for c in re.findall(r'<w:tc[^>]*>(.*?)</w:tc>', row, re.DOTALL)]
                        if i==1 and len(ct)>1: data['taskNo']=ct[1]
                        elif i==2 and len(ct)>1: data['company']=ct[1]
                        elif i==3 and len(ct)>1: data['address']=ct[1]
                        elif i==11 and len(ct)>1:
                            s=ct[1]; data['scope']=s[5:] if s.startswith('IATF:') else s
                        elif i==13 and len(ct)>2:
                            m=re.search(r'(\d{4})[年\-](\d{1,2})[月\-](\d{1,2})',ct[2])
                            if m: data['date']=m.group(1)+'-'+m.group(2).zfill(2)+'-'+m.group(3).zfill(2)
                        elif i==14 and len(ct)>1: data['auditType']=ct[1]
                        elif i==17 and len(ct)>1: data['leader']=ct[1]
                for ln in lines:
                    if ('公司名称' in ln or '组织名称' in ln) and not data.get('company'): data['company']=re.sub(r'.*[:：]','',ln).strip()
                    if ('任务编号' in ln or '任务号' in ln) and not data.get('taskNo'): data['taskNo']=re.sub(r'.*[:：]','',ln).strip()
                    if '审核组长' in ln and not data.get('leader'): data['leader']=re.sub(r'.*[:：]','',ln).strip()
                    if '审核地址' in ln and not data.get('address'): data['address']=re.sub(r'.*[:：]','',ln).strip()
                    if ('认证范围' in ln or 'IATF:' in ln) and not data.get('scope'):
                        s=re.sub(r'.*[:：]','',ln).strip(); data['scope']=s[5:] if s.startswith('IATF:') else s
                    if ('审核日期' in ln or re.search(r'\d{4}年\d{1,2}月\d{1,2}日',ln)) and not data.get('date'):
                        m=re.search(r'(\d{4})[年\-](\d{1,2})[月\-](\d{1,2})',ln)
                        if m: data['date']=m.group(1)+'-'+m.group(2).zfill(2)+'-'+m.group(3).zfill(2)
                    if ('审核性质' in ln or '审核类型' in ln) and not data.get('auditType'): data['auditType']=re.sub(r'.*[:：]','',ln).strip()
            except Exception as e:
                st.error('Parse failed: '+str(e)); data=None
        if data:
            st.markdown('### Extracted Fields')
            for k,v in data.items(): st.text('**'+k+'**: '+str(v or '(not found)'))
            if st.button('Generate Report', type='primary'):
                try:
                    r = fill_template(tpl_file.getvalue(), data)
                    ts = datetime.now().strftime('%Y%m%d_%H%M')
                    fn = sanitize(data.get('company','Report'))+'_'+ts+'.docx'
                    st.download_button('Download', data=r, file_name=fn,
                        mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                    st.success('Generated!')
                except Exception as e: st.error('Failed: '+str(e))
    elif form_file or tpl_file: st.info('Please upload both files')

with tab2:
    st.header('Batch Generation')
    col1, col2 = st.columns(2)
    with col1:
        excel_file = st.file_uploader('Cert Excel (.xlsx)', type=['xlsx'], key='b1')
    with col2:
        tpl_file2 = st.file_uploader('Report Template (.docx)', type=['docx'], key='b2')
    if excel_file and tpl_file2:
        try:
            import openpyxl
            wb = openpyxl.load_workbook(io.BytesIO(excel_file.getvalue()), data_only=True)
            ws = wb.active
            rows_data = []
            for row in ws.iter_rows(min_row=2):
                vals = [c.value for c in row]
                if vals[3]: rows_data.append(vals)
            st.success('Read '+str(len(rows_data))+' rows')
            if st.button('Generate All', type='primary'):
                results, errors = [], []
                for ri, rv in enumerate(rows_data):
                    company,audit_team,audit_type = rv[3],rv[4],rv[5]
                    evaluate_date,audit_address,cert_scope,task_no = rv[9],rv[12],rv[13],rv[14]
                    ds = evaluate_date.strftime('%Y-%m-%d') if isinstance(evaluate_date, datetime) else str(evaluate_date)
                    leader = str(audit_team).split('+')[0].strip() if audit_team else ''
                    d = {'company':str(company),'taskNo':str(task_no) if task_no else '',
                         'leader':leader,'auditType':str(audit_type) if audit_type else '',
                         'address':str(audit_address) if audit_address else '',
                         'scope':str(cert_scope) if cert_scope else '','date':ds}
                    try:
                        rb = fill_template(tpl_file2.getvalue(), d)
                        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
                        results.append((sanitize(str(company))+'_'+ts+'.docx', rb))
                    except Exception as e: errors.append(str(company)+': '+str(e))
                if results:
                    buf = io.BytesIO()
                    with zipfile.ZipFile(buf, 'w', compression=zipfile.ZIP_DEFLATED) as zfout:
                        for fn, fd in results: zfout.writestr(fn, fd)
                    buf.seek(0)
                    st.download_button('Download All (ZIP)', data=buf,
                        file_name='reports_'+datetime.now().strftime('%Y%m%d_%H%M')+'.zip',
                        mime='application/zip')
                    st.success('Generated '+str(len(results))+' reports')
                if errors:
                    st.warning(str(len(errors))+' failed')
                    for e in errors[:5]: st.text('  - '+e)
        except ImportError: st.error('Need openpyxl')
    elif excel_file or tpl_file2: st.info('Please upload both files')
