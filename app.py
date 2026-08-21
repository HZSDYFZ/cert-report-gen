# -*- coding: utf-8 -*-
import streamlit as st
import zipfile, io
from datetime import datetime

st.title('认证报告生成器')
st.write('测试中...')

mode = st.radio('选择模式', ['Single Report', 'Batch Generation'], horizontal=True)

if mode == 'Single Report':
    st.header('单报告生成')
    st.write('这是测试页面，请检查是否能正常显示')
    st.success('单报告模式 - 正常')
    
else:
    st.header('批量生成')
    st.write('这是测试页面，请检查是否能正常显示')
    st.success('批量生成模式 - 正常')
    
    if st.button('测试生成'):
        st.write('按钮点击成功！')
        st.success('生成完成')
