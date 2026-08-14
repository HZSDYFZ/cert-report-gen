# 认证报告生成器

上传 FORM6101 文件和报告模板，自动生成认证报告。

## 本地运行

```powershell
.\start_streamlit.ps1
```

然后访问 http://localhost:8501

## 部署到 Streamlit Cloud

1. 在 GitHub 创建新仓库 `cert-report-gen`
2. 推送代码：
   ```bash
   git remote add origin https://github.com/YOUR_USERNAME/cert-report-gen.git
   git push -u origin main
   ```
3. 在 [streamlit.io](https://streamlit.io) 连接 GitHub 仓库
4. 部署即可

## 依赖

- streamlit
- python-docx