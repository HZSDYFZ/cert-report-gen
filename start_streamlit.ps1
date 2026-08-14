$py = "C:\Users\31953\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
$env:PYTHONPATH = "C:\Users\31953\Documents\Codex\2026-08-12\wo\.streamlit-deps"
Write-Host "Starting Streamlit app..." -ForegroundColor Green
& $py -m streamlit run app.py --server.port 8501 --server.headless true