$py = 'C:\Users\31953\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
$deps = 'C:\Users\31953\Documents\Codex\2026-08-12\wo\.streamlit-deps'
$env:PYTHONPATH = $deps
Write-Host "Starting Streamlit..."
& $py -m streamlit run app.py --server.port 8501 --server.headless true