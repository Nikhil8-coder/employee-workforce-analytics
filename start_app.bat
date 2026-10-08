@echo off
TITLE Enterprise Workforce Analytics - Live Server
echo =========================================================================
echo   Launching Enterprise Workforce Intelligence & Compensation Analytics
echo =========================================================================
echo.
echo Local URL   : http://localhost:8501
echo Cloud Tunnel: Starting background services...
echo.

start "Streamlit Web App" python -m streamlit run app.py --server.port 8501 --server.headless true
timeout /t 3 >nul

if exist cloudflared.exe (
    echo Starting Public Internet Tunnel...
    start "Cloudflare Public Tunnel" cloudflared.exe tunnel --url http://127.0.0.1:8501
)

echo.
echo Both Local and Cloud services are online.
pause
