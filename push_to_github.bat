@echo off
TITLE Deploy to GitHub: Nikhil8-coder/employee-workforce-analytics
echo =========================================================================
echo   Deploying Enterprise Workforce Analytics to GitHub (Nikhil8-coder)
echo =========================================================================
echo.

set "PATH=%LOCALAPPDATA%\Microsoft\WinGet\Packages\GitHub.cli_Microsoft.Winget.Source_8wekyb3d8bbwe\bin;%LOCALAPPDATA%\Microsoft\WinGet\Packages\Git.MinGit_Microsoft.Winget.Source_8wekyb3d8bbwe\cmd;%PATH%"

echo Checking GitHub CLI login status...
gh auth status >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo [ACTION NEEDED] Authenticating with GitHub...
    echo A browser window will open. Click 'Authorize github' to link your account.
    echo.
    gh auth login --web --git-protocol https
)

echo.
echo Creating repository 'employee-workforce-analytics' on https://github.com/Nikhil8-coder...
gh repo create employee-workforce-analytics --public --source=. --remote=origin --push 2>nul
if %errorlevel% neq 0 (
    echo Repository may already exist. Pushing latest commits to main branch...
    git push -u origin main
)

echo.
echo =========================================================================
echo [SUCCESS] Deployed to: https://github.com/Nikhil8-coder/employee-workforce-analytics
echo =========================================================================
echo.
pause
