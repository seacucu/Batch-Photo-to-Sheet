@echo off
REM ---------------------------------------------------------------------------
REM clasp wrapper. Run from the project root:
REM
REM   clasp.cmd login     <- run this once yourself; it opens a browser
REM   clasp.cmd push      <- upload appsscript\ to Apps Script
REM   clasp.cmd status
REM
REM Resolution order:
REM   1. portable Node + clasp under tools\  (for machines where you cannot
REM      install software; tools\ is gitignored)
REM   2. clasp on PATH
REM   3. npx @google/clasp
REM
REM A fresh clone from GitHub has no tools\, so it takes path 2 or 3 and you
REM need Node + clasp installed first. See README.md ("deployment").
REM
REM ASCII only, on purpose: cmd.exe parses .cmd files using the system ANSI
REM codepage, so non-ASCII bytes here break the script on other machines.
REM ---------------------------------------------------------------------------
setlocal
set PROJ=%~dp0
set PROJ=%PROJ:~0,-1%
set NODE_DIR=%PROJ%\tools\node-v24.19.0-win-x64
set NODE_EXE=%NODE_DIR%\node.exe
set CLASP_JS=%PROJ%\tools\clasp\node_modules\@google\clasp\build\src\index.js

if not exist "%PROJ%\appsscript\.clasp.json" (
  echo [clasp.cmd] appsscript\.clasp.json not found.
  echo [clasp.cmd] Run this from the project root.
  exit /b 1
)

cd /d "%PROJ%\appsscript"

if exist "%NODE_EXE%" set PATH=%NODE_DIR%;%PATH%
if exist "%NODE_EXE%" if exist "%CLASP_JS%" goto portable

where clasp >nul 2>nul
if not errorlevel 1 goto systemClasp

where npx >nul 2>nul
if not errorlevel 1 goto viaNpx

echo [clasp.cmd] clasp not found. Set it up with:
echo.
echo   1. Install Node.js        https://nodejs.org/
echo   2. npm install -g @google/clasp
echo   3. clasp.cmd login
echo.
exit /b 1

:portable
"%NODE_EXE%" "%CLASP_JS%" %*
exit /b %ERRORLEVEL%

:systemClasp
call clasp %*
exit /b %ERRORLEVEL%

:viaNpx
call npx --yes @google/clasp %*
exit /b %ERRORLEVEL%
