@echo off
REM Portable clasp wrapper. Node and clasp both live under tools\ -- nothing is
REM installed system-wide and PATH is not modified outside this process.
REM
REM Usage:  clasp.cmd login          (you must run this one -- it opens a browser)
REM         clasp.cmd push
REM         clasp.cmd status
setlocal
set PROJ=%~dp0
set PROJ=%PROJ:~0,-1%
set NODE_DIR=%PROJ%\tools\node-v24.19.0-win-x64
set CLASP=%PROJ%\tools\clasp\node_modules\@google\clasp\build\src\index.js
set PATH=%NODE_DIR%;%PATH%
cd /d "%PROJ%\appsscript"
"%NODE_DIR%\node.exe" "%CLASP%" %*
