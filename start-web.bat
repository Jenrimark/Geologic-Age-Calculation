@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo ============================================================
echo   Geologic Age Auto-Alignment
echo   Web demo + HTTP/JSON API + Excel batch upload
echo ============================================================
echo.

REM ---- Locate Rscript: PATH first, then common install dirs ----
set "RSCRIPT="
where Rscript >nul 2>nul && set "RSCRIPT=Rscript"
if not defined RSCRIPT (
  for /d %%D in ("C:\Program Files\R\R-*") do if exist "%%D\bin\Rscript.exe" set "RSCRIPT=%%D\bin\Rscript.exe"
)
if not defined RSCRIPT (
  for /d %%D in ("%LOCALAPPDATA%\Programs\R\R-*") do if exist "%%D\bin\Rscript.exe" set "RSCRIPT=%%D\bin\Rscript.exe"
)
if not defined RSCRIPT (
  echo [ERROR] R not found. Install R from https://cran.r-project.org/
  echo         Tick "Add R to PATH" during setup, or add the R bin folder to PATH.
  echo.
  pause
  exit /b 1
)
echo Using R: !RSCRIPT!
echo.

REM ---- 1) Install missing packages (Tsinghua mirror, skip installed) ----
REM openxlsx is REQUIRED: the Excel upload feature uses it to preserve
REM the original workbook's fills / column widths when writing results back.
REM NOTE: do NOT put '!' in the R one-liner below - delayed expansion is
REM enabled and would strip it, corrupting the code. setdiff() avoids '!'.
echo [1/3] Checking packages: httpuv jsonlite readxl openxlsx writexl dplyr tibble
"!RSCRIPT!" -e "options(repos=c(CRAN='https://mirrors.tuna.tsinghua.edu.cn/CRAN/')); pkgs=c('httpuv','jsonlite','readxl','openxlsx','writexl','dplyr','tibble'); miss=setdiff(pkgs, rownames(installed.packages())); if(length(miss)>0){cat('Installing:', miss, fill=TRUE); install.packages(miss)} else {cat('All packages already installed.')}"
if errorlevel 1 (
  echo.
  echo [ERROR] Package installation failed. Check the network or mirror, then retry.
  echo         Note: if every package is already installed, the service starts
  echo         fine without network access.
  pause
  exit /b 1
)
echo.

REM ---- 2) Locate the service script ----
REM Detected by content, not by filename: the service script is the only .R
REM in this folder that calls startServer. This keeps working if it is renamed.
echo [2/3] Locating service script ...
set "APIR="
for %%F in ("%~dp0*.R") do (
  findstr /m /c:"startServer" "%%F" >nul 2>nul && set "APIR=%%~F"
)
if not defined APIR (
  echo [ERROR] No R service script found in this folder.
  echo         Expected an .R file that calls startServer.
  pause
  exit /b 1
)
echo       !APIR!
echo.

REM ---- 3) Start service ----
REM No --xlsx needed: the script auto-discovers docs\ in this folder.
REM Port 8787 is the default; if it is busy the service shifts to the next
REM free port automatically and prints the address it actually bound to.
echo [3/3] Starting service on port 8787 (auto-shifts if the port is busy).
echo       The browser opens automatically - watch this window for the URL.
echo       KEEP THIS WINDOW OPEN. Stop with Ctrl+C, or just close the window.
echo.
"!RSCRIPT!" "!APIR!" --port=8787

echo.
echo Service stopped.
pause
