@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo ============================================
echo   Geologic Age Auto-Alignment - Web Demo
echo ============================================
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
  echo Check "Add R to PATH" during setup, or add R bin folder to PATH.
  echo.
  pause
  exit /b 1
)
echo Using R: !RSCRIPT!
echo.

REM ---- 1) Install missing packages (Tsinghua mirror, skip installed) ----
echo [1/2] Checking packages: httpuv jsonlite readxl dplyr writexl tibble ...
"!RSCRIPT!" -e "options(repos=c(CRAN='https://mirrors.tuna.tsinghua.edu.cn/CRAN/')); pkgs=c('httpuv','jsonlite','readxl','dplyr','writexl','tibble'); for(p in pkgs) if(!requireNamespace(p,quietly=TRUE)){cat('Installing',p,'\n');install.packages(p)}"
if errorlevel 1 (
  echo.
  echo [ERROR] Package installation failed. Check network or mirror and retry.
  pause
  exit /b 1
)
echo.

REM ---- 2) Start service (R opens browser; close window or Ctrl+C to stop) ----
echo [2/2] Starting service; browser opens automatically. Close this window to stop.
echo.
set "APIR="
for %%F in ("%~dp0*API*.R") do set "APIR=%%F"
if not defined APIR (
  echo [ERROR] API R script not found in script folder.
  pause
  exit /b 1
)
"!RSCRIPT!" "!APIR!"

echo.
echo Service stopped.
pause
