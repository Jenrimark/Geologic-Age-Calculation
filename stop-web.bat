@echo off
REM ============================================================
REM  stop-web.bat - leftover process cleaner for this project
REM
REM  What it does:
REM    1. Finds Rscript.exe processes running the geologic-age
REM       web service. Matched two ways:
REM         a) command line contains the project script-name
REM            prefix (built at runtime from Unicode code points,
REM            so this file stays pure ASCII);
REM         b) any Rscript.exe still holding a TCP listener in
REM            the service port range 8787-8806 (the service
REM            auto-shifts into this range when ports are busy).
REM    2. Finds cmd.exe windows parked at the "Press any key"
REM       prompt of start-web.bat after the service already died.
REM    3. Lists everything found, asks for confirmation, then
REM       terminates the service processes first and the parked
REM       consoles last, and verifies the ports are released.
REM
REM  Safety:
REM    - R sessions NOT belonging to this project (your own R
REM      scripts, RStudio, plain R consoles) are never touched:
REM      name match requires the project prefix, and the port
REM      sweep only kills processes owned by Rscript.exe whose
REM      listener sits in 8787-8806.
REM    - The script that launched the service through
REM      start-web.bat has its own console closed only when the
REM      console command line contains start-web.bat.
REM    - Re-running this file when nothing is left over is a
REM      harmless no-op ("Nothing to clean").
REM
REM  Usage:
REM    stop-web.bat        interactive: preview list + confirm
REM    stop-web.bat -y     kill without asking (for scripting)
REM ============================================================

setlocal
set "PS1=%TEMP%\gts_stop_web_worker.ps1"

> "%PS1%" echo $ErrorActionPreference = 'SilentlyContinue'
>> "%PS1%" echo $autoYes = ($args[0] -eq '-y')
>> "%PS1%" echo # Project script-name prefix built from Unicode code points so
>> "%PS1%" echo # this worker file stays pure ASCII: Nian-Dai-Biao-Zhuan-Huan.
>> "%PS1%" echo $pat = [string]::Join('', [char]0x5E74, [char]0x4EE3, [char]0x8868, [char]0x8F6C, [char]0x6362)
>> "%PS1%" echo $batPat = 'start-web.bat'
>> "%PS1%" echo $ports = 8787..8806
>> "%PS1%" echo Write-Host '=== GTS web-service leftover cleaner ==='
>> "%PS1%" echo Write-Host ('script-name prefix: ' + $pat)
>> "%PS1%" echo Write-Host 'scanning: service Rscript processes, port range 8787-8806, parked start-web consoles'
>> "%PS1%" echo Write-Host ''
>> "%PS1%" echo $targets = @()
>> "%PS1%" echo $seen = @()
>> "%PS1%" echo $procs = Get-CimInstance Win32_Process -Filter "Name='Rscript.exe' OR Name='cmd.exe'"
>> "%PS1%" echo foreach ($p in $procs) {
>> "%PS1%" echo   $cl = '' + $p.CommandLine
>> "%PS1%" echo   if ($p.Name -eq 'Rscript.exe' -and $cl.IndexOf($pat) -ge 0) {
>> "%PS1%" echo     $targets += ,@('service', $p.ProcessId, $cl)
>> "%PS1%" echo     $seen += $p.ProcessId
>> "%PS1%" echo   } elseif ($p.Name -eq 'cmd.exe' -and $cl.IndexOf($batPat, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
>> "%PS1%" echo     $targets += ,@('console', $p.ProcessId, $cl)
>> "%PS1%" echo     $seen += $p.ProcessId
>> "%PS1%" echo   }
>> "%PS1%" echo }
>> "%PS1%" echo foreach ($port in $ports) {
>> "%PS1%" echo   $conns = Get-NetTCPConnection -LocalPort $port -State Listen
>> "%PS1%" echo   foreach ($c in $conns) {
>> "%PS1%" echo     $oid = $c.OwningProcess
>> "%PS1%" echo     if ($seen.IndexOf($oid) -ge 0) { continue }
>> "%PS1%" echo     $pr = Get-Process -Id $oid
>> "%PS1%" echo     if ($pr -and $pr.Name -eq 'Rscript') {
>> "%PS1%" echo       $targets += ,@('port', $oid, ('Rscript.exe holding port ' + $port))
>> "%PS1%" echo       $seen += $oid
>> "%PS1%" echo     }
>> "%PS1%" echo   }
>> "%PS1%" echo }
>> "%PS1%" echo if ($targets.Count -eq 0) {
>> "%PS1%" echo   Write-Host 'Nothing to clean - no leftover project processes found.'
>> "%PS1%" echo   exit 0
>> "%PS1%" echo }
>> "%PS1%" echo Write-Host ('Found ' + $targets.Count + ' leftover process(es):')
>> "%PS1%" echo foreach ($t in $targets) {
>> "%PS1%" echo   Write-Host ('  [' + $t[0] + '] PID ' + $t[1])
>> "%PS1%" echo   Write-Host ('      ' + $t[2])
>> "%PS1%" echo }
>> "%PS1%" echo Write-Host ''
>> "%PS1%" echo if (-not $autoYes) {
>> "%PS1%" echo   $ans = Read-Host 'Kill all of the above? y/N'
>> "%PS1%" echo   if ($ans -ne 'y' -and $ans -ne 'Y') {
>> "%PS1%" echo     Write-Host 'Aborted. Nothing was killed.'
>> "%PS1%" echo     exit 0
>> "%PS1%" echo   }
>> "%PS1%" echo }
>> "%PS1%" echo Write-Host 'Killing service processes ...'
>> "%PS1%" echo foreach ($t in $targets) {
>> "%PS1%" echo   if ($t[0] -ne 'console') {
>> "%PS1%" echo     Stop-Process -Id $t[1] -Force
>> "%PS1%" echo     Write-Host ('  killed PID ' + $t[1] + ' [' + $t[0] + ']')
>> "%PS1%" echo   }
>> "%PS1%" echo }
>> "%PS1%" echo Write-Host 'Closing parked start-web consoles ...'
>> "%PS1%" echo foreach ($t in $targets) {
>> "%PS1%" echo   if ($t[0] -eq 'console') {
>> "%PS1%" echo     Stop-Process -Id $t[1] -Force
>> "%PS1%" echo     Write-Host ('  killed PID ' + $t[1] + ' [console]')
>> "%PS1%" echo   }
>> "%PS1%" echo }
>> "%PS1%" echo Start-Sleep -Milliseconds 800
>> "%PS1%" echo Write-Host ''
>> "%PS1%" echo Write-Host '=== verification ==='
>> "%PS1%" echo $remain = 0
>> "%PS1%" echo $procs2 = Get-CimInstance Win32_Process -Filter "Name='Rscript.exe' OR Name='cmd.exe'"
>> "%PS1%" echo foreach ($p in $procs2) {
>> "%PS1%" echo   $cl = '' + $p.CommandLine
>> "%PS1%" echo   if ($p.Name -eq 'Rscript.exe' -and $cl.IndexOf($pat) -ge 0) { $remain++ }
>> "%PS1%" echo   if ($p.Name -eq 'cmd.exe' -and $cl.IndexOf($batPat, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { $remain++ }
>> "%PS1%" echo }
>> "%PS1%" echo foreach ($port in $ports) {
>> "%PS1%" echo   $conns = Get-NetTCPConnection -LocalPort $port -State Listen
>> "%PS1%" echo   foreach ($c in $conns) {
>> "%PS1%" echo     $pr = Get-Process -Id $c.OwningProcess
>> "%PS1%" echo     if ($pr -and $pr.Name -eq 'Rscript') { $remain++ }
>> "%PS1%" echo   }
>> "%PS1%" echo }
>> "%PS1%" echo if ($remain -eq 0) {
>> "%PS1%" echo   Write-Host 'CLEAN. All leftover project processes gone; ports 8787-8806 free.'
>> "%PS1%" echo } else {
>> "%PS1%" echo   Write-Host ('Some processes survived (' + $remain + '). Re-run stop-web.bat -y or close the windows manually.')
>> "%PS1%" echo }
>> "%PS1%" echo exit 0

powershell -NoProfile -ExecutionPolicy Bypass -File "%PS1%" %*
set "RC=%ERRORLEVEL%"
del "%PS1%" >nul 2>nul
echo.
if /I "%~1" NEQ "-y" pause
endlocal
exit /b %RC%
