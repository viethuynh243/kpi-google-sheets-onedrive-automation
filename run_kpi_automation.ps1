$ErrorActionPreference = "Stop"

$WorkDir = "C:\Users\admin\OneDrive\Documents\Excel"
$LogDir  = Join-Path $WorkDir "logs"
$LogFile = Join-Path $LogDir ("kpi_automation_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))

New-Item -ItemType Directory -Force -Path $LogDir | Out-Null
Set-Location $WorkDir

$env:PYTHONIOENCODING = "utf-8"

"Started at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" | Tee-Object -FilePath $LogFile
python "$WorkDir\automate_kpi.py" --download 2>&1 | Tee-Object -FilePath $LogFile -Append
$ExitCode = $LASTEXITCODE
"Finished at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') with exit code $ExitCode" | Tee-Object -FilePath $LogFile -Append

exit $ExitCode
