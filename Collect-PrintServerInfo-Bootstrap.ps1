# Bootstrap: download & run the print export, producing a ZIP for handoff
$ScriptUrl   = "https://raw.githubusercontent.com/viyusmanji/temp/refs/heads/main/Collect-PrintServerInfo.ps1"
$LocalScript = "C:\Temp\Collect-PrintServerInfo.ps1"
$CopyToShare = ""   # or "" to skip copying

New-Item -ItemType Directory -Path (Split-Path $LocalScript) -Force | Out-Null
Set-ExecutionPolicy Bypass -Scope Process -Force

Invoke-WebRequest -Uri $ScriptUrl -OutFile $LocalScript -UseBasicParsing

# Run with optional GPO report
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File $LocalScript -IncludeGpoReport

# Find newest ZIP and copy (optional)
$zip = Get-ChildItem "C:\ProgramData\Viyu\PrintServer" -Filter '*-PrintExport.zip' -Recurse |
       Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($zip -and $CopyToShare) {
  if (-not (Test-Path $CopyToShare)) { New-Item -ItemType Directory -Path $CopyToShare -Force | Out-Null }
  Copy-Item $zip.FullName -Destination $CopyToShare -Force
  Write-Host "Copied ZIP to $CopyToShare"
} else {
  Write-Host "ZIP ready for transfer: $($zip.FullName)"
}
