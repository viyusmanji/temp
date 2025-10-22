<#
.SYNOPSIS
  Collects full Print Server configuration and writes a single ZIP for handoff.
.DESCRIPTION
  Gathers queues, ports, drivers, security ACLs (+ optional GPO report),
  saves JSON/CSV artifacts to C:\ProgramData\Viyu\PrintServer\<Server>\,
  then compresses the folder to a timestamped ZIP.
.NOTES
  Run as local admin on each print server (or with equivalent remote rights).
#>

[CmdletBinding()] Param(
  [string]$OutRoot = "C:\ProgramData\Viyu\PrintServer",
  [switch]$IncludeGpoReport   # adds GPO XML reports if GroupPolicy module present
)

$ErrorActionPreference = 'Stop'
$server = $env:COMPUTERNAME
$ts     = (Get-Date).ToString('yyyyMMdd-HHmmss')
$work   = Join-Path $OutRoot $server
$null = New-Item -ItemType Directory -Path $work -Force

# Load modules if available
try { Import-Module PrintManagement -ErrorAction Stop } catch { throw "PrintManagement module is required." }
try { Import-Module GroupPolicy -ErrorAction SilentlyContinue } catch {}

Write-Host "[$server] Collecting printers/ports/drivers…"

# --- Collect ---
$printers = Get-Printer | Select Name,ShareName,DriverName,PortName,Location,Comment,Published,Type
$ports    = Get-PrinterPort | Select Name,PrinterHostAddress,PortNumber,SNMP,Protocol
$drivers  = Get-PrinterDriver | Select Name,Manufacturer,DriverVersion,MajorVersion,InfPath,IsXPSDriver

$security = foreach($p in Get-Printer){
  try{
    [pscustomobject]@{
      Printer = $p.Name
      SDDL    = (Get-PrinterProperty -PrinterName $p.Name -PropertyName "SecurityDescriptorSDDL").Value
    }
  } catch {
    [pscustomobject]@{ Printer=$p.Name; SDDL="<unreadable>" }
  }
}

# --- Save JSON for fidelity ---
$meta = [pscustomobject]@{
  Server    = $server
  TimeStamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
}
$meta        | ConvertTo-Json -Depth 6 | Out-File (Join-Path $work 'ServerMeta.json') -Encoding UTF8
$printers    | ConvertTo-Json -Depth 6 | Out-File (Join-Path $work 'Printers.json')   -Encoding UTF8
$ports       | ConvertTo-Json -Depth 6 | Out-File (Join-Path $work 'Ports.json')      -Encoding UTF8
$drivers     | ConvertTo-Json -Depth 6 | Out-File (Join-Path $work 'Drivers.json')    -Encoding UTF8
$security    | ConvertTo-Json -Depth 6 | Out-File (Join-Path $work 'Security.json')   -Encoding UTF8

# --- Save CSV for quick review ---
$printers | Export-Csv (Join-Path $work 'Printers.csv') -NoTypeInformation -Encoding UTF8
$ports    | Export-Csv (Join-Path $work 'Ports.csv')    -NoTypeInformation -Encoding UTF8
$drivers  | Export-Csv (Join-Path $work 'Drivers.csv')  -NoTypeInformation -Encoding UTF8
$security | Export-Csv (Join-Path $work 'Security.csv') -NoTypeInformation -Encoding UTF8

# --- Optional: GPO XML report (deployed printers, etc.) ---
if ($IncludeGpoReport -and (Get-Module -ListAvailable GroupPolicy)) {
  $gpoDir = Join-Path $work "GPOReports"
  $null = New-Item -ItemType Directory -Path $gpoDir -Force
  try {
    Get-GPO -All | ForEach-Object {
      Get-GPOReport -Guid $_.Id -ReportType Xml -Path (Join-Path $gpoDir "$($_.DisplayName).xml")
    }
  } catch {
    Write-Warning "GPO export issue: $_"
  }
}

# --- Compress to ZIP ---
$zip = Join-Path $OutRoot ("{0}-{1}-PrintExport.zip" -f $server,$ts)

# Prefer Compress-Archive; fall back to .NET if needed
try {
  if (Get-Command Compress-Archive -ErrorAction SilentlyContinue) {
    if (Test-Path $zip) { Remove-Item $zip -Force }
    Compress-Archive -Path (Join-Path $work '*') -DestinationPath $zip
  } else {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::CreateFromDirectory($work, $zip)
  }
  Write-Host "ZIP created: $zip"
} catch {
  throw "Failed to create ZIP: $_"
}

Write-Host "Done."
