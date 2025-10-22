<#
.SYNOPSIS
    Collects full Print Server configuration and uploads results to Viyu n8n.
.DESCRIPTION
    - Enumerates all printer queues, ports, drivers, and permissions.
    - Packages results into a single JSON payload per server.
    - Sends the payload to a secure webhook for centralized ingestion.
    - Requires local admin on the print server (or remote admin rights).
.NOTES
    Author: Viyu Network Solutions
    Project: Impact Floors – Pinnacle Migration
#>

# ---------- CONFIGURATION ----------
$WebhookUri = "https://n8n.viyu.network/webhook/printserver_ingest"  # <-- replace with your endpoint
$ServerName = $env:COMPUTERNAME
$OutDir     = "C:\ProgramData\Viyu\PrintServer"
New-Item -ItemType Directory -Path $OutDir -Force | Out-Null

Import-Module PrintManagement -ErrorAction Stop

Write-Host "Collecting printer queues, ports, and drivers from $ServerName..."

# ---------- DATA COLLECTION ----------
$printers = Get-Printer | Select-Object Name,ShareName,DriverName,PortName,Location,Comment,Published
$ports    = Get-PrinterPort | Select-Object Name,PrinterHostAddress,PortNumber,SNMP,Protocol
$drivers  = Get-PrinterDriver | Select-Object Name,Manufacturer,DriverVersion,MajorVersion,InfPath
$security = foreach($p in Get-Printer){
    try{
        [pscustomobject]@{
            Printer  = $p.Name
            SDDL     = (Get-PrinterProperty -PrinterName $p.Name -PropertyName "SecurityDescriptorSDDL").Value
        }
    }catch{}
}

$result = [pscustomobject]@{
    Server     = $ServerName
    TimeStamp  = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    Printers   = $printers
    Ports      = $ports
    Drivers    = $drivers
    Security   = $security
}

# ---------- SAVE + UPLOAD ----------
$OutFile = Join-Path $OutDir "$ServerName-PrintExport.json"
$result | ConvertTo-Json -Depth 6 | Out-File $OutFile -Encoding UTF8

try {
    Invoke-RestMethod -Uri $WebhookUri -Method Post -InFile $OutFile -ContentType "application/json" -ErrorAction Stop
    Write-Host "Upload complete for $ServerName"
} catch {
    Write-Warning "Upload failed: $_"
}

Remove-Item $OutDir -Recurse -Force -ErrorAction SilentlyContinue
