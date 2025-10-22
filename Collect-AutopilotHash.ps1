<#
.SYNOPSIS
    Collects Windows Autopilot hardware hash information and securely sends it to an n8n webhook.
.DESCRIPTION
    - Downloads the official Microsoft Get-WindowsAutoPilotInfo.ps1 script from the PowerShell Gallery.
    - Executes the script to collect device hardware hashes.
    - Generates a CSV with standard Autopilot data (DeviceName, Serial, Hash, etc.).
    - POSTs the CSV content along with metadata (DeviceName, Username, Timestamp) to the configured n8n webhook.
    - Cleans up temporary files automatically.
.NOTES
    Author: Viyu Network Solutions
    For use by: Pinnacle IT team
    Context: Impact Floors M365/Intune Migration (Autopilot-based)
#>

# ---------- CONFIGURATION ----------
$WebhookUri = "https://n8n.botstuff.org/webhook/autopilot_ingest_5F4F8CF4-5BC4-463D-ACA2-8A640B0C1067"  # <-- replace with your actual n8n endpoint
$OutDir     = "C:\ProgramData\Viyu\Autopilot"
$OutFile    = Join-Path $OutDir "$($env:COMPUTERNAME)_AutopilotHash.csv"

# ---------- PREPARE ENVIRONMENT ----------
New-Item -Path $OutDir -ItemType Directory -Force | Out-Null
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force

# ---------- DOWNLOAD / VALIDATE SCRIPT ----------
try {
    Write-Host "Downloading Get-WindowsAutoPilotInfo.ps1 from PowerShell Gallery..."
    Install-Script -Name Get-WindowsAutoPilotInfo -Force -ErrorAction Stop
} catch {
    Write-Host "Failed to install script. Attempting manual download..."
    $scriptUrl = "https://raw.githubusercontent.com/Microsoft/WindowsAutopilotInfo/master/Get-WindowsAutoPilotInfo.ps1"
    $localPath = "$OutDir\Get-WindowsAutoPilotInfo.ps1"
    Invoke-WebRequest -Uri $scriptUrl -OutFile $localPath -UseBasicParsing
}

# ---------- RUN SCRIPT ----------
try {
    Write-Host "Running Autopilot hardware hash collection..."
    Get-WindowsAutoPilotInfo -OutputFile $OutFile -Append -ErrorAction Stop
} catch {
    Write-Host "Error during hash collection: $_"
}

# ---------- POST RESULTS TO N8N ----------
if (Test-Path $OutFile) {
    $FileContent = Get-Content $OutFile -Raw
    $Payload = @{
        ComputerName = $env:COMPUTERNAME
        UserName     = (Get-WmiObject Win32_ComputerSystem).UserName
        TimeStamp    = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        FileName     = (Split-Path $OutFile -Leaf)
        CSVContent   = $FileContent
    } | ConvertTo-Json -Depth 5

    try {
        Write-Host "Uploading Autopilot data to webhook..."
        Invoke-RestMethod -Uri $WebhookUri -Method Post -Body $Payload -ContentType "application/json" -ErrorAction Stop
        Write-Host "Upload successful."
    } catch {
        Write-Host "Upload failed: $_"
    }
} else {
    Write-Host "No Autopilot hash file found. Skipping upload."
}

# ---------- CLEANUP ----------
try {
    Remove-Item $OutDir -Recurse -Force
} catch {
    Write-Host "Cleanup skipped or failed: $_"
}

Write-Host "Autopilot collection complete."
