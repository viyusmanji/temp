<#
.SYNOPSIS
  Performs file server discovery and produces a SharePoint migration planning CSV.

.DESCRIPTION
  Enumerates SMB shares, folder sizes, last modified times, and permission owners.
  Suggests SharePoint structure based on folder names and recency of data.
  Outputs a single CSV summarizing suggested SharePoint sites, libraries, and migration actions.

.PARAMETER ServerList
  Optional. Either:
    - A CSV or TXT file containing a list of servers, or
    - A single server name or IP address.
  If omitted, the script defaults to localhost.

.PARAMETER OutputCsv
  Path to output CSV file for planning results.

.EXAMPLE
  .\Discover-FileShares-Plan.ps1 -ServerList "servers.csv" -OutputCsv "C:\Reports\SPO_Plan.csv"
  .\Discover-FileShares-Plan.ps1 -ServerList "FS01" -OutputCsv "C:\Reports\SPO_Plan.csv"
  .\Discover-FileShares-Plan.ps1 -OutputCsv "C:\Reports\SPO_Plan.csv"
#>

param(
    [Parameter(Mandatory = $true)] [string] $OutputCsv,
    [Parameter(Mandatory = $false)] [string] $ServerList
)

# Determine server list intelligently
if ([string]::IsNullOrWhiteSpace($ServerList)) {
    Write-Host "No server list specified. Defaulting to localhost." -ForegroundColor Yellow
    $servers = @($env:COMPUTERNAME)
}
elseif (Test-Path $ServerList) {
    if ($ServerList.EndsWith(".csv")) {
        try {
            $servers = Import-Csv $ServerList | Select-Object -ExpandProperty ServerName
            Write-Host "Loaded $($servers.Count) servers from CSV file: $ServerList" -ForegroundColor Cyan
        }
        catch {
            Write-Warning "Failed to read CSV file. Ensure it has a column named 'ServerName'."
            exit
        }
    }
    else {
        $servers = Get-Content $ServerList | Where-Object { $_.Trim() -ne "" }
        Write-Host "Loaded $($servers.Count) servers from text file: $ServerList" -ForegroundColor Cyan
    }
}
else {
    # Treat argument as a single hostname
    Write-Host "Using single server target: $ServerList" -ForegroundColor Cyan
    $servers = @($ServerList)
}

# Container for results
$results = @()

Write-Host "Starting File Server Discovery and Planning..." -ForegroundColor Cyan

foreach ($srv in $servers) {
    Write-Host "Scanning $srv ..." -ForegroundColor Yellow
    try {
        # Establish CIM session for remote scanning
        $session = New-CimSession -ComputerName $srv -ErrorAction Stop

        # Enumerate non-system SMB shares
        $shares = Get-SmbShare -CimSession $session | Where-Object { $_.Name -notmatch "ADMIN\$|IPC\$|C\$" }

        foreach ($sh in $shares) {
            Write-Host "Processing Share: $($sh.Name)" -ForegroundColor White
            $path = $sh.Path

            if (-not (Test-Path "\\$srv\$($sh.Name)")) {
                Write-Warning "Path unavailable: \\$srv\$($sh.Name)"
                continue
            }

            # Determine NTFS owner
            try { 
                $owner = (Get-Acl $path).Owner 
            } catch { 
                $owner = "Unknown" 
            }

            # Gather file details
            $files = Get-ChildItem -Path $path -Recurse -File -ErrorAction SilentlyContinue
            $totalSize = ($files | Measure-Object -Property Length -Sum).Sum
            $fileCount = $files.Count
            $lastWrite = ($files | Sort-Object LastWriteTime -Descending | Select-Object -First 1).LastWriteTime
            $lastAccess = ($files | Sort-Object LastAccessTime -Descending | Select-Object -First 1).LastAccessTime

            # Suggest SharePoint site by name heuristic
            $folderName = Split-Path $path -Leaf
            switch -regex ($folderName) {
                "Finance|Accounting" { $site = "Finance"; break }
                "HR|People|Payroll"  { $site = "HR"; break }
                "Sales|Customers"    { $site = "Sales"; break }
                "IT|Projects"        { $site = "IT"; break }
                default               { $site = "General"; break }
            }

            $targetURL = "https://tenant.sharepoint.com/sites/$site"
            $targetLib = "Shared Documents"

            # Migration recommendation logic
            if ($lastWrite -lt (Get-Date).AddYears(-2)) {
                $move = "Archive"
            }
            elseif ($totalSize -gt 500GB) {
                $move = "Review (Large Volume)"
            }
            else {
                $move = "Migrate"
            }

            # Warning flags
            $longPaths = ($files | Where-Object { $_.FullName.Length -gt 400 }).Count
            $badNames  = ($files | Where-Object { $_.Name -match "[~#%&*{}<>?|]" }).Count
            $warns = @()
            if ($longPaths -gt 0) { $warns += "LongPaths:$longPaths" }
            if ($badNames -gt 0)  { $warns += "InvalidNames:$badNames" }

            # Build output record
            $results += [pscustomobject]@{
                Server             = $srv
                ShareName          = $sh.Name
                SourcePath         = $path
                SizeGB             = [math]::Round($totalSize / 1GB, 2)
                FileCount          = $fileCount
                LastModified       = $lastWrite
                LastAccessed       = $lastAccess
                Owner              = $owner
                TargetSiteURL      = $targetURL
                TargetLibrary      = $targetLib
                MoveRecommendation = $move
                Notes              = ($warns -join "; ")
            }
        }

        $session | Remove-CimSession
    }
    catch {
        Write-Warning "Error processing ${srv}: $_"
    }
}

# Export results
$results | Sort-Object Server, ShareName | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8
Write-Host "`nDiscovery and Planning Complete. Results saved to $OutputCsv" -ForegroundColor Green
