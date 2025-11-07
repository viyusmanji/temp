<#
.SYNOPSIS
    Complete Shared Mailbox Export - All Mailboxes, Members, Full Access, and Send As Permissions

.DESCRIPTION
    Exports all shared mailboxes from the source tenant with complete permission details:
    - Shared mailbox properties (name, email, display name)
    - Members with Full Access permissions
    - Members with Send As permissions
    
    This script is designed to be run by the source tenant administrator (ILG/Pinnacle)
    and exported to CSV for migration planning.
    
    Cross-platform compatible: Works on macOS, Linux, and Windows using PowerShell Core (pwsh).

.PARAMETER OutputPath
    Directory for exports (default: ./exports/shared-mailboxes or $HOME/exports/shared-mailboxes)

.EXAMPLE
    pwsh .\Export-SharedMailboxesComplete.ps1
    pwsh .\Export-SharedMailboxesComplete.ps1 -OutputPath "./my-exports"

.NOTES
    Requires:
    - PowerShell Core (pwsh) 7.0 or later
    - ExchangeOnlineManagement PowerShell module
    - Exchange Administrator or Global Administrator role
    - Connection to Exchange Online (will prompt if not connected)
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputPath
)

#region Configuration
$ErrorActionPreference = "Stop"
$ProgressPreference = "Continue"

# Set default output path based on platform
if (-not $OutputPath) {
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    if ($scriptDir) {
        $OutputPath = Join-Path $scriptDir "exports" "shared-mailboxes"
    } else {
        $OutputPath = if ($env:HOME) {
            Join-Path $env:HOME "exports" "shared-mailboxes"
        } else {
            Join-Path $env:USERPROFILE "exports" "shared-mailboxes"
        }
    }
}
#endregion

#region Utility Functions
function Write-Status {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [ValidateSet("Cyan", "Green", "Yellow", "Red", "Gray", "DarkGray", "White")]
        [string]$Color = "Cyan"
    )
    Write-Host $Message -ForegroundColor $Color
}

function Write-ErrorWithContext {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [string]$Context = ""
    )
    $fullMessage = if ($Context) { "$Message - Context: $Context" } else { $Message }
    Write-Warning $fullMessage
}

function Test-PowerShellCore {
    <#
    .SYNOPSIS
        Verifies PowerShell Core (pwsh) is being used
    #>
    if ($PSVersionTable.PSVersion.Major -lt 7) {
        Write-Error "This script requires PowerShell Core 7.0 or later. Please use 'pwsh' instead of 'powershell'."
        Write-Host "Install PowerShell Core: https://aka.ms/powershell" -ForegroundColor Yellow
        throw "PowerShell Core required"
    }
    return $true
}
#endregion

#region Module Management
function Import-ExchangeOnlineModule {
    <#
    .SYNOPSIS
        Imports or installs the ExchangeOnlineManagement module
    #>
    [CmdletBinding()]
    param()

    Write-Status "Checking ExchangeOnlineManagement module..." "Cyan"
    
    try {
        $module = Get-Module -ListAvailable -Name ExchangeOnlineManagement
        if (-not $module) {
            Write-Status "Module not found. Installing ExchangeOnlineManagement..." "Yellow"
            Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            Write-Status "Module installed successfully" "Green"
        } else {
            Write-Status "Module found (Version: $($module.Version))" "Green"
        }
        
        Import-Module -Name ExchangeOnlineManagement -Force -ErrorAction Stop
        Write-Status "ExchangeOnlineManagement module loaded" "Green"
        return $true
    } catch {
        Write-Error "Failed to load ExchangeOnlineManagement module: $($_.Exception.Message)"
        Write-Host "Please install with: Install-Module ExchangeOnlineManagement -Scope CurrentUser" -ForegroundColor Red
        throw
    }
}

function Connect-ExchangeOnlineSession {
    <#
    .SYNOPSIS
        Connects to Exchange Online if not already connected
    #>
    [CmdletBinding()]
    param()

    Write-Status "Checking Exchange Online connection..." "Cyan"
    
    try {
        $exoSession = Get-ConnectionInformation -ErrorAction SilentlyContinue
        if (-not $exoSession) {
            Write-Status "Not connected to Exchange Online. Connecting..." "Yellow"
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
            Write-Status "Connected to Exchange Online" "Green"
        } else {
            Write-Status "Already connected to Exchange Online" "Green"
        }
        return $true
    } catch {
        Write-Error "Failed to connect to Exchange Online: $($_.Exception.Message)"
        Write-Host "Please ensure you have Exchange Administrator or Global Administrator permissions." -ForegroundColor Red
        throw
    }
}
#endregion

#region Data Retrieval Functions
function Get-AllSharedMailboxes {
    <#
    .SYNOPSIS
        Retrieves all shared mailboxes from the tenant
    #>
    [CmdletBinding()]
    param()

    Write-Status "Retrieving all shared mailboxes..." "Cyan"
    
    try {
        $allSharedMailboxes = Get-ExoMailbox -ResultSize Unlimited -RecipientTypeDetails SharedMailbox -ErrorAction Stop
        $mailboxCount = ($allSharedMailboxes | Measure-Object).Count
        Write-Status "Found $mailboxCount shared mailbox(es)" "Green"
        return $allSharedMailboxes
    } catch {
        Write-Error "Failed to retrieve shared mailboxes: $($_.Exception.Message)"
        throw
    }
}

function Get-SharedMailboxDetails {
    <#
    .SYNOPSIS
        Extracts details from shared mailbox objects
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Mailboxes
    )

    $details = @()
    foreach ($mailbox in $Mailboxes) {
        $details += [PSCustomObject]@{
            SharedMailboxIdentity = $mailbox.Identity
            SharedMailboxEmail = $mailbox.PrimarySmtpAddress
            SharedMailboxDisplayName = $mailbox.DisplayName
            Alias = $mailbox.Alias
            EmailAddresses = ($mailbox.EmailAddresses -join ';')
            WhenCreated = $mailbox.WhenCreated
            WhenChanged = $mailbox.WhenChanged
            RecipientTypeDetails = $mailbox.RecipientTypeDetails
            ArchiveStatus = $mailbox.ArchiveStatus
            RetentionHoldEnabled = $mailbox.RetentionHoldEnabled
        }
    }
    return $details
}

function Get-FullAccessPermissions {
    <#
    .SYNOPSIS
        Retrieves Full Access permissions for shared mailboxes
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Mailboxes
    )

    $permissions = @()
    $totalCount = $Mailboxes.Count
    $processedCount = 0

    foreach ($mailbox in $Mailboxes) {
        $processedCount++
        $mailboxIdentity = $mailbox.Identity
        $mailboxEmail = $mailbox.PrimarySmtpAddress
        $mailboxDisplayName = $mailbox.DisplayName
        
        Write-Status "[$processedCount/$totalCount] Processing Full Access: $mailboxDisplayName" "DarkGray"
        
        try {
            $fullAccessPerms = Get-ExoMailboxPermission -Identity $mailboxIdentity -ErrorAction SilentlyContinue | 
                Where-Object { 
                    -not $_.IsInherited -and 
                    $_.User -notlike "NT AUTHORITY\*" -and 
                    $_.User -notlike "S-1-*" -and
                    $_.User -ne "Anonymous" -and
                    $null -ne $_.User -and
                    ($_.AccessRights -contains "FullAccess" -or $_.AccessRights -contains "ReadPermission")
                }
            
            foreach ($perm in $fullAccessPerms) {
                $permissions += [PSCustomObject]@{
                    SharedMailboxEmail = $mailboxEmail
                    SharedMailboxDisplayName = $mailboxDisplayName
                    User = $perm.User
                    AccessRights = ($perm.AccessRights -join ',')
                    Deny = $perm.Deny
                    IsInherited = $perm.IsInherited
                }
            }
            
            if ($fullAccessPerms) {
                Write-Status "  -> Found $($fullAccessPerms.Count) Full Access permission(s)" "DarkGray"
            }
        } catch {
            Write-ErrorWithContext "Failed to retrieve Full Access permissions for $mailboxEmail" $_.Exception.Message
        }
    }
    
    return $permissions
}

function Get-SendAsPermissions {
    <#
    .SYNOPSIS
        Retrieves Send As permissions for shared mailboxes
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Mailboxes
    )

    $permissions = @()
    $totalCount = $Mailboxes.Count
    $processedCount = 0

    foreach ($mailbox in $Mailboxes) {
        $processedCount++
        $mailboxIdentity = $mailbox.Identity
        $mailboxEmail = $mailbox.PrimarySmtpAddress
        $mailboxDisplayName = $mailbox.DisplayName
        
        Write-Status "[$processedCount/$totalCount] Processing Send As: $mailboxDisplayName" "DarkGray"
        
        try {
            $sendAsPerms = Get-ExoRecipientPermission -Identity $mailboxIdentity -ErrorAction SilentlyContinue | 
                Where-Object { 
                    $_.Trustee -notlike "NT AUTHORITY\*" -and 
                    $_.Trustee -notlike "S-1-*" -and
                    $_.Trustee -ne "Anonymous" -and
                    $null -ne $_.Trustee -and
                    ($_.AccessRights -eq "SendAs" -or $_.AccessRights -contains "SendAs")
                }
            
            foreach ($perm in $sendAsPerms) {
                $permissions += [PSCustomObject]@{
                    SharedMailboxEmail = $mailboxEmail
                    SharedMailboxDisplayName = $mailboxDisplayName
                    Trustee = $perm.Trustee
                    AccessRights = $perm.AccessRights
                    InheritanceType = $perm.InheritanceType
                }
            }
            
            if ($sendAsPerms) {
                Write-Status "  -> Found $($sendAsPerms.Count) Send As permission(s)" "DarkGray"
            }
        } catch {
            Write-ErrorWithContext "Failed to retrieve Send As permissions for $mailboxEmail" $_.Exception.Message
        }
    }
    
    return $permissions
}
#endregion

#region Export Functions
function Export-SharedMailboxData {
    <#
    .SYNOPSIS
        Exports shared mailbox data to CSV files
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$OutputPath,
        [Parameter(Mandatory = $true)]
        [object[]]$MailboxDetails,
        [Parameter(Mandatory = $true)]
        [object[]]$FullAccessPermissions,
        [Parameter(Mandatory = $true)]
        [object[]]$SendAsPermissions
    )

    # Ensure output directory exists
    if (-not (Test-Path -Path $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
        Write-Status "Created output directory: $OutputPath" "Green"
    } else {
        Write-Status "Using output directory: $OutputPath" "Green"
    }

    # Export Shared Mailbox Details
    Write-Status "`nExporting Shared Mailbox Details..." "Cyan"
    $mailboxDetailsFile = Join-Path $OutputPath "SharedMailboxes_Details.csv"
    try {
        $MailboxDetails | Export-Csv -Path $mailboxDetailsFile -NoTypeInformation -Encoding UTF8 -Force
        Write-Status "Exported $($MailboxDetails.Count) shared mailbox(es) to: $mailboxDetailsFile" "Green"
    } catch {
        Write-Error "Failed to export mailbox details: $($_.Exception.Message)"
        throw
    }

    # Export Full Access Permissions
    Write-Status "Exporting Full Access Permissions..." "Cyan"
    $fullAccessFile = Join-Path $OutputPath "SharedMailboxes_FullAccess.csv"
    try {
        if ($FullAccessPermissions.Count -gt 0) {
            $FullAccessPermissions | Export-Csv -Path $fullAccessFile -NoTypeInformation -Encoding UTF8 -Force
            Write-Status "Exported $($FullAccessPermissions.Count) Full Access permission(s) to: $fullAccessFile" "Green"
        } else {
            Write-Status "No Full Access permissions found. Creating empty file." "Yellow"
            "SharedMailboxEmail,SharedMailboxDisplayName,User,AccessRights,Deny,IsInherited" | Out-File -FilePath $fullAccessFile -Encoding UTF8 -Force
        }
    } catch {
        Write-Error "Failed to export Full Access permissions: $($_.Exception.Message)"
        throw
    }

    # Export Send As Permissions
    Write-Status "Exporting Send As Permissions..." "Cyan"
    $sendAsFile = Join-Path $OutputPath "SharedMailboxes_SendAs.csv"
    try {
        if ($SendAsPermissions.Count -gt 0) {
            $SendAsPermissions | Export-Csv -Path $sendAsFile -NoTypeInformation -Encoding UTF8 -Force
            Write-Status "Exported $($SendAsPermissions.Count) Send As permission(s) to: $sendAsFile" "Green"
        } else {
            Write-Status "No Send As permissions found. Creating empty file." "Yellow"
            "SharedMailboxEmail,SharedMailboxDisplayName,Trustee,AccessRights,InheritanceType" | Out-File -FilePath $sendAsFile -Encoding UTF8 -Force
        }
    } catch {
        Write-Error "Failed to export Send As permissions: $($_.Exception.Message)"
        throw
    }

    return @{
        DetailsFile = $mailboxDetailsFile
        FullAccessFile = $fullAccessFile
        SendAsFile = $sendAsFile
    }
}

function Export-SummaryReport {
    <#
    .SYNOPSIS
        Generates a summary report of the export
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$OutputPath,
        [Parameter(Mandatory = $true)]
        [int]$MailboxCount,
        [Parameter(Mandatory = $true)]
        [int]$FullAccessCount,
        [Parameter(Mandatory = $true)]
        [int]$SendAsCount
    )

    Write-Status "`nGenerating summary report..." "Cyan"
    $summaryFile = Join-Path $OutputPath "SharedMailboxes_ExportSummary.txt"
    
    try {
        $tenantName = (Get-OrganizationConfig | Select-Object -ExpandProperty Name -ErrorAction SilentlyContinue)
    } catch {
        $tenantName = "Unable to retrieve"
    }
    
    $summary = @"
Shared Mailbox Export Summary
==============================
Export Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
PowerShell Version: $($PSVersionTable.PSVersion)
Platform: $($PSVersionTable.Platform)
Tenant: $tenantName

Total Shared Mailboxes: $MailboxCount
Total Full Access Permissions: $FullAccessCount
Total Send As Permissions: $SendAsCount

Exported Files:
- Shared Mailbox Details: SharedMailboxes_Details.csv
- Full Access Permissions: SharedMailboxes_FullAccess.csv
- Send As Permissions: SharedMailboxes_SendAs.csv

Output Location: $OutputPath
"@

    $summary | Out-File -FilePath $summaryFile -Encoding UTF8 -Force
    Write-Status "Summary report saved to: $summaryFile" "Green"
    
    return $summaryFile
}
#endregion

#region Main Execution
function Start-SharedMailboxExport {
    <#
    .SYNOPSIS
        Main orchestration function for the export process
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    Write-Status "=== Shared Mailbox Complete Export ===" "Cyan"
    Write-Status "Starting export at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" "Gray"
    Write-Status "PowerShell Version: $($PSVersionTable.PSVersion)" "Gray"
    Write-Status "Platform: $($PSVersionTable.Platform)" "Gray"

    # Verify PowerShell Core
    Test-PowerShellCore | Out-Null

    # Import and connect
    Import-ExchangeOnlineModule | Out-Null
    Connect-ExchangeOnlineSession | Out-Null

    # Retrieve data
    $allSharedMailboxes = Get-AllSharedMailboxes
    
    if (($allSharedMailboxes | Measure-Object).Count -eq 0) {
        Write-Status "No shared mailboxes found in the tenant." "Yellow"
        Write-Status "Export complete. No data to export." "Green"
        return
    }

    # Process mailboxes
    Write-Status "`nProcessing shared mailboxes..." "Cyan"
    $mailboxDetails = Get-SharedMailboxDetails -Mailboxes $allSharedMailboxes
    $fullAccessPermissions = Get-FullAccessPermissions -Mailboxes $allSharedMailboxes
    $sendAsPermissions = Get-SendAsPermissions -Mailboxes $allSharedMailboxes

    # Export data
    Export-SharedMailboxData -OutputPath $OutputPath `
        -MailboxDetails $mailboxDetails `
        -FullAccessPermissions $fullAccessPermissions `
        -SendAsPermissions $sendAsPermissions | Out-Null

    # Generate summary
    Export-SummaryReport -OutputPath $OutputPath `
        -MailboxCount $mailboxDetails.Count `
        -FullAccessCount $fullAccessPermissions.Count `
        -SendAsCount $sendAsPermissions.Count | Out-Null

    # Display final summary
    Write-Status "`n=== Export Complete ===" "Cyan"
    Write-Status "Shared Mailboxes: $($mailboxDetails.Count)" "Green"
    Write-Status "Full Access Permissions: $($fullAccessPermissions.Count)" "Green"
    Write-Status "Send As Permissions: $($sendAsPermissions.Count)" "Green"
    Write-Status "`nAll files exported to: $OutputPath" "Green"
    Write-Status "Export completed at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" "Gray"
}
#endregion

# Execute main function
try {
    Start-SharedMailboxExport -OutputPath $OutputPath
} catch {
    Write-Error "Export failed: $($_.Exception.Message)"
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
    exit 1
}

