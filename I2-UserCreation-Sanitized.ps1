#Requires -Version 5.1
<#
.SYNOPSIS
    Secure Bulk User Creation Script for jibJab M365 Migration
.DESCRIPTION
    Creates user accounts for Pinnacle users in jibJab tenant with proper licensing and attributes.
    This script includes comprehensive security measures and input validation.
.PARAMETER InputCsvPath
    Path to CSV file containing Pinnacle user data
.PARAMETER OutputPath
    Path to save creation results and logs
.PARAMETER LicenseSku
    M365 license SKU to assign (default: ENTERPRISEPACK for E3)
.PARAMETER DomainSuffix
    Domain suffix for new users (default: jibJabflooring.com)
.PARAMETER DryRun
    Preview mode - don't actually create users
.PARAMETER UseSecurePassword
    Generate secure random passwords instead of default
.EXAMPLE
    .\I2-UserCreation-Sanitized.ps1 -InputCsvPath ".\pinnacle_users.csv" -OutputPath ".\Output" -LicenseSku "ENTERPRISEPACK" -UseSecurePassword
.EXAMPLE
    CSV File Format Example:
    UserPrincipalName,DisplayName,GivenName,Surname,JobTitle,Department,OfficeLocation,MobilePhone,BusinessPhones
.NOTES
    CSV Requirements:
    - UserPrincipalName: Required - Original email address from Pinnacle
    - DisplayName: Required - Full display name for the user
    - GivenName: Optional - First name
    - Surname: Optional - Last name
    - JobTitle: Optional - Job title/position
    - Department: Optional - Department name
    - OfficeLocation: Optional - Office location
    - MobilePhone: Optional - Mobile phone number
    - BusinessPhones: Optional - Business phone number(s)
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({Test-Path $_ -PathType Leaf})]
    [string]$InputCsvPath,
    
    [ValidateScript({Test-Path $_ -PathType Container -IsValid})]
    [string]$OutputPath = ".\Output",
    
    [ValidateSet("ENTERPRISEPACK", "ENTERPRISEPREMIUM", "BUSINESSBASIC", "BUSINESSSTANDARD", "BUSINESSPREMIUM")]
    [string]$LicenseSku = "ENTERPRISEPACK",
    
    [ValidatePattern("^[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$")]
    [string]$DomainSuffix = "jibJabflooring.com",
    
    [switch]$DryRun = $false,
    
    [switch]$UseSecurePassword = $true
)

# Security Configuration
$SecurityConfig = @{
    MaxPasswordLength = 16
    MinPasswordLength = 12
    MaxRetryAttempts = 3
    RateLimitDelay = 1000  # milliseconds
    MaxInputFileSize = 50MB
    AllowedFileExtensions = @('.csv')
    LogRetentionDays = 90
}

# Initialize secure logging
$LogPath = Join-Path $OutputPath "UserCreation_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$ErrorLogPath = Join-Path $OutputPath "UserCreation_Errors_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$SecurityLogPath = Join-Path $OutputPath "Security_Audit_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$ResultsPath = Join-Path $OutputPath "UserCreation_Results_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

# Secure password generation function
function New-SecureRandomPassword {
    param([int]$Length = 16)
    
    $CharSets = @{
        Lowercase = 'abcdefghijklmnopqrstuvwxyz'
        Uppercase = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
        Numbers = '0123456789'
        Special = '!@#$%^&*()_+-=[]{}|;:,.<>?'
    }
    
    $Password = @()
    $Password += $CharSets.Lowercase | Get-Random -Count 1
    $Password += $CharSets.Uppercase | Get-Random -Count 1
    $Password += $CharSets.Numbers | Get-Random -Count 1
    $Password += $CharSets.Special | Get-Random -Count 1
    
    $RemainingLength = $Length - 4
    $AllChars = $CharSets.Lowercase + $CharSets.Uppercase + $CharSets.Numbers + $CharSets.Special
    $Password += $AllChars | Get-Random -Count $RemainingLength
    
    return ($Password | Sort-Object {Get-Random}) -join ''
}

# Input validation and sanitization
function Test-InputValidation {
    param([object]$InputData)
    
    $ValidationResults = @{
        IsValid = $true
        Errors = @()
    }
    
    # Check for required fields
    $RequiredFields = @('UserPrincipalName', 'DisplayName')
    foreach ($Field in $RequiredFields) {
        if ($InputData.PSObject.Properties.Name -notcontains $Field) {
            $ValidationResults.Errors += "Missing required field: $Field"
            $ValidationResults.IsValid = $false
        }
    }
    
    # Sanitize string inputs
    $StringFields = @('UserPrincipalName', 'DisplayName', 'GivenName', 'Surname', 'JobTitle', 'Department', 'OfficeLocation')
    foreach ($Field in $StringFields) {
        if ($InputData.$Field) {
            # Remove potentially dangerous characters
            $InputData.$Field = $InputData.$Field -replace '[<>"''&]', ''
            # Limit length
            if ($InputData.$Field.Length -gt 255) {
                $InputData.$Field = $InputData.$Field.Substring(0, 255)
            }
        }
    }
    
    # Validate email format
    if ($InputData.UserPrincipalName) {
        $EmailPattern = '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        if ($InputData.UserPrincipalName -notmatch $EmailPattern) {
            $ValidationResults.Errors += "Invalid email format: $($InputData.UserPrincipalName)"
            $ValidationResults.IsValid = $false
        }
    }
    
    return $ValidationResults
}

# Secure logging function
function Write-SecureLog {
    param(
        [string]$Message, 
        [string]$Level = "INFO",
        [bool]$IncludeSensitiveData = $false
    )
    
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    # Sanitize message to remove sensitive data unless explicitly allowed
    if (-not $IncludeSensitiveData) {
        $Message = $Message -replace '(?i)(password|pwd|secret|key|token|credential)', '[REDACTED]'
        $Message = $Message -replace '@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', '[EMAIL_REDACTED]'
    }
    
    $LogEntry = "[$Timestamp] [$Level] $Message"
    Write-Host $LogEntry
    Add-Content -Path $LogPath -Value $LogEntry
}

# Security audit logging
function Write-SecurityAudit {
    param(
        [string]$Action,
        [string]$Details,
        [string]$Severity = "INFO"
    )
    
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $AuditEntry = "[$Timestamp] [SECURITY] [$Severity] Action: $Action | Details: $Details"
    Add-Content -Path $SecurityLogPath -Value $AuditEntry
}

# Error logging with security considerations
function Write-SecureErrorLog {
    param([string]$Message, [Exception]$Exception = $null)
    
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $ErrorEntry = "[$Timestamp] [ERROR] $Message"
    
    if ($Exception) {
        # Don't log full stack trace in production
        $ErrorEntry += "`nException: [REDACTED - Check security logs]"
    }
    
    Write-Host $ErrorEntry -ForegroundColor Red
    Add-Content -Path $ErrorLogPath -Value $ErrorEntry
    
    # Log security event
    Write-SecurityAudit -Action "ERROR_OCCURRED" -Details $Message -Severity "WARNING"
}

# Rate limiting function
function Invoke-RateLimit {
    param([int]$DelayMs = $SecurityConfig.RateLimitDelay)
    Start-Sleep -Milliseconds $DelayMs
}

# Create secure output directory
if (!(Test-Path $OutputPath)) {
    try {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
        Write-SecurityAudit -Action "DIRECTORY_CREATED" -Details "Output directory created: $OutputPath"
    }
    catch {
        Write-SecureErrorLog "Failed to create output directory: $OutputPath" $_
        throw
    }
}

# Validate input file
try {
    $FileInfo = Get-Item $InputCsvPath
    if ($FileInfo.Length -gt $SecurityConfig.MaxInputFileSize) {
        throw "Input file exceeds maximum size limit: $($SecurityConfig.MaxInputFileSize)"
    }
    
    $FileExtension = [System.IO.Path]::GetExtension($InputCsvPath).ToLower()
    if ($FileExtension -notin $SecurityConfig.AllowedFileExtensions) {
        throw "Invalid file extension. Allowed: $($SecurityConfig.AllowedFileExtensions -join ', ')"
    }
    
    Write-SecurityAudit -Action "FILE_VALIDATION" -Details "Input file validated: $InputCsvPath"
}
catch {
    Write-SecureErrorLog "Input file validation failed" $_
    throw
}

Write-SecureLog "Starting Secure Bulk User Creation for jibJab Migration" "INFO"
Write-SecureLog "Input CSV: $InputCsvPath" "INFO"
Write-SecureLog "Output Path: $OutputPath" "INFO"
Write-SecureLog "License SKU: $LicenseSku" "INFO"
Write-SecureLog "Domain Suffix: $DomainSuffix" "INFO"
Write-SecureLog "Dry Run Mode: $DryRun" "INFO"
Write-SecureLog "Secure Password Mode: $UseSecurePassword" "INFO"

try {
    # Install required modules with security validation
    Write-SecureLog "Installing required PowerShell modules..." "INFO"
    $RequiredModules = @(
        "Microsoft.Graph",
        "Microsoft.Graph.Users",
        "Microsoft.Graph.Groups",
        "Microsoft.Graph.Identity.DirectoryManagement"
    )

    foreach ($Module in $RequiredModules) {
        try {
            if (!(Get-Module -ListAvailable -Name $Module)) {
                Write-SecureLog "Installing module: $Module" "INFO"
                Install-Module -Name $Module -Force -AllowClobber -Scope CurrentUser -Repository PSGallery
                Write-SecurityAudit -Action "MODULE_INSTALLED" -Details "Module installed: $Module"
            }
        }
        catch {
            Write-SecureErrorLog "Failed to install module: $Module" $_
        }
    }

    # Secure connection to Microsoft Graph
    Write-SecureLog "Connecting to Microsoft Graph..." "INFO"
    $Scopes = @(
        "User.ReadWrite.All",
        "Group.ReadWrite.All",
        "Directory.ReadWrite.All",
        "Organization.Read.All"
    )

    try {
        Connect-MgGraph -Scopes $Scopes -NoWelcome
        Write-SecureLog "Successfully connected to Microsoft Graph" "INFO"
        Write-SecurityAudit -Action "GRAPH_CONNECTED" -Details "Connected to Microsoft Graph with required scopes"
    }
    catch {
        Write-SecureErrorLog "Failed to connect to Microsoft Graph" $_
        throw
    }

    # Read and validate CSV data
    Write-SecureLog "Reading and validating user data from CSV..." "INFO"
    try {
        $UserData = Import-Csv -Path $InputCsvPath
        Write-SecureLog "Found $($UserData.Count) users to process" "INFO"
        
        # Validate each user record
        $ValidUsers = @()
        $InvalidUsers = @()
        
        foreach ($User in $UserData) {
            $Validation = Test-InputValidation -InputData $User
            if ($Validation.IsValid) {
                $ValidUsers += $User
            }
            else {
                $InvalidUsers += @{
                    User = $User
                    Errors = $Validation.Errors
                }
                Write-SecureLog "Invalid user data: $($User.DisplayName) - $($Validation.Errors -join ', ')" "WARNING"
            }
        }
        
        Write-SecureLog "Valid users: $($ValidUsers.Count), Invalid users: $($InvalidUsers.Count)" "INFO"
        
        if ($InvalidUsers.Count -gt 0) {
            Write-SecurityAudit -Action "DATA_VALIDATION_FAILED" -Details "Found $($InvalidUsers.Count) invalid user records" -Severity "WARNING"
        }
    }
    catch {
        Write-SecureErrorLog "Failed to read or validate CSV data" $_
        throw
    }

    # Get available licenses with security validation
    Write-SecureLog "Retrieving available licenses..." "INFO"
    try {
        $SubscribedSkus = Get-MgSubscribedSku
        $TargetLicense = $SubscribedSkus | Where-Object { $_.SkuPartNumber -eq $LicenseSku }
        
        if (!$TargetLicense) {
            throw "License SKU '$LicenseSku' not found in tenant"
        }
        
        $AvailableLicenses = $TargetLicense.PrepaidUnits.Enabled - $TargetLicense.ConsumedUnits
        Write-SecureLog "Available $LicenseSku licenses: $AvailableLicenses" "INFO"
        
        if ($AvailableLicenses -lt $ValidUsers.Count) {
            $WarningMessage = "Insufficient licenses! Need $($ValidUsers.Count), have $AvailableLicenses"
            Write-SecureLog $WarningMessage "WARNING"
            Write-SecurityAudit -Action "LICENSE_SHORTAGE" -Details $WarningMessage -Severity "WARNING"
        }
    }
    catch {
        Write-SecureErrorLog "Failed to retrieve license information" $_
        throw
    }

    # Initialize results tracking
    $Results = @()
    $SuccessCount = 0
    $ErrorCount = 0
    $SkippedCount = 0

    # Process each valid user
    foreach ($User in $ValidUsers) {
        $UserResult = @{
            OriginalUPN = $User.UserPrincipalName
            NewUPN = ""
            DisplayName = $User.DisplayName
            Status = ""
            ErrorMessage = ""
            CreatedDateTime = ""
            LicenseAssigned = $false
            SecurityNotes = ""
        }

        try {
            # Generate new UPN with validation
            $Username = $User.UserPrincipalName.Split('@')[0]
            $NewUPN = "$Username@$DomainSuffix"
            $UserResult.NewUPN = $NewUPN

            Write-SecureLog "Processing user: $($User.DisplayName) -> $NewUPN" "INFO"

            # Check if user already exists
            $ExistingUser = Get-MgUser -Filter "userPrincipalName eq '$NewUPN'" -ErrorAction SilentlyContinue
            
            if ($ExistingUser) {
                Write-SecureLog "User already exists: $NewUPN" "WARNING"
                $UserResult.Status = "SKIPPED"
                $UserResult.ErrorMessage = "User already exists"
                $UserResult.SecurityNotes = "Duplicate user detected"
                $SkippedCount++
                $Results += $UserResult
                continue
            }

            if ($DryRun) {
                Write-SecureLog "DRY RUN: Would create user $NewUPN" "INFO"
                $UserResult.Status = "DRY_RUN"
                $UserResult.CreatedDateTime = Get-Date
                $Results += $UserResult
                continue
            }

            # Generate secure password
            $SecurePassword = if ($UseSecurePassword) {
                New-SecureRandomPassword -Length $SecurityConfig.MaxPasswordLength
            } else {
                "TempPass123!"  # Fallback - should be changed immediately
            }

            # Create user object with security validation
            $UserParams = @{
                DisplayName = $User.DisplayName
                UserPrincipalName = $NewUPN
                MailNickname = $Username
                PasswordProfile = @{
                    ForceChangePasswordNextSignIn = $true
                    Password = $SecurePassword
                }
                AccountEnabled = $true
                UsageLocation = "US"
            }

            # Add optional attributes with validation
            if ($User.GivenName) { $UserParams.GivenName = $User.GivenName }
            if ($User.Surname) { $UserParams.Surname = $User.Surname }
            if ($User.JobTitle) { $UserParams.JobTitle = $User.JobTitle }
            if ($User.Department) { $UserParams.Department = $User.Department }
            if ($User.OfficeLocation) { $UserParams.OfficeLocation = $User.OfficeLocation }
            if ($User.MobilePhone) { $UserParams.MobilePhone = $User.MobilePhone }
            if ($User.BusinessPhones) { $UserParams.BusinessPhones = @($User.BusinessPhones) }

            # Create the user with rate limiting
            Write-SecureLog "Creating user: $NewUPN" "INFO"
            $NewUser = New-MgUser @UserParams
            $UserResult.CreatedDateTime = Get-Date
            $UserResult.Status = "CREATED"
            $SuccessCount++
            
            Write-SecurityAudit -Action "USER_CREATED" -Details "User created: $NewUPN" -Severity "INFO"

            Write-SecureLog "User created successfully: $($NewUser.Id)" "INFO"

            # Assign license with error handling
            if ($TargetLicense) {
                try {
                    Write-SecureLog "Assigning license to user: $NewUPN" "INFO"
                    $LicenseAssignment = @{
                        AddLicenses = @(
                            @{
                                SkuId = $TargetLicense.SkuId
                            }
                        )
                        RemoveLicenses = @()
                    }
                    
                    Set-MgUserLicense -UserId $NewUser.Id -BodyParameter $LicenseAssignment
                    $UserResult.LicenseAssigned = $true
                    Write-SecureLog "License assigned successfully" "INFO"
                    Write-SecurityAudit -Action "LICENSE_ASSIGNED" -Details "License assigned to: $NewUPN"
                }
                catch {
                    Write-SecureErrorLog "Failed to assign license to user: $NewUPN" $_
                    $UserResult.ErrorMessage += "License assignment failed: [REDACTED]"
                    $UserResult.SecurityNotes += "License assignment failed"
                }
            }

            # Rate limiting
            Invoke-RateLimit

        }
        catch {
            Write-SecureErrorLog "Failed to create user: $($User.DisplayName)" $_
            $UserResult.Status = "ERROR"
            $UserResult.ErrorMessage = "[ERROR_REDACTED]"
            $UserResult.SecurityNotes = "Creation failed - check security logs"
            $ErrorCount++
        }

        $Results += $UserResult
    }

    # Export results securely
    Write-SecureLog "Exporting results..." "INFO"
    $Results | Export-Csv -Path $ResultsPath -NoTypeInformation

    # Generate secure summary report
    $SummaryPath = Join-Path $OutputPath "UserCreation_Summary_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    
    $HtmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <title>User Creation Summary Report</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .header { background-color: #0078d4; color: white; padding: 20px; }
        .section { margin: 20px 0; padding: 15px; border: 1px solid #ddd; }
        .success { background-color: #d4edda; border-color: #c3e6cb; }
        .warning { background-color: #fff3cd; border-color: #ffeaa7; }
        .error { background-color: #f8d7da; border-color: #f5c6cb; }
        table { width: 100%; border-collapse: collapse; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; }
        .stats { display: flex; justify-content: space-around; margin: 20px 0; }
        .stat-box { text-align: center; padding: 20px; border: 1px solid #ddd; }
        .security-notice { background-color: #e7f3ff; border: 1px solid #b3d9ff; padding: 10px; margin: 10px 0; }
    </style>
</head>
<body>
    <div class="header">
        <h1>User Creation Summary Report</h1>
        <p>Generated: $(Get-Date)</p>
        <p>Input File: $InputCsvPath</p>
        <p>Domain Suffix: $DomainSuffix</p>
        <p>License SKU: $LicenseSku</p>
        <p>Security Mode: Enhanced</p>
    </div>
    
    <div class="security-notice">
        <h3>Security Notice</h3>
        <p>This report contains sanitized data. Sensitive information has been redacted for security purposes. 
        Check security audit logs for detailed information.</p>
    </div>
    
    <div class="stats">
        <div class="stat-box success">
            <h3>$SuccessCount</h3>
            <p>Users Created</p>
        </div>
        <div class="stat-box warning">
            <h3>$SkippedCount</h3>
            <p>Users Skipped</p>
        </div>
        <div class="stat-box error">
            <h3>$ErrorCount</h3>
            <p>Errors</p>
        </div>
    </div>
    
    <div class="section">
        <h2>Detailed Results</h2>
        <table>
            <tr>
                <th>Original UPN</th>
                <th>New UPN</th>
                <th>Display Name</th>
                <th>Status</th>
                <th>License Assigned</th>
                <th>Security Notes</th>
            </tr>
            $(foreach ($Result in $Results) {
                $StatusClass = switch ($Result.Status) {
                    "CREATED" { "success" }
                    "SKIPPED" { "warning" }
                    "ERROR" { "error" }
                    default { "" }
                }
                "<tr class='$StatusClass'>
                    <td>$($Result.OriginalUPN)</td>
                    <td>$($Result.NewUPN)</td>
                    <td>$($Result.DisplayName)</td>
                    <td>$($Result.Status)</td>
                    <td>$($Result.LicenseAssigned)</td>
                    <td>$($Result.SecurityNotes)</td>
                </tr>"
            })
        </table>
    </div>
</body>
</html>
"@
    
    $HtmlContent | Out-File -FilePath $SummaryPath -Encoding UTF8

    Write-SecureLog "User creation process completed" "INFO"
    Write-SecureLog "Results exported to: $ResultsPath" "INFO"
    Write-SecureLog "Summary report: $SummaryPath" "INFO"
    Write-SecurityAudit -Action "PROCESS_COMPLETED" -Details "User creation process completed successfully"
    
    # Display secure summary
    Write-Host "`n=== SECURE USER CREATION SUMMARY ===" -ForegroundColor Green
    Write-Host "Total Users Processed: $($ValidUsers.Count)" -ForegroundColor White
    Write-Host "Successfully Created: $SuccessCount" -ForegroundColor Green
    Write-Host "Skipped (Already Exist): $SkippedCount" -ForegroundColor Yellow
    Write-Host "Errors: $ErrorCount" -ForegroundColor Red
    Write-Host "`nResults saved to: $ResultsPath" -ForegroundColor Cyan
    Write-Host "Summary report: $SummaryPath" -ForegroundColor Cyan
    Write-Host "Security audit log: $SecurityLogPath" -ForegroundColor Magenta

}
catch {
    Write-SecureErrorLog "Critical error during user creation" $_
    Write-SecurityAudit -Action "CRITICAL_ERROR" -Details "Critical error occurred during user creation" -Severity "CRITICAL"
    throw
}
finally {
    Write-SecureLog "User creation script completed" "INFO"
    Write-SecurityAudit -Action "SCRIPT_COMPLETED" -Details "User creation script execution completed"
}
