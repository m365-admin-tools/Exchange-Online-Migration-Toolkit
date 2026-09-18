<#
.SYNOPSIS
    Exchange Online mailbox migration helper - interactive menu for common migration tasks.

.DESCRIPTION
    Provides an interactive menu for completing migrated mailboxes in AD, syncing
    mailboxes, updating UPNs, checking licenses, and monitoring migration status.

    All org-specific values are defined in the CONFIGURATION section below.
    Update that block once per engagement and the rest of the script adapts.

.NOTES
    - Bulk operations read from $AccountListFile (one sAMAccountName per line).
    - All output is logged to $LogFile.
    - Requires: ExchangeOnlineManagement, ActiveDirectory, Microsoft.Graph.Users modules.
    - Run from an elevated PowerShell 7 session on a domain-joined machine (or with
      AD RSAT tools and line-of-sight to a DC).
#>

# ==================================================================================================
# CONFIGURATION -- UPDATE THESE VALUES FOR EACH MIGRATION - AUTHOR: CHARLES ARCONI updated 5-13-2026
# ==================================================================================================

$Config = @{
    # Admin UPN used to connect to Exchange Online / Graph
    AdminUPN            = "admin@contoso.com"

    # Primary vanity domain (user-facing)
    PrimaryDomain       = "contoso.com"

    # Microsoft 365 routing domain (<tenant>.mail.onmicrosoft.com)
    RoutingDomain       = "contoso.mail.onmicrosoft.com"

    # AD security group that grants the target license (e.g., "O365_E3_License")
    LicenseGroup        = "O365_E3_License"

    # Path to the text file containing sAMAccountNames (one per line)
    AccountListFile     = "C:\Temp\Account_List.txt"

    # Log / output file for all operations
    LogFile             = "C:\Temp\Migration_Output.txt"

    # Default completion time for move requests (24-hr format, local time)
    CompletionTime      = "05:00 AM"

    # Bad-item limit for Set-MoveRequest -CompleteAfter
    BadItemLimit        = 1000
}

# ============================================================================
# MODULE CHECKS
# ============================================================================

function Assert-ModuleAvailable {
    param ([string]$ModuleName)
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Warning "Module '$ModuleName' is not installed. Install it with:"
        Write-Warning "  Install-Module -Name $ModuleName -Scope CurrentUser"
        return $false
    }
    Import-Module $ModuleName -ErrorAction SilentlyContinue
    return $true
}

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

function Write-Log {
    param (
        [string]$Message,
        [string]$LogPath = $Config.LogFile
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry     = "[$timestamp] $Message"
    Write-Host $entry
    $entry | Out-File -FilePath $LogPath -Append -Encoding utf8
}

function Get-AccountList {
    $path = $Config.AccountListFile
    if (-not (Test-Path $path)) {
        Write-Warning "Account list not found at '$path'. Create it first."
        return $null
    }
    $accounts = Get-Content $path | Where-Object { $_.Trim() -ne '' }
    if ($accounts.Count -eq 0) {
        Write-Warning "Account list at '$path' is empty."
        return $null
    }
    Write-Host "Loaded $($accounts.Count) account(s) from '$path'."
    return $accounts
}

function Get-CompletionDateUTC {
    $today    = Get-Date -Format "MM/dd/yyyy"
    $combined = "$today $($Config.CompletionTime)"
    return (Get-Date $combined).ToUniversalTime()
}

function Build-UserStrings {
    param ([string]$SamAccount)
    return @{
        UPN           = "$SamAccount@$($Config.PrimaryDomain)"
        TargetAddress = "$SamAccount@$($Config.RoutingDomain)"
        ProxyAddress  = "smtp:$SamAccount@$($Config.RoutingDomain)"
    }
}

# ============================================================================
# MENU ACTIONS
# ============================================================================

function Connect-MigrationServices {
    Write-Host "`nConnecting to Exchange Online..." -ForegroundColor Cyan
    Connect-ExchangeOnline -UserPrincipalName $Config.AdminUPN -ShowProgress $true

    Write-Host "Connecting to Security and Compliance (IPPS)..." -ForegroundColor Cyan
    Connect-IPPSSession -UserPrincipalName $Config.AdminUPN

    Write-Host "Connecting to Microsoft Graph (replaces MSOnline)..." -ForegroundColor Cyan
    Connect-MgGraph -Scopes "User.Read.All" -NoWelcome
    Write-Log "Connected to Exchange Online, IPPS, and Microsoft Graph."
}

function Complete-MailboxAD {
    param ([string]$SamAccount)
    $strings = Build-UserStrings -SamAccount $SamAccount
    try {
        Set-ADUser -Identity $SamAccount -UserPrincipalName $strings.UPN
        Set-ADUser -Identity $SamAccount -Add @{ proxyAddresses = $strings.ProxyAddress }
        Set-ADUser -Identity $SamAccount -Replace @{ targetAddress = $strings.TargetAddress }
        Write-Log "AD complete: $SamAccount | UPN=$($strings.UPN) | target=$($strings.TargetAddress)"
    }
    catch {
        Write-Log "ERROR completing AD for $SamAccount -- $($_.Exception.Message)"
    }
}

function Complete-MailboxAD-Bulk {
    $accounts = Get-AccountList
    if (-not $accounts) { return }
    foreach ($acct in $accounts) {
        Complete-MailboxAD -SamAccount $acct
    }
}

function Complete-MailboxAD-Single {
    $user = Read-Host -Prompt "Enter sAMAccountName"
    try {
        Add-ADGroupMember -Identity $Config.LicenseGroup -Members $user -ErrorAction Stop
        Write-Log "Added $user to group '$($Config.LicenseGroup)'."
    }
    catch {
        Write-Log "ERROR adding $user to license group -- $($_.Exception.Message)"
    }
    Complete-MailboxAD -SamAccount $user
}

function Complete-SyncedMailbox {
    param ([string]$Identity)
    $completeAfter = Get-CompletionDateUTC
    try {
        Set-MoveRequest -Identity $Identity `
            -CompleteAfter $completeAfter `
            -BadItemLimit $Config.BadItemLimit `
            -AcceptLargeDataLoss `
            -Confirm:$false
        Write-Log "MoveRequest scheduled: $Identity | CompleteAfter=$completeAfter UTC"
    }
    catch {
        Write-Log "ERROR scheduling MoveRequest for $Identity -- $($_.Exception.Message)"
    }
}

function Complete-SyncedMailbox-Bulk {
    $accounts = Get-AccountList
    if (-not $accounts) { return }
    foreach ($acct in $accounts) {
        Complete-SyncedMailbox -Identity $acct
    }
}

function Complete-SyncedMailbox-Single {
    $user = Read-Host -Prompt "Enter sAMAccountName"
    Complete-SyncedMailbox -Identity $user
}

function Update-UPN-Bulk {
    $accounts = Get-AccountList
    if (-not $accounts) { return }
    foreach ($acct in $accounts) {
        $upn = "$acct@$($Config.PrimaryDomain)"
        try {
            Set-ADUser -Identity $acct -UserPrincipalName $upn
            Write-Log "UPN updated: $acct -> $upn"
        }
        catch {
            Write-Log "ERROR updating UPN for $acct -- $($_.Exception.Message)"
        }
    }
}

function Update-UPN-Single {
    $user = Read-Host -Prompt "Enter sAMAccountName"
    $upn  = "$user@$($Config.PrimaryDomain)"
    try {
        Set-ADUser -Identity $user -UserPrincipalName $upn
        Write-Log "UPN updated: $user -> $upn"
    }
    catch {
        Write-Log "ERROR updating UPN for $user -- $($_.Exception.Message)"
    }
}

function Check-License-Single {
    $user = Read-Host -Prompt "Enter sAMAccountName"
    $upn  = "$user@$($Config.PrimaryDomain)"
    try {
        $mgUser = Get-MgUser -UserId $upn -Property DisplayName, AssignedLicenses
        Write-Host "`nDisplayName : $($mgUser.DisplayName)"
        Write-Host "Licenses    : $($mgUser.AssignedLicenses.SkuId -join ', ')"
        Write-Log "License check: $upn -- $($mgUser.AssignedLicenses.Count) license(s)"
    }
    catch {
        Write-Log "ERROR checking license for $upn -- $($_.Exception.Message)"
    }
}

function Check-License-Bulk {
    $accounts = Get-AccountList
    if (-not $accounts) { return }
    foreach ($acct in $accounts) {
        $upn = "$acct@$($Config.PrimaryDomain)"
        try {
            $mgUser = Get-MgUser -UserId $upn -Property DisplayName, AssignedLicenses
            Write-Host "$($mgUser.DisplayName) -- $($mgUser.AssignedLicenses.Count) license(s)"
        }
        catch {
            Write-Log "ERROR checking license for $upn -- $($_.Exception.Message)"
        }
    }
}

function Check-MigrationStatus-Bulk {
    $accounts = Get-AccountList
    if (-not $accounts) { return }
    foreach ($acct in $accounts) {
        try {
            Get-MoveRequest $acct | Format-List Name, Status
        }
        catch {
            Write-Log "ERROR getting MoveRequest for $acct -- $($_.Exception.Message)"
        }
    }
}

function Check-MigrationStatus-Single {
    $user = Read-Host -Prompt "Enter sAMAccountName"
    try {
        Get-MoveRequest -Identity $user | Format-List Name, Status
    }
    catch {
        Write-Log "ERROR getting MoveRequest for $user -- $($_.Exception.Message)"
    }
}

# ============================================================================
# INTERACTIVE MENU
# ============================================================================

function Show-Menu {
    Clear-Host
    Write-Host ""
    Write-Host "  =================================================================" -ForegroundColor Cyan
    Write-Host "              EXCHANGE ONLINE MIGRATION TOOLKIT                     " -ForegroundColor Cyan
    Write-Host "  =================================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "   1  Connect to Exchange Online / Graph"
    Write-Host ""
    Write-Host "   2  Bulk   - Complete migrated mailboxes in AD"
    Write-Host "   3  Single - Complete migrated mailbox in AD"
    Write-Host ""
    Write-Host "   4  Bulk   - Schedule synced mailbox completion"
    Write-Host "   5  Single - Schedule synced mailbox completion"
    Write-Host ""
    Write-Host "   6  Bulk   - Update UPN to primary domain"
    Write-Host "   7  Single - Update UPN to primary domain"
    Write-Host ""
    Write-Host "   8  Single - Check license assignment"
    Write-Host "   9  Bulk   - Check license assignments"
    Write-Host ""
    Write-Host "   0  Bulk   - Check migration status"
    Write-Host "   Z  Single - Check migration status"
    Write-Host ""
    Write-Host "   Q  Quit"
    Write-Host ""
    Write-Host "  -----------------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host "   Tenant  : $($Config.PrimaryDomain)" -ForegroundColor DarkGray
    Write-Host "   Routing : $($Config.RoutingDomain)" -ForegroundColor DarkGray
    Write-Host "   Accounts: $($Config.AccountListFile)" -ForegroundColor DarkGray
    Write-Host "   Log     : $($Config.LogFile)" -ForegroundColor DarkGray
    Write-Host "  -----------------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host ""
}

# ============================================================================
# MAIN LOOP
# ============================================================================

# Pre-flight module checks (non-blocking warnings)
$requiredModules = @("ExchangeOnlineManagement", "ActiveDirectory", "Microsoft.Graph.Users")
foreach ($mod in $requiredModules) {
    if (-not (Assert-ModuleAvailable $mod)) {
        Write-Host "  ^ Needed for full functionality. Some menu options may fail." -ForegroundColor Yellow
        Write-Host ""
    }
}

do {
    Show-Menu
    $selection = Read-Host "  Select an option"

    switch ($selection) {
        '1' { Connect-MigrationServices }
        '2' { Complete-MailboxAD-Bulk }
        '3' { Complete-MailboxAD-Single }
        '4' { Complete-SyncedMailbox-Bulk }
        '5' { Complete-SyncedMailbox-Single }
        '6' { Update-UPN-Bulk }
        '7' { Update-UPN-Single }
        '8' { Check-License-Single }
        '9' { Check-License-Bulk }
        '0' { Check-MigrationStatus-Bulk }
        'Z' { Check-MigrationStatus-Single }
        'Q' {
            Write-Host "`nExiting..." -ForegroundColor Yellow
            break
        }
        default {
            Write-Host "`nInvalid selection. Please try again." -ForegroundColor Red
        }
    }

    if ($selection -ne 'Q') { pause }
}
until ($selection -eq 'Q')
