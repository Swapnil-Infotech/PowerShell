# ============================================================
#  Active Directory - New Forest Deployment Script
#  Author  : Swapnil Lohot
#  Version : 2.0
#  Tested  : Windows Server 2019/2022
# ============================================================

#STEP 1: Check if Running as Administrator ---

Write-Host "`n[STEP 1] Checking Administrator Privileges..." -ForegroundColor Cyan

$CurrentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
$Principal   = New-Object Security.Principal.WindowsPrincipal($CurrentUser)

if (-not $Principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "[ERROR] Please run this script as Administrator!" -ForegroundColor Red
    Exit 1
} else {
    Write-Host "[OK] Running as Administrator." -ForegroundColor Green
}

#StepEND

#STEP 2: Collect User Inputs ---

Write-Host "`n[STEP 2] Please provide the following details to configure your Domain...`n" -ForegroundColor Cyan

# Domain FQDN
do {
    $DomainFQDN = Read-Host "  Enter Domain FQDN (e.g., corp.example.com)"
} while ([string]::IsNullOrWhiteSpace($DomainFQDN))

# NetBIOS Name
do {
    $DomainNetBIOS = Read-Host "  Enter NetBIOS Name (e.g., CORP)"
} while ([string]::IsNullOrWhiteSpace($DomainNetBIOS))

# DSRM Password (Encrypted / SecureString)
do {
    $SecureDSRMPassword  = Read-Host "  Enter DSRM Password" -AsSecureString
    $ConfirmDSRMPassword = Read-Host "  Confirm DSRM Password" -AsSecureString

    # Convert both to plain text just for comparison, then discard
    $Pass1 = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
                [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureDSRMPassword))
    $Pass2 = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
                [Runtime.InteropServices.Marshal]::SecureStringToBSTR($ConfirmDSRMPassword))

    if ($Pass1 -ne $Pass2) {
        Write-Host "  [WARNING] Passwords do not match. Please try again.`n" -ForegroundColor Yellow
    }
} while ($Pass1 -ne $Pass2)

Write-Host "  [OK] Password confirmed.`n" -ForegroundColor Green

# Clear plain text password variables immediately from memory
$Pass1 = $null
$Pass2 = $null

#StepEND

#STEP 3: Confirm Details Before Proceeding ---

Write-Host "`n[STEP 3] Please confirm the details below before proceeding:" -ForegroundColor Cyan
Write-Host "  ----------------------------------------" -ForegroundColor Gray
Write-Host "  Domain FQDN   : $DomainFQDN"  -ForegroundColor Yellow
Write-Host "  NetBIOS Name  : $DomainNetBIOS" -ForegroundColor Yellow
Write-Host "  DSRM Password : ************" -ForegroundColor Yellow
Write-Host "  Forest Mode   : WinThreshold (Server 2016/2019/2022)" -ForegroundColor Yellow
Write-Host "  Domain Mode   : WinThreshold (Server 2016/2019/2022)" -ForegroundColor Yellow
Write-Host "  ----------------------------------------`n" -ForegroundColor Gray

$Confirm = Read-Host "  Do you want to proceed? (Y/N)"

if ($Confirm -notmatch "^[Yy]$") {
    Write-Host "`n[INFO] Operation cancelled by user." -ForegroundColor Red
    Exit 0
}

#StepEND

#STEP 4: Install ADDS Role ---

Write-Host "`n[STEP 4] Installing Active Directory Domain Services Role..." -ForegroundColor Cyan

try {
    Install-WindowsFeature -Name AD-Domain-Services -IncludeManagementTools -ErrorAction Stop
    Write-Host "[OK] ADDS Role installed successfully." -ForegroundColor Green
} catch {
    Write-Host "[ERROR] Failed to install ADDS Role: $_" -ForegroundColor Red
    Exit 1
}

#StepEND

#STEP 5: Import ADDS Deployment Module ---

Write-Host "`n[STEP 5] Importing ADDSDeployment Module..." -ForegroundColor Cyan

try {
    Import-Module ADDSDeployment -ErrorAction Stop
    Write-Host "[OK] Module imported successfully." -ForegroundColor Green
} catch {
    Write-Host "[ERROR] Failed to import ADDSDeployment module: $_" -ForegroundColor Red
    Exit 1
}

#StepEND

#STEP 6: Promote Server to Domain Controller (New Forest) ---

Write-Host "`n[STEP 6] Promoting server to Domain Controller..." -ForegroundColor Cyan

try {
    Install-ADDSForest `
        -DomainName                    $DomainFQDN `
        -DomainNetbiosName             $DomainNetBIOS `
        -ForestMode                    "WinThreshold" `
        -DomainMode                    "WinThreshold" `
        -DatabasePath                  "C:\Windows\NTDS" `
        -LogPath                       "C:\Windows\NTDS" `
        -SysvolPath                    "C:\Windows\SYSVOL" `
        -SafeModeAdministratorPassword $SecureDSRMPassword `
        -InstallDns:$true `
        -NoRebootOnCompletion:$false `
        -Force:$true `
        -ErrorAction Stop

} catch {
    Write-Host "[ERROR] Promotion failed: $_" -ForegroundColor Red
    Exit 1
}

#StepEND

# Note: Server will automatically reboot after successful promotion.
# After reboot, log in with: DOMAIN\Administrator