Add-Type -AssemblyName System.Windows.Forms

# Prompt for credentials once
$DomainCred = Get-Credential -Message "Enter DOMAIN credentials (with rights to join computers)"
$LocalCred  = Get-Credential -Message "Enter LOCAL admin credentials (for remote connection)"

# Prompt user to select CSV file
$FileDialog = New-Object System.Windows.Forms.OpenFileDialog
$FileDialog.Title = "Select CSV File (IP,DomainName,OUPath)"
$FileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
$null = $FileDialog.ShowDialog()

if (-not $FileDialog.FileName) {
    Write-Host "No file selected. Exiting..." -ForegroundColor Red
    exit
}

$CSVFile = $FileDialog.FileName

# Import CSV (must have headers: ComputerIP,DomainName,OUPath)
$Servers = Import-Csv $CSVFile

# Prepare log file
$LogFile = Join-Path (Split-Path $CSVFile) "DomainJoinLog.csv"
"ComputerIP,DomainName,OUPath,Status,Message" | Out-File $LogFile

# Backup current TrustedHosts
$OriginalTrustedHosts = (Get-Item WSMan:\localhost\Client\TrustedHosts).Value

# Build list of IPs from CSV
$IPs = $Servers.ComputerIP -join ","

# Temporarily add IPs to TrustedHosts
Set-Item WSMan:\localhost\Client\TrustedHosts -Value $IPs -Force
Write-Host "TrustedHosts temporarily set to: $IPs" -ForegroundColor Yellow

foreach ($Server in $Servers) {
    $ComputerIP = $Server.ComputerIP
    $DomainName = $Server.DomainName
    $OUPath     = $Server.OUPath

    Write-Host "Processing $ComputerIP for domain $DomainName ..." -ForegroundColor Cyan

    try {
        $Session = New-PSSession -ComputerName $ComputerIP -Credential $LocalCred -ErrorAction Stop

        Invoke-Command -Session $Session -ScriptBlock {
            param($DomainName, $DomainCred, $OUPath)
            if ([string]::IsNullOrWhiteSpace($OUPath)) {
                Add-Computer -DomainName $DomainName -Credential $DomainCred -Force -ErrorAction Stop
            }
            else {
                Add-Computer -DomainName $DomainName -Credential $DomainCred -OUPath $OUPath -Force -ErrorAction Stop
            }
            Restart-Computer -Force
        } -ArgumentList $DomainName, $DomainCred, $OUPath

        Remove-PSSession $Session
        Write-Host "$ComputerIP joined successfully." -ForegroundColor Green
        "$ComputerIP,$DomainName,$OUPath,Success,Joined domain" | Out-File $LogFile -Append
    }
    catch {
        Write-Host "Failed to process $ComputerIP : $_" -ForegroundColor Red
        "$ComputerIP,$DomainName,$OUPath,Failed,$($_.Exception.Message)" | Out-File $LogFile -Append
    }
}

# Restore original TrustedHosts
if ($OriginalTrustedHosts) {
    Set-Item WSMan:\localhost\Client\TrustedHosts -Value $OriginalTrustedHosts -Force
    Write-Host "TrustedHosts restored to original value: $OriginalTrustedHosts" -ForegroundColor Yellow
}
else {
    Clear-Item WSMan:\localhost\Client\TrustedHosts -Force
    Write-Host "TrustedHosts cleared back to empty." -ForegroundColor Yellow
}

Write-Host "Operation complete. Log saved to $LogFile" -ForegroundColor Yellow