# Prompt for credentials once
$Credential = Get-Credential -Message "Enter vCenter Login Credential"

# Import mapping file (VMName,vCenter)
$VMMapping = Import-Csv C:\Scripts\VMMapping.csv

# Group VMs by vCenter
$GroupedVMs = $VMMapping | Group-Object vCenter

foreach ($group in $GroupedVMs) {
    $vCenter = $group.Name
    $vmList  = $group.Group.VMName

    Write-Host "Connecting to $vCenter..."
    Connect-VIServer -Server $vCenter -Credential $Credential -ErrorAction Stop

    foreach ($vm in $vmList) {
        try {
            Write-Host "Processing VM: $vm on $vCenter"
            Get-VM -Name $vm -ErrorAction Stop | Get-Snapshot | Remove-Snapshot -Confirm:$false
        }
        catch {
            Write-Warning "Could not process VM $vm on $vCenter. Error: $_"
        }
    }

    Disconnect-VIServer -Server $vCenter -Confirm:$false
}
pause
