$Credential = Get-Credential -Message "Enter Login Credential"
$vCenter =  Read-Host -Prompt "Enter vCenter Name"
Connect-VIServer -Server $vCenter -Credential $Credential
$VMs = Get-Content C:\Scripts\Allservers.txt
foreach ($vm in $VMs){
Get-VM $vm | Get-Snapshot | Remove-Snapshot -Confirm:$false}
pause