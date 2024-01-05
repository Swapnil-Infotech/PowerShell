Import-Module BitsTransfer
$Servers = Get-content "C:\Scripts\livePCs.txt"
$Folder= "c:\Software\Firefox Setup 121.0.exe"

Foreach ($Server in $Servers) {
IF (Test-Connection -BufferSize 32 -Count 1 -ComputerName $Server -Quiet) {
$Test = Test-Path -path "\\$Server\c$\Temp\"

If ($Test -eq $True) {Write-Host "Path exists, hence installing softwares on $Server."}

Else {(Write-Host "Path doesnt exists, hence Creating foldet on $Server and starting installation") , (New-Item -ItemType Directory -Name Temp -Path "\\$Server\c$" | Out-Null)}
Start-BitsTransfer -Source $Folder -Destination \\$Server\c$\Temp -Description "Backup" -DisplayName "Copying file on $Server"

Invoke-Command -ComputerName $Server -ScriptBlock {
Start-Process -FilePath "c:\Temp\Firefox Setup 121.0.exe" -ArgumentList "/S"
Start-Sleep -s 35
rm -Force "c:\Temp\Firefox Setup 121.0.exe"

}}
else {Write-Host "The remote computer " $Server " is Offline"}
}
