#Provide KB file information by adding comma in between multiple KB files as below
$Folder= "C:\WINUpdate\windows10.0-kb5029924-x64.msu" , "C:\WINUpdate\windows10.0-kb5019182-x64.msu"

foreach ($server in $Servers){
#Check if "WINUpdate Folder is present in C drive on Remote Computer, If not then It will create it
$Test = Test-Path -path "\\$Server\c$\WINUpdate\"
If ($Test -eq $True) {Write-Host "Path exists, hence Copying Patches on $server."}
Else {(Write-Host "Path doesnt exists, hence Creating folder on $server ") , (New-Item -ItemType Directory -Name WINUpdate -Path "\\$Server\c$\")}

#Export MSU files into "WINUpdate folder on remote computers"
expand -F:* $Folder "\\$Server\c$\WINUpdate"

#Find how many CAB files has been extracted by above command and Select only CAB file for further process
icm -ComputerName $server -Scriptblock {

$CabFile = Get-ChildItem "c:\WINUpdate\Window*.cab" -Recurse | Select-Object -ExpandProperty VersionInfo 
$CabPath  = $CabFile | Select FileName -ExpandProperty FileName

#Run installation against each CAB file and Try to install it on remote computer using DISM Tool. If KB is not applicable on computer it will show error
Foreach ($cab in $CabPath){
& cmd /c DISM /Online /Add-Package /PackagePath:$cab /quiet /norestart 
}

#Wait for 5 Seconds to delete all extracted KB files from remote computer. Dont save any other data into "WINUpdate" as below command will delete all files present into it.
Start-Sleep -Seconds 5
DEL c:\WINUpdate\*.*
}}