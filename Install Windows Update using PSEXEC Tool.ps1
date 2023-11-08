# Here we are selecting multiple Patches
Add-Type -AssemblyName System.Windows.Forms
$FilePicker = New-Object -TypeName System.Windows.Forms.OpenFileDialog  -Property @{ Title = "Select Windows Patches files to Install (*.MSU)" 
Filter = '*.MSU |*.msu'}
$FilePicker.Multiselect = $true
$result = $FilePicker.ShowDialog()
$Hotfixes = $FilePicker.FileNames
if (!$Hotfixes){
Read-Host -Prompt "File not selected, Press any key to Exit the script" 
exit}
else {
$Patches = Split-Path $Hotfixes -leaf}


#Here we are selecting server list in TXT format
Add-Type -AssemblyName System.Windows.Forms
$ServersPicker = New-Object -TypeName System.Windows.Forms.OpenFileDialog -Property @{Title = "Select Servers List (*.TXT)" 
Filter = '*.TXT,*.CSV |*.txt;*.csv'}
$result = $ServersPicker.ShowDialog()
$ServerFile = $ServersPicker.FileName
if (!$ServerFile){
Read-Host -Prompt "File not selected, Press any key to Exit the script" 
exit}
else {
$Servers = Get-Content $ServerFile}

#Log file will be saved at below lcoation. You can change name of log file and change path as well.
$Log = Split-Path $ServerFile -Parent
$Logfile = $Log + '\Recent_OS_Updates.csv'

#Here we are showing Server list before execution
Write-host 'Server List' -ForegroundColor Green
$Servers
Write-Host "`n$line"
#Here we are showing Patches list before execution
Write-host 'Patches List' -ForegroundColor Green
$Hotfixes
Write-Host "`n$line"

#Confiramtion prompt to ensure you have selected correct computer list and Patches
$confirmation = Read-Host "Are you Sure ? You Want To Install Patches on above computer (Press 'Yes')"

if ($confirmation -eq 'yes') {

foreach ($Server in $Servers)
{
    $needsReboot = $False
    $remotePath = "\\$Server\c$\WINUpdate\"
    
        if( ! (Test-Connection $Server -Count 1 -Quiet)) 
    {
        Write-Warning "$Server is not accessible"
        continue
    }

        if(!(Test-Path $remotePath))
    {
        New-Item -ItemType Directory -Force -Path $remotePath | Out-Null
    }
    
    foreach ($Patch in $Patches)
    {

        Copy-Item $Hotfixes $remotePath
        # Run command as SYSTEM via PsExec (-s switch)
        & C:\PSTools\PsExec -s \\$Server wusa C:\WINUpdate\$Patch /quiet /norestart
        
        if ($LastExitCode -eq 3010) {
        Write-Host $Patch Installed on Server $Server
            $needsReboot = $true
        }
        if ($LastExitCode -eq -2145124329) {
        Write-Host $Patch is not applicable on Server $Server
            $needsReboot = $true
        }
        if ($LastExitCode -eq 123) {
        Write-Host $Patch not found on installation Path on Server $Server
            $needsReboot = $true
        }
    }

    # Delete patches copied on remote server
    Remove-Item $remotePath -Force -Recurse

    if($needsReboot)
    {
        Write-Host "Restarting $Server..."
        Restart-Computer -ComputerName $Server -Force -Confirm
    }
}}
else
{exit}

$confirmation = Read-Host "Do you want to run Patch Report for Today's Patch Installation ? (Press 'Yes')"

if ($confirmation -eq 'yes') {
foreach ($Server in $Servers){

(get-hotfix -ComputerName $Server |Where-Object { $_.InstalledOn -gt ((Get-Date).AddDays(-1)) } | sort source) | Export-CSV -Path $Logfile -NoTypeInformation -Append
    Start-Sleep -Seconds 10  } 
 Write-Host 'Detailed report has been saved at' $Logfile -ForegroundColor Green
}

else
{Exit}
Read-Host "Press Enter to Exit"