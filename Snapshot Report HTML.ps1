# Connect to vCenter
connect-viserver -Server mum-vc01.sitech.com -Verbose

#File Path to save report
$FilePath = "C:\Scripts\Test\SnapshotReport.html"

#Get all snapshots
$Snapshots = Get-VM | Get-Snapshot | select Description,Created,VM,SizeMB,SizeGB

# Format snapshot Size
function Get-SnapshotSize ($Snapshot)
{
    if ($snapshot.SizeGB -ge "1")
    {
        $Snapshotsize = [string]([math]::Round($snapshot.SizeGB,3)) + " GB"
        }
        else {
        $Snapshotsize = [string]([math]::Round($snapshot.SizeMB,3)) + " MB"
        }
   Return $Snapshotsize
}

#Generate HTML Report

$date = (get-date -Format d/M/yyyy)
$header =@"
 <Title>Snapshot Report - $date</Title>
<style>
body {   font-family: 'Helvetica Neue', Helvetica, Arial;
         font-size: 14px;
         line-height: 20px;
         font-weight: 400;
         color: black;
    }
table{
  margin: 0 0 40px 0;
  width: 100%;
  box-shadow: 0 1px 3px rgba(0,0,0,0.2);
  display: table;
  border-collapse: collapse;
  border: 1px solid black;
}
th {
    font-weight: 900;
    color: #ffffff;
    background: black;
   }
td {
    border: 0px;
    border-bottom: 1px solid black
    }
</style>
"@

# Convert to HTML
$HTMLmessage = $Snapshots | select VM,Created,@{Label="Size";Expression={Get-SnapshotSize($_)}},Description | sort Created -Descending | ConvertTo-Html -Head $header

# Format the HTML Report
$Report = Format-HTMLBody ($HTMLmessage)
$Report | out-file -FilePath $FilePath
Write-Host "Your snapshots report has been saved to: $FilePath"

# Disconnect from vCenter Server
Disconnect-VIServer -Confirm:$false