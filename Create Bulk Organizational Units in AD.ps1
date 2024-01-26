Import-Module activedirectory

$ADOU = Import-csv "C:\Scripts\OU.csv"

foreach ($ou in $ADOU)
{
#Map CSV coumn to variable
$name = $ou.name
$path = $ou.path

# Below command will create OU as per CSV
New-ADOrganizationalUnit `
-Name $name `
-path $path `

}
