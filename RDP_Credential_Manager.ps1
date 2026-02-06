Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic

# Path for metadata JSON
$MetadataFile = "$env:USERPROFILE\RDPManager\servers.json"
if (-not (Test-Path $MetadataFile)) {
    New-Item -ItemType File -Path $MetadataFile -Force | Out-Null
    Set-Content $MetadataFile "[]"
}

function Load-Metadata {
    if (Test-Path $MetadataFile) {
        $content = Get-Content $MetadataFile -Raw
        if ([string]::IsNullOrWhiteSpace($content)) { return @() }

        $data = $content | ConvertFrom-Json

        # Wrap single object into array
        if ($data -is [System.Collections.IEnumerable] -and -not ($data -is [string])) {
            return @($data)
        }
        elseif ($null -eq $data) {
            return @()
        }
        else {
            return @($data)
        }
    }
    else {
        return @()
    }
}



function Save-Metadata($data) {
    # Force array output and pretty formatting
    @($data) | ConvertTo-Json -Depth 3 | Set-Content $MetadataFile
}





# Helper: GUI password box
function Show-PasswordBox {
    param([string]$Message = "Enter Password",[string]$Title = "Password")
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.Size = New-Object System.Drawing.Size(300,150)
    $form.StartPosition = "CenterScreen"
	
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Message
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(10,20)
    $form.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(10,50)
    $textBox.Width = 260
    $textBox.UseSystemPasswordChar = $true
    $form.Controls.Add($textBox)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Location = New-Object System.Drawing.Point(200,80)
    $okButton.Add_Click({ $form.Tag = $textBox.Text; $form.Close() })
    $form.Controls.Add($okButton)

    $form.ShowDialog() | Out-Null
    return $form.Tag
}

# GUI setup
$form = New-Object System.Windows.Forms.Form
$form.Text = "RDP Credential Manager"
$form.Size = New-Object System.Drawing.Size(500,400)
$form.StartPosition = "CenterScreen"

# Add form background color here
$form.BackColor = [System.Drawing.Color]::LightSteelBlue

$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Size = New-Object System.Drawing.Size(300,200)
$listBox.Location = New-Object System.Drawing.Point(20,20)
$form.Controls.Add($listBox)

$serverLabel = New-Object System.Windows.Forms.Label
$serverLabel.Text = "Server Name:"
$serverLabel.Location = New-Object System.Drawing.Point(20,240)
$form.Controls.Add($serverLabel)

$serverBox = New-Object System.Windows.Forms.TextBox
$serverBox.Location = New-Object System.Drawing.Point(120,240)
$serverBox.Width = 200   # increase width to fit long FQDNs
$form.Controls.Add($serverBox)

$addButton = New-Object System.Windows.Forms.Button
$addButton.Text = "Add"
$addButton.Location = New-Object System.Drawing.Point(20,280)
$form.Controls.Add($addButton)

# Add button color here
$addButton.BackColor = [System.Drawing.Color]::LightGreen


$updateButton = New-Object System.Windows.Forms.Button
$updateButton.Text = "Update"
$updateButton.Location = New-Object System.Drawing.Point(100,280)
$form.Controls.Add($updateButton)

# Add button color here
$updateButton.BackColor = [System.Drawing.Color]::LightGreen

$deleteButton = New-Object System.Windows.Forms.Button
$deleteButton.Text = "Delete"
$deleteButton.Location = New-Object System.Drawing.Point(180,280)
$form.Controls.Add($deleteButton)

# Add button color here
$deleteButton.BackColor = [System.Drawing.Color]::LightGreen

$connectButton = New-Object System.Windows.Forms.Button
$connectButton.Text = "Connect"
$connectButton.Location = New-Object System.Drawing.Point(260,280)
$form.Controls.Add($connectButton)

# Add button color here
$connectButton.BackColor = [System.Drawing.Color]::LightGreen

# --- Event Handlers ---
$addButton.Add_Click({
    $username = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Username","Add Credential","")
    if (-not $username) { return }

    $plainPassword = Show-PasswordBox -Message "Enter Password" -Title "Add Credential"
    if (-not $plainPassword) { return }

    $serverName = $serverBox.Text
    if (-not $serverName) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a server name.")
        return
    }

    # Store in Credential Manager
    cmdkey /add:$serverName /user:$username /pass:$plainPassword | Out-Null

    # Update metadata
    $data = Load-Metadata
	if (-not $data) { $data = @() }
	$data = @($data)
	$data += [PSCustomObject]@{Server=$serverName;User=$username}
	Save-Metadata $data

    # Update GUI
    $listBox.Items.Add("$serverName - $username")
    [System.Windows.Forms.MessageBox]::Show("Credential added successfully.")
})

$updateButton.Add_Click({
    if (-not $listBox.SelectedItem) {
        [System.Windows.Forms.MessageBox]::Show("Select an entry to update.")
        return
    }
    $parts = $listBox.SelectedItem -split " - "
    $serverName = $parts[0]; $username = $parts[1]

    $plainPassword = Show-PasswordBox -Message "Enter new password for $username" -Title "Update Credential"
    if (-not $plainPassword) { return }

    cmdkey /add:$serverName /user:$username /pass:$plainPassword | Out-Null
    [System.Windows.Forms.MessageBox]::Show("Credential updated successfully.")
})

$deleteButton.Add_Click({
    if (-not $listBox.SelectedItem) {
        [System.Windows.Forms.MessageBox]::Show("Select an entry to delete.")
        return
    }

    $parts = $listBox.SelectedItem -split " - "
    $serverName = $parts[0]; $username = $parts[1]

    # Confirmation prompt
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Are you sure you want to delete the credential for $serverName - $username?",
        "Confirm Delete",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        cmdkey /delete:$serverName | Out-Null

        # Load metadata
        $data = Load-Metadata

        # Normalize into array
        $data = @($data)

        # Remove matching entry
        $newData = $data | Where-Object { -not (($_.Server -eq $serverName) -and ($_.User -eq $username)) }

        # If nothing left, explicitly write empty array
        if ($newData.Count -eq 0) {
            Set-Content $MetadataFile "[]"
        }
        else {
            $newData | ConvertTo-Json -Depth 3 | Set-Content $MetadataFile
        }

        $listBox.Items.Remove($listBox.SelectedItem)
        [System.Windows.Forms.MessageBox]::Show("Credential deleted successfully.")
    }
})


$connectButton.Add_Click({
    if (-not $listBox.SelectedItem) {
        [System.Windows.Forms.MessageBox]::Show("Please select a server entry.")
        return
    }

    $parts = $listBox.SelectedItem -split " - "
    $serverName = $parts[0]

    mstsc /v:$serverName
})

# --- Load saved entries on startup ---
$data = Load-Metadata
foreach ($entry in $data) {
    $listBox.Items.Add("$($entry.Server) - $($entry.User)")
}

$form.ShowDialog()