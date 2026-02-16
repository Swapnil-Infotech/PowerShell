# Hide PowerShell console window
Add-Type -Name Window -Namespace Console -MemberDefinition @"
    [DllImport("kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
"@
$consolePtr = [Console.Window]::GetConsoleWindow()
[Console.Window]::ShowWindow($consolePtr, 0)   # 0 = Hide, 6 = Minimize

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Import-Module ImportExcel

# Create Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Excel Merger Tool by Swapnil InfoTech"
$form.Size = New-Object System.Drawing.Size(750,550)
$form.StartPosition = "CenterScreen"

# DataGridView for files
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Size = New-Object System.Drawing.Size(500,250)
$grid.Location = New-Object System.Drawing.Point(20,20)
$grid.AllowUserToAddRows = $false
$grid.RowHeadersVisible = $false

# Add Columns
$colCheck = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$colCheck.HeaderText = "Include"
$colCheck.Width = 60
$grid.Columns.Add($colCheck)

$colName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colName.HeaderText = "File Name"
$colName.Width = 200
$grid.Columns.Add($colName)

$colPath = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colPath.HeaderText = "Path"
$colPath.Width = 220
$grid.Columns.Add($colPath)

$form.Controls.Add($grid)

# Buttons
$btnSelectFiles = New-Object System.Windows.Forms.Button
$btnSelectFiles.Text = "Select Files"
$btnSelectFiles.Location = New-Object System.Drawing.Point(540,20)
$btnSelectFiles.Size = New-Object System.Drawing.Size(150,25)
$form.Controls.Add($btnSelectFiles)

$btnSelectFolder = New-Object System.Windows.Forms.Button
$btnSelectFolder.Text = "Select Folder"
$btnSelectFolder.Location = New-Object System.Drawing.Point(540,60)
$btnSelectFolder.Size = New-Object System.Drawing.Size(150,25)
$form.Controls.Add($btnSelectFolder)

$btnReset = New-Object System.Windows.Forms.Button
$btnReset.Text = "Reset"
$btnReset.Location = New-Object System.Drawing.Point(540,100)
$btnReset.Size = New-Object System.Drawing.Size(150,25)
$form.Controls.Add($btnReset)

$btnMerge = New-Object System.Windows.Forms.Button
$btnMerge.Text = "Merge Files"
$btnMerge.Location = New-Object System.Drawing.Point(540,140)
$btnMerge.Size = New-Object System.Drawing.Size(150,25)
$form.Controls.Add($btnMerge)

# Timestamp option
$chkTimestamp = New-Object System.Windows.Forms.CheckBox
$chkTimestamp.Text = "Append Timestamp To `nFile Name"
$chkTimestamp.Location = New-Object System.Drawing.Point(540,180)
$chkTimestamp.Size = New-Object System.Drawing.Size(150,50)
$chkTimestamp.Checked = $true
$form.Controls.Add($chkTimestamp)

# Status Label
$status = New-Object System.Windows.Forms.Label
$status.Text = "Ready"
$status.Location = New-Object System.Drawing.Point(20,280)
$status.Size = New-Object System.Drawing.Size(500,30)
$form.Controls.Add($status)

# Progress Bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(20,320)
$progressBar.Size = New-Object System.Drawing.Size(500,25)
$progressBar.Minimum = 0
$form.Controls.Add($progressBar)

# Column Selection Box
$chkColumns = New-Object System.Windows.Forms.CheckedListBox
$chkColumns.Location = New-Object System.Drawing.Point(20,360)
$chkColumns.Size = New-Object System.Drawing.Size(500,120)
$form.Controls.Add($chkColumns)

# Function to load headers preserving natural order
function Load-Headers {
    $chkColumns.Items.Clear()
    $orderedHeaders = New-Object System.Collections.Generic.List[string]

    if ($grid.Rows.Count -gt 0) {
        # Start with the first file's headers in natural order
        $firstFile = $grid.Rows[0].Cells[2].Value
        try {
            $excel = Open-ExcelPackage -Path $firstFile
            $ws = $excel.Workbook.Worksheets[1]   # first worksheet
            $colCount = $ws.Dimension.End.Column

            # Read header row (row 1) cell by cell
            for ($i = 1; $i -le $colCount; $i++) {
                $header = $ws.Cells[1,$i].Text
                if (-not [string]::IsNullOrWhiteSpace($header)) {
                    if (-not $orderedHeaders.Contains($header)) {
                        $orderedHeaders.Add($header)
                    }
                }
            }
        } catch {
            $status.Text = "Error reading headers from $firstFile"
        }

        # Append headers from other files (preserve encounter order)
        foreach ($row in $grid.Rows) {
            $file = $row.Cells[2].Value
            try {
                $excel = Open-ExcelPackage -Path $file
                $ws = $excel.Workbook.Worksheets[1]
                $colCount = $ws.Dimension.End.Column

                for ($i = 1; $i -le $colCount; $i++) {
                    $header = $ws.Cells[1,$i].Text
                    if (-not [string]::IsNullOrWhiteSpace($header)) {
                        if (-not $orderedHeaders.Contains($header)) {
                            $orderedHeaders.Add($header)
                        }
                    }
                }
            } catch {
                $status.Text = "Error reading headers from $file"
            }
        }
    }

    # Populate the checkbox list in preserved order
    foreach ($h in $orderedHeaders) {
        [void]$chkColumns.Items.Add($h, $true)
    }
}

# Event Handlers
$btnSelectFiles.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "Excel Files|*.xlsx"
    $dialog.Multiselect = $true
    if ($dialog.ShowDialog() -eq "OK") {
        foreach ($file in $dialog.FileNames) {
            $row = $grid.Rows.Add()
            $grid.Rows[$row].Cells[0].Value = $true
            $grid.Rows[$row].Cells[1].Value = [System.IO.Path]::GetFileName($file)
            $grid.Rows[$row].Cells[2].Value = $file
        }
        Load-Headers
    }
})

$btnSelectFolder.Add_Click({
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($folderDialog.ShowDialog() -eq "OK") {
        Get-ChildItem $folderDialog.SelectedPath -Filter *.xlsx | ForEach-Object {
            $row = $grid.Rows.Add()
            $grid.Rows[$row].Cells[0].Value = $true
            $grid.Rows[$row].Cells[1].Value = $_.Name
            $grid.Rows[$row].Cells[2].Value = $_.FullName
        }
        Load-Headers
    }
})

$btnReset.Add_Click({
    $grid.Rows.Clear()
    $progressBar.Value = 0
    $chkColumns.Items.Clear()
    $status.Text = "Ready"
})

$btnMerge.Add_Click({
    $selectedFiles = @()
    foreach ($row in $grid.Rows) {
        if ($row.Cells[0].Value -eq $true) {
            $selectedFiles += $row.Cells[2].Value
        }
    }

    if ($selectedFiles.Count -eq 0) {
        $status.Text = "No files selected!"
        return
    }

    # Get selected columns in the order they appear in CheckedListBox
    $selectedColumns = @()
    foreach ($item in $chkColumns.CheckedItems) {
        $selectedColumns += $item
    }

    if ($selectedColumns.Count -eq 0) {
        $status.Text = "No columns selected!"
        return
    }

    # Build default filename
    $defaultName = "Merged.xlsx"
    if ($chkTimestamp.Checked) {
        $timestamp = (Get-Date).ToString("yyyy-MM-dd_HHmm")
        $defaultName = "Merged_$timestamp.xlsx"
    }

    # Save File Dialog
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "Excel Workbook|*.xlsx"
    $saveDialog.Title = "Save Merged Excel File"
    $saveDialog.FileName = $defaultName

    if ($saveDialog.ShowDialog() -eq "OK") {
        $output = $saveDialog.FileName
        $status.Text = "Merging..."
        $progressBar.Maximum = $selectedFiles.Count
        $progressBar.Value = 0

        $data = @()
        foreach ($file in $selectedFiles) {
            try {
                # Preserve column order by using header row order
                $imported = Import-Excel $file | Select-Object ($selectedColumns)
                $data += $imported
                $progressBar.Value += 1
                $form.Refresh()
            } catch {
                $status.Text = "Error importing $file"
            }
        }

        try {
            $data | Export-Excel $output -WorksheetName "Combined"
            $status.Text = "Merge complete! Saved to $output"
        } catch {
            $status.Text = "Error saving merged file!"
        }
    }
})

# Show Form
$form.ShowDialog()