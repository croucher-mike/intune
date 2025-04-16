# Add-Type for Windows Forms
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "MSI to IntuneWin Converter"
$form.Size = New-Object System.Drawing.Size(600, 400)
$form.StartPosition = "CenterScreen"

# Folder Browser Dialog
$folderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog

# Define Controls
$labelMsiFolder = New-Object System.Windows.Forms.Label
$labelMsiFolder.Text = "MSI Folder:"
$labelMsiFolder.Location = New-Object System.Drawing.Point(20, 20)
$labelMsiFolder.Size = New-Object System.Drawing.Size(100, 20)

$textBoxMsiFolder = New-Object System.Windows.Forms.TextBox
$textBoxMsiFolder.Location = New-Object System.Drawing.Point(120, 20)
$textBoxMsiFolder.Size = New-Object System.Drawing.Size(350, 20)

$buttonBrowseMsiFolder = New-Object System.Windows.Forms.Button
$buttonBrowseMsiFolder.Text = "Browse"
$buttonBrowseMsiFolder.Location = New-Object System.Drawing.Point(480, 20)
$buttonBrowseMsiFolder.Size = New-Object System.Drawing.Size(75, 20)
$buttonBrowseMsiFolder.Add_Click({
    if ($folderBrowserDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textBoxMsiFolder.Text = $folderBrowserDialog.SelectedPath
    }
})

$labelOutputFolder = New-Object System.Windows.Forms.Label
$labelOutputFolder.Text = "Output Folder:"
$labelOutputFolder.Location = New-Object System.Drawing.Point(20, 60)
$labelOutputFolder.Size = New-Object System.Drawing.Size(100, 20)

$textBoxOutputFolder = New-Object System.Windows.Forms.TextBox
$textBoxOutputFolder.Location = New-Object System.Drawing.Point(120, 60)
$textBoxOutputFolder.Size = New-Object System.Drawing.Size(350, 20)

$buttonBrowseOutputFolder = New-Object System.Windows.Forms.Button
$buttonBrowseOutputFolder.Text = "Browse"
$buttonBrowseOutputFolder.Location = New-Object System.Drawing.Point(480, 60)
$buttonBrowseOutputFolder.Size = New-Object System.Drawing.Size(75, 20)
$buttonBrowseOutputFolder.Add_Click({
    if ($folderBrowserDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textBoxOutputFolder.Text = $folderBrowserDialog.SelectedPath
    }
})

$buttonStart = New-Object System.Windows.Forms.Button
$buttonStart.Text = "Start Conversion"
$buttonStart.Location = New-Object System.Drawing.Point(20, 100)
$buttonStart.Size = New-Object System.Drawing.Size(200, 30)
$buttonStart.Add_Click({
    $msiFolder = $textBoxMsiFolder.Text
    $outputFolder = $textBoxOutputFolder.Text

    if (-not (Test-Path $msiFolder)) {
        [System.Windows.Forms.MessageBox]::Show("Please select a valid MSI folder.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    if (-not (Test-Path $outputFolder)) {
        [System.Windows.Forms.MessageBox]::Show("Please select a valid output folder.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # Run the packaging script logic
    Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSScriptRoot\MSI_2_Intunewin.ps1`" -msiFolder `"$msiFolder`" -outputFolder `"$outputFolder`"" -NoNewWindow
    [System.Windows.Forms.MessageBox]::Show("Conversion started!", "Info", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})

# Add Controls to Form
$form.Controls.Add($labelMsiFolder)
$form.Controls.Add($textBoxMsiFolder)
$form.Controls.Add($buttonBrowseMsiFolder)
$form.Controls.Add($labelOutputFolder)
$form.Controls.Add($textBoxOutputFolder)
$form.Controls.Add($buttonBrowseOutputFolder)
$form.Controls.Add($buttonStart)

# Show Form
$form.Add_Shown({ $form.Activate() })
[void] $form.ShowDialog()
