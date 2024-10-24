$settingsPopup = New-Object System.Windows.Forms.Form
$settingsPopup.Text ='Settings'
$settingsPopup.Width = 600
$settingsPopup.Height = 210
$settingsPopup.AutoSize = $true
$settingsPopup.ShowInTaskbar = $false
$settingsPopup.MaximizeBox = $false
$settingsPopup.FormBorderStyle = 'FixedSingle'
$settingsPopup.StartPosition = 1 # centre
$settingsPopup.Padding = 3

$folderL = New-Object System.Windows.Forms.Label
$folderL.Text = "Default folder:"
$folderL.Location = New-Object System.Drawing.Point(65, 12)
$folderL.AutoSize = $true

$folderBox = New-Object System.Windows.Forms.TextBox
$folderBox.Width = 350
$folderBox.Location  = New-Object System.Drawing.Point(140, 10)
$folderBox.ReadOnly = $true
$folderBox.add_KeyDown($textBox_KeyDown)

$folderBrowse = New-Object System.Windows.Forms.Button
$folderBrowse.Location = New-Object System.Drawing.Size(500, 8)
$folderBrowse.Size = New-Object System.Drawing.Size(75, 23)
$folderBrowse.Text = "Browse"
$folderBrowse.Add_Click({
    Set-AppData-Default-Directory
    Set-Default-Directory-Global
    $folderBox.Text = "$global:defaultDir"
})

$shortcutsL = New-Object System.Windows.Forms.Label
$shortcutsL.Text ="
If any of the words 'password', 'unprotected' or 'zip' are in a patients general or appointment notes, this app will change the email to be unprotected automatically

If you rename a file in the folder to the any of the shortcuts below, it will rename it into the full version
coc.pdf = Certificate of Capacity.pdf
ir.pdf = Imaging Request.pdf
mc.pdf = Medical Certificate.pdf
pr.pdf = Pathology Request.pdf
ref.pdf = Referral.pdf
res.pdf = Results.pdf
"
$shortcutsL.Location = New-Object System.Drawing.Point(30, 100)
$shortcutsL.AutoSize = $false
$shortcutsL.TextAlign = 512 # BottomCenter
$shortcutsL.Dock = 5 # Bottom

$ocrCheckBox = New-Object System.Windows.Forms.CheckBox
$ocrCheckBox.Location = New-Object System.Drawing.Size(10, 10)
$ocrCheckBox.Text = "OCR?"
$ocrCheckBox.AutoSize = $true

$settingsPopup.Controls.Add($folderL)
$settingsPopup.Controls.Add($folderBox)
$settingsPopup.Controls.Add($folderBrowse)
$settingsPopup.Controls.Add($ocrCheckBox)
$settingsPopup.Controls.Add($shortcutsL)

$settingsButton = New-Object System.Windows.Forms.Button
$settingsButton.Location = New-Object System.Drawing.Size(10, 170)
$settingsButton.Size = New-Object System.Drawing.Size(75, 23)
$settingsButton.Text = "Settings"
$settingsButton.Add_Click({
    $folderBox.Text = "$global:defaultDir"
    $settingsPopup.ShowDialog()
    $settingsPopup.Focus()
})