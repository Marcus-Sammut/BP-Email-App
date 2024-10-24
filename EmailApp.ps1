# C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -window hidden -ExecutionPolicy RemoteSigned -command "& 'C:\Users\recep\OneDrive\Desktop\Test\New folder\email.ps1'"
. .\PatientScrape.ps1
. .\Helper.ps1

Set-Variable -Name lastEditHwnd -Value $null -Scope global
Set-Variable -Name closeAllPatientsFirst -Value $false -Scope global

function Autofill-Form {
    $patientDetails = Get-Patient
    if ($patientDetails.length -gt 1 -and -not $closeAllPatientsFirst) {
        $emailBox.Enabled = $dobBox.Enabled = $subjBox.Enabled = $zipBox.Enabled = $true
        $AddFileButton.Enabled = $ClearButton.Enabled = $UnprotectedCheckbox.Enabled = $autofillCheckbox.Enabled = $true
        $patientEmail, $patientDOB, $firstName, $lastName, $editHwnd, $noZip = $patientDetails
        if ($editHwnd -ne $lastEditHwnd) {
            Clear-Form
            $emailBox.Text = $patientEmail
            $dobBox.Text = $patientDOB
            $zipBox.Text = '{0} {1}' -f $firstName, $lastName
            if ($noZip) {
                $checkNoZipL.Text = "CHECK IF PATIENT DOESN'T WANT ZIP FILE"
                $UnprotectedCheckbox.Checked = $true
            } else {
                $checkNoZipL.Text = ""
            }
            $subjBox.Focus()
            Set-Variable -Name lastEditHwnd -Value $editHwnd -Scope global
        }
    } elseif ($patientDetails -gt 0) { # if window count is at least 1
        Clear-Form
        Set-Variable -Name lastEditHwnd -Value $null -Scope global
        Set-Variable -Name closeAllPatientsFirst -Value $true -Scope global
        $emailBox.Enabled = $dobBox.Enabled = $subjBox.Enabled = $zipBox.Enabled = $false
        $AddFileButton.Enabled = $ClearButton.Enabled = $UnprotectedCheckbox.Enabled = $autofillCheckbox.Enabled = $false
        $checkNoZipL.Text = "Multiple patient files opened, close all first"
    } elseif ($patientDetails -eq 0 -and $closeAllPatientsFirst) {
        Set-Variable -Name closeAllPatientsFirst -Value $false -Scope global
    }
}

function Clear-Form {
    Clear-Files
    $UnprotectedCheckbox.checked = $false
    $emailBox.Text = $dobBox.Text = $dobErrorL.Text = $subjBox.Text = $zipBox.Text = $checkNoZipL.Text = ""
    $emailBox.Focus()
}

function Clear-Files {
    #TODO
}

function Enable-Emailing {
    #TODO
}

$global:zipFileNames = @()

function Add-FileToZip {
    # Don't autofill while fileDialog is open
    $autofillCheckbox.Checked = $false

    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
        Title = Set-FileDialog-Title
        InitialDirectory = $global:defaultDir
        MultiSelect = $false
        Filter = 'PDF Files (*.pdf)|*.pdf|All Files (*.*)|*.*'
    }

    $FileBrowser.add_FileOk({
        param($s, $e)
        $filename = Split-Path $FileBrowser.FileName -leaf
        if ($filename -match "^Scan\d{4}-\d\d-\d\d_\d{6}\.pdf$") {
            $e.Cancel = $true
            [System.Windows.MessageBox]::Show("Please rename the selected file: {0}" -f $filename, "Error")
        } elseif ($global:zipFileNames -contains $filename) {
            $e.Cancel = $true
            [System.Windows.MessageBox]::Show("You have already selected this file: {0}" -f $filename, "Error")
        }
    })
    if ($FileBrowser.ShowDialog() -eq 1) { # sys.win.forms.dialogresult OK == 1
        $global:zipFileNames += Split-Path $FileBrowser.FileName -leaf
        $mainForm.Text = $global:zipFileNames
        $filesListedL.Text = $global:zipFileNames -join ', '
        $filesAddedL.Text = "Files added: {0}" -f $global:zipFileNames.Count
        $ResetFilesButton.Enabled = $true
        $SendButton.Enabled = $true
    }
}

function Send-Email {
    # Delete previous zip
    Remove-Item .\*.zip

    # Cleanup variables/parameters
    $dob = $dobBox.Text
    $email = $emailBox.Text.trim()
    $subj = (Get-Culture).TextInfo.ToTitleCase($subjBox.Text).trim()
    $zipName = $zipBox.Text.trim()
    Set-Location $global:defaultDir

    # Complete zip file name
    if ($zipName -eq '') {
        $zipName = "Attachment_"+[datetime]::Now.ToString('MM-dd-yy_hh-mm-ss')+".zip"
    } elseif ($global:zipFileNames.Count -gt 1) {
        $zipName = "{0} - {1} files.zip" -f $zipName, $global:zipFileNames.Count
    } else {
        $zipName = $zipName + ".zip"
    }
    #TODO check that the files are VALID aka havent been renamed/don't exist
    if ($UnprotectedCheckbox.checked) {
        $ol = New-Object -ComObject Outlook.Application
        $new = $ol.CreateItem(0)
        $new.To = $email
        $new.Subject = $subj
        $inspector = $new.GetInspector
        foreach($file in $global:zipFileNames) {
            $new.Attachments.Add($file)
        }
        $inspector.Activate()
        Remove-Item $global:zipFileNames
    } else {
        &"C:\Program Files\7-Zip\7z.exe" a $zipName "-p$dob" $global:zipFileNames -sdel
        $oLookPath = "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
        $zipPath = "`"$pwd\$zipName`""
        $escaped_subj = [uri]::EscapeDataString($subj)
        #$zipPath = "$pwd\$zipName"; while (-Not (Test-Path $zipPath)) {Start-Sleep -Seconds 0.1}; $new.Attachments.Add($zipPath); $inspector.Activate()
        &$oLookPath /c ipm.note /m "$email`?subject=$escaped_subj" /a $zipPath
    }
    Clear-Form
    $global:zipFileNames = @()
    $SendButton.Enabled = $false
    $filesAddedL.Text = "Files added: 0"
    $filesListedL.Text = ""
    $ResetFilesButton.Enabled = $false
    $autofillCheckbox.Checked = $true
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationCore,PresentationFramework # for messageBox

$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Text ='NEW VERSION NEW VERSION Create New Email - Press F3 to copy DOB'
$mainForm.Width = 460
$mainForm.Height = 260
$mainForm.AutoSize = $true
$mainForm.MaximizeBox = $False
$mainForm.FormBorderStyle = 'FixedSingle'
$mainForm.StartPosition = 1 #centre

$templateFont = [System.Drawing.Font]::new("Arial", 12, [System.Drawing.FontStyle]::Bold)

$copiedLFlashScript = {
    Set-Clipboard -Value $dobBox.Text
    $dobL.Visible = $false
    $copiedL.Visible = $true
    $dobL.Cursor = [System.Windows.Forms.Cursors]::Pointer
    Start-Sleep -Milliseconds 600
    $dobL.Visible = $true
    $dobL.Cursor = [System.Windows.Forms.Cursors]::Hand
    $copiedL.Visible = $false
}

$textBox_KeyDown = [System.Windows.Forms.KeyEventHandler] {
    # https://learn.microsoft.com/en-us/dotnet/api/system.windows.forms.keys
    $F2_code = 113
    $F3_code = 114
    $F4_code = 115
    if ($_.KeyCode -eq $F2_code) {
        Autofill-Form
        $_.SuppressKeyPress = $true
    } elseif ($_.KeyCode -eq $F3_code) {
        & $copiedLFlashScript
        $_.SuppressKeyPress = $true
    } elseif ($_.KeyCode -eq $F4_code) {
        Clear-Form
        $_.SuppressKeyPress = $true
    } elseif ($_.KeyCode -eq 'Enter') {
        #$AddFileButton.PerformClick()
        $_.SuppressKeyPress = $true
    } elseif ($_.KeyCode -eq 'Escape') {
        #$ClearButton.PerformClick()
        $_.SuppressKeyPress = $true
    } elseif ($_.Control -and $_.KeyCode -eq 'A') {
        foreach ($box in $emailBox, $dobBox, $subjBox, $zipBox) {
            if ($box.ContainsFocus) {
                $box.SelectAll()
                $_.SuppressKeyPress = $true
            }
        }
    }
}

# Email label and box
$emailL = New-Object System.Windows.Forms.Label
$emailL.Text = "Email:"
$emailL.font = $templateFont
$emailL.Location = New-Object System.Drawing.Point(10, 10)
$emailL.AutoSize = $true
$mainForm.Controls.Add($emailL)

$emailBox = New-Object System.Windows.Forms.TextBox
$emailBox.Width = 300
$emailBox.Location  = New-Object System.Drawing.Point(155, 10)
$emailBox.AutoCompleteSource = 'CustomSource'
$emailBox.AutoCompleteMode='SuggestAppend'
$emailBox.add_KeyDown($textBox_KeyDown)
$mainForm.Controls.Add($emailBox)

# DOB label and mouse events
$dobL = New-Object System.Windows.Forms.Label
$dobLDefaultText = "DOB(ddmmyy):"
$dobL.Text = $dobLDefaultText
$dobL.Font = $templateFont
$dobL.Location = New-Object System.Drawing.Point(10, 50)
$dobL.AutoSize = $true

# dobL mouse hover events
$dobL.Cursor = [System.Windows.Forms.Cursors]::Hand
$dobL.Add_MouseEnter({
    $dobL.font = [System.Drawing.Font]::new("Arial", 12, [System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Underline)
})
$dobL.Add_MouseLeave({
    $dobL.Font = $templateFont
})

# dobL click event
$copiedL = New-Object System.Windows.Forms.Label
$copiedL.Text = "Copied!"
$copiedL.Font = $templateFont
$copiedL.Location = New-Object System.Drawing.Point(10, 50)
$copiedL.AutoSize = $true
$copiedL.Visible = $false
$mainForm.Controls.Add($copiedL)

$dobL.Add_Click($copiedLFlashScript)
$mainForm.Controls.Add($dobL)

# DOB Box
$dobBox = New-Object System.Windows.Forms.TextBox
$dobBox.Width = 300
$dobBox.Location = New-Object System.Drawing.Point(155, 50)
$dobBox.MaxLength = 6
$dobBox.add_KeyDown($textBox_KeyDown)
$dobBox.Text = "111111"
$mainForm.Controls.Add($dobBox)

#DOB error label
$dobErrorL = New-Object System.Windows.Forms.Label
$dobErrorL.Location = New-Object System.Drawing.Point(155, 70)
$dobErrorL.Size = New-Object System.Drawing.Size(150, 15)
$dobErrorL.ForeColor = "red"
$mainForm.Controls.Add($dobErrorL)

# Subject label and box
$subjL = New-Object System.Windows.Forms.Label
$subjL.Text = "Subject:"
$subjL.font = $templateFont
$subjL.Location = New-Object System.Drawing.Point(10, 90)
$subjL.AutoSize = $true
$mainForm.Controls.Add($subjL)

$subjBox = New-Object System.Windows.Forms.ComboBox
$subjBox.Width = 300
$subjBox.Location  = New-Object System.Drawing.Point(155, 90)
$subjectList = 'Attachments','Certificate of Capacity', 'Documents', 'Imaging Request', 'Medical Certificate', 'Pathology Request', 'Referral', 'Results'
$subjBox.AutoCompleteSource = 'CustomSource'
$subjBox.AutoCompleteMode='SuggestAppend'
$subjBox.AutoCompleteCustomSource=$autocomplete
$subjBox.Items.AddRange($subjectList) # add to combo box
$subjBox.AutoCompleteCustomSource.AddRange($subjectList) # add to tab autocomplete
$subjBox.add_KeyDown($textBox_KeyDown)
$mainForm.Controls.Add($subjBox)

# Zip Name label, optional label and text box
$zipL = New-Object System.Windows.Forms.Label
$zipL.Text = "Zip Filename:"
$zipL.font = $templateFont
$zipL.Location = New-Object System.Drawing.Point(10, 130)
$zipL.AutoSize = $true
$mainForm.Controls.Add($zipL)

$zipBox = New-Object System.Windows.Forms.TextBox
$zipBox.Width = 300
$zipBox.Location  = New-Object System.Drawing.Point(155, 130)
$zipBox.AutoCompleteSource = 'CustomSource'
$zipBox.AutoCompleteMode='SuggestAppend'
$zipBox.add_KeyDown($textBox_KeyDown)
$mainForm.Controls.Add($zipBox)

# Warning label to notify if pt wants zip or unzipped
$checkNoZipL = New-Object System.Windows.Forms.Label
$checkNoZipL.Location = New-Object System.Drawing.Point(155, 150)
$checkNoZipL.Size = New-Object System.Drawing.Size(300, 15)
$checkNoZipL.Text = ""
$checkNoZipL.ForeColor = "red"
$checkNoZipL.Font = [System.Drawing.Font]::new("Tahoma", 9, [System.Drawing.FontStyle]::Bold)
$mainForm.Controls.Add($checkNoZipL)

# OK button
. ".\AddFileButton.ps1"

$mainForm.Controls.Add($AddFileButton)

$SendButton = New-Object System.Windows.Forms.Button
$SendButton.Location = New-Object System.Drawing.Size(380, 196)
$SendButton.Size = New-Object System.Drawing.Size(75, 23)
$SendButton.Text = "Send email"
$SendButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#99f0b2")
$SendButton.Enabled = $false
$SendButton.Add_Click({
    if (-Not $UnprotectedCheckbox.checked) {
        $isValidDob, $dobErrorMessage = Test-Dob -dob $dobBox.Text
        if (-Not $isValidDob) {
            $dobErrorL.text = $dobErrorMessage
            return
        }
    }
    Send-Email
})
$mainForm.Controls.Add($SendButton)

$ResetFilesButton = New-Object System.Windows.Forms.Button
$ResetFilesButton.Location = New-Object System.Drawing.Size(301, 170)
$ResetFilesButton.Size = New-Object System.Drawing.Size(75, 23)
$ResetFilesButton.Text = "Reset Files"
$ResetFilesButton.Enabled = $false
$ResetFilesButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#99c1f0")
$ResetFilesButton.Add_Click({
    $global:zipFileNames = @()
    $filesAddedL.Text = "Files added: 0"
    $filesListedL.Text = ""
    $ResetFilesButton.Enabled = $false
    $SendButton.Enabled = $false
})
$mainForm.Controls.Add($ResetFilesButton)

$filesAddedL = New-Object System.Windows.Forms.Label
$filesAddedL.Text = "Files added: 0"
$filesAddedL.Location = New-Object System.Drawing.Point(301, 200)
$filesAddedL.AutoSize = $true
$mainForm.Controls.Add($filesAddedL)

$filesListedL = New-Object System.Windows.Forms.Label
$filesListedL.Text = ""
$filesListedL.Location = New-Object System.Drawing.Point(10, 200)
$filesListedL.AutoSize = $true
$mainForm.Controls.Add($filesListedL)

# Clear Form button
$ClearButton = New-Object System.Windows.Forms.Button
$ClearButton.Location = New-Object System.Drawing.Size(155, 170)
$ClearButton.Size = New-Object System.Drawing.Size(75, 23)
$ClearButton.Text = "Clear form"
$ClearButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#e87c86")
$ClearButton.Add_Click({
    Clear-Form
})
$mainForm.Controls.Add($ClearButton)

# Checkbox to send files unprotected
$UnprotectedCheckbox = New-Object System.Windows.Forms.CheckBox
$UnprotectedCheckbox.Location = New-Object System.Drawing.Size(240, 172)
$UnprotectedCheckbox.Text = "No Zip"
$UnprotectedCheckbox.AutoSize = $true

$global:prevDob = ""
$global:prevZipName = ""

$UnprotectedCheckbox.Add_CheckStateChanged({
    if($UnprotectedCheckbox.checked) {
        $dobErrorL.Text = ""
        $global:prevDob = $dobBox.Text
        $dobBox.Text = ""
        $dobBox.ReadOnly = $true
        $dobBox.Enabled = $false

        $global:prevZipName = $zipBox.Text
        $zipBox.Text = ""
        $zipBox.ReadOnly = $true
        $zipBox.Enabled = $false
    } else {
        $dobBox.Text = $global:prevDob
        $dobBox.ReadOnly = $false
        $dobBox.Enabled = $true

        $zipBox.Text = $global:prevZipName
        $zipBox.ReadOnly = $false
        $zipBox.Enabled = $true
    }
})
$mainForm.Controls.Add($UnprotectedCheckbox)

$autofillCheckbox = New-Object System.Windows.Forms.CheckBox
$autofillCheckbox.Location = New-Object System.Drawing.Size(90, 172)
$autofillCheckbox.Text = "Autofill?"
$autofillCheckbox.Checked = $true
$autofillCheckbox.AutoSize = $true
$mainForm.Controls.Add($autofillCheckbox)

. ".\SettingsPopup.ps1"
$mainForm.Controls.Add($settingsButton)

# Timers
$autofillTimer = New-Object System.Windows.Forms.Timer
$autofillTimer.Interval = 500
$autofillTimer.add_tick({
    if ($autofillCheckbox.Checked) {
        Autofill-Form
    }
})
$autofillTimer.Start()

$shortcutTimer = New-Object System.Windows.Forms.Timer
$shortcutTimer.Interval = 100
$shortcutTimer.add_tick({

    foreach ($file in  Get-ChildItem -Path $global:defaultDir) {
        $fileName = (Split-Path $file -Leaf)
        if ($fileName -match "\bcoc\b") {
            $fileName = $fileName.Replace("coc", "Certificate of Capacity")
            Rename-Item $file $fileName
        } elseif ($fileName -match "\bir\b") {
            $fileName = $fileName.Replace("ir", "Imaging Request")
            Rename-Item $file $fileName
        } elseif ($fileName -match "\bmc\b") {
            $fileName = $fileName.Replace("mc", "Medical Certificate")
            Rename-Item $file $fileName
        } elseif ($fileName -match "\bpr\b") {
            $fileName = $fileName.Replace("pr", "Pathology Request")
            Rename-Item $file $fileName
        } elseif ($fileName -match "\bref\b") {
            $fileName = $fileName.Replace("ref", "Referral")
            Rename-Item $file $fileName
        } elseif ($fileName -match "\bres\b") {
            $fileName = $fileName.Replace("res", "Results")
            Rename-Item $file $fileName
        }
    }

    <#
    foreach ($file in  Get-ChildItem -Path $global:defaultDir) {
        switch ((Split-Path $file -Leaf).ToLower()) {
            "coc.pdf"  {Rename-Item $file "Certificate of Capacity.pdf"}
            "ir.pdf"   {Rename-Item $file "Imaging Request.pdf"}
            "mc.pdf"   {Rename-Item $file "Medical Certificate.pdf"}
            "pr.pdf"   {Rename-Item $file "Pathology Request.pdf"}
            "ref.pdf"  {Rename-Item $file "Referral.pdf"}
            "res.pdf"  {Rename-Item $file "Results.pdf"}
        }
    }#>
})
$shortcutTimer.Start()

#DOES NOT WORK because threads
#. "C:\Users\recep\OneDrive\Desktop\Registration Forms\Old Registration Form\Email App\OCRFileRenamer.ps1"

## TODO: CREATE A MAIN FUNCTION, MAINLY TO GROUP SETGLOBALDEFAULT DIR AND START TIMERS

# Make powershell console disappear
$windowcode = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
$asyncwindow = Add-Type -MemberDefinition $windowcode -name Win32ShowWindowAsync -namespace Win32Functions -PassThru
$null = $asyncwindow::ShowWindowAsync((Get-Process -PID $pid).MainWindowHandle, 0)

Set-Default-Directory-Global
Set-Location $global:defaultDir

$mainForm.ShowDialog()
