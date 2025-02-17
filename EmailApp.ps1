﻿# C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -window hidden -ExecutionPolicy RemoteSigned -command "& 'C:\Users\recep\OneDrive\Desktop\Test\New folder\email.ps1'"
# TODO reorganise imports
# TODO either add more separate files, or just bring back the add files button into main file
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

function Clear-Files {
    $global:zipFileNames = @()
    $filesListedL.Text = ""
    $fileCountL.Text = "Files added: 0"
    $pageCountL.Text = "Total pages: 0"
    $global:zipPageCount = 0
    $SendButton.Enabled = $false
    $ResetFilesButton.Enabled = $false
    $autofillCheckbox.Checked = $true
}

function Clear-Form {
    Clear-Files
    $UnprotectedCheckbox.checked = $false
    $emailBox.Text = $dobBox.Text = $dobErrorL.Text = $subjBox.Text = $zipBox.Text = $checkNoZipL.Text = ""
    $emailBox.Focus()
}

#TODO move this to helper or own file
function Check-DetailsOCR {
    param($file)
    try {
        &magick convert -density 192 "$file[0]" -quality 100 -alpha remove $env:temp\magickoutput.png
        &tesseract $env:temp\magickoutput.png $env:temp\tesseractOCR
        # delete the output png
        $OCRtext = Get-Content $env:temp\tesseractOCR.txt
        $dob = $dobBox.Text.trim()
        $zipName = $zipbox.Text.trim().ToLower()
        $dob19 = $dob[0]+$dob[1]+'/'+$dob[2]+$dob[3]+'/'+'19'+$dob[4]+$dob[5]
        $dob20 = $dob[0]+$dob[1]+'/'+$dob[2]+$dob[3]+'/'+'20'+$dob[4]+$dob[5]
        $foundPatientDetails = $false
        foreach ($line in $OCRtext) {
            $line = $line.ToLower()
            if ($line.Contains($dob) -or $line.Contains($dob19) -or $line.Contains($dob20) -or $line.Contains($zipName)) {
                $foundPatientDetails = $true
                break
            }
        }
        Remove-Item $env:temp\magickoutput.png
        Remove-Item $env:temp\tesseractOCR.txt
        return $foundPatientDetails
    } catch {
        Remove-Item $env:temp\magickoutput.png
        Remove-Item $env:temp\tesseractOCR.txt
        return $false
    }
}


$global:zipFileNames = @()
$global:zipPageCount = 0
$cdfPath = $PWD.ToString() + "\cpdf.exe"

function Add-FileToZip {
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
        Title = Set-FileDialog-Title
        InitialDirectory = $global:defaultDir
        MultiSelect = $false
        Filter = 'PDF Files (*.pdf)|*.pdf|All Files (*.*)|*.*'
    }
    $FileBrowser.add_FileOk({
        param($s, $e)
        $filename = Split-Path $FileBrowser.FileName -leaf
        if ($filename -match "^Scan\d{4}.*\.pdf$") {
            $e.Cancel = $true
            [System.Windows.MessageBox]::Show("Please rename the selected file: {0}" -f $filename, "Error")
        } elseif ($global:zipFileNames -contains $filename) {
            $e.Cancel = $true
            [System.Windows.MessageBox]::Show("You have already selected this file: {0}" -f $filename, "Error")
        } elseif ((Check-DetailsOCR -file $FileBrowser.FileName) -match $false) {
            $e.Cancel = $true
            [System.Windows.MessageBox]::Show("Patient's DOB or Name was not found in this file: {0}. If you are sure this file belongs to the patient, add an underscore '_' to the end of the file name" -f $filename, "Error")
        }
    })
    
    # Don't autofill while fileDialog is open
    $autofillCheckbox.Checked = $false
    if ($FileBrowser.ShowDialog() -eq 1) { # sys.win.forms.dialogresult OK == 1
        $global:zipFileNames += $FileBrowser.FileName
        $filesListedL.Text += Split-Path $FileBrowser.FileName -leaf
        $fileCountL.Text = "Files added: {0}" -f $global:zipFileNames.Count
        $global:zipPageCount += (&$cdfPath -pages -gs-malformed $FileBrowser.FileName)
        $pageCountL.Text = "Total pages: {0}" -f $global:zipPageCount
        $ResetFilesButton.Enabled = $true
        $SendButton.Enabled = $true
    } elseif ($global:zipFileNames.Count -eq 0) {
        # re-enable autofill if no file was added
        $autofillCheckbox.Checked = $true
    }
}

function Send-Email {
    if (($global:zipFileNames | Test-Path) -notcontains $true) {
        [System.Windows.MessageBox]::Show("One or more of the selected files is now invalid, selected files have been reset." -f $filename, "Error")
        Clear-Files
        return
    }

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
    $autofillCheckbox.Checked = $true
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationCore,PresentationFramework # for messageBox

$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Text ='Create New Email - Press F3 to copy DOB'
$mainForm.Width = 460
$mainForm.Height = 280
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
        $AddFileButton.PerformClick()
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
    Clear-Files
})
$mainForm.Controls.Add($ResetFilesButton)


$fileCountL = New-Object System.Windows.Forms.Label
$fileCountL.Text = "Files added: 0"
$fileCountL.Location = New-Object System.Drawing.Point(298, 200)
$fileCountL.AutoSize = $true
$mainForm.Controls.Add($fileCountL)

$pageCountL = New-Object System.Windows.Forms.Label
$pageCountL.Text = "Total pages: 0"
$pageCountL.Location = New-Object System.Drawing.Point(295, 218)
$pageCountL.TextAlign = 64
$pageCountL.AutoSize = $true
$mainForm.Controls.Add($pageCountL)

$filesListedL = New-Object System.Windows.Forms.Label
$filesListedL.Text = ""
$filesListedL.Location = New-Object System.Drawing.Point(10, 200)
$filesListedL.MaximumSize = New-Object System.Drawing.Size(280, 38)
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
