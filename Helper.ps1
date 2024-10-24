function Test-Dob {
    param (
        [string]$dob
    )

    $badFormatErrorMessage = "Incorrect format for DOB!"
    $invalidDobErrorMessage = $dob[0]+$dob[1]+'/'+$dob[2]+$dob[3]+'/'+$dob[4]+$dob[5] + " is not a valid DOB!"
    
    if ($dob -notmatch '^[0-9]{6}$') {
        return $false, $badFormatErrorMessage
    }

    $day = [int]$dob.SubString(0,2)
    $month = [int]$dob.SubString(2,2)
    if ($day -lt 1 -or $day -gt 31 -or $month -lt 1 -or $month -gt 12) {
        return $false, $invalidDobErrorMessage
    }
    if ($month -in @(4, 6, 9, 11)) {
        if ($day -gt 30) {
            return $false, $invalidDobErrorMessage
        }
    } elseif ($month -eq 2) {
        if ($day -gt 29) {
            return $false, $invalidDobErrorMessage
        }
    }
    $true
}

function Set-FileDialog-Title {
    $dob = $dobBox.Text
    $dobFmt = $dob[0]+$dob[1]+'/'+$dob[2]+$dob[3]+'/'+$dob[4]+$dob[5]
    $subj = (Get-Culture).TextInfo.ToTitleCase($subjBox.Text).trim()
    $zipName = $zipBox.Text
    $email = $emailBox.Text.trim()
    if ($email.length -gt 50) {
        $email = $email.substring(0, 49) + '...'
    }
    if ($UnprotectedCheckbox.checked) {
        return "Shortcuts: coc ir mc pr ref res        Email: $email        |        Subject: $subj"
    }
    return "Shortcuts: coc ir mc pr ref res        Email: $email        |        Zip name: $zipName        |        Subject: $subj        |        DOB: $dobFmt"
}

function Set-Default-Directory-Global {
    if (Test-Path $env:APPDATA\MSEmailForm\DefaultFolder.txt) {
        $default = Get-Content -Path $env:APPDATA\MSEmailForm\DefaultFolder.txt -TotalCount 1
    } else {
        $default = [Environment]::GetFolderPath("MyDocuments")
    }
    $global:defaultDir = $default
}

function Set-AppData-Default-Directory($initialDirectory=$env:USERPROFILE) {
    $folderName = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderName.Description = "Select a folder"
    $folderName.rootfolder = "MyComputer"
    $folderName.SelectedPath = $initialDirectory

    #https://learn.microsoft.com/en-us/dotnet/api/system.windows.forms.dialogresult
    $OKResult = 1
    if($folderName.ShowDialog() -ne $OKResult) {
        return
    }

    if (-Not (Test-Path $env:APPDATA\MSEmailForm)) {
        mkdir $env:APPDATA\MSEmailForm
    }
    
    Write-Output $folderName.SelectedPath > $env:APPDATA\MSEmailForm\DefaultFolder.txt
}