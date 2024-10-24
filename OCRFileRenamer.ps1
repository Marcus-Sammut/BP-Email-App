#DOES NOT WORK because threads
<#
$ocrRenameTimer = New-Object System.Timers.Timer
$ocrRenameTimer.Interval = 200
$global:hashed = @()
$ocrRenameTick = {
    if (!(Test-Path $env:temp\full.uzn -PathType Leaf)) {
        $mainForm.Text = "MISSING full.uzn"
        $ocrRenameTimer.Stop()
    }
    
    
    if (-not $ocrCheckBox.Checked) {
        return
    }
    
    # todo change to arraylist for $currHashed
    # https://learn.microsoft.com/en-us/powershell/scripting/learn/deep-dives/everything-about-arrays?view=powershell-7.4
    foreach ($file in Get-ChildItem -Filter "*.pdf" -Path $global:defaultDir) {
        $hash = Get-FileHash $file | Select-Object -expand Hash | Out-String
        if ($global:hashed -match $hash) {
            $currHashed += $hash
            continue
        }
        $global:hashed += $hash
        &magick convert -density 192 $file -quality 100 -alpha remove $env:temp\full.png
        &tesseract $env:temp\full.png $env:temp\uznOut --psm 4
        $uznOCRtext = [System.IO.File]::ReadAllText("$env:temp\uznOut.txt").ToLower()
        #TODO change to switch statement
        if     ($uznOCRtext -match "pathology request")   {Rename-Item $file "Pathology Request.pdf"}
        elseif ($uznOCRtext -match "imaging request")     {Rename-Item $file "Imaging Request.pdf"}
        elseif ($uznOCRtext -match "medical certificate") {Rename-Item $file "Medical Certificate.pdf"}
        elseif ($uznOCRtext -match "255170bx")            {Rename-Item $file "Referral.pdf"}
        elseif ($uznOCRtext -match "abn 94 143 690 564")  {Rename-Item $file "Referral.pdf"}
    }
}
Register-ObjectEvent -InputObject $ocrRenameTimer -EventName Elapsed -Action $ocrRenameTick
$ocrRenameTimer.Start()
#>