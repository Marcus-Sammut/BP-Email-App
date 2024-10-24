$AddFileButton = New-Object System.Windows.Forms.Button
$AddFileButton.Location = New-Object System.Drawing.Size(380, 170)
$AddFileButton.Size = New-Object System.Drawing.Size(75, 23)
$AddFileButton.Text = "Add File"
$AddFileButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#edf099")
$AddFileButton.Add_Click({
    if (-Not $UnprotectedCheckbox.checked) {
        $isValidDob, $dobErrorMessage = Test-Dob -dob $dobBox.Text
        if (-Not $isValidDob) {
            $dobErrorL.text = $dobErrorMessage
            return
        }
    }
    Add-FileToZip
})