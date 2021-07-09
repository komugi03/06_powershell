# PowerShellでユーザーフォームを作る　- サブフォーム編 -

Add-Type -AssemblyName System.Windows.Forms

# メインフォーム
$formA = New-Object System.Windows.Forms.Form
$formA.Size = "200,200"
$formA.StartPosition = "Manual"
$formA.Location = "0,0"
$formA.text = "メインフォーム"
$formA.MinimizeBox = $False
$formA.MaximizeBox = $False

#フォーム表示ボタン
$Button = New-Object System.Windows.Forms.Button
$Button.Location = "50,50"
$Button.size = "100,30"
$Button.text  =　"フォーム表示"
$formA.Controls.Add($Button)

# サブフォーム
$formB = New-Object System.Windows.Forms.Form
$subForm.Size = "300,200"
$subForm.StartPosition = "CenterScreen"
$subForm.MaximizeBox = $False
$subForm.MinimizeBox = $false
$subForm.text = "登録済みの勤務地から選択"
$formB.Owner = $formA
$subForm.DialogResult = [System.Windows.Forms.DialogResult]::No


# サブフォームのクロージングイベント
$Close = {
    $_.Cancel = $True
    $formB.Visible = $false
}
$formB.Add_Closing($Close)

# フォーム表示ボタンのクリックイベント
$Click = {
    $formB.Show()
}
$Button.Add_Click($Click)

#閉じるボタン
$CloseButton = New-Object System.Windows.Forms.Button
$CloseButton.Location = "50,100"
$CloseButton.size = "80,30"
$CloseButton.text  =　"閉じる"
$CloseButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$formA.Controls.Add($CloseButton)

$FormA.Showdialog()