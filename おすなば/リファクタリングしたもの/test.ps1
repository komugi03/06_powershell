# # �Ζ��n��񃊃X�g�������Ă���e�L�X�g
# $infoTextFileName = ".\resources\�c�[���p����.txt"
# $infoTextFileFullpath = "$PWD\$infoTextFileName"

# # ###################
# $infoTextFileFullpath

# $registeredWorkPlaceArray = @()

# # �Ζ��n��񃊃X�g�e�L�X�g�����݂����Ƃ��̏���
# if (Test-Path $infoTextFileFullpath) {
    
#     # �Ζ��n��񃊃X�g�e�L�X�g�̓��e���擾
#     $argumentText = (Get-Content $infoTextFileFullpath)
    
#     # # �u�Ζ����e�v���̕�����Ƀ}�b�`�����Ζ��n�̏����A���X�g����擾
#     # # �z��̒��g�@[0]:�K�p�@[1]:��ԁ@[2]:��ʋ@�ց@[3]:���z
#     for ($i = 0; $i -lt $argumentText.Length; $i++) {
#         # $argumentText[$i]
#         $argumentText[$i] -Match "(?<workplace>.+?)_" | Out-Null
#         # workPlaceArray�̓��e�����Ȃ��悤�ɂ���
#         # $registeredWorkPlaceArray
#         # $Matches.workplace
#         # $registeredWorkPlaceArray.Contains($Matches.workplace)
#         if (!$registeredWorkPlaceArray.Contains($Matches.workplace)) {
#             $registeredWorkPlaceArray += $Matches.workplace
#         }
#     }
#     $registeredWorkPlaceArray
# }


# $array = @("a","W","asdf")
# $array


# # �I�Icontinue������������I�I
# $OK = $true

# for($i = 0; $i < 3; $i++){
#     if($OK){
#         write-host "OK"
#     }
# }

# $wsobj = new-object -comobject wscript.shell
# $wsobj.popup("�������ł�`r`n���΂炭���҂���������", 3, "�����ݒ�", 0)



Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Data Entry Form'
$form.Size = New-Object System.Drawing.Size(300,200)
$form.StartPosition = 'CenterScreen'

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(75,120)
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(150,120)
$cancelButton.Size = New-Object System.Drawing.Size(75,23)
$cancelButton.Text = 'Cancel'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = 'Please enter the information in the space below:'
$form.Controls.Add($label)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10,40)
$textBox.Size = New-Object System.Drawing.Size(260,20)
$textBox.ReadOnly = $true
$textBox.Text = "wawa"
$textBox.BackColor = "#EEEEEE"
$textBox.BorderStyle = 0
$form.Controls.Add($textBox)

$form.Topmost = $true

$form.Add_Shown({$textBox.Select()})
$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $x = $textBox.Text
    $x
}

