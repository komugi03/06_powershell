# PowerShell�Ń��[�U�[�t�H�[�������@- �T�u�t�H�[���� -

Add-Type -AssemblyName System.Windows.Forms

# ���C���t�H�[��
$formA = New-Object System.Windows.Forms.Form
$formA.Size = "200,200"
$formA.StartPosition = "Manual"
$formA.Location = "0,0"
$formA.text = "���C���t�H�[��"
$formA.MinimizeBox = $False
$formA.MaximizeBox = $False

#�t�H�[���\���{�^��
$Button = New-Object System.Windows.Forms.Button
$Button.Location = "50,50"
$Button.size = "100,30"
$Button.text  =�@"�t�H�[���\��"
$formA.Controls.Add($Button)

# �T�u�t�H�[��
$formB = New-Object System.Windows.Forms.Form
$subForm.Size = "300,200"
$subForm.StartPosition = "CenterScreen"
$subForm.MaximizeBox = $False
$subForm.MinimizeBox = $false
$subForm.text = "�o�^�ς݂̋Ζ��n����I��"
$formB.Owner = $formA
$subForm.DialogResult = [System.Windows.Forms.DialogResult]::No


# �T�u�t�H�[���̃N���[�W���O�C�x���g
$Close = {
    $_.Cancel = $True
    $formB.Visible = $false
}
$formB.Add_Closing($Close)

# �t�H�[���\���{�^���̃N���b�N�C�x���g
$Click = {
    $formB.Show()
}
$Button.Add_Click($Click)

#����{�^��
$CloseButton = New-Object System.Windows.Forms.Button
$CloseButton.Location = "50,100"
$CloseButton.size = "80,30"
$CloseButton.text  =�@"����"
$CloseButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$formA.Controls.Add($CloseButton)

$FormA.Showdialog()