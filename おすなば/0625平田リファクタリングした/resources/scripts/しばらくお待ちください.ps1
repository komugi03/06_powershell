#
# �L�����΂炭���҂����������t�H�[����񋟂���
#

write-host "���΂炭���҂���������.ps1���Ăяo����܂���"

function pleaseWait($catPath) {
    .{
# �t�H���g�̎w��
$font = New-Object System.Drawing.Font("���C���I", 8)

# �t�H�[���̐ݒ�
$waitForm = New-Object System.Windows.Forms.Form
$waitForm.Text = "�����ݒ�"
$waitForm.Size = New-Object System.Drawing.Size(265, 170)
$waitForm.StartPosition = "CenterScreen"
$waitForm.font = $font
# �t�H�[�����őO�ʂɕ\��
$waitForm.Topmost = $True

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(70, 30)
$label.Size = New-Object System.Drawing.Size(270, 30)
$label.Text = "�������ł�`r`n���΂炭���҂���������"
$label.font = $font
$waitForm.Controls.Add($label)

###### �摜�̃t���p�X�ύX���Ăق����ł��� #########
#PictureBox
$pic = New-Object System.Windows.Forms.PictureBox
$pic.Size = New-Object System.Drawing.Size(50, 50)
# �������L�摜�̃t���p�X
$catFullPath = join-path -path $PWD.path -childpath $catPath # "resources/pictures/���҂����������L.png"
$pic.Image = [System.Drawing.Image]::FromFile($catFullPath)
$pic.Location = New-Object System.Drawing.Point(20,20) 
$pic.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
$waitForm.Controls.Add($pic)


return

} | Out-Null


# �t�H�[�������^�[������
return $waitForm

}