# $wsobj = new-object -comobject wscript.shell
# $result = $wsobj.popup("Hello!Project",0,"�����Ƃ�",0)

# function add ($x, $y) {
#     $result = $x + $y
#     $result    
# }

# add 1 2


# $koguchi = Join-Path $PWD "�쐬����������ʔ����" | Join-Path -ChildPath "������ʔ�E�o������Z���׏�_�R�s�[.xlsx"
# Remove-Item -Path $koguchi


# if(Test-Path $PWD"\�쐬����������ʔ����"){
    #     echo "OK"
    # }
    # else{
        #     New-Item -Path $PWD"\�쐬����������ʔ����" -ItemType Directory | Out-Null
        # }
        
        # $i = 1
        # $i
        # # $koutsukikan1.text�ɂ�����
        # $koi = ('$koutsukikan' + $i)
        # $koi
        
        # $targetMonth = (Get-Date).Month
        # $fileNameMonth = "$targetMonth ��"
        # $kinmuhyou = Get-ChildItem -Recurse -File | ? Name -Match "[0-9]{3}_�Ζ��\_$fileNameMonth_.+"
        # $kinmuhyou
        
        
        # $workPlace = "�����"
        # if ($workPlace -ne "" -and $workPlace -ne '�ݑ�') {
        #     echo "OkKoguchi"
        # }else{
        #     echo "NoKoguchi"
        # }

# if(Test-Path $PWD"\�c�[���p����.txt"){
#     $argumentText = (Get-Content $PWD"\�c�[���p����.txt")[0..3]
#     $argumentText[0]
# }else{
#     Write-Output "�t�@�C���͂���܂���"
# }

# $workPlace = '��'

# if(Test-Path $PWD"\�c�[���p����.txt"){
#     $argumentText = (Get-Content $PWD"\�c�[���p����.txt")
#     # �Ζ��n�̏������X�g����擾 ( �z��̒��g�@[0]:�K�p�@[1]:��ԁ@[2]:��ʋ@�ց@[3]:���z )
#     $workPlaceInfo = $argumentText | Select-String -Pattern ($workPlace + '_')
#     if($workPlaceInfo -eq $null){
    
#         echo "catch!"
#     }

#     # $workPlaceInfo

#     # $workPlaceInfo[0]

# }else{
#     Write-Output "�t�@�C���͂���܂���"
# }

# $infomationTextFileName = "�c�[���p����.txt"
# Test-Path $PWD"\"$infomationTextFileName

# $infoTextFileName = "�c�[���p����.txt"
# $infoTextFileFullpath = "$PWD\$infoTextFileName"
# $infoTextFileFullpath
        
# if(Test-Path $infoTextFileFullpath){echo "OK"}


# $infoTextFileName = "�c�[���p����.txt"
# $infoTextFileFullpath = "$PWD\$infoTextFileName"
# $argumentText = (Get-Content $infoTextFileFullpath)

# $workPlace = "�����"
# $workPlaceInfo = $argumentText | Select-String -Pattern ($workPlace + '_')
# Write-Host ("�Ζ��nlist�F" + $workPlaceInfo)

# Write-Host ("�K�p�̍s�F" + $workPlaceInfo[0])


# $tekiyouText = [String]$workPlaceInfo[0]
# $tekiyouText

# $tekiyouText = $tekiyouText.Substring(4, $tekiyouText.Length - 4)
# Write-Host ("�K�p�F" + $tekiyouText)

# $tekiyouText = ([String]$workPlaceInfo[0]).Substring(4, ([String]$workPlaceInfo[0]).Length - 4)
# Write-Host ("�K�p�F" + $tekiyouText)


# �v���O���X�o�[
Add-Type -AssemblyName System.Windows.Forms

$formProgressBar = New-Object System.Windows.Forms.Form
$formProgressBar.Size = "300,200"
$formProgressBar.Startposition = "CenterScreen"
$formProgressBar.Text = "�쐬���c"

# $Button = New-Object System.Windows.Forms.Button
# $Button.Location = "110,20"
# $Button.Size = "80,30"
# $Button.Text = "�J�n"

# # �{�^���̃N���b�N�C�x���g
# $Start = {
#     # For ( $i = 0 ; $i -lt 10 ; $i++ )
#     # {
#     #     $progressBar.Value = $i+1
#     #     start-sleep -s 1
#     # }
#     [System.Windows.Forms.MessageBox]::Show("���҂������܂����I�����ł��I", "info")
# }
# $Button.Add_Click($Start)

# # �v���O���X�o�[
# $progressBar = New-Object System.Windows.Forms.ProgressBar
# $progressBar.Location = "10,100"
# $progressBar.Size = "260,30"
# $progressBar.Maximum = "10"
# $progressBar.Minimum = "0"
# $progressBar.Style = "Continuous"

# $progressBar.Value = 1


# # $formProgressBar.Controls.AddRange(@($progressBar,$Button))
# $formProgressBar.Controls.AddRange($progressBar)

# $formProgressBar.Topmost = $True

# $formProgressBar.Show()


# $progressBar.Value++

# $progressBar.Value++

# $formProgressBar.Show()
# # Start-Sleep -milliSeconds 300
# # $formProgressBar.Close()
# $formProgressBar
# $progressBar.Value++


# $progressBar.Value++
# $formProgressBar.Show()

# Start-Sleep -milliSeconds 500
# $progressBar.Value += 5

# # $formProgressBar.Visible = $false
# $finish = $formProgressBar.Show()
# $formProgressBar.Close()
# write-host $finish

# Write-Progress -Activity "������" -Status "���݂̏��" -PercentComplete 20 -SecondsRemaining 5



# # �v���O���X�o�[�@�T���v��

# For($b = 1 ; $b -le 10000 ; $b++)
# {
    # $c = $b / 100
# write-progress -Activity "���l�����Z���Ă��܂�" -Status "���΂炭���҂���������"�@-PercentComplete $c -CurrentOperation "$c % ����"
# }



# Function Invoke-ProgressBar {
#     Param([Int]$Minimum, [Int]$Maximum, [ScriptBlock]$ScriptBlock)

#     Add-Type -AssemblyName System.Windows.Forms

#     $formProgressBar = New-Object System.Windows.Forms.Form
#     #$formProgressBar.Text = 'Progress Bar'
#     $formProgressBar.Width = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Width / 2
#     $formProgressBar.Height = $formProgressBar.Width / 5
#     $formProgressBar.StartPosition = 'CenterScreen'
#     $formProgressBar.TopMost = $true

#     $progressBar = New-Object System.Windows.Forms.ProgressBar
#     $progressBar.Width = $formProgressBar.ClientRectangle.Width * 0.9
#     $progressBar.Height = $formProgressBar.ClientRectangle.Height * 0.3
#     $progressBar.Left = ($formProgressBar.ClientRectangle.Width - $progressBar.Width) / 2
#     $progressBar.Top = ($formProgressBar.ClientRectangle.Height - $progressBar.Height) / 2
#     $progressBar.Visible = $true
#     $progressBar.Style = 'Continuous'
#     $progressBar.Minimum = $Minimum
#     $progressBar.Maximum = $Maximum
#     $progressBar.Value = $Minimum
#     $progressBar.Step = 1
#     $formProgressBar.Controls.Add($progressBar)

#     $formProgressBar.Add_Shown({
#         For ($i = $Minimum; $i -le $Maximum; $i++) {
#             & $ScriptBlock
#             $formProgressBar.Text = ($progressBar.Value / ($Maximum - $Minimum)).ToString('0%')
#             $progressBar.PerformStep()
#         }
#         $formProgressBar.Close()
#     })
#     [void]$formProgressBar.ShowDialog()
# }

# Invoke-ProgressBar -Minimum 0 -Maximum 42 -ScriptBlock {
#     # Some time-comsuming task ...
#     Start-Sleep -Milliseconds 100
#     Write-Host $i
# }


# $koutsu = '�����_���c�}��`r`nJR�R���`r`n��񂩂���'
# $koutsukikanText = $koutsu.Substring(4, $koutsu.length - 4)
# $koutsukikanText

# $koutsukikanArray = $koutsukikanText -split '`r`n'

# $koutsukikanArray.Length

# for ($i = 0; $i -lt $koutsukikanArray.Length; $i++) {
#     # ���s�R�[�h�𑫂�
#     $koutsukikanKaigyou += $koutsukikanArray[$i] + "`r`n"
# }

# # �Ō�̉��s���폜����
# $koutsukikanKaigyou = $koutsukikanKaigyou.Substring(0, $koutsukikanKaigyou.Length - 1)
# $koutsukikanKaigyou




# # --------�V���������t�@�C������p��---------
# # <�Ј��ԍ�>_������ʔ�E�o������Z���׏�_YYYYMM_<����>
# # '116_������ʔ�E�o������Z���׏�_202105_�u����.xlsx'
# $koguchiNewFileName = '116' + "_������ʔ�E�o������Z���׏�_" + '202105' + "_" + '�u����'
# # �t�@�C�����Ɏg���Ȃ������������Ă�����폜����(�����̊Ԃ̋󔒂Ȃ�)
# # $koguchiNewFileName = remove-invalidFileNameChars $koguchiNewFileName
# # �V���������t�@�C���̃t���p�X
# $koguchiNewfullPath = Join-Path $PWD "�쐬����������ʔ����" | Join-Path -ChildPath $koguchiNewFileName

# $koguchiNewfullPath

# # if (Test-Path ($koguchiNewfullPath + '_' + '[0-9][0-9]' + '.xlsx')) {
# #     echo "OK"
# # }

# Test-Path ($koguchiNewfullPath + "_.+" + '.xlsx')



# $targetMonth = "2"
# $fileNameMonth = "{0:D2}" -f [int]$targetMonth
# $fileNameMonth



# # ---------------�A�Z���u���̓ǂݍ���---------------
# Add-Type -AssemblyName System.Windows.Forms
# Add-Type -AssemblyName System.Drawing

# # �t�H�[���S�̂̐ݒ������֐�
# # formText : �t�H�[���̖{���i������j
# # formYoko : �t�H�[���̉���
# # formTate : �t�H�[���̏c��
# function makeForm ($formText, $formYoko, $formTate) {
#     $form = New-Object System.Windows.Forms.Form
#     $form.Text = $formText
#     $form.Size = New-Object System.Drawing.Size($formYoko,$formTate)
#     $form.StartPosition = "CenterScreen"
#     $form.font = $Font
#     return $form
# }

# # ���x����\������֐�
# # $labelText : ���x���ɏ������ޕ�����
# # $form : �t�H�[���I�u�W�F�N�g
# function makeLabel ($labelText, $form) {
#     $label = New-Object System.Windows.Forms.Label
#     $label.Location = New-Object System.Drawing.Point(10,10)
#     $label.Size = New-Object System.Drawing.Size(270,30)
#     $label.Text = $labelText
#     $form.Controls.Add($label)
#     return $form
# }


# # �t�H�[���S�̂̐ݒ�
# $form = makeForm "�쐬���鏬���̑Ώ۔N��" 265 200

# # ���x����\��
# $label = makeLabel "�쐬�����������̔N����I�����Ă�������" $form

# # OK�{�^���̐ݒ�
# $OKButton = New-Object System.Windows.Forms.Button
# $OKButton.Location = New-Object System.Drawing.Point(40,100)
# $OKButton.Size = New-Object System.Drawing.Size(75,30)
# $OKButton.Text = "OK"
# $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
# $form.AcceptButton = $OKButton
# $form.Controls.Add($OKButton)

# # �L�����Z���{�^���̐ݒ�
# $CancelButton = New-Object System.Windows.Forms.Button
# $CancelButton.Location = New-Object System.Drawing.Point(130,100)
# $CancelButton.Size = New-Object System.Drawing.Size(75,30)
# $CancelButton.Text = "Cancel"
# $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
# $form.CancelButton = $CancelButton
# $form.Controls.Add($CancelButton)

# # �t�H�[�����őO�ʂɕ\��
# $form.Topmost = $True

# # �t�H�[����\���{�I�����ʂ�ϐ��Ɋi�[
# $result = $form.ShowDialog()

# # �I����AOK�{�^���������ꂽ�ꍇ�A�I�����ڂ�\��
# if ($result -eq "OK"){
#     # ���[�U�[�̉񓚂�"�N"�ŋ�؂�
#     $Combo.Text -match "(?<year>.+?)�N(?<month>.+?)��" | out-null

#     # ���[�U�[�w��̔N�������쐬�̑Ώ۔N�Ƃ��ď㏑����
#     $targetYear = $Matches.year

#     # ���[�U�[�w��̌��������쐬�̑Ώی��Ƃ��ď㏑������
#     $targetMonth = $Matches.month

# }else{
#     # �������I������
#     exit
# }



# # �|�b�v�A�b�v���쐬
# $popup = new-object -comobject wscript.shell

# # ����ɏI�������Ƃ��|�b�v�A�b�v��\��
# $successEnd = $popup.popup("���҂������܂����I����ɏI�����܂���`r`n�d�オ����m�F���Ă�������",0,"����I��",64)     

# if($successEnd -eq '1'){
#     Start-Process $PWD"\�쐬����������ʔ����"
#     # Invoke-Item $PWD"\�쐬����������ʔ����"
# }


# $wsobj = new-object -comobject wscript.shell
# $result = $wsobj.popup("Hello!Project",0,"wa",64)
# $result


# �쐬���鏬���̔N���������Ă��邩�m�F����_�C�A���O��\��
# (���ݓ��ɂ���ĕς��̂ŁAget-date -Format Y �ɂ͂��Ă��Ȃ�)
# $yesNo_yearMonthAreCorrect = [System.Windows.Forms.MessageBox]::Show("�쐬����̂� �y $thisYear �N $targetMonth �� �z�̏����ł�낵���ł����H`r`n`r`n�u�������v�ő��̌���I���ł��܂�",'�쐬���鏬���̑Ώ۔N��','YesNo','Question')
# $yesNo_yearMonthAreCorrect

# [System.Drawing.FontFamily]::Families


# �|�b�v�A�b�v���쐬
$popup = new-object -comobject wscript.shell

$targetPersonName = "�u����"
# $targetPersonName = '���c ��'
# $targetPersonName = "���V�@�ĊC"

if($targetPersonName -match ' ' -or $targetPersonName -match '�@'){
    # $targetPersonName = $targetPersonName -replace '�@', '' -replace ' ', ''
    $targetPersonName = $targetPersonName.replace('�@', '  ')
    $targetPersonName = $targetPersonName.replace(' ', '')
}

$successEnd = $popup.popup($targetPersonName + "���� : )`r`nOK�������ĕs�����Ȃ����m�F���Ă�������",0,"���҂������܂����I",64)    

