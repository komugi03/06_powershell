# 
# �Ζ��\�����Ƃɏ�����ʔ�������쐬����Powershell
# 
# �Ζ��\�̃t�@�C�����F<3���̎Ј��ԍ�>_�Ζ��\_m��_<����>.xlsx
# 

# ----------------- �֐���` ---------------------

# �Ζ��\�Ə�����ۑ������ɕ��āAExcel�𒆒f����֐�
function breakExcel {
    # Book�����
    $kinmuhyouBook.close()
    $koguchiBook.close()
    # �g�p���Ă����v���Z�X�̉��
    $excel = $null
    $kinmuhyouBook = $null
    $kinmuhyouSheet = $null
    $koguchiBook = $null
    $koguchiSheet = $null
    $koguchiCell = $null
    # �K�x�[�W�R���N�g
    [GC]::Collect()
    # �������I������
    exit
}

# �V���[�v���g�������b�Z�[�W�̕\��������֐�
# �ő啶��������ɃV���[�v�̒��������肷��
# ����1 : �����F
# ����2�ȍ~ : ���b�Z�[�W
function displayMessagesSurroundedBySharp {
    # �ϐ��̏�����
    $maxLengths = 0
    for ($i = 1; $i -lt $Args.length; $i++) {
        # ���b�Z�[�W�̒��ň�Ԓ������������擾����
        if ( $maxLengths -lt $Args[$i].length) {
            $maxLengths = $Args[$i].length
        }
    }
    # ���b�Z�[�W�̕\��
    Write-Host ("`r`n" + '#' * ($maxLengths * 2 + 6) + "`r`n") -ForegroundColor $Args[0]
    for ($i = 1; $i -lt $Args.length; $i++) {
        Write-Host ('�@�@' + $Args[$i] + "�@�@`r`n") -ForegroundColor $Args[0]
    }
    Write-Host ('#' * ($maxLengths * 2 + 6) + "`r`n") -ForegroundColor $Args[0]
}

# �����̋󔒂������t�@�C�����Ƃ��Ďg���Ȃ������������֐�
# fileName : �t�@�C����
function removeInvalidFileNameChars ($fileName) {
    $fileNameRemovedSpace = $fileName -replace "�@", ""�@-replace " ", ""
    $invalidChars = [IO.Path]::GetInvalidFileNameChars() -join ''
    $regex = "[{0}]" -f [RegEx]::Escape($invalidChars)
    return $fileNameRemovedSpace -replace $regex
}

# -------------------- �又�� --------------------------

##### ���ӏ�����\���B���Ȃ��ꍇ�ɂ�Enter����������B#####

# ���݂̔N�������擾����
$thisYear = (Get-Date).Year
$thisMonth = (Get-Date).Month
$today = (Get-Date).Day

# ���ݓ�������쐬����ׂ��Ζ��\�̌����𔻒�
# 24���܂ł͓����������
if ($today -le 24) {
    $month = $thisMonth - 1
}
else {
    $month = $thisMonth
}

# ---------------�A�Z���u���̓ǂݍ���---------------
Add-Type -Assembly System.Windows.Forms
Add-Type -AssemblyName System.Drawing
# # INPUT�̂��߂ɕK�v
# [void][System.Reflection.Assembly]::Load("Microsoft.VisualBasic, Version=8.0.0.0, Culture=Neutral, PublicKeyToken=b03f5f7f11d50a3a")

$yesNo_yearMonthAreCorrect = [System.Windows.Forms.MessageBox]::Show("�쐬����̂� �y $thisYear �N $month �� �z�̏����ł�낵���ł����H",'�쐬���鏬���̑Ώ۔N��','YesNo','Question')

if($yesNo_yearMonthAreCorrect -eq 'No'){
    
    # �t�H���g�̎w��
    $Font = New-Object System.Drawing.Font("���C���I",8)

    # �t�H�[���S�̂̐ݒ�
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "�쐬���鏬���̑Ώ۔N��"
    $form.Size = New-Object System.Drawing.Size(300,200)
    $form.StartPosition = "CenterScreen"
    $form.font = $Font

    # ���x����\��
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,10)
    $label.Size = New-Object System.Drawing.Size(270,30)
    $label.Text = "�쐬���鏬���̔N����I�����Ă�������`r`n���O���`�������I���ł��܂���"
    $form.Controls.Add($label)

    # OK�{�^���̐ݒ�
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Point(40,100)
    $OKButton.Size = New-Object System.Drawing.Size(75,30)
    $OKButton.Text = "OK"
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $OKButton
    $form.Controls.Add($OKButton)

    # �L�����Z���{�^���̐ݒ�
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Point(130,100)

    $CancelButton.Size = New-Object System.Drawing.Size(75,30)
    $CancelButton.Text = "Cancel"
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $CancelButton
    $form.Controls.Add($CancelButton)

    # �R���{�{�b�N�X���쐬
    $Combo = New-Object System.Windows.Forms.Combobox
    $Combo.Location = New-Object System.Drawing.Point(50,50)
    $Combo.size = New-Object System.Drawing.Size(150,30)
    # ���X�g�ȊO�̓��͂������Ȃ�
    $Combo.DropDownStyle = "DropDownList"
    $Combo.FlatStyle = "standard"
    # $Combo.font = $Font
    $Combo.BackColor = "#005050"
    $Combo.ForeColor = "white"
        
    # -----------�R���{�{�b�N�X�ɍ��ڂ�ǉ�-----------
    for($counterForMove = (-6); $counterForMove -le 6; $counterForMove++){
        $date = get-date (get-date).AddMonths($counterForMove) -Format Y
        [void] $Combo.Items.Add("$date")
    }
    
    # �t�H�[���ɃR���{�{�b�N�X��ǉ�
    $form.Controls.Add($Combo)
    $Combo.SelectedIndex = 6
    
    # �t�H�[�����őO�ʂɕ\��
    $form.Topmost = $True
    
    # �t�H�[����\���{�I�����ʂ�ϐ��Ɋi�[
    $result = $form.ShowDialog()

    # �I����AOK�{�^���������ꂽ�ꍇ�A�I�����ڂ�\��
    if ($result -eq "OK"){
        $comboAnswer = $Combo.Text
        # �I�����ʂ̔N���擾
        # ���e�L�X�g����Ƃ遙
        $targetYear = $comboAnswer -split "�N"
        $targetYear

    }else{
    exit
    }

    Write-Output $comboAnswer



}