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
    # �O�̌��������쐬�̑Ώی��Ƃ���
    $targetMonth = (Get-date).AddMonths(-1).month
}
else {
    # �����������쐬�̑Ώی��Ƃ���
    $targetMonth = $thisMonth
}

# ---------------�A�Z���u���̓ǂݍ���---------------
Add-Type -Assembly System.Windows.Forms
Add-Type -AssemblyName System.Drawing
# # INPUT�̂��߂ɕK�v
# [void][System.Reflection.Assembly]::Load("Microsoft.VisualBasic, Version=8.0.0.0, Culture=Neutral, PublicKeyToken=b03f5f7f11d50a3a")

# (���ݓ��ɂ���ĕς��̂ŁAget-date -Format Y �ɂ͂��Ă��Ȃ�)
$yesNo_yearMonthAreCorrect = [System.Windows.Forms.MessageBox]::Show("�쐬����̂� �y $thisYear �N $targetMonth �� �z�̏����ł�낵���ł����H`r`n`r`n�u�������v�ő��̌���I���ł��܂�",'�쐬���鏬���̑Ώ۔N��','YesNo','Question')

# ���N�������쐬�̑Ώ۔N�Ƃ���
$targetYear = $thisYear

# ��$yesNo_yearMonthAreCorrect -eq 'No'���[�v�J�n��
if($yesNo_yearMonthAreCorrect -eq 'No'){
    
    # �t�H���g�̎w��
    $Font = New-Object System.Drawing.Font("���C���I",8)

    # �t�H�[���S�̂̐ݒ�
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "�쐬���鏬���̑Ώ۔N��"
    $form.Size = New-Object System.Drawing.Size(265,200)
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
        # ���[�U�[�̉񓚂�"�N"�ŋ�؂�
        $comboAnswer = $Combo.Text -split "�N"

        # ���[�U�[�w��̔N�������쐬�̑Ώ۔N�Ƃ��ď㏑����
        $targetYear = $comboAnswer[0]

        # ���[�U�[�w��̌��������쐬�̑Ώی��Ƃ��ď㏑������
        $targetMonth = $comboAnswer[1] -split "��"

    }else{
        # �������I������
        exit
    }

# ��$yesNo_yearMonthAreCorrect -eq 'No'���[�v�I����
}


echo "$targetYear �N��"
echo "$targetMonth ���̏������쐬���܂�"



# -------�i�ꏊ�������j---------------�����e���v�����擾------------------------
$koguchiTemplate = Get-ChildItem -Recurse -File | ? Name -Match "������ʔ�E�o������Z���׏�_�e���v��.xlsx"
# �Y�������t�@�C���̌��m�F
if ($koguchiTemplate.Count -lt 1) {
    Write-Host "`r`n�Y�����鏬���t�@�C�������݂��܂���`r`n`r`n�_�E�����[�h�������Ă�������`r`n" -ForegroundColor Red
    exit
}
elseif ($koguchiTemplate.Count -gt 1) {
    Write-Host "`r`n�Y�����鏬���t�@�C�����������܂�`r`n`r`n�_�E�����[�h�������Ă�������`r`n" -ForegroundColor Red
    exit
}

# ------�i���[�U�[�w��̌����K�v������A�R���{�{�b�N�X����j----------�e���v���[�g���珬����ʔ�������쐬����---------------------
# �쐬�����������i�[����t�H���_�ɁA�e���v���[�g���R�s�[����
# ���t�H���_�����݂��Ă��Ȃ��ƃG���[���o��
$koguchi = Join-Path $PWD "�쐬����������ʔ����" | Join-Path -ChildPath "������ʔ�E�o������Z���׏�_�R�s�[.xlsx"
Copy-Item -path $koguchiTemplate.FullName -Destination $koguchi

# �Ζ��\�t�@�C�����擾
$kinmuhyou = Get-ChildItem -Recurse -File | ? Name -Match "[0-9]{3}_�Ζ��\_($targetMonth)��_.+"

# �Y���Ζ��\�t�@�C���̌��m�F
if ($kinmuhyou.Count -lt 1) {
    Write-Host "`r`n�Y������Ζ��\�t�@�C�������݂��܂���`r`n" -ForegroundColor Red
    exit
}
elseif ($kinmuhyou.Count -gt 1) {
    Write-Host "`r`n�Y������Ζ��\�t�@�C�����������܂�`r`n" -ForegroundColor Red
    exit
}










# �Ζ��\����Ƃ��Ă���Ζ��n�̏��́u�Ζ����e�v�̗񂩂炾����OK