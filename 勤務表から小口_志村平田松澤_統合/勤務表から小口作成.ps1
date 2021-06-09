# 
# �Ζ��\�����Ƃɏ�����ʔ�������쐬����Powershell
# 
# �Ζ��\�̃t�@�C�����F<3���̎Ј��ԍ�>_�Ζ��\_m��_<����>.xlsx
# 

# ---------------�A�Z���u���̓ǂݍ���---------------
Add-Type -Assembly System.Windows.Forms
Add-Type -AssemblyName System.Drawing
# # INPUT�̂��߂ɕK�v?
# [void][System.Reflection.Assembly]::Load("Microsoft.VisualBasic, Version=8.0.0.0, Culture=Neutral, PublicKeyToken=b03f5f7f11d50a3a")


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

# �t�H�[���S�̂̐ݒ������֐�
# formText : �t�H�[���̖{���i������j
# formYoko : �t�H�[���̉���
# formTate : �t�H�[���̏c��
function makeForm ($formText, $formYoko, $formTate) {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $formText
    $form.Size = New-Object System.Drawing.Size($formYoko,$formTate)
    $form.StartPosition = "CenterScreen"
    $form.font = $Font
}

# ���x����\������֐�
# $labelText : ���x���ɏ������ޕ�����
# $form : �t�H�[���I�u�W�F�N�g
function makeLabel ($labelText, $form) {
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,10)
    $label.Size = New-Object System.Drawing.Size(270,30)
    $label.Text = $labelText
    $form.Controls.Add($label)
    return $form
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

# �|�b�v�A�b�v���쐬
$popup = new-object -comobject wscript.shell

# -------�i�ꏊ�������j---------------�����e���v�����擾------------------------
$koguchiTemplate = Get-ChildItem -Recurse -File | ? Name -Match "������ʔ�E�o������Z���׏�_�e���v��.xlsx"
# �Y�������t�@�C���̌��m�F
if ($koguchiTemplate.Count -lt 1) {
    # �|�b�v�A�b�v��\��
    $popup.popup("�Y�����鏬���t�@�C���̃e���v���[�g�����݂��܂���`r`n`r`n�_�E�����[�h�������Ă�������",0,"��蒼���Ă�������",48) | Out-Null
    exit
}
elseif ($koguchiTemplate.Count -gt 1) {
    # �|�b�v�A�b�v��\��
    $popup.popup("�Y�����鏬���t�@�C���̃e���v���[�g���������܂�`r`n`r`n�_�E�����[�h�������Ă�������",0,"��蒼���Ă�������",48) | Out-Null
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
    
    # �t�H�[�����쐬
    Write-Host "`r`n�Y������Ζ��\�t�@�C�������݂��܂���`r`n" -ForegroundColor Red
    exit
}
elseif ($kinmuhyou.Count -gt 1) {
    Write-Host "`r`n�Y������Ζ��\�t�@�C�����������܂�`r`n" -ForegroundColor Red
    exit
}

# �������n�߂�O�ɁA�t�@�C���̑��݃`�F�b�N�ƃt�@�C�����̃`�F�b�N���s��
if ( $kinmuhyou.Name -match "[0-9]{3}_�Ζ��\_([1-9]|1[12])��_.+\.xlsx" ) {
    Start-Sleep -milliSeconds 300

    try {
        # �Ζ��\�t�@�C���̃t���p�X�擾
        $kinmuhyouFullPath = $kinmuhyou.FullName 
    }
    catch [Exception] {
        # �Ζ��\�����݂��Ă��邩�`�F�b�N
        Write-Host ($targetMonth + "���̋Ζ��\�t�@�C�������݂��܂���B`r`n�_�E�����[�h���Ă�������`r`n") -ForegroundColor Red
        exit
    }

    displaySharpMessage "White" ([string]$targetMonth + " ���̏�����ʔ�������쐬���܂��B") "���΂炭���҂����������B"
}
else {
    # �Ζ��\�t�@�C���̃t�H�[�}�b�g���Ⴄ�ꍇ�͏C��������
    Write-Host " ######### <�Ј��ԍ�>_�Ζ��\_m��_<����>.xlsx �̌`���Ƀt�@�C�������C�����Ă������� #########`r`n" -ForegroundColor Red
    exit
}

# ----------------------Excel���N������--------------------------------
try {
    # �N������Excel�v���Z�X���擾
    $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
}
catch {
    $excel = New-Object -ComObject "Excel.Application" 
}

# Excel�����b�Z�[�W�_�C�A���O��\�����Ȃ��悤�ɂ���
$excel.DisplayAlerts = $false
$excel.visible = $false

# �Ζ��\�u�b�N���J��
$kinmuhyouBook = $excel.workbooks.open($kinmuhyouFullPath)
$kinmuhyouSheet = $kinmuhyouBook.sheets( "$targetMonth" + '��')

# �����u�b�N���J��
$koguchiBook = $excel.workbooks.open($koguchi)
$koguchiSheet = $koguchiBook.sheets(1)







# �Ō�́u�J���v�u�I���v��2��
# �J�����ł����������Ƃ���̃G�N�X�v���[���[��\������

# �Ζ��\����Ƃ��Ă���Ζ��n�̏��́u�Ζ����e�v�̗񂩂炾����OK

# �e�L�X�g�͑S���ǂݍ���ŁA�z��ɓ��ꂿ�Ⴄ
# �K���I������A�K�����ɂ����Ă���Ă�
# 1�s�ڂ̕i��A5�s�ڂ̂���ꂾ�������Ă���H
# �Ζ��\�̓��e�ƃ}�b�`���邩���؂��āA�}�b�`���Ă��珬���ɔz��̓��e���R�s�[����B
# �݂����ȁI

# �������̃_�C�A���O��\��������i�o�[�Ƃ��ł�Ƃ����ˁj

# �ŏI�I�ɁA�o�b�`�t�@�C���̌`�ɂ���i.bat�ɂ���j
# �o�b�`�t�@�C�����������Ă�powershell�ۂ���ʂ��o�Ȃ��悤�ɂ���B

# �u���̃e�L�X�g�쐬�o�b�`�Ŋe��Əꏊ�̏ڍאݒ� �� ���V�̃o�b�`�@���@
# ��README������      �ǂ������`���ɂ��邩�͖������B
# ���V���[�g�J�b�g�����    �o�b�`�t�@�C���̃V���[�g�J�b�g���쐬�B�ȒP�ɍ���̂ł���΍��Ȃ��B