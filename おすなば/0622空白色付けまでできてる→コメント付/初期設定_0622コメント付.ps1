# ---------------�A�Z���u���̓ǂݍ���---------------
Add-Type -Assembly System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# �t�H�[���S�̂̐ݒ������֐�
# formText : �t�H�[���̖{���i������j
# formYoko : �t�H�[���̉���
# formTate : �t�H�[���̏c��
function makeForm ($formText, $formYoko, $formTate) {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $formText
    $form.Size = New-Object System.Drawing.Size($formYoko, $formTate)
    $form.StartPosition = "CenterScreen"
    $form.font = $font
}

# ���x����\������֐�
# $labelText : ���x���ɏ������ޕ�����
# $form : �t�H�[���I�u�W�F�N�g
function makeLabel ($labelText, $form) {
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10, 10)
    $label.Size = New-Object System.Drawing.Size(270, 30)
    $label.Text = $labelText
    $form.Controls.Add($label)
    return $form
}

# �Ζ��\��ۑ������ɕ��āAExcel�𒆒f����֐�
function breakExcel {
    # Book�����
    $kinmuhyouBook.close()
    # �g�p���Ă����v���Z�X�̉��
    $excel = $null
    $kinmuhyouBook = $null
    $kinmuhyouSheet = $null
    # �K�x�[�W�R���N�g
    [GC]::Collect()
    # �������I������
    exit
}

# -------------------- �又���̏��� --------------------------

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
$yesNo_yearMonthAreCorrect = [System.Windows.Forms.MessageBox]::Show("�y $thisYear �N $targetMonth �� �z�̋Ζ��\�����Ƃɏ����ݒ�����܂����H`r`n`r`n�u�������v�ő��̌���I���ł��܂�", '�쐬���鏬���̑Ώ۔N��', 'YesNo', 'Question')

# ���N�������쐬�̑Ώ۔N�Ƃ���
$targetYear = $thisYear

# ��$yesNo_yearMonthAreCorrect -eq 'No'���[�v�J�n��
if ($yesNo_yearMonthAreCorrect -eq 'No') {
    
    # �t�H���g�̎w��
    $font = New-Object System.Drawing.Font("���C���I", 8)

    # �t�H�[���S�̂̐ݒ�
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "�쐬���鏬���̑Ώ۔N��"
    $form.Size = New-Object System.Drawing.Size(265, 200)
    $form.StartPosition = "CenterScreen"
    $form.font = $font

    # ���x����\��
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10, 10)
    $label.Size = New-Object System.Drawing.Size(270, 30)
    $label.Text = "�쐬�����������̔N����I�����Ă�������"
    $form.Controls.Add($label)

    # OK�{�^���̐ݒ�
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Point(40, 100)
    $OKButton.Size = New-Object System.Drawing.Size(75, 30)
    $OKButton.Text = "OK"
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $OKButton
    $form.Controls.Add($OKButton)

    # �L�����Z���{�^���̐ݒ�
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Point(130, 100)

    $CancelButton.Size = New-Object System.Drawing.Size(75, 30)
    $CancelButton.Text = "Cancel"
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $CancelButton
    $form.Controls.Add($CancelButton)

    # �R���{�{�b�N�X���쐬
    $Combo = New-Object System.Windows.Forms.Combobox
    $Combo.Location = New-Object System.Drawing.Point(50, 50)
    $Combo.size = New-Object System.Drawing.Size(150, 30)
    # ���X�g�ȊO�̓��͂������Ȃ�
    $Combo.DropDownStyle = "DropDownList"
    $Combo.FlatStyle = "standard"
    # $Combo.font = $font
    $Combo.BackColor = "#005050"
    $Combo.ForeColor = "white"
        
    # -----------�R���{�{�b�N�X�ɍ��ڂ�ǉ�-----------
    for ($counterForMove = (-6); $counterForMove -le 6; $counterForMove++) {
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
    if ($result -eq "OK") {
        # ���[�U�[�̉񓚂�"�N"�ŋ�؂�
        $Combo.Text -match "(?<year>.+?)�N(?<month>.+?)��" | out-null

        # ���[�U�[�w��̔N�������쐬�̑Ώ۔N�Ƃ��ď㏑����
        $targetYear = $Matches.year

        # ���[�U�[�w��̌��������쐬�̑Ώی��Ƃ��ď㏑������
        $targetMonth = $Matches.month

    }
    else {
        # �������I������
        exit
    }

    # ��$yesNo_yearMonthAreCorrect -eq 'No'���[�v�I����
}

# ----------- ���΂炭���҂������������ -----------

# �t�H���g�̎w��
$font = New-Object System.Drawing.Font("���C���I", 8)

# �t�H�[���̐ݒ�
$waitForm = New-Object System.Windows.Forms.Form
$waitForm.Text = "�����ݒ�"
$waitForm.Size = New-Object System.Drawing.Size(265, 170)
$waitForm.StartPosition = "CenterScreen"
$waitForm.font = $font

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(70, 30)
$label.Size = New-Object System.Drawing.Size(270, 30)
$label.Text = "�������ł�`r`n���΂炭���҂���������"
$waitForm.Controls.Add($label)

###### �摜�̃t���p�X�ύX���Ăق����ł��� #########
#PictureBox
$pic = New-Object System.Windows.Forms.PictureBox
$pic.Size = New-Object System.Drawing.Size(50, 50)
$pic.Image = [System.Drawing.Image]::FromFile($PWD.Path + "\resources\picutures\���҂����������L.png")
$pic.Location = New-Object System.Drawing.Point(20,20) 
$pic.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
$waitForm.Controls.Add($pic)

# ����
$waitResult = $waitForm.Show()

# ------------------------------------------------------



Write-Host "$targetYear �N��"
Write-Host "$targetMonth ���̏������쐬���܂�"

# �|�b�v�A�b�v���쐬
$popup = new-object -comobject wscript.shell

# �t�@�C�����̋Ζ��\_�̂��Ƃ̕\�L
$fileNameMonth = [string]$targetMonth + "��"
# �Ζ��\�t�@�C�����擾
$kinmuhyou = Get-ChildItem -Recurse -File | ? Name -Match ("[0-9]{3}_�Ζ��\_" + $fileNameMonth + "_.+") 
# �Y���Ζ��\�t�@�C���̌��m�F
if ($kinmuhyou.Count -lt 1) {
    
    # �|�b�v�A�b�v��\��
    $popup.popup("$targetMonth ���̋Ζ��\�t�@�C�������݂��܂���", 0, "��蒼���Ă�������", 48) | Out-Null
    exit
}
elseif ($kinmuhyou.Count -gt 1) {
    # �|�b�v�A�b�v��\��
    $popup.popup("$targetMonth ���̋Ζ��\�t�@�C�����������܂�`r`n1�ɂ��Ă�������", 0, "��蒼���Ă�������", 48) | Out-Null
    exit
}

# ----------------------Excel���N������--------------------------------
try {
    # �N������Excel�v���Z�X���擾
    $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
}
catch {
    # Excel�v���Z�X���N�����ĂȂ���ΐV���ɋN������
    $excel = New-Object -ComObject "Excel.Application" 
}

# Excel�����b�Z�[�W�_�C�A���O��\�����Ȃ��悤�ɂ���
$excel.DisplayAlerts = $false
$excel.visible = $false

# �Ζ��\�̃t���p�X
$kinmuhyouFullPath = $kinmuhyou.FullName 

# �Ζ��\�u�b�N���J��
$kinmuhyouBook = $excel.workbooks.open($kinmuhyouFullPath)
$kinmuhyouSheet = $kinmuhyouBook.worksheets.item([String]$targetMonth + '��')

# ���͓��e���܂Ƃ߂ē���Ă������߂̔z��
$inputContentsArray = @()

# ����o�^����Ζ��n���i�[����z��
$workPlaceArray = @()

# ���łɋΖ��n��񃊃X�g�ɏ����Ă���Ζ��n���i�[����z��
$registeredWorkPlaceArray = @()

# ---------------�Ζ��n��񃊃X�g��ǂݍ���---------------------
# �Ζ��n��񃊃X�g�������Ă���e�L�X�g
$infoTextFileName = ".\resources\�c�[���p����.txt"
$infoTextFileFullpath = "$PWD\$infoTextFileName"

# �Ζ��n��񃊃X�g�e�L�X�g�����݂����Ƃ��̏���
if (Test-Path $infoTextFileFullpath) {

    # �Ζ��n��񃊃X�g�e�L�X�g�̓��e���擾
    $argumentText = (Get-Content $infoTextFileFullpath)
    
    # �Ζ��n��񃊃X�g�e�L�X�g�ɂ��łɏ�����Ă�������擾����
    for ($i = 0; $i -lt $argumentText.Length; $i++) {
        $argumentText[$i] -Match "(?<workplace>.+?)_" | Out-Null
        # ���łɔz��ɓ����Ă���Ζ��n�͒ǉ����Ȃ�
        if (!$registeredWorkPlaceArray.Contains($Matches.workplace)) {
            # �z��ɂȂ��Ζ��n��z��ɒǉ�����
            $registeredWorkPlaceArray += $Matches.workplace
        }
    }
}

# �Ζ��\����Ζ��n�ꗗ���擾����
# $kinmunaiyou : �Ζ����e��Z��
# $kinmujisseki : �Ζ����ї�Z��
# $sagyoubasho : ��Əꏊ�Z��
for ($Row = 14; $Row -le 44; $Row++) {
    # �u�Ζ����e�v���̕�������擾
    $kinmunaiyou = $kinmuhyouSheet.cells.item($Row, 26).text
    # �u�Ζ����сv���̏I�������̕�������擾
    $kinmujisseki = $kinmuhyouSheet.cells.item($Row, 7).text
    # �u��Əꏊ�v���̕�������擾
    $sagyoubasho = $kinmuhyouSheet.cells.item(7, 7).text

    # �Ζ����т���l�łȂ����o�΂��Ă��
    if ($kinmujisseki -ne "") {

        # �Ζ����e����l�łȂ����Ζ��n�Ȃǂ������Ă���
        if ($kinmunaiyou -ne "") {
            # �Ζ����e����Ζ��n�������Ă���
            $workPlace = $kinmunaiyou        
        }
        else {
            # �o�΂��Ă邯�ǋΖ����e�ɋΖ��n�������ĂȂ��ꍇ
            # ��Əꏊ����Ζ��n�������Ă���
            $workPlace = $sagyoubasho
        }   
    }

    # ����o�^����Ζ��n�ɂ܂��o�^����ĂȂ����A�c�[���p����.txt�ɂ܂��o�^����Ă��Ȃ��ꍇ�́A����o�^����Ζ��n�z��ɒǉ�����
    if (!$workPlaceArray.Contains($workPlace) -and !$registeredWorkPlaceArray.Contains($workPlace)) {
        $workPlaceArray += @($workPlace)
    }
}

# ����o�^������̂��Ȃ��ꍇ��popup��\�����ďI��
if ($workPlaceArray.Length -eq 0) {
    # �|�b�v�A�b�v��\��
    $popup.popup("$targetmonth ���̋Ζ��\�̋Ζ��n�͊��ɓo�^����Ă��܂��B", 0, "�o�^�ς�", 64) | Out-Null
    breakExcel    
    exit
}


# =========================== ���͉�� ===========================

# ---------------- �ϐ���` ----------------

# �t�H���g���w��
$font = New-Object System.Drawing.Font("�l�r �S�V�b�N", 11)

# �t�H�[�����Ƃ̗v�f���i�[����z��
$forms = @()
# �K�p
$outputTekiyous = @()
# ���
$outputKukans = @()
# ��ʋ@��
$outputKoutsukikans = @()
# ���z
$outputKingakus = @()

# ��ʋ@��
$koutsukikan1 = @()
$koutsukikan2 = @()
$koutsukikan3 = @()
$koutsukikan4 = @()
$koutsukikan5 = @()
$koutsukikan6 = @()

# �t�H�[������肷���Ȃ��悤�ɂ��邽�߂̃t���O
# $True : �V���Ƀt�H�[�������
# $False : �V���Ƀt�H�[�������Ȃ��i�㏑���̂݁j
# �ŏ��̃��[�v�͑��₵�����Ƃɂ���
$isAdded = $True


# ---------------- �֐���` ----------------

# �t�H�[�����쐬����֐�
# Args[0] : �^�C�g���ɕ\�����镶����
function drawForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "�����ݒ�  �y" + $Args[0] + "�z"
    $form.Size = New-Object System.Drawing.Size(660, 700)
    $form.StartPosition = "CenterScreen"
    return $form
}


# ���x�����쐬����֐�
# Args[0] : �t�H�[�����̐ݒ���W�i���̈ʒu�j
# Args[1] : �t�H�[�����̐ݒ���W�i�c�̈ʒu�B�����j
# Args[2] : ���x����\�����镝
# Args[3] : ���x���ɕ\�����镶����
# Args[4] : ���x����\������t�H�[��
# Args[5] : ���x���̃t�H���g
function drawLabel {
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point($Args[0], $Args[1])
    $label.Size = New-Object System.Drawing.Size($Args[2], 15)
    $label.Text = $Args[3]
    $label.forecolor = "black"
    $label.font = $Args[5]
    if ($Args[5] -ne $null) {
        $Args[5]
    }
    $Args[4].Controls.Add($label)
    return $label
}

# OK�{�^�����쐬����֐�
# Args[0] : �t�H�[�����̐ݒ���W�i�c�̈ʒu�B�����j
# Args[1] : OK�{�^���ɕ\�����镶����
# Args[2] : OK�{�^����\������t�H�[��
# result : OK
function drawOKButton {
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Point(20, $Args[0])
    $OKButton.Size = New-Object System.Drawing.Size(75, 30)
    $OKButton.Text = $Args[1]
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $Args[2].AcceptButton = $OKButton
    $Args[2].Controls.Add($OKButton)
}

# �ݑ�{�^�����쐬����֐�
# Args[0] : �t�H�[�����̐ݒ���W�i�c�̈ʒu�B�����j
# Args[1] : �ݑ�{�^���ɕ\�����镶����
# Args[2] : �ݑ�{�^����\������t�H�[��
# result : Yes
function drawAtHomeButton {
    $AtHomeButton = New-Object System.Windows.Forms.Button
    $AtHomeButton.Location = New-Object System.Drawing.Point(130, $Args[0])
    $AtHomeButton.Size = New-Object System.Drawing.Size(75, 30)
    $AtHomeButton.Text = $Args[1]
    $AtHomeButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $Args[2].Controls.Add($AtHomeButton)
}

# �߂�{�^�����쐬����֐�
# Args[0] : �t�H�[�����̐ݒ���W�i�c�̈ʒu�B�����j
# Args[1] : �߂�{�^���ɕ\�����镶����
# Args[2] : �߂�{�^����\������t�H�[��
# result : Retry
function drawReturnButton {
    $ReturnButton = New-Object System.Windows.Forms.Button
    $ReturnButton.Location = New-Object System.Drawing.Point(240, $Args[0])
    $ReturnButton.Size = New-Object System.Drawing.Size(75, 30)
    $ReturnButton.Text = $Args[1]
    $ReturnButton.DialogResult = [System.Windows.Forms.DialogResult]::Retry
    if ($i -eq 0) {
        $ReturnButton.Enabled = $false; 
    }
    else {
        $ReturnButton.Enabled = $True;
    }
    $Args[2].Controls.Add($ReturnButton)
}

# �o�^�ς݋Ζ��n����I���{�^�����쐬����֐�
# result : No
function drawregisteredButton {
    $registeredButton = New-Object System.Windows.Forms.Button
    $registeredButton.Location = New-Object System.Drawing.Point(350, 610)
    $registeredButton.Size = New-Object System.Drawing.Size(155, 30)
    $registeredButton.Text = "�o�^�ς݂̋Ζ��n����I������"
    $registeredButton.DialogResult = [System.Windows.Forms.DialogResult]::No
    # if ($i -eq 0) {
    #     $registeredButton.Enabled = $false; 
    # }else {
    #     $registeredButton.Enabled = $True;
    # }
    $forms[$i].Controls.Add($registeredButton)
}


# �e�L�X�g�{�b�N�X���쐬����֐�
# Args[0] : �t�H�[�����̐ݒ���W�i���̈ʒu�j
# Args[1] : �t�H�[�����̐ݒ���W�i�c�̈ʒu�B�����j
# Args[2] : �e�L�X�g�{�b�N�X�̉���
# Args[3] : �e�L�X�g�{�b�N�X�̍���
# Args[4] : �e�L�X�g�{�b�N�X��\������t�H�[��
function drawTextBox {
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point($Args[0], $Args[1])
    $textBox.Size = New-Object System.Drawing.Size($Args[2], $Args[3])
    $textBox.BackColor = "white"
    $Args[4].Controls.Add($textBox)
    return $textBox
}


# ���΂炭���҂�����������ʂ����
$waitForm.Close()


# ���͉�ʕ\��
# workPlaceArray : �Ζ��\����擾�����A����o�^����Ζ��n�ꗗ
:EMPTY for ($i = 0; $i -lt $workPlaceArray.Length; $i++) {

    # ---------------- Main ----------------- 

    # �߂�{�^��������A�G���[�̏ꍇ �ȊO�V�����t�H�[�����쐬����
    if ($isAdded) {
        # �t�H�[���쐬�֐��Ăяo��
        $forms += drawForm $workPlaceArray[$i]   
    }

    # OK�{�^���֐��Ăяo��
    drawOKButton 610 "OK" $forms[$i]

    # �ݑ�{�^���֐��Ăяo��
    drawAtHomeButton 610 "�ݑ�" $forms[$i]

    # �߂�{�^���֐��Ăяo��
    drawReturnButton 610 "�߂�" $forms[$i]

    # �o�^�ς݋Ζ��n����I���{�^���Ăяo��
    drawregisteredButton


    # =============================== input ===============================

    # �ݑ�Ζ����́u�ݑ�{�^���v�����������郁�b�Z�[�W���x�����쐬
    $atHomeLabel = drawLabel 10 10 470 "�� �ݑ�Ζ��̂Ƃ��́y�ݑ�z�{�^�����N���b�N ��" $forms[$i]
    # �����F�w��
    $atHomeLabel.forecolor = "red" 
    # �t�H���g�w��idrawLabel�֐��̑�5�����ɓ����ƁA�����F���Ԃ��Ȃ�Ȃ��j
    $atHomeLabel.font = $font 


    # ---------------- �K�p�i�s��A�v���j ----------------- 
    # �K�p���x���̃t�H�[�����̐ݒ���W�̍���
    $tekiyouLabelLocate = 50
    # �K�p�e�L�X�g�{�b�N�X�̃t�H�[�����̐ݒ���W�̍���
    $tekiyouTextBoxLocate = 108

    # ���x���֐��Ăяo��
    drawLabel 10 $tekiyouLabelLocate 470 ("�P�D�y �K�p �z �Ζ��n `"" + $workPlaceArray[$i] + "`" �̎��̓K�p����͂��Ă�������") $forms[$i] | Out-Null
    drawLabel 20 ($tekiyouLabelLocate + 20) 470 "ex.  ������c���{��" $forms[$i] | Out-Null
    drawLabel 20 ($tekiyouLabelLocate + 40) 470 "      ����i�쁨�����e���|�[�g������ (�Ζ��n�����̏ꍇ)" $forms[$i] | Out-Null

    # �e�L�X�g�{�b�N�X�֐��Ăяo��
    $outputTekiyou = drawTextBox 20 $tekiyouTextBoxLocate 300 20  $forms[$i]

    # �߂�{�^��������A�G���[�̏ꍇ �ȊO
    if ($isAdded) {
        # �K�p�e�L�X�g�{�b�N�X��z��ɒǉ�
        $outputTekiyous += $outputTekiyou    
    }

    # ---------------- ��� ----------------- 
    # ��ԃ��x���̃t�H�[�����̐ݒ���W�̍���
    $kukanLabelLocate = 150
    # ��ԃe�L�X�g�{�b�N�X�̃t�H�[�����̐ݒ���W�̍���
    $kukanTextBoxLocate = 208

    # ���x���֐��Ăяo��
    drawLabel  10 $kukanLabelLocate 550 ("�Q�D�y ��� �z �Ζ��n `"" + $workPlaceArray[$i] + "`" �̎��̋�ԁi����̍Ŋ��w�����Ζ��n�̍Ŋ��w�j����͂��Ă�������") $forms[$i] | Out-Null
    drawLabel 20 ($kukanLabelLocate + 20) 470 "ex.  <����̍Ŋ��w>�����c�� (�����̏ꍇ)" $forms[$i] | Out-Null
    drawLabel 20 ($kukanLabelLocate + 40) 670 "      <����̍Ŋ��w>���i�쁨�����e���|�[�g��<����̍Ŋ��w> (�Ζ��n�����̏ꍇ)" $forms[$i] | Out-Null

    # �e�L�X�g�{�b�N�X�֐��Ăяo��
    $outputKukan = drawTextBox 20 $kukanTextBoxLocate 430 20 $forms[$i]

    # �߂�{�^��������A�G���[�̏ꍇ �ȊO
    if ($isAdded) {
        # ��ԃe�L�X�g�{�b�N�X��z��ɒǉ�
        $outputKukans += $outputKukan    
    }


    # ---------------- ��ʋ@�� -----------------
    # ��ʋ@�փ��x���̃t�H�[�����̐ݒ���W�̍���
    $koutsukikanLabelLocate = 290
    # ��ʋ@�փe�L�X�g�{�b�N�X�̃t�H�[�����̐ݒ���W�̍���
    $koutsukikanTextBoxLocate = 288

    # ���x���֐��Ăяo��
    drawLabel 10 250 500 ("�R�D�y ��ʋ@�� �z �Ζ��n `"" + $workPlaceArray[$i] + "`" �̎��ɗ��p�����ʋ@�ւ���͂��Ă�������") $forms[$i] | Out-Null
    drawLabel 20 270 500 "ex. JR�R���" $forms[$i] | Out-Null
    drawLabel 10 $koutsukikanLabelLocate 70 "��ʋ@�ւP�F" $forms[$i] | Out-Null
    drawLabel 10 ($koutsukikanLabelLocate + 40) 70 "��ʋ@�ւQ�F" $forms[$i] | Out-Null
    drawLabel 10 ($koutsukikanLabelLocate + 80) 70 "��ʋ@�ւR�F" $forms[$i] | Out-Null
    drawLabel 10 ($koutsukikanLabelLocate + 120) 70 "��ʋ@�ւS�F" $forms[$i] | Out-Null
    drawLabel 10 ($koutsukikanLabelLocate + 160) 70 "��ʋ@�ւT�F" $forms[$i] | Out-Null
    drawLabel 10 ($koutsukikanLabelLocate + 200) 70 "��ʋ@�ւU�F" $forms[$i] | Out-Null

    # �e�L�X�g�{�b�N�X�֐��Ăяo��
    $koutsukikan1 = drawTextBox 90 $koutsukikanTextBoxLocate 200 20 $forms[$i]
    $koutsukikan2 = drawTextBox 90 ($koutsukikanTextBoxLocate + 40) 200 20 $forms[$i]
    $koutsukikan3 = drawTextBox 90 ($koutsukikanTextBoxLocate + 80) 200 20 $forms[$i]
    $koutsukikan4 = drawTextBox 90 ($koutsukikanTextBoxLocate + 120) 200 20 $forms[$i]
    $koutsukikan5 = drawTextBox 90 ($koutsukikanTextBoxLocate + 160) 200 20 $forms[$i]
    $koutsukikan6 = drawTextBox 90 ($koutsukikanTextBoxLocate + 200) 200 20 $forms[$i]

    # �߂�{�^��������A�G���[�̏ꍇ �ȊO
    if ($isAdded) {
        # ��̏����Ŏg���₷�����邽�߁A�e��ʋ@�ւ�z��Ɋi�[����
        $inputkoutsukikan = @($koutsukikan1, $koutsukikan2, $koutsukikan3, $koutsukikan4, $koutsukikan5, $koutsukikan6)
        $outputKoutsukikans+= , @($inputkoutsukikan)
    }
    

    # ---------------- ���z -----------------
    # ���z���x���̃t�H�[�����̐ݒ���W�̍���
    $kingakuLabelLocate = 530
    # ���z�e�L�X�g�{�b�N�X�̃t�H�[�����̐ݒ���W�̍���
    $kingakuTextBoxLocate = 570

    # ���x���֐��Ăяo��
    drawLabel 10 $kingakuLabelLocate 500 ("�S�D�y ���z �z �Ζ��n `"" + $workPlaceArray[$i] + "`" �̋��z�i��������j����͂��Ă�������") $forms[$i] | Out-Null
    drawLabel 20 ($kingakuLabelLocate + 20) 470 "ex.  750 �i���p�����j" $forms[$i] | Out-Null

    # �e�L�X�g�{�b�N�X�֐��Ăяo��
    $outputKingaku = drawTextBox 20 $kingakuTextBoxLocate 100 20 $forms[$i]

    # �߂�{�^��������A�G���[�̏ꍇ �ȊO
    if ($isAdded) {
        # ���z�e�L�X�g�{�b�N�X��z��ɒǉ�
        $outputKingakus += $outputKingaku   
    }

    # ����
    $inputContentsResult = $forms[$i].ShowDialog()


    # =============================== output ===============================
    # --------------- OK�{�^������������ ---------------
    if ($inputContentsResult -eq "OK") {

        #  ---------------- �󔒃G���[���� -----------------

        # �ȉ��̕ϐ������Z�b�g����
        #
        # nullOrEmptyCount : ��ʋ@�փe�L�X�g�{�b�N�X�̋�̌�
        # koutsukikans : �����̌�ʋ@�փe�L�X�g�{�b�N�X����ɂ܂Ƃ߂邽�߂̕ϐ�
        # outputKoutsukikan : �ҏW����koutsukikans��������
        # isEmpty : �󔒃G���[���N�������߂̃t���O
        #
        $nullOrEmptyCount = 0
        $koutsukikans = ""
        $outputKoutsukikan= ""
        $isEmpty = $false

        # �e�L�X�g�{�b�N�X�̐F�𔒂ɖ߂�
        $outputTekiyous[$i].BackColor = "white"
        $outputKukans[$i].BackColor = "white"
        $outputKingakus[$i].BackColor = "white"
        $outputKoutsukikans[$i][0].BackColor = "white"


        for ($l = 0; $l -lt $outputKoutsukikans[$i].length; $l++) {
            # ��ʋ@�փe�L�X�g�{�b�N�X�����ł͂Ȃ����̂𔲂��o��
            if ([string]::IsNullOrEmpty($outputKoutsukikans[$i][$l].text)) {
                # NULL �� '' �̏ꍇ
                $nullOrEmptyCount++
            }
            else {
                # ��L�ȊO�͐ݒ肳�ꂽ��������o��
                $koutsukikans += ($outputKoutsukikans[$i][$l].text + '`r`n')
            }
        }

        # �����́u`r`n�v������
        $outputKoutsukikan+= $koutsukikans.Substring(0, $koutsukikans.Length - 4)

        while ($True) {
            # ��ʋ@�ւ��S�ċ󂾂����ꍇ�̏���
            if ($nullOrEmptyCount -eq 6) {
                $outputKoutsukikans[$i][0].BackColor = "#ff99cc"
                $isEmpty = $True
            }
    
            # ���[�U���͂ɋ󔒂��������ꍇ�̏���
            $inputtedTextBoxes = @($outputTekiyous[$i], $outputKukans[$i], $outputKingakus[$i])
            # ���[�U���͂ɂP�ł��󔒂��������ꍇ�̏���
            foreach ($inputtedTextBox in $inputtedTextBoxes) {
                if ($inputtedTextBox.text -eq "") {
                    $inputtedTextBox.BackColor = "#ff99cc"
                    $isEmpty = $True
                }    
            }
            # �󔒂��������ꍇ����ȍ~�̏������X�L�b�v����
            if ($isEmpty) {
                # ��ʋ@�ւ̋󔒃J�E���g��������
                $nullOrEmptyCount = 0
                $i = $i - 1
                $isAdded = $false
                continue EMPTY
            }
            # �G���[���Ȃ��ꍇ�̓��[�v���甲����
            break
        }
        

        # ---------------- �K�p�i�s��A�v���j -----------------

        $inputContentsArray += @($workPlaceArray[$i] + "_" + $outputTekiyous[$i].text)

        # ---------------- ��� -----------------

        $inputContentsArray += @($workPlaceArray[$i] + "_" + $outputKukans[$i].text)

        # ---------------- ��ʋ@�� -----------------

        $inputContentsArray += @($workPlaceArray[$i] + "_" + $outputKoutsukikan)

        # ---------------- ���z -----------------

        $inputContentsArray += @($workPlaceArray[$i] + "_" + $outputKingakus[$i].text)

        # �t�H�[�����₵���t���O
        $isAdded = $True

    
        # --------------- �߂�{�^������������ ---------------
    }
    elseif ($inputContentsResult -eq "Retry") {
        
        # �J��Ԃ��̏�����2�߂�
        # �Ⴆ�΁A1��ʖڂ��c���i$i = 1�j2��ʖڂ������i$i = 2�j�������Ƃ��A�c���̉�ʂɖ߂肽���Ƃ��� $i = 1 �ɂ�����
        # for���̏������H�ŃC���N�������g����Ă��邽�߁A$i����2�������K�v������
        $i = $i - 2
        # �z��ɂȂɂ������Ă��Ȃ����i�Œ�z��Ȃ̂ŁA�ŏ��̗v�f�͋�ɂ��邾���ɂ����j
        if ($inputContentsArray.Length -le 4) {
            for ($j = 1; $j -lt 5; $j++) {
                $inputContentsArray[($inputContentsArray.Length - $j)] = ""
            }
            # �߂�{�^������������A�e�L�X�g�t�@�C���ɏo�͂���v�f���폜����    
        }
        else {
            $inputContentsArray = $inputContentsArray[0..($inputContentsArray.Length - 5)]
        }
        
        $isAdded = $false

        # --------------- �ݑ�{�^������������ ---------------
    }
    elseif ($inputContentsResult -eq "Yes") {
        $inputContentsArray += @($workPlaceArray[$i] + "_1")
        $inputContentsArray += @($workPlaceArray[$i] + "_1")
        $inputContentsArray += @($workPlaceArray[$i] + "_1")
        $inputContentsArray += @($workPlaceArray[$i] + "_1")
    }
    # �o�^�ς݋Ζ��n����I������ꍇ
    elseif ($inputContentsResult -eq "No") {
        # �o�^�ς݋Ζ��n�I��p�t�H�[�����쐬
        $selectForm = New-Object System.Windows.Forms.Form
        $selectForm.Text = "�o�^�ς݂̋Ζ��n����I��"
        $selectForm.Size = New-Object System.Drawing.Size(300, 200)
        $selectForm.StartPosition = "CenterScreen"
        
        # ����
        $selectResult = $selectForm.ShowDialog()
    
    }
    else {
        breakExcel
        exit
    }    
}

# ������ʂ��ق����Ȃ�

# �z����e�L�X�g�ɏo�͂���
foreach ($inputContent in $inputContentsArray) {
    $inputContent >> .\resources\�c�[���p����.txt
}

# �Ζ��\�t�@�C�������
breakExcel

# �ϐ��̉��
$outputTekiyou = $null
$outputKukan = $null
$koutsukikan1 = $null
$koutsukikan2 = $null
$koutsukikan3 = $null
$koutsukikan4 = $null
$koutsukikan5 = $null
$koutsukikan6 = $null
$inputKoutsukikan = $null
$koutsukikans = $null
$outputKoutsukikans = $null
$outputKingaku = $null