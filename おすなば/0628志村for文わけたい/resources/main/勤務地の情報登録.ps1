#
#
#



# --------------- �A�Z���u���̓ǂݍ��� ---------------
Add-Type -Assembly System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---------------- ���Ӊ�ʃX�N���v�g�̌Ăяo�� -------------------
. (Join-Path -path $PWD -childpath "..\scripts\���Ӊ��.ps1")

# ���Ӊ��.ps1�̊֐������s����
$attentionForm = attentionThisTool

# �t�H�[���̉���
$attentionResult = $attentionForm.ShowDialog()

if ($attentionResult -eq "Cancel") {
    exit 
}
# ---------------- ���Ӊ�ʃX�N���v�g�I�� -------------------

# ���݂̔N�������擾����
$targetYear = (Get-Date).Year
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
$yesNo_yearMonthAreCorrect = [System.Windows.Forms.MessageBox]::Show("�y $targetYear �N $targetMonth �� �z�̋Ζ��\�����Ƃɏ����ݒ�����܂����H`r`n`r`n�u�������v�ő��̌���I���ł��܂�", '�쐬���鏬���̑Ώ۔N��', 'YesNo', 'Question')

# ��$yesNo_yearMonthAreCorrect -eq 'No'���[�v�J�n��
if ($yesNo_yearMonthAreCorrect -eq 'No') {

    # ---------------- �Ζ��\�̌��I���X�N���v�g���Ăяo�� -------------------
    . (Join-Path -Path $PWD -ChildPath "..\scripts\�Ζ��\�̌��I��.ps1")

    # �Ζ��\�̌��I��.ps1�̊֐������s
    # choicedMonth[0] : �Ζ��\�̌��I���t�H�[��
    # choicedMonth[1] : �Ζ��\�̌��I���R���{�{�b�N�X
    $choicedMonth = choiceMonth

    # �Ζ��\�̌��I����ʂ�����
    $choiceMonthResult = $choicedMonth[0].ShowDialog()

    if ($choiceMonthResult -eq "OK") {
        # ���[�U�[�̉񓚂�"�N"�ŋ�؂�
        $choicedMonth[1].Text -match "(?<year>.+?)�N(?<month>.+?)��" | out-null

        # ���[�U�[�w��̔N�������쐬�̑Ώ۔N�Ƃ��ď㏑����
        $targetYear = $Matches.year

        # ���[�U�[�w��̌��������쐬�̑Ώی��Ƃ��ď㏑������
        $targetMonth = $Matches.month

    }
    else {   
        # �������I������
        exit
    }
}
# ---------------- �Ζ��\�̌��I���X�N���v�g�I�� -------------------

# ----------- ���΂炭���҂����������X�N���v�g�Ăяo�� -----------
. (Join-Path -Path $PWD -ChildPath ..\scripts\���΂炭���҂���������.ps1)

# ���΂炭���҂���������.ps1�̊֐������s
$waitCatForm = pleaseWait "..\images\���҂����������L.png"

# ���΂炭���҂����������t�H�[���̉���
$waitCatForm.show()

# ---------------- �Ζ��\�`�F�b�N���� -------------------
# �|�b�v�A�b�v���쐬
$popup = new-object -comobject wscript.shell

# �t�@�C�����̋Ζ��\_�̂��Ƃ̕\�L
$targetMonth00 = "{0:00}" -f [int]$targetMonth
$fileNameMonth = ("$targetYear" + "$targetMonth00")
# �Ζ��\�t�@�C�����擾
$kinmuhyou = Get-ChildItem "..\..\01_�_�E�����[�h�����Ζ��\" -Recurse -File | ? Name -Match ("[0-9]{3}_�Ζ��\_" + $fileNameMonth + "_.+") 
# �Y���Ζ��\�t�@�C���̌��m�F
if ($kinmuhyou.Count -lt 1) {
    
    # ���΂炭���҂�����������ʂ����
    $waitCatForm.Close()

    # �|�b�v�A�b�v��\��
    $popup.popup("$targetMonth ���̋Ζ��\�t�@�C�������݂��܂���", 0, "��蒼���Ă�������", 48) | Out-Null
    exit
}
elseif ($kinmuhyou.Count -gt 1) {
    
    # ���΂炭���҂�����������ʂ����
    $waitCatForm.Close()

    # �|�b�v�A�b�v��\��
    $popup.popup("$targetMonth ���̋Ζ��\�t�@�C�����������܂�`r`n1�ɂ��Ă�������", 0, "��蒼���Ă�������", 48) | Out-Null
    exit
}
# ---------------- �Ζ��\�`�F�b�N�����I�� -------------------

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

# ----------------------Excel���N������--------------------------------

# Excel�v���Z�X���N�����ĂȂ���ΐV���ɋN������
$excel = New-Object -ComObject "Excel.Application" 

# Excel�����b�Z�[�W�_�C�A���O��\�����Ȃ��悤�ɂ���
$excel.DisplayAlerts = $false
$excel.visible = $false

# �Ζ��\�̃t���p�X
$kinmuhyouFullPath = $kinmuhyou.FullName 

# �Ζ��\�u�b�N���J��
$kinmuhyouBook = $excel.workbooks.open($kinmuhyouFullPath)
$kinmuhyouSheet = $kinmuhyouBook.worksheets.item([String]$targetMonth + '��')


# ---------------�Ζ��n��񃊃X�g��ǂݍ��ޏ���---------------------

# ---------------�z���`---------------------
# ���͓��e���܂Ƃ߂ē���Ă������߂̔z��
$inputContentsArray = @()
# ����o�^����Ζ��n���i�[����z��
$workPlaceArray = @()
# ���łɋΖ��n��񃊃X�g�ɏ����Ă���Ζ��n���i�[����z��
$registeredWorkPlaceArray = @()


# �Ζ��n��񃊃X�g�������Ă���e�L�X�g
$infoTextFileName = "..\user_info\�c�[���p����.txt"
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
    
    # ���΂炭���҂�����������ʂ����
    $waitCatForm.Close()
    
    # �|�b�v�A�b�v��\��
    $popup.popup("$targetmonth ���̋Ζ��n�͂��ׂēo�^����Ă��܂��B", 0, "�o�^�ς�", 64) | Out-Null
    breakExcel    
    exit
}

# =========================== ���͉�� ===========================

# ---------------- �ϐ���` ----------------

# �t�H���g���w��
$bigFont = New-Object System.Drawing.Font("Yu Gothic UI", 20)
$blankErrorFont = New-Object System.Drawing.Font("Yu Gothic UI", 9)

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

# ���z�����p�����łȂ��������̃G���[���b�Z�[�W
$kingakuErrorMessages = @()

# ���[�U���͂��󔒂��������̃G���[���b�Z�[�W
$blankErrorMessages = @()

# �t�H�[������肷���Ȃ��悤�ɂ��邽�߂̃t���O
# $True : �V���Ƀt�H�[�������
# $False : �V���Ƀt�H�[�������Ȃ��i�㏑���̂݁j
# �ŏ��̃��[�v�͑��₵�����Ƃɂ���
$isAdded = $True

# ---------------- �又�� ----------------

# ���C�������p�֐�.ps1�̓ǂݍ���
. (Join-Path -Path $PWD -ChildPath ..\scripts\���C�������p�֐�.ps1)

# ---------------- �֐���` ----------------

# �t�H�[�����쐬����֐�
# Args[0] : �^�C�g���ɕ\�����镶����
function drawForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "�Ζ��n�̏���o�^"
    $form.Size = New-Object System.Drawing.Size(650, 730)
    $form.StartPosition = "CenterScreen"
    $form.font = $font
    $form.formborderstyle = "FixedSingle"
    return $form
}


# ���x�����쐬����֐�
# Args[0] : �t�H�[�����̐ݒ���W�i���̈ʒu�j
# Args[1] : �t�H�[�����̐ݒ���W�i�c�̈ʒu�B�����j
# Args[2] : ���x����\�����镝
# Args[3] : ���x����\�����鍂��
# Args[4] : ���x���ɕ\�����镶����
# Args[5] : ���x����\������t�H�[��
# Args[6] : ���x���̃t�H���g
function drawLabel {
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point($Args[0], $Args[1])
    $label.Size = New-Object System.Drawing.Size($Args[2], $Args[3])
    $label.Text = $Args[4]
    $label.forecolor = "black"
    $label.font = $Args[6]
    if ($Args[6] -ne $null) {
        $Args[6]
    }
    $Args[5].Controls.Add($label)
    return $label
}

# OK/�o�^�{�^�����쐬����֐�
# Args[0] : �t�H�[�����̐ݒ���W�i���̈ʒu�j
# Args[1] : �t�H�[�����̐ݒ���W�i�c�̈ʒu�B�����j
# Args[2] : �{�^����\�����鉡��
# Args[3] : �{�^����\������c��
# Args[4] : OK/�o�^�{�^���ɕ\�����镶����
# Args[5] : OK/�o�^�{�^����\������t�H�[��
# result : OK
function drawOKButton {
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Point($Args[0], $Args[1])
    $OKButton.Size = New-Object System.Drawing.Size($Args[2], $Args[3])
    $OKButton.Text = $Args[4]
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $Args[5].AcceptButton = $OKButton
    $Args[5].Controls.Add($OKButton)
}

# �ݑ�{�^�����쐬����֐�
# Args[0] : �t�H�[�����̐ݒ���W�i�c�̈ʒu�B�����j
# Args[1] : �ݑ�{�^���ɕ\�����镶����
# Args[2] : �ݑ�{�^����\������t�H�[��
# result : Yes
function drawAtHomeButton {
    $AtHomeButton = New-Object System.Windows.Forms.Button
    $AtHomeButton.Location = New-Object System.Drawing.Point(10, $Args[0])
    $AtHomeButton.Size = New-Object System.Drawing.Size(300, 30)
    $AtHomeButton.Text = $Args[1]
    $AtHomeButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $AtHomeButton.Backcolor = "paleturquoise"
    $AtHomeButton.Forecolor = "Blue"
    $Args[2].Controls.Add($AtHomeButton)
}

# �߂�{�^�����쐬����֐�
# Args[0] : �t�H�[�����̐ݒ���W�i�c�̈ʒu�B�����j
# Args[1] : �߂�{�^���ɕ\�����镶����
# Args[2] : �߂�{�^����\������t�H�[��
# result : Retry
function drawReturnButton {
    $ReturnButton = New-Object System.Windows.Forms.Button
    $ReturnButton.Location = New-Object System.Drawing.Point(500, $Args[0])
    $ReturnButton.Size = New-Object System.Drawing.Size(90, 30)
    $ReturnButton.Text = $Args[1]
    $ReturnButton.DialogResult = [System.Windows.Forms.DialogResult]::Retry
    # 1�Ԗڂ̃t�H�[���ł̓{�^����񊈐��ɂ���
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
function drawRegisteredButton {
    $registeredButton = New-Object System.Windows.Forms.Button
    $registeredButton.Location = New-Object System.Drawing.Point(320, $Args[0])
    $registeredButton.Size = New-Object System.Drawing.Size(300, 30)
    $registeredButton.Text = $Args[1]
    $registeredButton.DialogResult = [System.Windows.Forms.DialogResult]::No
    $registeredButton.Backcolor = "palegreen"
    $registeredButton.Forecolor = "darkgreen"
    # �c�[���p����.txt �����݂��Ă��Ȃ� or ���g����̎��̓{�^����񊈐��ɂ���
    if (!(Test-Path $infoTextFileFullpath) -or ($argumentText.Length -eq 0)) {
        $registeredButton.Enabled = $false; 
    }else {
        $registeredButton.Enabled = $True;
    }
    $Args[2].Controls.Add($registeredButton)
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
$waitCatForm.Close()

# ���͉�ʕ\��
# workPlaceArray : �Ζ��\����擾�����A����o�^����Ζ��n�ꗗ
:EMPTY for ($i = 0; $i -lt $workPlaceArray.Length; $i++) {

    # ---------------- Main ----------------- 

    # �߂�{�^��������A�G���[�̏ꍇ �ȊO�V�����t�H�[�����쐬����
    if ($isAdded) {
        # �t�H�[���쐬�֐��Ăяo��
        $forms += drawForm $workPlaceArray[$i]   
    }

    # �Ζ��n�\��
    drawLabel 15 5 550 40 ("�w" + $workPlaceArray[$i] + "�x�̏��������Ă�������")  $forms[$i] $bigFont | Out-Null

    # �o�^�{�^���쐬�֐��Ăяo��
    drawOKButton 250 645 130 30 "�o �^" $forms[$i]

    # �߂�{�^���쐬�֐��Ăяo��
    drawReturnButton 645 "�߂�" $forms[$i]

    # �ݑ�{�^���쐬�֐��Ăяo��
    drawAtHomeButton 50 "���ݑ�Ζ�/���/�o�^�ΏۊO�͂������N���b�N��" $forms[$i]

    # �o�^�ς݋Ζ��n����I���{�^���쐬�֐��Ăяo��
    drawRegisteredButton 50 "���o�^�ς݂̋Ζ��n����I������ꍇ�͂������N���b�N��" $forms[$i]

    # =============================== input ===============================

    # ---------------- �K�p�i�s��A�v���j ----------------- 
    # �K�p���x���̃t�H�[�����̐ݒ���W�̍���
    $tekiyouLabelLocate = 115
    # �K�p�e�L�X�g�{�b�N�X�̃t�H�[�����̐ݒ���W�̍���
    $tekiyouTextBoxLocate = 175

    # ���x���֐��Ăяo��
    drawLabel 10 $tekiyouLabelLocate 470 15 ("�P�D�y �K�p �z �K�p����͂��Ă�������") $forms[$i] | Out-Null
    drawLabel 30 ($tekiyouLabelLocate + 20) 470 15 "ex.  ������c���{��" $forms[$i] | Out-Null
    drawLabel 30 ($tekiyouLabelLocate + 40) 470 15 "      ����i�쁨�����e���|�[�g������ (�Ζ��n�����̏ꍇ)" $forms[$i] | Out-Null

    # �e�L�X�g�{�b�N�X�֐��Ăяo��
    $outputTekiyou = drawTextBox 30 $tekiyouTextBoxLocate 300 20  $forms[$i]

    # �߂�{�^��������A�G���[�̏ꍇ �ȊO
    if ($isAdded) {
        # �K�p�e�L�X�g�{�b�N�X��z��ɒǉ�
        $outputTekiyous += $outputTekiyou    
    }

    # ---------------- ��� ----------------- 
    # ��ԃ��x���̃t�H�[�����̐ݒ���W�̍���
    $kukanLabelLocate = 215
    # ��ԃe�L�X�g�{�b�N�X�̃t�H�[�����̐ݒ���W�̍���
    $kukanTextBoxLocate = 275

    # ���x���֐��Ăяo��
    drawLabel 10 $kukanLabelLocate 550 15 ("�Q�D�y ��� �z ��ԁi����̍Ŋ��w�����Ζ��n�̍Ŋ��w�j����͂��Ă�������") $forms[$i] | Out-Null
    drawLabel 30 ($kukanLabelLocate + 20) 470 15 "ex.  <����̍Ŋ��w>�����c�� (�����̏ꍇ)" $forms[$i] | Out-Null
    drawLabel 30 ($kukanLabelLocate + 40) 670 15 "      <����̍Ŋ��w>���i�쁨�����e���|�[�g��<����̍Ŋ��w> (�Ζ��n�����̏ꍇ)" $forms[$i] | Out-Null

    # �e�L�X�g�{�b�N�X�֐��Ăяo��
    $outputKukan = drawTextBox 30 $kukanTextBoxLocate 430 20 $forms[$i]

    # �߂�{�^��������A�G���[�̏ꍇ �ȊO
    if ($isAdded) {
        # ��ԃe�L�X�g�{�b�N�X��z��ɒǉ�
        $outputKukans += $outputKukan    
    }


    # ---------------- ��ʋ@�� -----------------
    # ��ʋ@�փ��x���̃t�H�[�����̐ݒ���W�̍���
    $koutsukikanLabelLocate = 360
    # ��ʋ@�փe�L�X�g�{�b�N�X�̃t�H�[�����̐ݒ���W�̍���
    $koutsukikanTextBoxLocate = 358

    # ���x���֐��Ăяo��
    drawLabel 10 315 500 15 ("�R�D�y ��ʋ@�� �z ���p�����ʋ@�ւ���͂��Ă�������") $forms[$i] | Out-Null
    drawLabel 30 335 500 15 "ex. JR�R���" $forms[$i] | Out-Null
    drawLabel 30 $koutsukikanLabelLocate 80 15 "��ʋ@�ւP�F" $forms[$i] | Out-Null
    drawLabel 30 ($koutsukikanLabelLocate + 35) 80 15 "��ʋ@�ւQ�F" $forms[$i] | Out-Null
    drawLabel 30 ($koutsukikanLabelLocate + 70) 80 15 "��ʋ@�ւR�F" $forms[$i] | Out-Null
    drawLabel 30 ($koutsukikanLabelLocate + 105) 80 15 "��ʋ@�ւS�F" $forms[$i] | Out-Null
    drawLabel 30 ($koutsukikanLabelLocate + 140) 80 15 "��ʋ@�ւT�F" $forms[$i] | Out-Null
    drawLabel 30 ($koutsukikanLabelLocate + 175) 80 15 "��ʋ@�ւU�F" $forms[$i] | Out-Null

    # �e�L�X�g�{�b�N�X�֐��Ăяo��
    $koutsukikan1 = drawTextBox 110 $koutsukikanTextBoxLocate 200 20 $forms[$i]
    $koutsukikan2 = drawTextBox 110 ($koutsukikanTextBoxLocate + 35) 200 20 $forms[$i]
    $koutsukikan3 = drawTextBox 110 ($koutsukikanTextBoxLocate + 70) 200 20 $forms[$i]
    $koutsukikan4 = drawTextBox 110 ($koutsukikanTextBoxLocate + 105) 200 20 $forms[$i]
    $koutsukikan5 = drawTextBox 110 ($koutsukikanTextBoxLocate + 140) 200 20 $forms[$i]
    $koutsukikan6 = drawTextBox 110 ($koutsukikanTextBoxLocate + 175) 200 20 $forms[$i]

    # �߂�{�^��������A�G���[�̏ꍇ �ȊO
    if ($isAdded) {
        # ��̏����Ŏg���₷�����邽�߁A�e��ʋ@�ւ�z��Ɋi�[����
        $inputkoutsukikan = @($koutsukikan1, $koutsukikan2, $koutsukikan3, $koutsukikan4, $koutsukikan5, $koutsukikan6)
        $outputKoutsukikans+= , @($inputkoutsukikan)
    }
    

    # ---------------- ���z -----------------
    # ���z���x���̃t�H�[�����̐ݒ���W�̍���
    $kingakuLabelLocate = 575
    # ���z�e�L�X�g�{�b�N�X�̃t�H�[�����̐ݒ���W�̍���
    $kingakuTextBoxLocate = 615

    # ���x���֐��Ăяo��
    drawLabel 10 $kingakuLabelLocate 500 15 ("�S�D�y ���z �z ��ʔ�i��������j����͂��Ă�������") $forms[$i] | Out-Null
    drawLabel 30 ($kingakuLabelLocate + 20) 470 15 "ex.  750 �i���p�����j" $forms[$i] | Out-Null

    # ���z�����p�����������ꍇ�ɕ\�������G���[���b�Z�[�W
    $kingakuErrorMessage = drawLabel 130 $kingakuTextBoxLocate 270 15 " " $forms[$i]
    $kingakuErrorMessage.foreColor = "red"

    # �G���[���b�Z�[�W��z��ɒǉ�
    if ($isadded) {
        $kingakuErrorMessages += $kingakuErrorMessage
    }

    # �󔒂������ꍇ�ɕ\�������G���[���b�Z�[�W
    $blankErrorMessage = drawLabel 15 90 270 15 " " $forms[$i] $blankErrorFont
    $blankErrorMessage.foreColor = "red"

    # �G���[���b�Z�[�W��z��ɒǉ�
    if ($isadded) {
        $blankErrorMessages += $blankErrorMessage
    }

    # �e�L�X�g�{�b�N�X�֐��Ăяo��
    $outputKingaku = drawTextBox 30 $kingakuTextBoxLocate 100 20 $forms[$i]

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
                    $checkBlank = 0
                    if (![int]::TryParse($outputKingakus[$i].text, [ref]$checkBlank)) {
                        $outputKingakus[$i].BackColor = "#ff99cc"
                        $blankErrorMessages[$i].text = "�����L���̍��ڂ�����܂�"
                    }

                    $isEmpty = $True
                }    
            }

            # ���z�������ł͂Ȃ��������̏���
            $checkKingaku = 0
            if (![int]::TryParse($outputKingakus[$i].text, [ref]$checkKingaku)) {
                $outputKingakus[$i].BackColor = "#ff99cc"
                $kingakuErrorMessages[$i].text = "�����p�����ŋL�����Ă�������"
                $isEmpty = $True
            }

            # �󔒂��������ꍇ����ȍ~�̏������X�L�b�v����
            if ($isEmpty) {
                # ��ʋ@�ւ̋󔒃J�E���g��������
                $nullOrEmptyCount = 0
                $i = $i - 1
                $isAdded = $false
                continue EMPTY
            }
            $kingakuErrorMessages[$i].text = "�@"
            $blankErrorMessages[$i].text = "�@"
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

        # --------------- �ݑ�/����{�^������������ ---------------
    }
    elseif ($inputContentsResult -eq "Yes") {
        $inputContentsArray += @($workPlaceArray[$i] + "_1")
        $inputContentsArray += @($workPlaceArray[$i] + "_1")
        $inputContentsArray += @($workPlaceArray[$i] + "_1")
        $inputContentsArray += @($workPlaceArray[$i] + "_1")

        $isAdded = $True
    }
    # �o�^�ς݋Ζ��n����I������ꍇ
    elseif ($inputContentsResult -eq "No") {
        # �o�^�ς݋Ζ��n�I��p�t�H�[�����쐬
        $selectForm = New-Object System.Windows.Forms.Form
        $selectForm.Text = "�o�^�ς݂̋Ζ��n����I��"
        $selectForm.Size = New-Object System.Drawing.Size(300, 200)
        $selectForm.StartPosition = "CenterScreen"

        # ���x���쐬�֐��Ăяo��
        drawLabel 10 10 550 15 ("�y " + $workPlaceArray[$i] + " �z�Ɠ����Ζ��n��") $selectForm | Out-Null
        drawLabel 10 27 550 15 ("�o�^�ς݂̋Ζ��n����I�����Ă�������") $selectForm | Out-Null


        # OK�{�^���쐬�֐��Ăяo��
        # Args[0] : �t�H�[�����̐ݒ���W�i�c�̈ʒu�B�����j
        # Args[1] : OK�{�^���ɕ\�����镶����
        # Args[2] : OK�{�^����\������t�H�[��
        # result : OK
        drawOKButton 20 100 75 30 "OK" $selectForm

        # �L�����Z���{�^���̐ݒ�
        $CancelButton = New-Object System.Windows.Forms.Button
        $CancelButton.Location = New-Object System.Drawing.Point(130,100)
        $CancelButton.Size = New-Object System.Drawing.Size(85,30)
        $CancelButton.Text = "�t�H�[���ɖ߂�"
        $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $selectForm.CancelButton = $CancelButton
        $selectForm.Controls.Add($CancelButton)
        
        # �R���{�{�b�N�X���쐬
        $Combo = New-Object System.Windows.Forms.Combobox
        $Combo.Location = New-Object System.Drawing.Point(50,50)
        $Combo.size = New-Object System.Drawing.Size(150,30)
        # ���X�g�ȊO�̓��͂������Ȃ�
        $Combo.DropDownStyle = "DropDownList"
        $Combo.FlatStyle = "standard"
        $Combo.BackColor = "#005050"
        $Combo.ForeColor = "white"
            
        # �R���{�{�b�N�X�ɍ��ڂ�ǉ�
        # ���ł� �c�[���p����.txt �ɋL�ڂ���Ă���Ζ��n���R���{�{�b�N�X�̍��ڂɒǉ�
        # for($counterForMove = (-6); $counterForMove -le 6; $counterForMove++){
        foreach ($registeredWorkPlace in $registeredWorkPlaceArray){
            [void] $Combo.Items.Add($registeredWorkPlace)
        }
        
        # �R���{�{�b�N�X�̏����l��z��̈�ԍŏ��ɂ��Ă���
        $Combo.SelectedIndex = 0

        # �t�H�[���ɃR���{�{�b�N�X��ǉ�
        $selectForm.Controls.Add($Combo)
        
        # ����
        $selectResult = $selectForm.ShowDialog()

        # �I����AOK�{�^���������ꂽ�ꍇ
        if ($selectResult -eq "OK") {
            $selectForm.Visible = $false
            $selectForm.Close()

            # ���[�U�[���I�������Ζ��n�̏����A���X�g����擾 ( �z��̒��g�@[0]:�K�p�@[1]:��ԁ@[2]:��ʋ@�ց@[3]:���z )
            $workPlaceInfo = $argumentText | Select-String -Pattern ($Combo.text + '_')

            # �u�I�����ꂽ�Ζ��n�̕����� + _ �v�̑�������
            $trimWordCount = $Combo.text.Length + 1

            # �K�p�i�s��A�v���j
            $inputContentsArray += @($workPlaceArray[$i] + "_" + ([String]$workPlaceInfo[0]).Substring($trimWordCount, ([String]$workPlaceInfo[0]).Length - $trimWordCount))
            # ���
            $inputContentsArray += @($workPlaceArray[$i] + "_" + ([String]$workPlaceInfo[1]).Substring($trimWordCount, ([String]$workPlaceInfo[1]).Length - $trimWordCount))
            # ��ʋ@��
            $inputContentsArray += @($workPlaceArray[$i] + "_" + ([String]$workPlaceInfo[2]).Substring($trimWordCount, ([String]$workPlaceInfo[2]).Length - $trimWordCount))
            # ���z
            $inputContentsArray += @($workPlaceArray[$i] + "_" + ([String]$workPlaceInfo[3]).Substring($trimWordCount, ([String]$workPlaceInfo[3]).Length - $trimWordCount))

            # �t�H�[�����₵���t���O
            $isAdded = $True
        
        }
        else {
            # OK�{�^���ȊO�������ꂽ�ꍇ
            # �J��Ԃ��̏�����1�߂�
            # �Ⴆ�΁A�I���{�^�����������Ƃ��̉�ʂ��c���i$i = 1�j�������Ƃ��A$i = 1 �̉�ʂ�\��������
            # for���̏������H�ŃC���N�������g����Ă��邽�߁A$i����1�������K�v������
            $i = $i - 1

            $isAdded = $false
        }
    
    }
    else {
        breakExcel
        exit
    }    
}

# �z����e�L�X�g�ɏo�͂���
foreach ($inputContent in $inputContentsArray) {
    $inputContent >> .\resources\�c�[���p����.txt
}

# �����ݒ芮�����
$popup.popup("�����ݒ肪�������܂���`r`n�����������̍쐬���s���Ă�������", 0, "�����ݒ肪�������܂���", 64)| Out-Null

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