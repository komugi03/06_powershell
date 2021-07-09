# ---------------�A�Z���u���̓ǂݍ���---------------
Add-Type -Assembly System.Windows.Forms
Add-Type -AssemblyName System.Drawing
# # INPUT�̂��߂ɕK�v?
# [void][System.Reflection.Assembly]::Load("Microsoft.VisualBasic, Version=8.0.0.0, Culture=Neutral, PublicKeyToken=b03f5f7f11d50a3a")

# �t�H�[���S�̂̐ݒ������֐�
# formText : �t�H�[���̖{���i������j
# formYoko : �t�H�[���̉���
# formTate : �t�H�[���̏c��
function makeForm ($formText, $formYoko, $formTate) {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $formText
    $form.Size = New-Object System.Drawing.Size($formYoko, $formTate)
    $form.StartPosition = "CenterScreen"
    $form.font = $Font
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
    $Font = New-Object System.Drawing.Font("���C���I", 8)

    # �t�H�[���S�̂̐ݒ�
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "�쐬���鏬���̑Ώ۔N��"
    $form.Size = New-Object System.Drawing.Size(265, 200)
    $form.StartPosition = "CenterScreen"
    $form.font = $Font

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
    # $Combo.font = $Font
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

echo "$targetYear �N��"
echo "$targetMonth ���̏������쐬���܂�"

# �|�b�v�A�b�v���쐬
$popup = new-object -comobject wscript.shell

# �t�@�C�����̋Ζ��\_�̂��Ƃ̕\�L
$fileNameMonth = [string]$targetMonth+"��"
# �Ζ��\�t�@�C�����擾
$kinmuhyou = Get-ChildItem -Recurse -File | ? Name -Match ("[0-9]{3}_�Ζ��\_"+$fileNameMonth+"_.+") 
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
    $excel = New-Object -ComObject "Excel.Application" 
}

# Excel�����b�Z�[�W�_�C�A���O��\�����Ȃ��悤�ɂ���
$excel.DisplayAlerts = $false
$excel.visible = $true

# �Ζ��\�̃t���p�X
$kinmuhyouFullPath = $kinmuhyou.FullName 

# �Ζ��\�u�b�N���J��
$kinmuhyouBook = $excel.workbooks.open($kinmuhyouFullPath)
$kinmuhyouSheet = $kinmuhyouBook.worksheets.item([String]$targetMonth + '��')

# ���͓��e���܂Ƃ߂ē���Ă���
$inputContentsArray = @()

# �Ζ����e�ƍ�Əꏊ�������Ă����Ζ��n�����Ȃ��悤�ɓ����Ă�
$workPlaceArray = @()

for ($Row = 14; $Row -le 44; $Row++) {
    $kinmunaiyou = $kinmuhyouSheet.cells.item($Row, 26).text
    $kinmujisseki = $kinmuhyouSheet.cells.item($Row, 7).text
    $sagyoubasho = $kinmuhyouSheet.cells.item(7, 7).text

    if ($kinmujisseki -ne "") {

        if ($kinmunaiyou -ne "") {
            # �Ζ����e����Ζ��n�������Ă���
            $workPlace = $kinmunaiyou        }
        else {
            # ��Əꏊ����Ζ��n�������Ă���
            $workPlace = $sagyoubasho
        }   
    }

    if (!$workPlaceArray.Contains($workPlace)) {
        $workPlaceArray += @($workPlace)
        
    }
}


# =========================== ���͉�� ===========================


# ---------------- �֐���` ----------------

# �t�H���g
$Font = New-Object System.Drawing.Font("�l�r �S�V�b�N",11)

# ���x����\��
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

# OK�{�^��
function drawOKButton {
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Point(20, $Args[0])
    $OKButton.Size = New-Object System.Drawing.Size(75, 30)
    $OKButton.Text = $Args[1]
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $Args[2].AcceptButton = $OKButton
    $Args[2].Controls.Add($OKButton)
}

# �ݑ�{�^��
function drawAtHomeButton {
    $AtHomeButton = New-Object System.Windows.Forms.Button
    $AtHomeButton.Location = New-Object System.Drawing.Point(130, $Args[0])
    $AtHomeButton.Size = New-Object System.Drawing.Size(75, 30)
    $AtHomeButton.Text = $Args[1]
    $AtHomeButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $Args[2].Controls.Add($AtHomeButton)
}

# �߂�{�^��
function drawReturnButton {
    $ReturnButton = New-Object System.Windows.Forms.Button
    $ReturnButton.Location = New-Object System.Drawing.Point(240, $Args[0])
    $ReturnButton.Size = New-Object System.Drawing.Size(75, 30)
    $ReturnButton.Text = $Args[1]
    $ReturnButton.DialogResult = [System.Windows.Forms.DialogResult]::Retry
    if ($i -eq 0) {
        $ReturnButton.Enabled = $false; 
    }else {
        $ReturnButton.Enabled = $True;
    }
    $Args[2].Controls.Add($ReturnButton)
}

# �e�L�X�g�{�b�N�X
function drawTextBox {
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point($Args[0], $Args[1])
    $textBox.Size = New-Object System.Drawing.Size($Args[2], $Args[3])
    $textBox.BackColor = "#ff99cc"
    $textBox.Text = $Args[4]
    $Args[5].Controls.Add($textBox)
    return $textBox
}

# �߂�{�^�������������p�l�ێ�
$tekiyouValue = ""
$kukanValue = ""
$koutsukikanValue = @()
$kingakuValue = ""


# ���͉�ʕ\��
for ($i = 0; $i -lt $workPlaceArray.Length; $i++) {

    # ---------------- Main ----------------- 

    # �t�H�[���S�̂̐ݒ�
    # �{�̂�����Ă���

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "�����ݒ�  �y"+$workPlaceArray[$i]+"�z"
    $form.Size = New-Object System.Drawing.Size(660, 700)
    $form.StartPosition = "CenterScreen"

    # OK�{�^���֐��Ăяo��
    drawOKButton 610 "OK" $form

    # �ݑ�{�^���֐��Ăяo��
    drawAtHomeButton 610 "�ݑ�" $form

    # �߂�{�^���֐��Ăяo��
    drawReturnButton 610 "�߂�" $form


    # =============================== input ===============================

    # �Ζ��n���u�ݑ�v�̏ꍇ�́u�ݑ�{�^���v����������
    $atHomeLabel = drawLabel 10 10 470 "�� �ݑ�Ζ��̂Ƃ��́y�ݑ�z�{�^�����N���b�N ��" $form
    $atHomeLabel.forecolor = "red" 
    $atHomeLabel.font = $Font 


    # ---------------- �K�p�i�s��A�v���j ----------------- 
    $tekiyouLabelLocate = 50
    $tekiyouTextBoxLocate = 108

    # ���x���֐��Ăяo��
    drawLabel 10 $tekiyouLabelLocate 470 ("�P�D�y �K�p �z �Ζ��n `""+$workPlaceArray[$i]+"`" �̎��̓K�p����͂��Ă�������") $form | Out-Null
    drawLabel 20 ($tekiyouLabelLocate + 20) 470 "ex.  ������c���{��" $form | Out-Null
    drawLabel 20 ($tekiyouLabelLocate + 40) 470 "      ����i�쁨�����e���|�[�g������ (�Ζ��n�����̏ꍇ)" $form | Out-Null


    # �e�L�X�g�{�b�N�X�֐��Ăяo��
    $outputTekiyou = drawTextBox 20 $tekiyouTextBoxLocate 300 20  $tekiyouValue $form

    # ---------------- ��� ----------------- 
    $kukanLabelLocate = 150
    $kukanTextBoxLocate = 208

    # ���x���֐��Ăяo��
    drawLabel  10 $kukanLabelLocate 550 ("�Q�D�y ��� �z �Ζ��n `""+$workPlaceArray[$i]+"`" �̎��̋�ԁi����̍Ŋ��w�����Ζ��n�̍Ŋ��w�j����͂��Ă�������") $form | Out-Null
    drawLabel 20 ($kukanLabelLocate + 20) 470 "ex.  <����̍Ŋ��w>�����c�� (�����̏ꍇ)" $form | Out-Null
    drawLabel 20 ($kukanLabelLocate + 40) 670 "      <����̍Ŋ��w>���i�쁨�����e���|�[�g��<����̍Ŋ��w> (�Ζ��n�����̏ꍇ)" $form | Out-Null


    # �e�L�X�g�{�b�N�X�֐��Ăяo��
    $outputKukan = drawTextBox 20 $kukanTextBoxLocate 430 20 $kukanValue $form



    # ---------------- ��ʋ@�� ----------------- 
    $koutsukikanLabelLocate = 290
    $koutsukikanTextBoxLocate = 288

    # ���x���֐��Ăяo��
    drawLabel 10 250 500 ("�R�D�y ��ʋ@�� �z �Ζ��n `""+$workPlaceArray[$i]+"`" �̎��ɗ��p�����ʋ@�ւ���͂��Ă�������") $form | Out-Null
    drawLabel 20 270 500 "ex. JR�R���" $form | Out-Null
    drawLabel 10 $koutsukikanLabelLocate 70 "��ʋ@�ւP�F" $form | Out-Null
    drawLabel 10 ($koutsukikanLabelLocate + 40) 70 "��ʋ@�ւQ�F" $form | Out-Null
    drawLabel 10 ($koutsukikanLabelLocate + 80) 70 "��ʋ@�ւR�F" $form | Out-Null
    drawLabel 10 ($koutsukikanLabelLocate + 120) 70 "��ʋ@�ւS�F" $form | Out-Null
    drawLabel 10 ($koutsukikanLabelLocate + 160) 70 "��ʋ@�ւT�F" $form | Out-Null
    drawLabel 10 ($koutsukikanLabelLocate + 200) 70 "��ʋ@�ւU�F" $form | Out-Null


    # �e�L�X�g�{�b�N�X�֐��Ăяo��
    $koutsukikan1 = drawTextBox 90 $koutsukikanTextBoxLocate 200 20 $koutsukikanValue[0] $form
    $koutsukikan2 = drawTextBox 90 ($koutsukikanTextBoxLocate + 40) 200 20 $koutsukikanValue[1] $form
    $koutsukikan3 = drawTextBox 90 ($koutsukikanTextBoxLocate + 80) 200 20 $koutsukikanValue[2] $form
    $koutsukikan4 = drawTextBox 90 ($koutsukikanTextBoxLocate + 120) 200 20 $koutsukikanValue[3] $form
    $koutsukikan5 = drawTextBox 90 ($koutsukikanTextBoxLocate + 160) 200 20 $koutsukikanValue[4] $form
    $koutsukikan6 = drawTextBox 90 ($koutsukikanTextBoxLocate + 200) 200 20 $koutsukikanValue[5] $form

    # ---------------- ���z -----------------
    $kingakuLabelLocate = 530
    $kingakuTextBoxLocate = 570

    # ���x���֐��Ăяo��
    drawLabel 10 $kingakuLabelLocate 500 ("�S�D�y ���z �z �Ζ��n `""+$workPlaceArray[$i]+"`" �̋��z�i��������j����͂��Ă�������") $form | Out-Null
    drawLabel 20 ($kingakuLabelLocate + 20) 470 "ex.  750 �i���p�����j" $form | Out-Null

    # �e�L�X�g�{�b�N�X�֐��Ăяo��
    $outputkingaku = drawTextBox 20 $kingakuTextBoxLocate 100 20 $kingakuValue $form



    # ����
    #$form.Add_Shown({$textBox.Select()})
    $inputContentsResult = $form.ShowDialog()




    # =============================== output ===============================

    if ($inputContentsResult -eq "OK") {

        # ---------------- �K�p�i�s��A�v���j ----------------- 

        $inputContentsArray += @($workPlaceArray[$i] + "_" + $outputTekiyou.text)

        # ---------------- ��� -----------------

        $inputContentsArray += @($workPlaceArray[$i] + "_" + $outputKukan.text)

        # ---------------- ��ʋ@�� ----------------- 
        $inputKoutsukikan = @($koutsukikan1.text, $koutsukikan2.text, $koutsukikan3.text, $koutsukikan4.text, $koutsukikan5.text, $koutsukikan6.text)

        foreach ($koutsukikan in $inputKoutsukikan) {
            if ([string]::IsNullOrEmpty($koutsukikan)) {
                # NULL �� '' �̏ꍇ
                Write-Host 'NULL or Empty'
            }
            else {
                # ��L�ȊO�͐ݒ肳�ꂽ��������o��
                $koutsukikans += $koutsukikan + '`r`n'
            }
        }

        $outputkoutsukikan = $koutsukikans.Substring(0, $koutsukikans.Length - 4)

        $inputContentsArray += @($workPlaceArray[$i] + "_" + $outputkoutsukikan)

        # ---------------- ���z -----------------

        $inputContentsArray += @($workPlaceArray[$i] + "_" + $outputkingaku.text)

        $tekiyouKeepValue = $outputTekiyou.text
        $kukanKeepValue = $outputKukan.text
        $koutsukikanKeepValue = @($koutsukikan1.text, $koutsukikan2.text, $koutsukikan3.text, $koutsukikan4.text, $koutsukikan5.text, $koutsukikan6.text)
        $kingakuKeepValue =$outputkingaku.text

        $tekiyouValue = ""
        $kukanValue = ""
        $koutsukikanValue = @()
        $kingakuValue = ""

    # �߂�{�^������������
    }elseif ($inputContentsResult -eq "Retry") {
        Write-Host "retry"
        $i = $i-2
        for ($j = 1; $j -lt 5; $j++) {
            $inputContentsArray[($inputContentsArray.Length-$j)] = ""
        }
        $tekiyouValue = $tekiyouKeepValue
        $kukanValue = $kukanKeepValue
        $koutsukikanValue = $koutsukikanKeepValue
        $kingakuValue = $kingakuKeepValue
        

    # �ݑ�{�^������������
    }elseif ($inputContentsResult -eq "Yes") {
        $inputContentsArray += @($workPlaceArray[$i] + "_1")
        $inputContentsArray += @($workPlaceArray[$i] + "_1")
        $inputContentsArray += @($workPlaceArray[$i] + "_1")
        $inputContentsArray += @($workPlaceArray[$i] + "_1")
    }else{
        break
    }    
}

# ������ʂ��ق����Ȃ�

foreach($inputContent in $inputContentsArray){
    $inputContent >> �c�[���p����.txt
}

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
$koutsukikan = $null
$koutsukikans = $null
$outputkoutsukikan = $null
$outputkingaku = $null

