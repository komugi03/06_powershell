# 
# �Ζ��\�����Ƃɏ�����ʔ�������쐬����Powershell
# 
# �Ζ��\�̃t�@�C�����F<3���̎Ј��ԍ�>_�Ζ��\_M��_<����>.xlsx
# 

# ---------------�A�Z���u���̓ǂݍ���---------------
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ----------------- �֐���` ---------------------

# �Ζ��\�Ə�����ۑ������ɕ��āAExcel�𒆒f����֐�
function breakExcel {
    # Book�����
    $kinmuhyouBook.close()
    $koguchiBook.close()
    Remove-Item -Path $koguchi
    # �g�p���Ă����v���Z�X�̉��
    $kinmuhyouBook = $null
    $kinmuhyouSheet = $null
    $koguchiBook = $null
    $koguchiSheet = $null
    $koguchiCell = $null
    # �K�x�[�W�R���N�g
    [GC]::Collect()
    # # �������I������
    # exit
}

# �����̋󔒂������t�@�C�����Ƃ��Ďg���Ȃ������������֐�
# fileName : �t�@�C����
function remove-invalidFileNameChars ($fileName) {
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

# �쐬���鏬���̔N���������Ă��邩�m�F����_�C�A���O��\��
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
    $label.Text = "�쐬�����������̔N����I�����Ă�������"
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
        $Combo.Text -match "(?<year>.+?)�N(?<month>.+?)��" | out-null

        # ���[�U�[�w��̔N�������쐬�̑Ώ۔N�Ƃ��ď㏑����
        $targetYear = $Matches.year

        # ���[�U�[�w��̌��������쐬�̑Ώی��Ƃ��ď㏑������
        $targetMonth = $Matches.month

    }else{
        # �������I������
        exit
    }

# ��$yesNo_yearMonthAreCorrect -eq 'No'���[�v�I����
}

Write-Host "$targetYear �N��"
Write-Host "$targetMonth ���̏������쐬���܂�"

# �|�b�v�A�b�v���쐬
$popup = new-object -comobject wscript.shell

# ----------------------�����e���v�����擾------------------------
$koguchiTemplate = Get-ChildItem -Recurse -File | ? Name -Match "������ʔ�E�o������Z���׏�_�e���v��.xlsx"
# �����e���v���̌��m�F
if ($koguchiTemplate.Count -lt 1) {
    # �|�b�v�A�b�v��\��
    $popup.popup("�����t�@�C���̃e���v���[�g�����݂��܂���`r`n�_�E�����[�h�������Ă�������",0,"��蒼���Ă�������",48) | Out-Null    
    exit
}
elseif ($koguchiTemplate.Count -gt 1) {
    # �|�b�v�A�b�v��\��
    $popup.popup("�����t�@�C���̃e���v���[�g���������܂�`r`n1�ɂ��Ă�������",0,"��蒼���Ă�������",48) | Out-Null
    exit
}

# -----------�쐬�����������i�[����t�H���_�ɁA�e���v���[�g���R�s�[����------------------

# �����i�[�t�H���_�����݂��Ă��Ȃ��ꍇ�͍쐬����
if(!(Test-Path $PWD"\�쐬����������ʔ����")){
    New-Item -Path $PWD"\�쐬����������ʔ����" -ItemType Directory | Out-Null
}

$koguchi = Join-Path $PWD "�쐬����������ʔ����" | Join-Path -ChildPath "������ʔ�E�o������Z���׏�_�R�s�[.xlsx"
Copy-Item -path $koguchiTemplate.FullName -Destination $koguchi

# ----------------�e���v���[�g���珬����ʔ�������쐬����---------------------

# �t�@�C�����̋Ζ��\_�̂��Ƃ̕\�L���uM���v�\�L�̏ꍇ
$fileNameMonth = [string]("$targetMonth" + "��")

# �����u�Ζ��\_YYYYMM�v�̂悤�ȕ\�L�ɂ���Ȃ� �� ���R�����g�A�E�g���� �� �̃R�����g�A�E�g���ʂ�
# $targetMonth00 = "{0:00}" -f [int]$targetMonth
# $fileNameMonth = ("$targetYear" + "$targetMonth00")

# �Ζ��\�t�@�C�����擾
$kinmuhyou = Get-ChildItem -Recurse -File | ? Name -Match ("[0-9]{3}_�Ζ��\_" + "$fileNameMonth" + "_.+")

# �Y���Ζ��\�t�@�C���̌��m�F
if ($kinmuhyou.Count -lt 1) {
    
    # �|�b�v�A�b�v��\��
    $popup.popup("$targetMonth ���̋Ζ��\�t�@�C�������݂��܂���",0,"��蒼���Ă�������",48) | Out-Null
    # �����̃e���v���̃R�s�[���폜����
    Remove-Item -Path $koguchi
    exit
}
elseif ($kinmuhyou.Count -gt 1) {
    # �|�b�v�A�b�v��\��
    $popup.popup("$targetMonth ���̋Ζ��\�t�@�C�����������܂�`r`n1�ɂ��Ă�������",0,"��蒼���Ă�������",48) | Out-Null
    # �����̃e���v���̃R�s�[���폜����
    Remove-Item -Path $koguchi
    exit
}


# --------------- �������̃v���O���X�o�[��\�� -------------

# �v���O���X�o�[�p�̃t�H�[����p��
$formProgressBar = New-Object System.Windows.Forms.Form
$formProgressBar.Size = "300,200"
$formProgressBar.Startposition = "CenterScreen"
$formProgressBar.Text = "�쐬���c"

# �v���O���X�o�[��p��
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = "10,100"
$progressBar.Size = "260,30"
$progressBar.Maximum = "10"
$progressBar.Minimum = "0"
$progressBar.Style = "Continuous"

# =========�v���O���X�o�[��i�߂�2/10 =======
$progressBar.Value = 2
$formProgressBar.Controls.AddRange($progressBar)
$formProgressBar.Topmost = $True
$formProgressBar.Show()


# displaySharpMessage "White" ([string]$targetMonth + " ���̏�����ʔ�������쐬���܂�") "���΂炭���҂����������B"

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

# �����u�b�N���J��
$koguchiBook = $excel.workbooks.open($koguchi)
$koguchiSheet = $koguchiBook.sheets(1)


# =========�v���O���X�o�[��i�߂�4/10 =======
$progressBar.Value += 2
$formProgressBar.Show()

# ------------- �Ζ��\�̒��g�������ɃR�s�[���� ----------------
# �u�Ζ����e�v���ɏ�����Ă���Ζ��n���Q�l�ɂ��āA�Ζ��n��񃊃X�g�e�L�X�g����Y�����������ɋL������

# �����̍s�J�E���^�[
$koguchiRowCounter = 11

# �Ζ��\��1���`�����܂�1�s���J��Ԃ�
for ($row = 14; $row -le 44; $row++) {
    # �Ζ��n����̂��߂Ɂu�Ζ����e�v���̕�������擾
    $workPlace = $kinmuhyouSheet.cells.item($row, 26).formula
    Write-Host ("�Ζ��n�F" + $workPlace)
    $workPlaceLength = [int]$workPlace.length + 1
    write-host ('$workPlace�ƁQ�̕������F' + $workPlaceLength)
    
    # �ݑ�x�݂̎��ȊO�̏ꍇ�A�����ɋL��
    if ($workPlace -ne "" -and $workPlace -ne '�ݑ�') {
        
        # ------------- �ϐ���` ---------------
        # �K�p(�J�n�ʒu)
        $tekiyou = 6
        # ���(�J�n�ʒu)
        $kukan = 18
        # ��ʋ@��(�J�n�ʒu)
        $koutsukikan = 26
        # ���z(�J�n�ʒu)
        $kingaku = 30
        
        # ---------------�Ζ��n��񃊃X�g��ǂݍ���---------------------
        # �Ζ��n��񃊃X�g�������Ă���e�L�X�g
        $infoTextFileName = "�c�[���p����.txt"
        $infoTextFileFullpath = "$PWD\$infoTextFileName"
        
        # �Ζ��n��񃊃X�g�e�L�X�g�����݂����Ƃ��̏���
        if(Test-Path $infoTextFileFullpath){
            
            $argumentText = (Get-Content $infoTextFileFullpath)
            
            # �u�Ζ����e�v���̕�����Ƀ}�b�`�����Ζ��n�̏����A���X�g����擾 ( �z��̒��g�@[0]:�K�p�@[1]:��ԁ@[2]:��ʋ@�ց@[3]:���z )
            $workPlaceInfo = $argumentText | Select-String -Pattern ($workPlace + '_')
            Write-Host ("�Ζ��nlist�F" + $workPlaceInfo)
            
            # �u�Ζ����e�v���̓��e���Ζ��̏�񃊃X�g�ɂȂ������ꍇ�A�|�b�v�A�b�v��\�����I������
            if($workPlaceInfo -eq $null){
                # �|�b�v�A�b�v��\��
                $popup.popup("�Ζ��n�̏�񂪓o�^����Ă��܂���`r`n�����ݒ�������͏㏑�����A��蒼���Ă�������",0,"��蒼���Ă�������",48) | Out-Null
                
                # �����𒆒f���A�I��
                breakExcel
                exit
                
            }
            
            # �ݑ�t���O(�K�p������1)�������Ă���ꍇ�A�����ɂ͋L�����Ȃ�
            elseif(([String]$workPlaceInfo[0]).Substring($workPlaceLength, ([String]$workPlaceInfo[0]).Length - $workPlaceLength) -eq '1'){
                # �����ɋL�����Ȃ�

                write-host "!!!!!!zaitaku!!!!!!"
            }
            
            # ��L�ȊO�̏ꍇ�A�����ɏ�������
            else{
                # �󔒂Ȃ�L���A���܂��Ă��牺�̒i�Ɉړ�����
                if($koguchiSheet.Cells.item($koguchiRowCounter,2).text -eq ""){
                    
                    # �u���v�ɋL��
                    # B11�A14�A17...�Ƀ��[�U�[�����͂����Ώی�������
                    $koguchiSheet.cells.item($koguchiRowCounter, 2) = $targetMonth
                    
                    # �u���v�ɋL��
                    # �Ζ��\��C����R�s�y
                    $koguchiSheet.cells.item($koguchiRowCounter, 4) = $kinmuhyouSheet.cells.item($row, 3).text
                    
                    # �u�K�p�i�s��A�v���j�v�ɋL��
                    $tekiyouText = ([String]$workPlaceInfo[0]).Substring($workPlaceLength, ([String]$workPlaceInfo[0]).Length - $workPlaceLength)
                    $koguchiSheet.Cells.item($koguchiRowCounter,$tekiyou) = $tekiyouText

                    # �u��ԁv�ɋL��
                    $kukanText = ([String]$workPlaceInfo[1]).Substring($workPlaceLength, ([String]$workPlaceInfo[1]).Length - $workPlaceLength)
                    $koguchiSheet.Cells.item($koguchiRowCounter,$kukan) = $kukanText

                    # �u��ʋ@�ցv�ɋL��

                    # �ŏ��̂����_����菜����������ɂ���
                    # ���c�}��`r`nJR�R���`r`n��񂩂����@�̏��
                    $koutsukikanText = ([String]$workPlaceInfo[2]).Substring($workPlaceLength, ([String]$workPlaceInfo[2]).Length - $workPlaceLength)
                    $koutsukikanArray = $koutsukikanText -split '`r`n'
                    
                    # �����ɋL�����镶������i�[����ϐ���p�ӂ��A����������
                    $koutsukikanKaigyou = $null

                    # �z��1�ȉ�����Ȃ��ԁA�J��Ԃ�
                    for ($i = 0; $i -lt $koutsukikanArray.Length; $i++) {
                        # ���s�R�[�h�𑫂�
                        $koutsukikanKaigyou += $koutsukikanArray[$i] + "`r`n"
                    }
                    
                    # �Ō�̉��s���폜����
                    $koutsukikanKaigyou = $koutsukikanKaigyou.Substring(0, $koutsukikanKaigyou.Length - 1)
                    $koguchiSheet.Cells.item($koguchiRowCounter,$koutsukikan) = $koutsukikanKaigyou
                    
                    # 4�s�ȏ�Ȃ��ʋ@�ւ̍s���𑝂₷(5�s�ڂ܂łȂ�ǂ߂鍂��)
                    if($koguchiSheet.Cells.item($koguchiRowCounter,$koutsukikan).text -match "^.+\n.+\n.+\n.+"){
                        $koguchiSheet.Range("Z$koguchiRowCounter").RowHeight = 40
                    }

                    # �u���z�v�ɋL��
                    $kingakuText = ([String]$workPlaceInfo[3]).Substring($workPlaceLength, ([String]$workPlaceInfo[3]).Length - $workPlaceLength)
                    $koguchiSheet.Cells.item($koguchiRowCounter,$kingaku) = $kingakuText

                }

                # �����̍s�J�E���^�[��3��ǉ����A���̍s�ɂ���
                $koguchiRowCounter = $koguchiRowCounter + 3

            }
            
        # �Ζ��n��񃊃X�g�e�L�X�g�����݂����Ƃ��̏����I��
        }else{
            # �|�b�v�A�b�v��\��
            $popup.popup("�Ζ��n�̏�񃊃X�g��������܂���`r`n��蒼���Ă�������",0,"��蒼���Ă�������",48) | Out-Null
        }
        
        # �u�Ζ����e�v������or�ݑ�̏����I��
    }

}

# =========�v���O���X�o�[��i�߂�6/10 =======
$progressBar.Value += 2
$formProgressBar.Show()

# ------------- �l��񗓂̃R�s�[ --------------
# --- �N�����̃R�s�[ ---
$koguchiSheet.cells.item(78, 4) = $targetYear
$koguchiSheet.cells.item(78, 8) = $targetMonth

# ���̍ŏI������t���ɐݒ�
$koguchiSheet.cells.item(78, 11) = [DateTime]::DaysInMonth($targetYear,$targetMonth)

# --- ���O�̃R�s�[ ---
$targetPersonName = $kinmuhyouSheet.cells.range("W7").text
$koguchiSheet.cells.item(82, 21) = $targetPersonName
# �Ζ��\�̖��O���󔒂������ꍇ�����𒆒f����
if ($koguchiSheet.cells.item(82, 21).text -eq "") {
    $popup.popup($targetMonth + "���̋Ζ��\�Ɂy���O�z���L�ڂ���Ă��܂���`r`n�����𒆒f���܂�",0,"��蒼���Ă�������",48) | Out-Null
    breakExcel
    exit
}

# --- �����̃R�s�[ ---
$affiliation = $kinmuhyouSheet.cells.range("W6").text
# "��" ���폜����
$affiliation -match "(?<affliationName>.+?)��" | Out-Null
$koguchiSheet.cells.item(80, 6) = $Matches.affliationName
# �Ζ��\�̏������󔒂������ꍇ�����𒆒f����
if ($koguchiSheet.cells.item(80, 6).text -eq "") {
    $popup.popup($targetMonth + "���̋Ζ��\�Ɂy�����z���L�ڂ���Ă��܂���`r`n�����𒆒f���܂�",0,"��蒼���Ă�������",48) | Out-Null
    breakExcel
    exit
}
# --- ��ӂ̃R�s�[ ---
# ��ӂ��R�s�y�������Z���̈ʒu
$targetStampCell = "AD82"

# ��ӂ��Ȃ���������Ȃ��t���O
$haveStamp = $true
# �Ζ��\�̈�ӂ̂���Z�����N���b�v�{�[�h�ɃR�s�[
$kinmuhyouSheet.range("AA7").copy() | Out-Null
# �����V�[�g�Ɉ�ӂ��y�[�X�g
$koguchiCell = $koguchiSheet.range($targetStampCell)
$koguchiSheet.paste($koguchiCell)
# �y�[�X�g���ҏW
$koguchiSheet.range($targetStampCell).formula = ""
$koguchiSheet.range($targetStampCell).interior.colorindex = 0
# �r����ҏW���邽�߂̐錾
$LineStyle = "microsoft.office.interop.excel.xlLineStyle" -as [type]
# �r�����Ȃ��ɂ���
$koguchiSheet.range($targetStampCell).borders.linestyle = $linestyle::xllinestylenone
# ��Ӂi�I�u�W�F�N�g�j�������ĂȂ������Ȃ�A���b�Z�[�W��\������
$numberOfObject = 79
if ($koguchiSheet.shapes.count -eq $numberOfObject) {
    $haveStamp = $false
}

# ��ӂ��Ȃ���������Ȃ��ꍇ���ӊ��N
if (!($haveStamp)) {

    $popup.popup("��ӂ��Ζ��\�ɓ����Ă��Ȃ�`r`n�܂��͈󂩂�啝�ɂ���Ă���\��������܂�`r`n��蒼���Ă�������",0,"��蒼���Ă�������",48) | Out-Null
    breakExcel
    exit

}

# �����F�̕ύX�i�S�����Ɂj
$koguchiSheet.range("A1:BN90").font.colorindex = 1

# �~�{�^�����������Ƃ��A�����r���̂��̂��폜���悤


# =========�v���O���X�o�[��i�߂�8/10 =======
$progressBar.Value += 2
$formProgressBar.Show()

# ---------------- �I������ ------------------

# ����1�� (ex 1��) �̏ꍇ2�� (ex 01) ��p�ӂ���
$fileNameMonth = "{0:D2}" -f [int]$targetMonth

# �����u�b�N�̕ۑ�
$koguchiBook.save()

# �Ζ��\�u�b�N�Ə����u�b�N�����
$kinmuhyouBook.close()
$koguchiBook.close()

# --------�V���������t�@�C������p��---------
# <�Ј��ԍ�>_������ʔ�E�o������Z���׏�_YYYYMM_<����>
$koguchiNewFileName = $kinmuhyou.name.Substring(0, 3) + "_������ʔ�E�o������Z���׏�_" + $targetYear + $fileNameMonth + "_" + $targetPersonName
# �t�@�C�����Ɏg���Ȃ������������Ă�����폜����(�����̊Ԃ̋󔒂Ȃ�)
$koguchiNewFileName = remove-invalidFileNameChars $koguchiNewFileName
# �V���������t�@�C���̃t���p�X
$koguchiNewfullPath = Join-Path $PWD "�쐬����������ʔ����" | Join-Path -ChildPath $koguchiNewFileName

# ------------�t�@�C������ύX----------------

# ���łɑΏی��̏���������Ă���Ƃ��̏���
# ��1���܂őΉ�
# if (Test-Path ($koguchiNewfullPath + "_$numberOfFiles.xlsx")) {
if (Test-Path ($koguchiNewfullPath + '_' + "[1-9]" + '.xlsx')) {

    
    # ------�Ώ۔N���̏�����2�ȏ㑶�݂��Ă�ꍇ--------
    # <�Ј��ԍ�>_������ʔ�E�o������Z���׏�_YYYYMM_<����>_<numberOfFiles>.xlsx�����݂���

    # �������̏����̃t�@�C�������擾(_1�Ȃǐ��������Ă���)
    $onajiFileName = Get-ChildItem -Recurse | Where-Object name -CMatch "[0-9]{3}_������ʔ�E�o������Z���׏�_.+_.+_"

    # �������̏����̃t�@�C������_�ŕ�����
    # [0]: <�Ј��ԍ�>
    # [1]: ������ʔ�E�o������Z���׏�
    # [2]: <���t>
    # [3]: <����>
    # [4]:�u1.xlsx�v�̐������� 
    $splitBy_FileName = $onajiFileName -split "_"
    
    # -----------�ő�̐�����T��--------------
    for($i = 4; $i -lt (($onajiFileName.count)*5); $i = $i + 5){

        # �u1.xlsx�v�̐��������𔲂��o���ăC���N�������g�ł���悤�ɐ����ɂ���
        $fileNameCountNumber = [int]($splitBy_FileName[$i].Substring(0,1))
        $fileNameCountNumber
        
        # ���������傫������������
        if($fileNameCount -lt $fileNameCountNumber){
            write-host "$fileNameCount ��"
            $fileNameCount = $fileNameCountNumber
            write-host "$fileNameCount �ɂ�����"

        }
        
    }

    # �t�@�C�����̖����̐����������C���N�������g
    $fileNameCount = $fileNameCount + 1

    # �t�@�C�����̕ύX�Ɏg�p���镶�����p��
    $koguchiNewFileName = ($koguchiNewFileName + '_' + $fileNameCount + '.xlsx')

} elseif (Test-Path ($koguchiNewfullPath + '.xlsx')) {
    
    # ------�Ώ۔N���̏�����1���݂��Ă�ꍇ--------
    # <�Ј��ԍ�>_������ʔ�E�o������Z���׏�_YYYYMM_<����>_<numberOfFiles>.xlsx�����݂��Ȃ�

    # �u_1.xlsx�v���t�@�C�����ɒǉ�����
    $koguchiNewFileName = $koguchiNewFileName + '_1.xlsx'

}else{
    # �g���q��ǉ�
    $koguchiNewFileName = $koguchiNewFileName + '.xlsx'
}


# �����t�@�C������ύX
Rename-Item -path $koguchi -NewName $koguchiNewFileName -ErrorAction:Stop


# =======�v���O���X�o�[�̏I��8/10========
$progressBar.Value += 2
$finish = $formProgressBar.Show()
$formProgressBar.Close()

# ����ɏI�������Ƃ��|�b�v�A�b�v��\��
$popup.popup("���҂������܂����I����ɏI�����܂���`r`n�d�オ����m�F���Ă�������",0,"����I��",64) | Out-Null    

# �g�p�����v���Z�X�̉��
$kinmuhyouBook = $null
$kinmuhyouSheet = $null
$koguchiBook = $null
$koguchiSheet = $null
$koguchiCell = $null
[GC]::Collect()


# �Ō�́u�J���v�u�I���v��2��
# �J�����ł����������Ƃ���̃G�N�X�v���[���[��\������