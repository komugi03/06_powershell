$kinmuhyoukaraKoguchi = @'
#
# �Ζ��\���珬����ʔ�������쐬����Powershell
# 
# �O����� : ���Ypowershell�Ɠ����t�H���_�Ƀt�H�[�}�b�g�ƃn���R���L�ڂ��ꂽ ������ʔ�E�o������Z���׏�Excel�t�@�C�� ��1���݂��邱��
#
# ���s�`�� : .\createInvoice.ps1 �Ζ��\Excel�t�@�C���@����Excel�t�@�C��
#
# �Ζ��\�̌`�� : <�Ј��ԍ�>_�Ζ��\_m��_<����>.xlsx
#

# ----------------- �֐���` ---------------------

# �Ζ��\�Ə�����ۑ������ɕ��āAExcel�𒆒f����֐�
function endExcel {
    # Excel�̏I��
    $excel.quit()
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
# ����1 : �����F
# ����2�ȍ~ : ���b�Z�[�W
function displaySharpMessage {
    # �ϐ��̏�����
    $maxLengths = 0
    for($i=1;$i -lt $Args.length;$i++){
        # ���b�Z�[�W�̒��ň�Ԓ������������擾����
        if( $maxLengths -lt $Args[$i].length){
            $maxLengths = $Args[$i].length
        }
    }
    # ���b�Z�[�W�̕\��
    Write-Host ("`r`n" + '#' * ($maxLengths*2+6) + "`r`n") -ForegroundColor $Args[0]
    for($i=1;$i -lt $Args.length;$i++){
        Write-Host ('�@�@' + $Args[$i] + "�@�@`r`n") -ForegroundColor $Args[0]
    }
    Write-Host ('#' * ($maxLengths*2+6) + "`r`n") -ForegroundColor $Args[0]
}

# -------------------- �又�� ----------------------------

#=====================================================================
########################## ���ӏ�����\���B���Ȃ��ꍇ�ɂ�Enter����������B
#=========================================================================

# ���ݓ������擾����
$thisMonth = (Get-Date).Month
$today = (Get-Date).Day

# ���ݓ�������쐬����ׂ��Ζ��\�̌����𔻒�
if ($today -le 24) {
    $month = $thisMonth -1
} else {
    $month = $thisMonth
}

# �����e���v�����擾
$koguchiTemplate = Get-ChildItem -Recurse -File |? Name -Match "������ʔ�E�o������Z���׏�_�e���v��.xlsx"
# �Y�������t�@�C���̌��m�F
if ($koguchiTemplate.Count -lt 1) {
    Write-Host "`r`n�Y�����鏬���t�@�C�������݂��܂���`r`n`r`n�_�E�����[�h�������Ă�������`r`n" -ForegroundColor Red
    exit
} elseif ($koguchiTemplate.Count -gt 1) {
    Write-Host "`r`n�Y�����鏬���t�@�C�����������܂�`r`n`r`n�_�E�����[�h�������Ă�������`r`n" -ForegroundColor Red
    exit
}

# �e���v���[�g���珬����ʔ�������쐬����
$koguchi = Join-Path $PWD "�쐬�����������׏�" | Join-Path -ChildPath "������ʔ�E�o������Z���׏�_�R�s�[��.xlsx"
Copy-Item -path $koguchiTemplate.FullName -Destination $koguchi

# �Ζ��\�t�@�C�����擾
$kinmuhyou = Get-ChildItem -Recurse -File |? Name -Match "[0-9]{3}_�Ζ��\_($month)��_.+"

# �Y���Ζ��\�t�@�C���̌��m�F
if ($kinmuhyou.Count -lt 1) {
    Write-Host "`r`n�Y������Ζ��\�t�@�C�������݂��܂���`r`n" -ForegroundColor Red
    exit
} elseif ($kinmuhyou.Count -gt 1) {
    Write-Host "`r`n�Y������Ζ��\�t�@�C�����������܂�`r`n" -ForegroundColor Red
    exit
}

# �������n�߂�O�ɁA�t�@�C���̑��݃`�F�b�N�ƃt�@�C�����̃`�F�b�N���s��
if ( $kinmuhyou.Name  -match "[0-9]{3}_�Ζ��\_([1-9]|1[12])��_.+\.xlsx" ) {
    Start-Sleep -milliSeconds 300

    try {
    # �Ζ��\�t�@�C���̃t���p�X�擾
    $kinmuhyouFullPath = $kinmuhyou.FullName 
    } catch [Exception] {
        # �Ζ��\�����݂��Ă��邩�`�F�b�N
        Write-Host ($month + "���̋Ζ��\�t�@�C�������݂��܂���B`r`n�_�E�����[�h���Ă�������`r`n") -ForegroundColor Red
        exit
    }

    displaySharpMessage "White" ([string]$month + " ���̏�����ʔ�������쐬���܂��B") "���΂炭���҂����������B"
}else {
    # �Ζ��\�t�@�C���̃t�H�[�}�b�g���Ⴄ�ꍇ�͏C��������
    Write-Host " ######### <�Ј��ԍ�>_�Ζ��\_m��_<����>.xlsx �̌`���Ƀt�@�C�������C�����Ă������� #########`r`n" -ForegroundColor Red
    exit
}

# Excel���N������
try {
    # �N������Excel�v���Z�X���擾
    $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
} catch {
    $excel = New-Object -ComObject "Excel.Application" 
}

# Excel�����b�Z�[�W�_�C�A���O��\�����Ȃ��悤�ɂ���
$excel.DisplayAlerts = $false
$excel.visible = $true

# �Ζ��\�u�b�N���J��
$kinmuhyouBook = $excel.workbooks.open($kinmuhyouFullPath)
$kinmuhyouSheet = $kinmuhyouBook.sheets( "$month"+'��')

# �����u�b�N���J��
$koguchiBook = $excel.workbooks.open($koguchi)
$koguchiSheet = $koguchiBook.sheets(1)


# ------------- �Ζ��\�̒��g�������ɃR�s�[���� ----------------

# ------------- �l��񗓂̃R�s�[ --------------

# �����̏c��J�E���^�[
$rowCounter = 11

# ���l�ɏ�����Ă���Ζ��n���Q�l�ɏ����ɋL��
for ($row = 14; $row -le 44; $row++) {
    # ���l���̕�����
    $workPlace = $kinmuhyouSheet.cells.item($row,27).text

    # �ݑ�x�݂̎��ȊO
    if ($workPlace -ne "" -and $workPlace -ne '�ݑ�') {
        # 1. �����̋L��
        $koguchiSheet.cells.item($rowCounter,2) = $month
        $koguchiSheet.cells.item($rowCounter,4) = $kinmuhyouSheet.cells.item($row,3).text

        # ------------- �ϐ���` ---------------
        # �K�p�Z��(��)
        $tekiyou = 6
        # ��ԃZ��(��)
        $kukan = 18
        # ��ʋ@�փZ��(��)
        $koutsukikan = 26
        # ���z(��)
        $kingaku = 30





        switch -regex ($workPlace) {
            "^�V�q��$" {
                 # 2. �K�p�̋L��
                $koguchiSheet.cells.item($rowCounter,$tekiyou).formula = "������c��"
                # 3. ��Ԃ̋L��
                $koguchiSheet.cells.item($rowCounter,$kukan).formula = "��쁩���c��"
                # 4. ��ʋ@�ւ̋L��
                $koguchiSheet.cells.item($rowCounter,$koutsukikan).formula = "������`r`nJR�R���"
                # 5. ���z�̋L��
                $koguchiSheet.cells.item($rowCounter,$kingaku).formula = "=376*2"
            }
            "^�����$"{
                # 2. �K�p�̋L��
                $koguchiSheet.cells.item($rowCounter,$tekiyou).formula = "����������"
                # 3. ��Ԃ̋L��
                $koguchiSheet.cells.item($rowCounter,$kukan).formula = "��쁩�������e���|�[�g"
                # 4. ��ʋ@�ւ̋L��
                $koguchiSheet.cells.item($rowCounter,$koutsukikan).formula = "������`r`nJR�鋞��`r`n��񂩂���"
                # 5. ���z�̋L��
                $koguchiSheet.cells.item($rowCounter,$kingaku).formula = "=681*2"
                # 6. 3�s�ȏ�̗�������ꍇ�͍s�̍�����ύX����
                $koguchiSheet.cells.item($rowCounter,1).rowheight = 20
            }
            "^�i��$"{
                # 2. �K�p�̋L��
                $koguchiSheet.cells.item($rowCounter,$tekiyou).formula = "������i��"
                # 3. ��Ԃ̋L��
                $koguchiSheet.cells.item($rowCounter,$kukan).formula = "��쁩���i��"
                # 4. ��ʋ@�ւ̋L��
                $koguchiSheet.cells.item($rowCounter,$koutsukikan).formula = "������`r`nJR�R���"
                # 5. ���z�̋L��
                $koguchiSheet.cells.item($rowCounter,$kingaku).formula = "=376*2"
            }
            "^�i��/�����$"{
                # 2. �K�p�̋L��
                $koguchiSheet.cells.item($rowCounter,$tekiyou).formula = "����i�쁨����ꁨ����"
                # 3. ��Ԃ̋L��
                $koguchiSheet.cells.item($rowCounter,$kukan).formula = "��쁨�i��`r`n�������e���|�[�g�����"
                # 4. ��ʋ@�ւ̋L��
                $koguchiSheet.cells.item($rowCounter,$koutsukikan).formula = "������`r`nJR�R���`r`n���C���{�[�o�X"
                # 5. ���z�̋L��
                $koguchiSheet.cells.item($rowCounter,$kingaku).formula = "=376+220+681"
                # 6. 3�s�ȏ�̗�������ꍇ�͍s�̍�����ύX����
                $koguchiSheet.cells.item($rowCounter,1).rowheight = 20
            }
            "^�����/�i��$"{
                # 2. �K�p�̋L��
                $koguchiSheet.cells.item($rowCounter,$tekiyou).formula = "�������ꁨ�i�쁨����"
                # 3. ��Ԃ̋L��
                $koguchiSheet.cells.item($rowCounter,$kukan).formula = "��쁨�����e���|�[�g`r`n���i�쁨���"
                # 4. ��ʋ@�ւ̋L��
                $koguchiSheet.cells.item($rowCounter,$koutsukikan).formula = "������`r`nJR�R���`r`n���C���{�[�o�X"
                # 5. ���z�̋L��
                $koguchiSheet.cells.item($rowCounter,$kingaku).formula = "=681+220+376"
                # 6. 3�s�ȏ�̗�������ꍇ�͍s�̍�����ύX����
                $koguchiSheet.cells.item($rowCounter,1).rowheight = 20
            }
            # �ǂ��ɂ��Y�����Ȃ������ꍇ
            Default {
                displaySharpMessage "Red" ([string]$month + "��" + $kinmuhyouSheet.cells.item($row,3).text + "���̋Ζ��n���������F���ł��܂���ł����B") "����I����Ɋm�F���Ă�������"
            }
        }

        # �c��J�E���^�[�̃J�E���g�A�b�v
        $rowCounter = $rowCounter + 3
    }
}

# ------------- �l��񗓂̃R�s�[ --------------

# ���݂̔N���擾
$thisYear = (Get-Date).Year
# 1����12���̏�������낤�Ƃ��Ă�����N����N�߂�
if ($month -eq 1 -and (Get-Date).day -le 24) {
    $thisYear = (Get-Date).AddYears(-1).Year
}

# 1. �N�����̃R�s�[
$koguchiSheet.cells.item(78,4) = $thisYear
$koguchiSheet.cells.item(78,8) = $month

# ���̍ŏI������t���ɐݒ�
$koguchiSheet.cells.item(78,11) = (Get-Date "$thisYear/$month/1").AddMonths(1).AddDays(-1).Day

# 2. ���O�̃R�s�[
$koguchiSheet.cells.item(82,21) = $kinmuhyouSheet.cells.range("W7").text
# �Ζ��\�̖��O���󔒂������ꍇ�����𒆒f����
if ($koguchiSheet.cells.item(82,21).text -eq "") {
    Write-Host ("`r`n" + $month + "���̋Ζ��\�ɖ��O���L�ڂ���Ă��܂���`r`n�����𒆒f���܂�`r`n") -ForegroundColor Red
    endExcel
}

# 3. �����̃R�s�[
$affiliation = $kinmuhyouSheet.cells.range("W6").text
# "��" ���폜����
$affiliation -match "(?<affliationName>.+?)��" | Out-Null
$koguchiSheet.cells.item(80,6) = $Matches.affliationName
# �Ζ��\�̏������󔒂������ꍇ�����𒆒f����
if ($koguchiSheet.cells.item(80,6).text -eq "") {
    Write-Host ("`r`n" + $month + "���̋Ζ��\�ɏ������L�ڂ���Ă��܂���`r`n�����𒆒f���܂�`r`n") -ForegroundColor Red
    endExcel
}

# 4. ��ӂ̃R�s�[
# ��ӂ��Ȃ���������Ȃ��t���O
$haveNotStamp = $false
# �Ζ��\�̈�ӂ̂���Z�����N���b�v�{�[�h�ɃR�s�[
$kinmuhyouSheet.range("AA7").copy() | Out-Null
# �����V�[�g�Ɉ�ӂ��y�[�X�g
$koguchiCell=$koguchiSheet.range("AD82")
$koguchiSheet.paste($koguchiCell)
# �y�[�X�g���ҏW
$koguchiSheet.range("AD82").formula = ""
$koguchiSheet.range("AD82").interior.colorindex = 0
# �r����ҏW���邽�߂̐錾
$LineStyle = "microsoft.office.interop.excel.xlLineStyle" -as [type]
# �r�����Ȃ��ɂ���
$koguchiSheet.range("AD82").borders.linestyle = $linestyle::xllinestylenone
# ��Ӂi�I�u�W�F�N�g�j�������ĂȂ������Ȃ�A���b�Z�[�W��\������
$numberOfObject = 79
if ($koguchiSheet.shapes.count -eq $numberOfObject) {
    $haveNotStamp = $true
}

# �����F�̕ύX�i�S�����Ɂj
$koguchiSheet.range("A1:BN90").font.colorindex = 1

# ---------------- �I������ ------------------
# �V���������t�@�C����
$koguchiNewName = $kinmuhyou.name.Substring(0,3) + "_������ʔ�E�o������Z���׏�_" + $kinmuhyouSheet.cells.range("W7").text + ".xlsx"
# �t�@�C�������t�@�C�����Ƃ��Ďg����`�ɕҏW
$koguchiName -replace "�@",""�@-replace " ",""
$koguchiNewPath = Join-Path $PWD "�쐬�����������׏�" | Join-Path -ChildPath $koguchiNewName
$invalidChars = [IO.Path]::GetInvalidFileNameChars() -join ''
$re = "[{0}]" -f [RegEx]::Escape($invalidChars)
return ($Name -replace $re)
# Book�̕ۑ�
$koguchiBook.save()
# Book�����
$kinmuhyouBook.close()
$koguchiBook.close()
# Excel�̏I��
$excel.quit()
# �g�p���Ă����v���Z�X�̉��
$excel = $null
$kinmuhyouBook = $null
$kinmuhyouSheet = $null
$koguchiBook = $null
$koguchiSheet = $null
$koguchiCell = $null
[GC]::Collect()
# �쐬���������̃t�@�C�����ύX
Rename-Item -path $koguchi -NewName $koguchiNewPath

# ��ӂ��Ȃ���������Ȃ��ꍇ���ӊ��N
if ($haveNotStamp) {
    displaySharpMessage "Blue" "��ӂ��Ζ��\�ɓ����Ă��Ȃ��A�܂��͊���̃Z�����炸��Ă���\��������܂�" "�m�F���Ă�������"
}'@