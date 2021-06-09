#
# �Ζ��\�����Ƃɏ������쐬����PowerShell
# 
# �Ζ��\�̃t�@�C����:<�Ј��ԍ�>_�Ζ��\_m��_<���O>.xlsx
#

# ���N
$thisYear = (Get-Date).Year

# �Θb�����Ői�߂Ă���

# �����̏������쐬���邩�ǂ����̌���
$month = $input | Read-Host "�����̏������쐬���܂����H[���p����]"

if (-not([int]::TryParse($month,[ref]$null))) {
    Write-host "���p�����ȊO�̒l�������Ă��邽�߁A�������I�����܂��B"
    break
}

# �Ζ��\���󂯎��
$kinmuhyou = gci -Recurse |? name -cmatch ("^[0-9]{2,3}_�Ζ��\_")

# --------�Ζ��\�̃t�@�C��������<���O>��<�Ј��ԍ�>�擾--------
$fileName = $kinmuhyou.name.Split("_")
# <���O>
$name_Extension = $fileName[3].Split(".")
$lastName = $name_Extension[0]
# <�Ј��ԍ�>
$employeeNumber = $fileName[0]

$koguchiFileName = $employeeNumber+"_������ʔ�E�o������Z���׏�_"+$month+"��_"+$lastName+".xlsx"
#---------------------------------------------------------

if ($kinmuhyou -cmatch $month+"��") {
    $YorN = Read-Host `r`n$lastName"�����"$thisYear"�N"$month"���̏������쐬���܂��B`r`n��낵���ł����H[Y/N]"
}else {
    Write-host `r`n$month"���̋Ζ��\��������܂���B"
    break
}

if ($YorN -match "n") {
    $thisYear = Read-Host "`r`n���N�̏������쐬���܂����H[���p����]"
    if (-not([int]::TryParse($month,[ref]$null))) {
        Write-host "���p�����ȊO�̒l�������Ă��邽�߁A�������I�����܂��B"
        break
    }
}
Write-host "`r`n*************************************************`r`n"`t$lastName"�����"$thisYear"�N"$month"���̏������쐬��...`r`n`t���΂炭���҂���������`r`n*************************************************"


# ------------------------�����ݒ�------------------------

#$moyoriStation = Read-Host "����̍Ŋ��w����͂��Ă��������B"

#$transportation = Read-Host "��ʋ@�ւ���͂��Ă��������B"

#$WorkPlace1_Fee = Read-Host "����-[$WorkPlace1]�Ԃ̌�ʔ�(������)����͂��Ă��������B[���p����]"

# --------------------------------------------------------

# Excel���J��
$excel = New-Object -ComObject Excel.Application

# Excel��������悤�ɂ���
$excel.visible = $true

# �Ζ��\���J��
$kinmuhyouBook = $excel.workbooks.Open($kinmuhyou.fullname)
$kinmuhyouSheet = $kinmuhyouBook.sheets($month+'��')


# �������J��
$koguchi = gci -Recurse |? name -cmatch '������ʔ�E�o������Z���׏�'
$koguchiBook = $excel.workbooks.Open($koguchi.fullname)
$koguchiSheet = $koguchiBook.sheets('������ʔ�')

# �Ζ��\�Ə������K�v�B�Ζ��\����K�v�ȏ��������ɓn���B�n���l�̉��H���K�v
# �u���l�v����Ζ��n���擾

# �s�J�E���g
$countRow = 11

for ($Row = 14; $Row -le 44; $Row++) {
        
    $workPlace = $kinmuhyouBook.sheets($month+'��').cells.item($Row,27).formula



    if (($workPlace -ne "") -And ($workPlace -ne "�ݑ�")){

        # �K�p�Z��
        $tekiyo = 6

        # ��ԃZ��
        $kukan = 18

        # ��ʋ@�փZ��
        $koutsukikan = 26

        # ���z�Z��
        $kingaku = 30

        $day = $kinmuhyouSheet.cells.item($Row,3).text


        switch -Regex ($workPlace) {
            "^�����$" {
                # ��
                $koguchiSheet.cells.item($countRow,2).formula = $month
                # ���t
                $koguchiSheet.cells.item($countRow,4).formula = $day
                # �K�p�i�s��A�v���j 
                $koguchiSheet.cells.item($countRow,$tekiyo).formula = "����������"
                # ���  ="$moyoriStation���������"
                $koguchiSheet.cells.item($countRow,$kukan).formula = "������ّO���������e���|�[�g"
                # ��ʋ@�� ="$transportation(���������ɂȂ�)"�����[�U�̓��͂����鎞�ǂ����悤�H
                $koguchiSheet.cells.item($countRow,$koutsukikan).formula = "�_�ޒ��o�X`r`n���c�}��`r`nJR�鋞��-��񂩂���" 
                # ���z ="$WorkPlace1_Fee"
                $koguchiSheet.cells.item($countRow,$kingaku).formula = "2640"
            }
            "^�c��$"{
                # ���t
                $koguchiSheet.cells.item($countRow,2).formula = $month
                $koguchiSheet.cells.item($countRow,4).formula = $day
                # �K�p�i�s��A�v���j 
                $koguchiSheet.cells.item($countRow,$tekiyo).formula = "������c��"
                # ���  ="$moyoriStation�����c��"
                $koguchiSheet.cells.item($countRow,$kukan).formula = "������ّO�����c��"
                # ��ʋ@�� ="$transportation(���������ɂȂ�)"�����[�U�̓��͂����鎞�ǂ����悤�H
                $koguchiSheet.cells.item($countRow,$koutsukikan).formula = "�_�ޒ��o�X`r`n���c�}��`r`nJR�R���" 
                # ���z ="$WorkPlace2_Fee"
                $koguchiSheet.cells.item($countRow,$kingaku).formula = "2030"
            }
            # �ǂ��ɂ��Y�����Ȃ������ꍇ
            Default {
                Write-Host "`r`n====================== ���� ======================`r`n"     $day"���̋Ζ��n�F[ "$workPlace" ] �͓o�^����Ă��܂���`r`n�o�^���Ă�蒼���Ă�������`r`n`r`n���͌��ݍ쐬���ꂽ�����ł�"$day"���̍s�͋󔒂ł��邽��`r`n"$day"���݂̂����g�ł��L����������`r`n=================================================="
                $koguchiSheet.cells.item($countRow,2).formula = $month
                $koguchiSheet.cells.item($countRow,4).formula = $day
            }
        }
        # �s�J�E���g�̃J�E���g
        $countRow = $countRow + 3
    }
}

# �L��������������Ō�Ɂu�ȉ��]���v���L��
$koguchiSheet.cells.item($countRow,$tekiyo).formula = "�ȉ��]��"

# �N�L��
$koguchiSheet.cells.item(66,4).formula = [string]$thisYear
# ���L��
$koguchiSheet.cells.item(66,8).formula = $month
# �������L��
$lastDayOftheMonth = [DateTime]::DaysInMonth($thisYear,$month)
$koguchiSheet.cells.item(66,11).formula = [string]$lastDayOftheMonth


# ���O��t���ĕۑ�
# �����̃t�@�C�����F<�Ј��ԍ�>_������ʔ�E�o������Z���׏�_m��_<���O>.xlsx
$koguchiBook.SaveAs("C:\Users\bvs20005\Documents\03_��Սu�K\PowerShell-Lesson\�Ζ��\���珬���쐬�c�[��\�쐬�ςݏ���\"+$koguchiFileName)

# Excel���I��
$Excel.Quit()

# �ϐ��̉��
$excel = $null
$koguchiBook = $null
$kinmuhyouBook = $null


