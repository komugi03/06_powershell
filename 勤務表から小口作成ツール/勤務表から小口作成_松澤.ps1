#
# �Ζ��\���珬�����쐬����c�[��
# 
# �����A�c���̏ꍇ�̂ݏ������L��
# 15�����L���\
#
# �p�ӂ������
# �E�Ώۂ̋Ζ��\
#       �`��: <3���̎Ј��ԍ�>_�Ζ��\_m��_<���O>.xlsx
# �E�����v�Z�̃e���v���[�g
#       �`��: <3���̎Ј��ԍ�>_������ʔ�E�o������Z���׏�_<����>_�e���v��.xlsx
#       �����A�����A��ӂ͋L�����Ă���
# 

# =======1.���[�U�[�Ɂu�����̂ɂ��܂��H�v�Θb�^�ŕ���=======
# =======2.�Ώی������=======
while($nanngatsu -notmatch '^([1-9]|1[0-2])$'){
    $nanngatsu = Read-Host '�����̏������쐬���܂����H( �� ���p�����œ��� �� )'

    if($nanngatsu -match '^([1-9]|1[0-2])$'){
        break
    }elseif($nanngatsu -match '^[0-9]{1,2}$'){
        Write-Output @"
����������������������������������������������������������������������������������������
        
    1~12���̊Ԃœ��͂��Ă�������
    OK: 4    NG: 04
        
����������������������������������������������������������������������������������������
"@
    }else{
        Write-Output @"
����������������������������������������������������������������������������������������

    ���p�����݂̂œ��͂��Ă�������
    OK: 4    NG: 4��

����������������������������������������������������������������������������������������
"@
    }
}

# =======3.�Ώۂ̋Ζ��\��INPUT�Ƃ��Ď󂯎��=======
$kinmuhyou = Get-ChildItem -Recurse | Where-Object name -CMatch "[0-9]{3}_�Ζ��\_($nanngatsu)��_.+"

if(!($null -eq $kinmuhyou)){

} else {
    Write-Output @"
������������������������������������������������������������������������������������������������������������������������������������������

    $nanngatsu ���̋Ζ��\��������܂���ł���
    $nanngatsu ���̋Ζ��\��p�ӂ��Aps1�t�@�C�������s���Ȃ����Ă�������

������������������������������������������������������������������������������������������������������������������������������������������
"@
    # �������I��������
    exit
}

# ���݂̔N�ł��������m�F
$thisYear = (get-date).year

while(($nannnen -ne 'y') -or ($nannnen -ne 'n')){
    
    $nannnen = Read-Host "$thisYear �N $nanngatsu ���ł�낵���ł����H [ y or n ]"

    if($nannnen -eq 'y'){
        $targetYear = $thisYear
        break

    } elseif($nannnen -eq 'n') {
        $targetYear = Read-Host '�N����͂��Ă�������( �� ���p�����œ��� �� )'
        break

    } else {
        Write-Output @"

    y �������� n ����͂��Ă�������

"@
    }
}

Write-Output @"
����-------------------------------------------

    ����ł�    
    $thisYear �N $nanngatsu ���̏������쐬���܂�

-------------------------------------------����
"@

# =======4.�����ɓ���=======
# ���łɂ���Excel�̃v���Z�X������
Add-Type -AssemblyName Microsoft.office.Interop.Excel
try{
    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
} catch {
    # �Ȃ���ΐV�K�����
    $excel = New-Object -ComObject Excel.Application
}

$excel.visible = $true

# �Ζ��\�̃e���v��������
$kinmuhyouBook = $excel.workbooks.open($kinmuhyou.fullname)

# �����̃e���v��������
$koguchiTemple = Get-ChildItem -Recurse | ? name -CMatch '[0-9]{3}_������ʔ�E�o������Z���׏�_.+_�e���v��'
$koguchiTempleBook = $excel.workbooks.open($koguchiTemple.fullname)

# �����𕡐�
$koguchiFullpath = 'C:\Users\bvs20002\Documents\010_���K�̉�\06_powershell-lesson\�Ζ��\���珬���쐬�c�[��\��ƒ�.xlsx'
copy-item -Path $koguchiTemple.fullname -Destination $koguchiFullpath
$koguchiBook = $excel.workbooks.open($koguchiFullpath)

# �f�[�^�擾�ΏۃV�[�g���w�肷��
$kinmuSheet = $kinmuhyouBook.worksheets.item("$nanngatsu" + '��')
Write-Output @"

    �쐬���ł�...
    ���΂炭���҂���������...

"@

$koguchiSheet = $koguchiBook.worksheets.item(1)
$koguchiMonthRow = 11

# �����A�c���̏ꍇ�̂ݏ������L��
# ���u�Ζ����e�vor�u���l�v���[�v�J�n��
for($row = 14; $row -le 44; $row++){

    # �󔒂łȂ����ݑ�ȊO
    if((($kinmuSheet.Cells.item($row,26).text -ne '�ݑ�') -and !([String]::IsNullOrEmpty($kinmuSheet.Cells.item($row,26).text))) -or ((($kinmuSheet.Cells.item($row,27).text -ne '�ݑ�') -and !([String]::IsNullOrEmpty($kinmuSheet.Cells.item($row,27).text))))){
        
        # ����ꂪ�������珬���ɋL��
        if(($kinmuSheet.Cells.item($row,26).text -eq '�����') -or ($kinmuSheet.Cells.item($row,27).text -eq '�����')){
            
            # �󔒂Ȃ�L���A���܂��Ă��牺�̒i�Ɉړ�����
            if($koguchiSheet.Cells.item($koguchiMonthRow,2).text -eq ""){

                # �u���v�ɋL��
                # B11�A14�A17...�Ƀ��[�U�[�����͂����Ώی�������
                $koguchiSheet.Cells.item($koguchiMonthRow,2) = $nanngatsu

                # �u���v�ɋL��
                # �Ζ��\��C����R�s�y
                $koguchiSheet.Cells.item($koguchiMonthRow,4) = $kinmuSheet.Cells.item($row,3).text

                # �u�K�p�i�s��A�v���j�v�ɋL��
                # �c���F����i���c�j�����c��
                # �����F����i���c�j������Ɓi�����j
                $koguchiSheet.Cells.item($koguchiMonthRow,6) = '����i���c�j������Ɓi�����j'

                # �u��ԁv�ɋL��
                $koguchiSheet.Cells.item($koguchiMonthRow,18) = '���c���������e���|�[�g'

                # �u��ʋ@�ցv�ɋL��
                $koguchiSheet.Cells.item($koguchiMonthRow,26) = "���c�}��`r`nJR�鋞��`r`n��񂩂���`r`n���C���{�[�o�X"
                
                # 4�s�ȏ�Ȃ��ʋ@�ւ̍s���𑝂₷(5�s�ڂ܂łȂ�ǂ߂鍂��)
                if($koguchiSheet.Cells.item($koguchiMonthRow,26).text -match "^.+\n.+\n.+\n.+"){
                    $koguchiSheet.Range("Z$koguchiMonthRow").RowHeight = 40
                }

                # �u���z�v�ɋL��
                $koguchiSheet.Cells.item($koguchiMonthRow,30) = '1572'

            }

            $koguchiMonthRow = $koguchiMonthRow + 3

        }

        # �c�����������珬���ɋL��
        # ����ꂪ�������珬���ɋL��
        elseif(($kinmuSheet.Cells.item($row,26).text -eq '�c��') -or ($kinmuSheet.Cells.item($row,27).text -eq '�c��')){
            
            # ���󔒂Ȃ�L���A���܂��Ă��牺�̒i�Ɉړ����遙
            if($koguchiSheet.Cells.item($koguchiMonthRow,2).text -eq ""){

                # �u���v�ɋL��
                # B11�A14�A17...�Ƀ��[�U�[�����͂����Ώی�������
                $koguchiSheet.Cells.item($koguchiMonthRow,2) = $nanngatsu

                # �u���v�ɋL��
                # �Ζ��\��C����R�s�y
                $koguchiSheet.Cells.item($koguchiMonthRow,4) = $kinmuSheet.Cells.item($row,3).text

                # �u�K�p�i�s��A�v���j�v�ɋL��
                # �c���F����i���c�j�����c��
                $koguchiSheet.Cells.item($koguchiMonthRow,6) = '����i���c�j�����c��'

                # �u��ԁv�ɋL��
                $koguchiSheet.Cells.item($koguchiMonthRow,18) = '���c�����c��'

                # �u��ʋ@�ցv�ɋL��
                $koguchiSheet.Cells.item($koguchiMonthRow,26) = "���c�}��`r`nJR�R���"
                
                # 4�s�ȏ�Ȃ��ʋ@�ւ̍s���𑝂₷(5�s�ڂ܂łȂ�ǂ߂鍂��)
                if($koguchiSheet.Cells.item($koguchiMonthRow,26).text -match "^.+\n.+\n.+\n.+"){
                    $koguchiSheet.Range("Z$koguchiMonthRow").RowHeight = 40
                }

                # �u���z�v�ɋL��
                $koguchiSheet.Cells.item($koguchiMonthRow,30) = '962'

            }

            # �s�J�E���^�̃J�E���g�A�b�v
            $koguchiMonthRow = $koguchiMonthRow + 3

        }
        else{
            # �u���v�ɋL��
            $koguchiSheet.Cells.item($koguchiMonthRow,2) = $nanngatsu

            # �u���v�ɋL��
            $koguchiSheet.Cells.item($koguchiMonthRow,4) = $kinmuSheet.Cells.item($row,3).text

            # �s�J�E���^�̃J�E���g�A�b�v
            $koguchiMonthRow = $koguchiMonthRow + 3
        }
    }

# �����[�v�I����
}

# 53�s�ڂ���Ȃ�������u�K�p�i�s��A�v���j�v�Ɂu�ȉ��]���v�L��
if($koguchiMonthRow -lt 53){
    $koguchiSheet.Cells.item($koguchiMonthRow,6) = '�ȉ��]��'
}

$targetDateRow = 60

# D60�ɑΏ۔N�����
$koguchiSheet.Cells.item($targetDateRow,4) = $targetYear

# H60�ɑΏی�
$koguchiSheet.Cells.item($targetDateRow,8) = $nanngatsu

# K60�Ɍ����������
$koguchiSheet.Cells.item($targetDateRow,11) = (Get-Date -month $nanngatsu -day 1).AddMonths(1).AddDays(-1).day

# book��ۑ�
$kinmuhyouBook.save()
$koguchiTempleBook.save()
$koguchiBook.save()

# ----------------------�t�@�C�����ύX�̂��߂̏����W----------------------
# �e���v���̃t�@�C�������O���[�v��
$koguchiTempleBook.name -match '([0-9]{3}_������ʔ�E�o������Z���׏�_)(.+)_' | Out-Null
# 4 �� 04 �ɂ���悤�ȃt�H�[�}�b�g�ɕύX
$gatsu = "{0:00}" -f [int]$nanngatsu

# �t�@�C�����̕ύX�Ɏg�p���镶�����p��
# $matches[1]: <�ԍ�>_������ʔ�E�o������Z���׏�_
# $matches[2]: <����>
$rename = ($matches[1] + $thisYear + $gatsu + '_' + $matches[2])

# �������̏����̑��݃`�F�b�N
if(Test-path ($rename +�@'_[0-9]' + '.xlsx')){

    # 2�ȏ㑶�݂��Ă�ꍇ
    # 119_������ʔ�E�o������Z���׏�_202104_���V_1�i�����j������

    # �������̏����̃t�@�C�������擾(_1�Ȃǐ��������Ă���)
    $onajiFileName = Get-ChildItem -Recurse | Where-Object name -CMatch "[0-9]{3}_������ʔ�E�o������Z���׏�_.+_.+_"
    
    # �ő�̐�����T��
    $splitBy_FileName = $onajiFileName -split "_"

    for($i = 4; $i -lt (($onajiFileName.count)*5); $i = $i + 5){

        # �u1.xlsx�v�̐��������𔲂��o���ăC���N�������g�ł���悤�ɐ����ɂ���
        $fileNameCount = [int]($splitBy_FileName[$i].Substring(0,1))
        $fileNameCount
    }

    # �t�@�C�����̖����̐����������C���N�������g
    $fileNameCount = $fileNameCount + 1


    # �t�@�C�����̕ύX�Ɏg�p���镶�����p��
    $rename = ($rename + '_' + $fileNameCount)

}elseif(Test-path ($rename + '.xlsx')){
    # if($matches[3] -match '[0-9]'){
        
        # ���łɓ������̏��������݂��Ă�(1����)
        # 119_������ʔ�E�o������Z���׏�_202104_���V������
        $rename = ($rename + '_1')
    
}

$kinmuhyouBook.close()
$koguchiTempleBook.close()
$koguchiBook.close()

# �t�@�C�����ύX
Rename-Item -Path '��ƒ�.xlsx' -NewName ($rename + '.xlsx')

Write-Output @"
---------------------------------------------------------------------------

    ���҂������܂����I
    $rename.xlsx ���쐬���܂���

---------------------------------------------------------------------------
"@ 

# Excel�����(���̑��ɊJ���Ă���Excel�������Ⴄ����v����)

# �ϐ��̉��