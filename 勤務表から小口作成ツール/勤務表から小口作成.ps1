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
$nanngatsu = Read-Host '�����̏������쐬���܂����H(���p�����œ���)'


# =======3.�Ώۂ̋Ζ��\��INPUT�Ƃ��Ď󂯎��=======
$kinmuhyou = Get-ChildItem -Recurse | ? name -CMatch "[0-9]{3}_�Ζ��\_($nanngatsu)��_.+"

if($kinmuhyou -eq $null){
    echo ($nanngatsu + '���̋Ζ��\��p�ӂ��Ă�������')
} else {
    echo ($nanngatsu +'���̏������쐬���܂�')
}

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
echo ($koguchiTemple.name + ' ���e���v���[�g�Ƃ��܂�')

$koguchiTempleBook = $excel.workbooks.open($koguchiTemple.fullname)

# �����𕡐�
$koguchiFullpath = 'C:\Users\bvs20002\Documents\010_���K�̉�\06_powershell-lesson\�Ζ��\���珬���쐬�c�[��\�b�肾��.xlsx'
copy-item -Path $koguchiTemple.fullname -Destination $koguchiFullpath
$koguchiBook = $excel.workbooks.open($koguchiFullpath)

# �f�[�^�擾�ΏۃV�[�g���w�肷��
$kinmuSheet = $kinmuhyouBook.worksheets.item("$nanngatsu" + '��')
echo ('�u' + $kinmuSheet.name + '�v�V�[�g��ǂݍ���ł��܂�...')

$koguchiSheet = $koguchiBook.worksheets.item(1)
$koguchiMonthRow = 11

# �����A�c���̏ꍇ�̂ݏ������L��
# ���u�Ζ����e�vor�u���l�v���[�v�J�n��
for($row = 14; $row -le 15; $row++){

    # ����ꂪ�������珬���ɋL��
    if(($kinmuSheet.Cells.item($row,26).text -eq '�����') -or ($kinmuSheet.Cells.item($row,27).text -eq '�����')){
        
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
            # �����F����i���c�j������Ɓi�����j
            $koguchiSheet.Cells.item($koguchiMonthRow,6) = '����i���c�j������Ɓi�����j'

            # �u��ԁv�ɋL��
            $koguchiSheet.Cells.item($koguchiMonthRow,18) = '���c���������e���|�[�g'

            # �u��ʋ@�ցv�ɋL��
            $koguchiSheet.Cells.item($koguchiMonthRow,26) = "���c�}��`r`nJR�鋞��`r`n��񂩂���"

            # �u���z�v�ɋL��
            $koguchiSheet.Cells.item($koguchiMonthRow,30) = '1572'

        }

        $koguchiMonthRow = $koguchiMonthRow + 3

    }

    # �c�����������珬���ɋL��
    
        # �u���v�ɋL��
        # B11�A14�A17...�Ƀ��[�U�[�����͂����Ώی�������

        # �u���v�ɋL��
        # C����R�s�y

        # �u�K�p�i�s��A�v���j�v�ɋL��
        # �c���F����i���c�j?�c��
        # �����F����i���c�j?��Ɓi�����j

        # �u��ԁv�ɋL��

        # �u��ʋ@�ցv�ɋL��

        # �u���z�v�ɋL��

# �����[�v�I����
}

# 53�s�ڂ���Ȃ�������u�K�p�i�s��A�v���j�v�Ɂu�ȉ��]���v�L��


# H60�ɑΏی������

# K60�Ɍ����������


# book��ۑ�
$kinmuhyouBook.save()
$koguchiTempleBook.save()
$koguchiBook.save()

# �t�@�C�����ύX�̂��߂̏����W
$koguchiTempleBook.name -match '([0-9]{3}_������ʔ�E�o������Z���׏�_)(.+)_' | Out-Null
$gatsu = "{0:00}" -f [int]$nanngatsu
$rename = ($matches[1] + (get-date).year + $gatsu + '_' + $matches[2])


$kinmuhyouBook.close()
$koguchiTempleBook.close()
$koguchiBook.close()

# �t�@�C�����ύX
Rename-Item -Path '�b�肾��.xlsx' -NewName ($rename + '.xlsx')

# Excel�����

# �ϐ��̉��