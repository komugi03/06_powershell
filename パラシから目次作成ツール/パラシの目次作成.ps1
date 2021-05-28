# 
# �p���V�̖ڎ�������ς�[������ 
# 


# �p�C�v���C������p���V���擾
# ���K�\���Ńt�B���^�����O
$pypeKaraUketori = $rowCountnput | ? Name -CMatch '�p���V���ǂ�_.+'

# ���łɂ���Excel�̃v���Z�X������
Add-Type -AssemblyName Microsoft.office.Interop.Excel
try{
    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
} catch {
    # �Ȃ���ΐV�K�����
    $excel = New-Object -ComObject Excel.Application
}

# ���ꂼ��̃p���V�ɑ΂��ď���
$pypeKaraUketori | %{

    # �t���p�X���擾
    $fullPath = $_.fullname
    $fullPath

    # Excel���J��
    $book = $excel.workbooks.open($fullPath)

    # 3�V�[�g�ڂ̑O�ɐV�����V�[�g��ǉ����A�u�ڎ��v�Ɩ��Â���
    $book.worksheets.add($book.sheets(3)) | Out-Null
    $book.sheets(3).name = "�ڎ�"
    $book.sheets(3).Range("A1") = "���o��"
    $book.sheets(3).Range("A1").font.bold = $true
    $book.sheets(3).Range("B1") = "�����o��"

    # �ォ�珇�ɏ������ނ��߂̕ϐ���p��
    $gyouCount = 2

    # �e�V�[�g�ɑ���
    for($sheetCount = 4; $sheetCount -le $book.worksheets.count;$sheetCount++){
    
        # �V�[�g���擾
        $sheet = $book.worksheets($sheetCount)
        echo ($sheet.name + " ���R�s�[���Ă��܂��B�B�B")
        
        # �e�V�[�g��B2�i���o���j�Z���̒��g���擾
        $midashi = $sheet.range("B2")

        # �ڎ��V�[�g��A��Ɍ��o���𒣂�t��
        $book.worksheets("�ڎ�").cells.item($gyouCount,1) = $midashi.text
        # ���o���𑾎��ɂ���
        $book.worksheets("�ڎ�").cells.item($gyouCount,1).font.bold = $true
        $gyouCount++

        # �e�V�[�g��C�񂩂珬���o�����擾
        for($rowCount = 1; $rowCount -le 100; $rowCount++){
            
            if($sheet.cells.item($rowCount,3).text -cmatch '^[0-9]{1,2}-[0-9]{1,2}'){
                $komidashi = $sheet.cells.item($rowCount,3).text
                 $book.worksheets("�ڎ�").cells.item($gyouCount,2) = $komidashi
                 $gyouCount++
            }
        }
    }

    # Excel��ۑ����ăN���[�Y
    $book.save()
    $book.close()

}

# Excel�����
$excel.quit()

# �ϐ��̉��
$excel = $null
$book = $null