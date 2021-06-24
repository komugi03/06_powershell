Add-Type -AssemblyName Microsoft.office.Interop.Excel
try{
    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
} catch {
    # �Ȃ���ΐV�K�����
    $excel = New-Object -ComObject Excel.Application
}

$excel.visible = $true

# �Ζ��\�̃e���v��������
$book = $excel.workbooks.open("C:\Users\bvs20002\Documents\010_���K�̉�\06_powershell-lesson\�Ζ��\���珬���쐬�c�[��\�߂�.xlsx")

$sheet = $book.worksheets.item(2)

# ����������聚
# 
$sheet.Range("B9").text

$row = 9

if($sheet.Range("B9").text -match "^.+\n.+\n.+\n.+"){

    # $sheet.Range("B9") | get-member
# $sheet.Range("B9").Columns.Autofit
$sheet.Range("B$row").RowHeight = 10

}


$book.save()
