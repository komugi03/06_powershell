Add-Type -AssemblyName Microsoft.office.Interop.Excel
try{
    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
} catch {
    # なければ新規を作る
    $excel = New-Object -ComObject Excel.Application
}

$excel.visible = $true

# 勤務表のテンプレをつかむ
$book = $excel.workbooks.open("C:\Users\bvs20002\Documents\010_自習の回\06_powershell-lesson\勤務表から小口作成ツール\めも.xlsx")

$sheet = $book.worksheets.item(2)

# ★ここが問題★
# 
$sheet.Range("B9").text

$row = 9

if($sheet.Range("B9").text -match "^.+\n.+\n.+\n.+"){

    # $sheet.Range("B9") | get-member
# $sheet.Range("B9").Columns.Autofit
$sheet.Range("B$row").RowHeight = 10

}


$book.save()
