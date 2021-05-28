$pypeKaraUketori = $input | ? Name -CMatch 'パラシもどき_.+'

# すでにあるExcelのプロセスをつかむ
Add-Type -AssemblyName Microsoft.office.Interop.Excel
try{
    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
} catch {
    # なければ新規を作る
    $excel = New-Object -ComObject Excel.Application

}


$pypeKaraUketori | %{

    # フルパスを取得
    $fullPath = $_.fullname
    $fullPath

    # Excelを開く
    $book = $excel.workbooks.open($fullPath)

    for($sheetCount = 4; $sheetCount -lt $book.worksheets.count;$sheetCount++){
    
        # 各シートのB2（見出し）セルの中身を取得
        $sheet = $book.worksheets($sheetCount)
        $sheet.name

        # 各シートのB2（見出し）セルの中身を取得
        $midashi = $sheet.range("B2")
        $midashi.text


        
        # 各シートのC列から小見出しを取得
        for($i = 1; $i -le 10; $i++){
            
            if($sheet.cells.item($i,3).text -cmatch '^[0-9]{1,2}-[0-9]{1,2}'){
                $komidashi = $sheet.cells.item($i,3).text
                $komidashi
            }
        }
    }

    $book.save()
    $book.close()
}

$excel.quit()