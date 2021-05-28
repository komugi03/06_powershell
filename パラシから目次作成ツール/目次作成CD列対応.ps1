# 
# パラシの目次CD列対応をつくるぱわーしぇる 
# 


# パイプラインからパラシを取得
# 正規表現でフィルタリング
$pypeKaraUketori = $input | ? Name -CMatch 'パラシもどき_.+'

# すでにあるExcelのプロセスをつかむ
Add-Type -AssemblyName Microsoft.office.Interop.Excel
try{
    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
} catch {
    # なければ新規を作る
    $excel = New-Object -ComObject Excel.Application
}

# それぞれのパラシに対して処理
$pypeKaraUketori | %{

    # フルパスを取得
    $fullPath = $_.fullname
    $fullPath

    # Excelを開く
    $book = $excel.workbooks.open($fullPath)

    # 3シート目の前に新しいシートを追加し、「目次CD列対応」と名づける
    $book.worksheets.add($book.sheets(3)) | Out-Null
    $book.sheets(3).name = "目次CD列対応"
    $book.sheets(3).Range("A1") = "見出し"
    $book.sheets(3).Range("A1").font.bold = $true
    $book.sheets(3).Range("B1") = "小見出し"

    # 上から順に書き込むための変数を用意
    $gyouCount = 2

    # 各シートに操作
    for($sheetCount = 4; $sheetCount -lt $book.worksheets.count;$sheetCount++){
    
        # シート名取得
        $sheet = $book.worksheets($sheetCount)
        echo ($sheet.name + " をコピーしています。。。")
        
        # 各シートのB2（見出し）セルの中身を取得
        $midashi = $sheet.range("B2")

        # 目次CD列対応シートのA列に見出しを張り付け
        $book.worksheets("目次CD列対応").cells.item($gyouCount,1) = $midashi.text
        # 見出しを太字にする
        $book.worksheets("目次CD列対応").cells.item($gyouCount,1).font.bold = $true
        $gyouCount++

        # 各シートのC列から小見出しを取得
        for($i = 1; $i -le 100; $i++){
            
            if($sheet.cells.item($i,3).text -cmatch '^[0-9]{1,2}-[0-9]{1,2}'){
                $komidashi = $sheet.cells.item($i,3).text
                 $book.worksheets("目次CD列対応").cells.item($gyouCount,2) = $komidashi
                 $gyouCount++
            }
            if($sheet.cells.item($i,4).text -cmatch '^[0-9]{1,2}-[0-9]{1,2}'){
                $komidashi = $sheet.cells.item($i,4).text
                 $book.worksheets("目次CD列対応").cells.item($gyouCount,2) = $komidashi
                 $gyouCount++
            }
        }
    }

    # Excelを保存してクローズ
    $book.save()
    $book.close()

}

# Excelを閉じる
$excel.quit()

# 変数の解放
$excel = $null
$book = $null