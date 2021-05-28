# パラシから送られてきたオブジェクトの中でパラシだけ受け取る
$parashies = $input |? name -Match "パラシもどき"
 
# Excelを開く
$excel = New-Object -ComObject Excel.Application

 

# Excelを見えるようにする
$excel.visible = $true

 

# 繰り返し
foreach($parashi in $parashies){

 

  # パラシを開く
  $book = $excel.workbooks.Open($parashi.fullname)  

 

  # シートを追加する

 

  # 目次の縦列カウンタ
  $rowcount = 2

 

  # 目次作成（繰り返し）
  for($i = 4;$i -le 1000;$i++){

 

    # 大見出しをコピペ
    $book.sheets(3).cells.item($rowcount,2) = $book.sheets($i).cells.item(2,2)

 

    $rowcount++

 

    # 大見出しを太字にする
    $book.sheets(3).range("B1:B1000").font.bold = $true
    
    # 小見出しをコピペ（繰り返し）
    for($j = 1;$j -le 1000;$j++){
      if($book.sheets($i).cells.item($j,3) -cmatch '^[0-9]{1,2}-[0-9]{1,2}'){
        $book.sheets(3).cells.item($rowcount,3) = $book.sheets($i).cells.item($j,3)
        $rowcount++
    }
  }
}
}