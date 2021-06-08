#
# 勤務表から小口を作成するツール
# 
# お台場、田町の場合のみ小口を記入
# 15日分記入可能
#
# 用意するもの
# ・対象の勤務表
#       形式: <3桁の社員番号>_勤務表_m月_<名前>.xlsx
# ・小口計算のテンプレート
#       形式: <3桁の社員番号>_小口交通費・出張旅費精算明細書_<氏名>_テンプレ.xlsx
#       所属、氏名、印鑑は記入しておく
# 

# =======1.ユーザーに「何月のにします？」対話型で聞く=======
# =======2.対象月を入力=======
while($nanngatsu -notmatch '^([1-9]|1[0-2])$'){
    $nanngatsu = Read-Host '何月の小口を作成しますか？( ※ 半角数字で入力 ※ )'

    if($nanngatsu -match '^([1-9]|1[0-2])$'){
        break
    }elseif($nanngatsu -match '^[0-9]{1,2}$'){
        Write-Output @"
■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
        
    1~12月の間で入力してください
    OK: 4    NG: 04
        
■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
"@
    }else{
        Write-Output @"
■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■

    半角数字のみで入力してください
    OK: 4    NG: 4月

■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
"@
    }
}

# =======3.対象の勤務表をINPUTとして受け取る=======
$kinmuhyou = Get-ChildItem -Recurse | Where-Object name -CMatch "[0-9]{3}_勤務表_($nanngatsu)月_.+"

if(!($null -eq $kinmuhyou)){

} else {
    Write-Output @"
■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■

    $nanngatsu 月の勤務表が見つかりませんでした
    $nanngatsu 月の勤務表を用意し、ps1ファイルを実行しなおしてください

■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
"@
    # 処理を終了させる
    exit
}

# 現在の年でいいかを確認
$thisYear = (get-date).year

while(($nannnen -ne 'y') -or ($nannnen -ne 'n')){
    
    $nannnen = Read-Host "$thisYear 年 $nanngatsu 月でよろしいですか？ [ y or n ]"

    if($nannnen -eq 'y'){
        $targetYear = $thisYear
        break

    } elseif($nannnen -eq 'n') {
        $targetYear = Read-Host '年を入力してください( ※ 半角数字で入力 ※ )'
        break

    } else {
        Write-Output @"

    y もしくは n を入力してください

"@
    }
}

Write-Output @"
■■-------------------------------------------

    それでは    
    $thisYear 年 $nanngatsu 月の小口を作成します

-------------------------------------------■■
"@

# =======4.小口に入力=======
# すでにあるExcelのプロセスをつかむ
Add-Type -AssemblyName Microsoft.office.Interop.Excel
try{
    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
} catch {
    # なければ新規を作る
    $excel = New-Object -ComObject Excel.Application
}

$excel.visible = $true

# 勤務表のテンプレをつかむ
$kinmuhyouBook = $excel.workbooks.open($kinmuhyou.fullname)

# 小口のテンプレをつかむ
$koguchiTemple = Get-ChildItem -Recurse | ? name -CMatch '[0-9]{3}_小口交通費・出張旅費精算明細書_.+_テンプレ'
$koguchiTempleBook = $excel.workbooks.open($koguchiTemple.fullname)

# 小口を複製
$koguchiFullpath = 'C:\Users\bvs20002\Documents\010_自習の回\06_powershell-lesson\勤務表から小口作成ツール\作業中.xlsx'
copy-item -Path $koguchiTemple.fullname -Destination $koguchiFullpath
$koguchiBook = $excel.workbooks.open($koguchiFullpath)

# データ取得対象シートを指定する
$kinmuSheet = $kinmuhyouBook.worksheets.item("$nanngatsu" + '月')
Write-Output @"

    作成中です...
    しばらくお待ちください...

"@

$koguchiSheet = $koguchiBook.worksheets.item(1)
$koguchiMonthRow = 11

# お台場、田町の場合のみ小口を記入
# ★「勤務内容」or「備考」ループ開始★
for($row = 14; $row -le 44; $row++){

    # 空白でないかつ在宅以外
    if((($kinmuSheet.Cells.item($row,26).text -ne '在宅') -and !([String]::IsNullOrEmpty($kinmuSheet.Cells.item($row,26).text))) -or ((($kinmuSheet.Cells.item($row,27).text -ne '在宅') -and !([String]::IsNullOrEmpty($kinmuSheet.Cells.item($row,27).text))))){
        
        # お台場があったら小口に記入
        if(($kinmuSheet.Cells.item($row,26).text -eq 'お台場') -or ($kinmuSheet.Cells.item($row,27).text -eq 'お台場')){
            
            # 空白なら記入、埋まってたら下の段に移動する
            if($koguchiSheet.Cells.item($koguchiMonthRow,2).text -eq ""){

                # 「月」に記入
                # B11、14、17...にユーザーが入力した対象月を入れる
                $koguchiSheet.Cells.item($koguchiMonthRow,2) = $nanngatsu

                # 「日」に記入
                # 勤務表のC列をコピペ
                $koguchiSheet.Cells.item($koguchiMonthRow,4) = $kinmuSheet.Cells.item($row,3).text

                # 「適用（行先、要件）」に記入
                # 田町：自宅（生田）←→田町
                # お台場：自宅（生田）←→作業（お台場）
                $koguchiSheet.Cells.item($koguchiMonthRow,6) = '自宅（生田）←→作業（お台場）'

                # 「区間」に記入
                $koguchiSheet.Cells.item($koguchiMonthRow,18) = '生田←→東京テレポート'

                # 「交通機関」に記入
                $koguchiSheet.Cells.item($koguchiMonthRow,26) = "小田急線`r`nJR埼京線`r`nりんかい線`r`nレインボーバス"
                
                # 4行以上なら交通機関の行幅を増やす(5行目までなら読める高さ)
                if($koguchiSheet.Cells.item($koguchiMonthRow,26).text -match "^.+\n.+\n.+\n.+"){
                    $koguchiSheet.Range("Z$koguchiMonthRow").RowHeight = 40
                }

                # 「金額」に記入
                $koguchiSheet.Cells.item($koguchiMonthRow,30) = '1572'

            }

            $koguchiMonthRow = $koguchiMonthRow + 3

        }

        # 田町があったら小口に記入
        # お台場があったら小口に記入
        elseif(($kinmuSheet.Cells.item($row,26).text -eq '田町') -or ($kinmuSheet.Cells.item($row,27).text -eq '田町')){
            
            # ☆空白なら記入、埋まってたら下の段に移動する☆
            if($koguchiSheet.Cells.item($koguchiMonthRow,2).text -eq ""){

                # 「月」に記入
                # B11、14、17...にユーザーが入力した対象月を入れる
                $koguchiSheet.Cells.item($koguchiMonthRow,2) = $nanngatsu

                # 「日」に記入
                # 勤務表のC列をコピペ
                $koguchiSheet.Cells.item($koguchiMonthRow,4) = $kinmuSheet.Cells.item($row,3).text

                # 「適用（行先、要件）」に記入
                # 田町：自宅（生田）←→田町
                $koguchiSheet.Cells.item($koguchiMonthRow,6) = '自宅（生田）←→田町'

                # 「区間」に記入
                $koguchiSheet.Cells.item($koguchiMonthRow,18) = '生田←→田町'

                # 「交通機関」に記入
                $koguchiSheet.Cells.item($koguchiMonthRow,26) = "小田急線`r`nJR山手線"
                
                # 4行以上なら交通機関の行幅を増やす(5行目までなら読める高さ)
                if($koguchiSheet.Cells.item($koguchiMonthRow,26).text -match "^.+\n.+\n.+\n.+"){
                    $koguchiSheet.Range("Z$koguchiMonthRow").RowHeight = 40
                }

                # 「金額」に記入
                $koguchiSheet.Cells.item($koguchiMonthRow,30) = '962'

            }

            # 行カウンタのカウントアップ
            $koguchiMonthRow = $koguchiMonthRow + 3

        }
        else{
            # 「月」に記入
            $koguchiSheet.Cells.item($koguchiMonthRow,2) = $nanngatsu

            # 「日」に記入
            $koguchiSheet.Cells.item($koguchiMonthRow,4) = $kinmuSheet.Cells.item($row,3).text

            # 行カウンタのカウントアップ
            $koguchiMonthRow = $koguchiMonthRow + 3
        }
    }

# ★ループ終了★
}

# 53行目じゃなかったら「適用（行先、要件）」に「以下余白」記入
if($koguchiMonthRow -lt 53){
    $koguchiSheet.Cells.item($koguchiMonthRow,6) = '以下余白'
}

$targetDateRow = 60

# D60に対象年を入力
$koguchiSheet.Cells.item($targetDateRow,4) = $targetYear

# H60に対象月
$koguchiSheet.Cells.item($targetDateRow,8) = $nanngatsu

# K60に月末日を入力
$koguchiSheet.Cells.item($targetDateRow,11) = (Get-Date -month $nanngatsu -day 1).AddMonths(1).AddDays(-1).day

# bookを保存
$kinmuhyouBook.save()
$koguchiTempleBook.save()
$koguchiBook.save()

# ----------------------ファイル名変更のための情報収集----------------------
# テンプレのファイル名をグループ化
$koguchiTempleBook.name -match '([0-9]{3}_小口交通費・出張旅費精算明細書_)(.+)_' | Out-Null
# 4 を 04 にするようなフォーマットに変更
$gatsu = "{0:00}" -f [int]$nanngatsu

# ファイル名の変更に使用する文字列を用意
# $matches[1]: <番号>_小口交通費・出張旅費精算明細書_
# $matches[2]: <氏名>
$rename = ($matches[1] + $thisYear + $gatsu + '_' + $matches[2])

# 同じ月の小口の存在チェック
if(Test-path ($rename +　'_[0-9]' + '.xlsx')){

    # 2つ以上存在してる場合
    # 119_小口交通費・出張旅費精算明細書_202104_松澤_1（数字）がある

    # 同じ月の小口のファイル名を取得(_1など数字がついている)
    $onajiFileName = Get-ChildItem -Recurse | Where-Object name -CMatch "[0-9]{3}_小口交通費・出張旅費精算明細書_.+_.+_"
    
    # 最大の数字を探す
    $splitBy_FileName = $onajiFileName -split "_"

    for($i = 4; $i -lt (($onajiFileName.count)*5); $i = $i + 5){

        # 「1.xlsx」の数字部分を抜き出してインクリメントできるように数字にする
        $fileNameCount = [int]($splitBy_FileName[$i].Substring(0,1))
        $fileNameCount
    }

    # ファイル名の末尾の数字部分をインクリメント
    $fileNameCount = $fileNameCount + 1


    # ファイル名の変更に使用する文字列を用意
    $rename = ($rename + '_' + $fileNameCount)

}elseif(Test-path ($rename + '.xlsx')){
    # if($matches[3] -match '[0-9]'){
        
        # すでに同じ月の小口が存在してる(1つだけ)
        # 119_小口交通費・出張旅費精算明細書_202104_松澤がある
        $rename = ($rename + '_1')
    
}

$kinmuhyouBook.close()
$koguchiTempleBook.close()
$koguchiBook.close()

# ファイル名変更
Rename-Item -Path '作業中.xlsx' -NewName ($rename + '.xlsx')

Write-Output @"
---------------------------------------------------------------------------

    お待たせしました！
    $rename.xlsx を作成しました

---------------------------------------------------------------------------
"@ 

# Excelを閉じる(その他に開いているExcelも閉じちゃうから要検討)

# 変数の解放