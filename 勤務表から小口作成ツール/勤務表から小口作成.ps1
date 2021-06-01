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
$nanngatsu = Read-Host '何月の小口を作成しますか？(半角数字で入力)'


# =======3.対象の勤務表をINPUTとして受け取る=======
$kinmuhyou = Get-ChildItem -Recurse | ? name -CMatch "[0-9]{3}_勤務表_($nanngatsu)月_.+"

if($kinmuhyou -eq $null){
    echo ($nanngatsu + '月の勤務表を用意してください')
} else {
    echo ($nanngatsu +'月の小口を作成します')
}

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
echo ($koguchiTemple.name + ' をテンプレートとします')

$koguchiTempleBook = $excel.workbooks.open($koguchiTemple.fullname)

# 小口を複製
$koguchiFullpath = 'C:\Users\bvs20002\Documents\010_自習の回\06_powershell-lesson\勤務表から小口作成ツール\暫定だよ.xlsx'
copy-item -Path $koguchiTemple.fullname -Destination $koguchiFullpath
$koguchiBook = $excel.workbooks.open($koguchiFullpath)

# データ取得対象シートを指定する
$kinmuSheet = $kinmuhyouBook.worksheets.item("$nanngatsu" + '月')
echo ('「' + $kinmuSheet.name + '」シートを読み込んでいます...')

$koguchiSheet = $koguchiBook.worksheets.item(1)
$koguchiMonthRow = 11

# お台場、田町の場合のみ小口を記入
# ★「勤務内容」or「備考」ループ開始★
for($row = 14; $row -le 15; $row++){

    # お台場があったら小口に記入
    if(($kinmuSheet.Cells.item($row,26).text -eq 'お台場') -or ($kinmuSheet.Cells.item($row,27).text -eq 'お台場')){
        
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
            # お台場：自宅（生田）←→作業（お台場）
            $koguchiSheet.Cells.item($koguchiMonthRow,6) = '自宅（生田）←→作業（お台場）'

            # 「区間」に記入
            $koguchiSheet.Cells.item($koguchiMonthRow,18) = '生田←→東京テレポート'

            # 「交通機関」に記入
            $koguchiSheet.Cells.item($koguchiMonthRow,26) = "小田急線`r`nJR埼京線`r`nりんかい線"

            # 「金額」に記入
            $koguchiSheet.Cells.item($koguchiMonthRow,30) = '1572'

        }

        $koguchiMonthRow = $koguchiMonthRow + 3

    }

    # 田町があったら小口に記入
    
        # 「月」に記入
        # B11、14、17...にユーザーが入力した対象月を入れる

        # 「日」に記入
        # C列をコピペ

        # 「適用（行先、要件）」に記入
        # 田町：自宅（生田）?田町
        # お台場：自宅（生田）?作業（お台場）

        # 「区間」に記入

        # 「交通機関」に記入

        # 「金額」に記入

# ★ループ終了★
}

# 53行目じゃなかったら「適用（行先、要件）」に「以下余白」記入


# H60に対象月を入力

# K60に月末日を入力


# bookを保存
$kinmuhyouBook.save()
$koguchiTempleBook.save()
$koguchiBook.save()

# ファイル名変更のための情報収集
$koguchiTempleBook.name -match '([0-9]{3}_小口交通費・出張旅費精算明細書_)(.+)_' | Out-Null
$gatsu = "{0:00}" -f [int]$nanngatsu
$rename = ($matches[1] + (get-date).year + $gatsu + '_' + $matches[2])


$kinmuhyouBook.close()
$koguchiTempleBook.close()
$koguchiBook.close()

# ファイル名変更
Rename-Item -Path '暫定だよ.xlsx' -NewName ($rename + '.xlsx')

# Excelを閉じる

# 変数の解放