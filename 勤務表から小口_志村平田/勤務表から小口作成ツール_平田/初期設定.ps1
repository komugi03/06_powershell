$kinmuhyoukaraKoguchi = @'
#
# 勤務表から小口交通費請求書を作成するPowershell
# 
# 前提条件 : 当該powershellと同じフォルダにフォーマットとハンコが記載された 小口交通費・出張旅費精算明細書Excelファイル が1つ存在すること
#
# 実行形式 : .\createInvoice.ps1 勤務表Excelファイル　小口Excelファイル
#
# 勤務表の形式 : <社員番号>_勤務表_m月_<氏名>.xlsx
#

# ----------------- 関数定義 ---------------------

# 勤務表と小口を保存せずに閉じて、Excelを中断する関数
function endExcel {
    # Excelの終了
    $excel.quit()
    # 使用していたプロセスの解放
    $excel = $null
    $kinmuhyouBook = $null
    $kinmuhyouSheet = $null
    $koguchiBook = $null
    $koguchiSheet = $null
    $koguchiCell = $null
    # ガベージコレクト
    [GC]::Collect()
    # 処理を終了する
    exit
}

# シャープを使ったメッセージの表示をする関数
# 引数1 : 文字色
# 引数2以降 : メッセージ
function displaySharpMessage {
    # 変数の初期化
    $maxLengths = 0
    for($i=1;$i -lt $Args.length;$i++){
        # メッセージの中で一番長い文字数を取得する
        if( $maxLengths -lt $Args[$i].length){
            $maxLengths = $Args[$i].length
        }
    }
    # メッセージの表示
    Write-Host ("`r`n" + '#' * ($maxLengths*2+6) + "`r`n") -ForegroundColor $Args[0]
    for($i=1;$i -lt $Args.length;$i++){
        Write-Host ('　　' + $Args[$i] + "　　`r`n") -ForegroundColor $Args[0]
    }
    Write-Host ('#' * ($maxLengths*2+6) + "`r`n") -ForegroundColor $Args[0]
}

# -------------------- 主処理 ----------------------------

#=====================================================================
########################## 注意書きを表示。問題ない場合にはEnterを押させる。
#=========================================================================

# 現在日時を取得する
$thisMonth = (Get-Date).Month
$today = (Get-Date).Day

# 現在日時から作成するべき勤務表の月次を判定
if ($today -le 24) {
    $month = $thisMonth -1
} else {
    $month = $thisMonth
}

# 小口テンプレを取得
$koguchiTemplate = Get-ChildItem -Recurse -File |? Name -Match "小口交通費・出張旅費精算明細書_テンプレ.xlsx"
# 該当小口ファイルの個数確認
if ($koguchiTemplate.Count -lt 1) {
    Write-Host "`r`n該当する小口ファイルが存在しません`r`n`r`nダウンロードし直してください`r`n" -ForegroundColor Red
    exit
} elseif ($koguchiTemplate.Count -gt 1) {
    Write-Host "`r`n該当する小口ファイルが多すぎます`r`n`r`nダウンロードし直してください`r`n" -ForegroundColor Red
    exit
}

# テンプレートから小口交通費請求書を作成する
$koguchi = Join-Path $PWD "作成した小口明細書" | Join-Path -ChildPath "小口交通費・出張旅費精算明細書_コピー先.xlsx"
Copy-Item -path $koguchiTemplate.FullName -Destination $koguchi

# 勤務表ファイルを取得
$kinmuhyou = Get-ChildItem -Recurse -File |? Name -Match "[0-9]{3}_勤務表_($month)月_.+"

# 該当勤務表ファイルの個数確認
if ($kinmuhyou.Count -lt 1) {
    Write-Host "`r`n該当する勤務表ファイルが存在しません`r`n" -ForegroundColor Red
    exit
} elseif ($kinmuhyou.Count -gt 1) {
    Write-Host "`r`n該当する勤務表ファイルが多すぎます`r`n" -ForegroundColor Red
    exit
}

# 処理を始める前に、ファイルの存在チェックとファイル名のチェックを行う
if ( $kinmuhyou.Name  -match "[0-9]{3}_勤務表_([1-9]|1[12])月_.+\.xlsx" ) {
    Start-Sleep -milliSeconds 300

    try {
    # 勤務表ファイルのフルパス取得
    $kinmuhyouFullPath = $kinmuhyou.FullName 
    } catch [Exception] {
        # 勤務表が存在しているかチェック
        Write-Host ($month + "月の勤務表ファイルが存在しません。`r`nダウンロードしてください`r`n") -ForegroundColor Red
        exit
    }

    displaySharpMessage "White" ([string]$month + " 月の小口交通費請求書を作成します。") "しばらくお待ちください。"
}else {
    # 勤務表ファイルのフォーマットが違う場合は修正させる
    Write-Host " ######### <社員番号>_勤務表_m月_<氏名>.xlsx の形式にファイル名を修正してください #########`r`n" -ForegroundColor Red
    exit
}

# Excelを起動する
try {
    # 起動中のExcelプロセスを取得
    $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
} catch {
    $excel = New-Object -ComObject "Excel.Application" 
}

# Excelがメッセージダイアログを表示しないようにする
$excel.DisplayAlerts = $false
$excel.visible = $true

# 勤務表ブックを開く
$kinmuhyouBook = $excel.workbooks.open($kinmuhyouFullPath)
$kinmuhyouSheet = $kinmuhyouBook.sheets( "$month"+'月')

# 小口ブックを開く
$koguchiBook = $excel.workbooks.open($koguchi)
$koguchiSheet = $koguchiBook.sheets(1)


# ------------- 勤務表の中身を小口にコピーする ----------------

# ------------- 個人情報欄のコピー --------------

# 小口の縦列カウンター
$rowCounter = 11

# 備考に書かれている勤務地を参考に小口に記入
for ($row = 14; $row -le 44; $row++) {
    # 備考欄の文字列
    $workPlace = $kinmuhyouSheet.cells.item($row,27).text

    # 在宅か休みの時以外
    if ($workPlace -ne "" -and $workPlace -ne '在宅') {
        # 1. 月日の記入
        $koguchiSheet.cells.item($rowCounter,2) = $month
        $koguchiSheet.cells.item($rowCounter,4) = $kinmuhyouSheet.cells.item($row,3).text

        # ------------- 変数定義 ---------------
        # 適用セル(横)
        $tekiyou = 6
        # 区間セル(横)
        $kukan = 18
        # 交通機関セル(横)
        $koutsukikan = 26
        # 金額(横)
        $kingaku = 30





        switch -regex ($workPlace) {
            "^新子安$" {
                 # 2. 適用の記入
                $koguchiSheet.cells.item($rowCounter,$tekiyou).formula = "自宅←→田町"
                # 3. 区間の記入
                $koguchiSheet.cells.item($rowCounter,$kukan).formula = "仙川←→田町"
                # 4. 交通機関の記入
                $koguchiSheet.cells.item($rowCounter,$koutsukikan).formula = "京王線`r`nJR山手線"
                # 5. 金額の記入
                $koguchiSheet.cells.item($rowCounter,$kingaku).formula = "=376*2"
            }
            "^お台場$"{
                # 2. 適用の記入
                $koguchiSheet.cells.item($rowCounter,$tekiyou).formula = "自宅←→お台場"
                # 3. 区間の記入
                $koguchiSheet.cells.item($rowCounter,$kukan).formula = "仙川←→東京テレポート"
                # 4. 交通機関の記入
                $koguchiSheet.cells.item($rowCounter,$koutsukikan).formula = "京王線`r`nJR埼京線`r`nりんかい線"
                # 5. 金額の記入
                $koguchiSheet.cells.item($rowCounter,$kingaku).formula = "=681*2"
                # 6. 3行以上の欄がある場合は行の高さを変更する
                $koguchiSheet.cells.item($rowCounter,1).rowheight = 20
            }
            "^品川$"{
                # 2. 適用の記入
                $koguchiSheet.cells.item($rowCounter,$tekiyou).formula = "自宅←→品川"
                # 3. 区間の記入
                $koguchiSheet.cells.item($rowCounter,$kukan).formula = "仙川←→品川"
                # 4. 交通機関の記入
                $koguchiSheet.cells.item($rowCounter,$koutsukikan).formula = "京王線`r`nJR山手線"
                # 5. 金額の記入
                $koguchiSheet.cells.item($rowCounter,$kingaku).formula = "=376*2"
            }
            "^品川/お台場$"{
                # 2. 適用の記入
                $koguchiSheet.cells.item($rowCounter,$tekiyou).formula = "自宅→品川→お台場→自宅"
                # 3. 区間の記入
                $koguchiSheet.cells.item($rowCounter,$kukan).formula = "仙川→品川`r`n→東京テレポート→仙川"
                # 4. 交通機関の記入
                $koguchiSheet.cells.item($rowCounter,$koutsukikan).formula = "京王線`r`nJR山手線`r`nレインボーバス"
                # 5. 金額の記入
                $koguchiSheet.cells.item($rowCounter,$kingaku).formula = "=376+220+681"
                # 6. 3行以上の欄がある場合は行の高さを変更する
                $koguchiSheet.cells.item($rowCounter,1).rowheight = 20
            }
            "^お台場/品川$"{
                # 2. 適用の記入
                $koguchiSheet.cells.item($rowCounter,$tekiyou).formula = "自宅→お台場→品川→自宅"
                # 3. 区間の記入
                $koguchiSheet.cells.item($rowCounter,$kukan).formula = "仙川→東京テレポート`r`n→品川→仙川"
                # 4. 交通機関の記入
                $koguchiSheet.cells.item($rowCounter,$koutsukikan).formula = "京王線`r`nJR山手線`r`nレインボーバス"
                # 5. 金額の記入
                $koguchiSheet.cells.item($rowCounter,$kingaku).formula = "=681+220+376"
                # 6. 3行以上の欄がある場合は行の高さを変更する
                $koguchiSheet.cells.item($rowCounter,1).rowheight = 20
            }
            # どこにも該当しなかった場合
            Default {
                displaySharpMessage "Red" ([string]$month + "月" + $kinmuhyouSheet.cells.item($row,3).text + "日の勤務地が正しく認識できませんでした。") "動作終了後に確認してください"
            }
        }

        # 縦列カウンターのカウントアップ
        $rowCounter = $rowCounter + 3
    }
}

# ------------- 個人情報欄のコピー --------------

# 現在の年を取得
$thisYear = (Get-Date).Year
# 1月に12月の小口を作ろうとしていたら年を一年戻す
if ($month -eq 1 -and (Get-Date).day -le 24) {
    $thisYear = (Get-Date).AddYears(-1).Year
}

# 1. 年月日のコピー
$koguchiSheet.cells.item(78,4) = $thisYear
$koguchiSheet.cells.item(78,8) = $month

# 月の最終日を日付欄に設定
$koguchiSheet.cells.item(78,11) = (Get-Date "$thisYear/$month/1").AddMonths(1).AddDays(-1).Day

# 2. 名前のコピー
$koguchiSheet.cells.item(82,21) = $kinmuhyouSheet.cells.range("W7").text
# 勤務表の名前が空白だった場合処理を中断する
if ($koguchiSheet.cells.item(82,21).text -eq "") {
    Write-Host ("`r`n" + $month + "月の勤務表に名前が記載されていません`r`n処理を中断します`r`n") -ForegroundColor Red
    endExcel
}

# 3. 所属のコピー
$affiliation = $kinmuhyouSheet.cells.range("W6").text
# "部" を削除する
$affiliation -match "(?<affliationName>.+?)部" | Out-Null
$koguchiSheet.cells.item(80,6) = $Matches.affliationName
# 勤務表の所属が空白だった場合処理を中断する
if ($koguchiSheet.cells.item(80,6).text -eq "") {
    Write-Host ("`r`n" + $month + "月の勤務表に所属が記載されていません`r`n処理を中断します`r`n") -ForegroundColor Red
    endExcel
}

# 4. 印鑑のコピー
# 印鑑がないかもしれないフラグ
$haveNotStamp = $false
# 勤務表の印鑑のあるセルをクリップボードにコピー
$kinmuhyouSheet.range("AA7").copy() | Out-Null
# 小口シートに印鑑をペースト
$koguchiCell=$koguchiSheet.range("AD82")
$koguchiSheet.paste($koguchiCell)
# ペースト先を編集
$koguchiSheet.range("AD82").formula = ""
$koguchiSheet.range("AD82").interior.colorindex = 0
# 罫線を編集するための宣言
$LineStyle = "microsoft.office.interop.excel.xlLineStyle" -as [type]
# 罫線をなしにする
$koguchiSheet.range("AD82").borders.linestyle = $linestyle::xllinestylenone
# 印鑑（オブジェクト）が増えてなさそうなら、メッセージを表示する
$numberOfObject = 79
if ($koguchiSheet.shapes.count -eq $numberOfObject) {
    $haveNotStamp = $true
}

# 文字色の変更（全部黒に）
$koguchiSheet.range("A1:BN90").font.colorindex = 1

# ---------------- 終了処理 ------------------
# 新しい小口ファイル名
$koguchiNewName = $kinmuhyou.name.Substring(0,3) + "_小口交通費・出張旅費精算明細書_" + $kinmuhyouSheet.cells.range("W7").text + ".xlsx"
# ファイル名をファイル名として使える形に編集
$koguchiName -replace "　",""　-replace " ",""
$koguchiNewPath = Join-Path $PWD "作成した小口明細書" | Join-Path -ChildPath $koguchiNewName
$invalidChars = [IO.Path]::GetInvalidFileNameChars() -join ''
$re = "[{0}]" -f [RegEx]::Escape($invalidChars)
return ($Name -replace $re)
# Bookの保存
$koguchiBook.save()
# Bookを閉じる
$kinmuhyouBook.close()
$koguchiBook.close()
# Excelの終了
$excel.quit()
# 使用していたプロセスの解放
$excel = $null
$kinmuhyouBook = $null
$kinmuhyouSheet = $null
$koguchiBook = $null
$koguchiSheet = $null
$koguchiCell = $null
[GC]::Collect()
# 作成した小口のファイル名変更
Rename-Item -path $koguchi -NewName $koguchiNewPath

# 印鑑がないかもしれない場合注意喚起
if ($haveNotStamp) {
    displaySharpMessage "Blue" "印鑑が勤務表に入っていない、または既定のセルからずれている可能性があります" "確認してください"
}'@