# 
# 勤務表をもとに小口交通費請求書を作成するPowershell
# 

# ---------------アセンブリの読み込み---------------
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ----------------- 関数定義 ---------------------

# 勤務表と小口を保存せずに閉じて、Excelを中断する関数
function breakExcel {
    # Bookを閉じる
    $kinmuhyouBook.close()
    $koguchiBook.close()
    Remove-Item -Path $koguchi
    # 使用していたプロセスの解放
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

# 引数の空白を除きファイル名として使えない文字を消す関数
# fileName : ファイル名
function remove-invalidFileNameChars ($fileName) {
    $fileNameRemovedSpace = $fileName -replace "　", "" -replace " ", ""
    $invalidChars = [IO.Path]::GetInvalidFileNameChars() -join ''
    $regex = "[{0}]" -f [RegEx]::Escape($invalidChars)
    return $fileNameRemovedSpace -replace $regex
}

# -------------------- 主処理の準備 --------------------------

# 現在の年月日を取得する
$targetYear = (Get-Date).Year
$thisMonth = (Get-Date).Month
$today = (Get-Date).Day

# 現在日時から作成するべき勤務表の月次を判定
# 24日までは先月分を作る
if ($today -le 24) {
    # 前の月を小口作成の対象月とする
    $targetMonth = (Get-date).AddMonths(-1).month
}
else {
    # 今月を小口作成の対象月とする
    $targetMonth = $thisMonth
}

# 作成する小口の年月が合っているか確認するダイアログを表示
# (現在日によって変わるので、get-date -Format Y にはしていない)
$yesNo_yearMonthAreCorrect = [System.Windows.Forms.MessageBox]::Show("作成するのは 【 $targetYear 年 $targetMonth 月 】の小口でよろしいですか？`r`n`r`n「いいえ」で他の月を選択できます",'作成する小口の対象年月','YesNo','Question')

# ☆$yesNo_yearMonthAreCorrect -eq 'No'ループ開始☆
if($yesNo_yearMonthAreCorrect -eq 'No'){

    # ---------------- 勤務表の月選択スクリプトを呼び出し -------------------
    . (Join-Path -Path $PWD -ChildPath "..\scripts\勤務表の月選択.ps1")
 
    # 勤務表の月選択.ps1の関数を実行
    # choicedMonth[0] : 勤務表の月選択フォーム
    # choicedMonth[1] : 勤務表の月選択コンボボックス
    $choicedMonth = choiceMonth
 
    # 勤務表の月選択画面を可視化
    $choiceMonthResult = $choicedMonth[0].ShowDialog()
 
    if ($choiceMonthResult -eq "OK") {
        # ユーザーの回答を"年"で区切る
        $choicedMonth[1].Text -match "(?<year>.+?)年(?<month>.+?)月" | out-null

        # ユーザー指定の年を小口作成の対象年として上書する
        $targetYear = $Matches.year

        # ユーザー指定の月を小口作成の対象月として上書きする
        $targetMonth = $Matches.month

    }else{
        # 処理を終了する
        exit
    }
}

# ポップアップを作成
$popup = new-object -comobject wscript.shell

# ----------------------小口テンプレを取得------------------------
$koguchiTemplate = Get-ChildItem -Recurse -File -path ..\ | ? Name -Match "^小口交通費・出張旅費精算明細書_テンプレ.xlsx$"
# 小口テンプレの個数確認
if ($koguchiTemplate.Count -lt 1) {
    # ポップアップを表示
    $popup.popup("小口ファイルのテンプレートが存在しません`r`nダウンロードし直してください",0,"やり直してください",48) | Out-Null    
    exit
}
elseif ($koguchiTemplate.Count -gt 1) {
    # ポップアップを表示
    $popup.popup("小口ファイルのテンプレートが多すぎます`r`n1つにしてください",0,"やり直してください",48) | Out-Null
    exit
}

# -----------作成した小口を格納するフォルダに、テンプレートをコピーする------------------

# 小口格納フォルダ名（変更があった場合はこの文字列を変更する）
$koguchiKakunousakiName = "04_作成済小口明細書"
$koguchiFullPath = join-path -Path $PWD -ChildPath ..\..\$koguchiKakunousakiName

# 小口格納フォルダが存在していない場合は作成する
if(!(Test-Path $koguchiFullPath)){
    New-Item -Path $koguchiFullPath -ItemType Directory | Out-Null
}

$koguchi = Join-Path -Path $koguchiFullPath -ChildPath "小口交通費出張旅費精算明細書_コピー.xlsx"
Copy-Item -path $koguchiTemplate.FullName -Destination $koguchi

# ----------------テンプレートから小口交通費請求書を作成する---------------------

# 勤務表ファイル名に使うYYYYMMの年月の形式を変数化
$targetMonth00 = "{0:00}" -f [int]$targetMonth
$fileNameMonth = ("$targetYear" + "$targetMonth00")

# 勤務表ファイルを取得
$kinmuhyou = Get-ChildItem -Recurse -File -path ..\..\ | ? Name -Match ("[0-9]{3}_勤務表_" + $fileNameMonth + "_.+")

# 該当勤務表ファイルの個数確認
if ($kinmuhyou.Count -lt 1) {
    
    # ポップアップを表示
    $popup.popup("$targetMonth 月の勤務表ファイルが存在しません",0,"やり直してください",48) | Out-Null
    # 小口のテンプレのコピーを削除する
    Remove-Item -Path $koguchi
    exit
}
elseif ($kinmuhyou.Count -gt 1) {
    # ポップアップを表示
    $popup.popup("$targetMonth 月の勤務表ファイルが多すぎます`r`n1つにしてください",0,"やり直してください",48) | Out-Null
    # 小口のテンプレのコピーを削除する
    Remove-Item -Path $koguchi
    exit
}

# --------------- 仮払いの有無を聞く -----------------
# フォーム全体の設定
$karibaraiForm = New-Object System.Windows.Forms.Form
$karibaraiForm.Text = "仮払いの有無"
$karibaraiForm.Size = New-Object System.Drawing.Size(265,200)
$karibaraiForm.StartPosition = "CenterScreen"
$karibaraiForm.formborderstyle = "FixedSingle"
$karibaraiForm.font = $Font
$karibaraiForm.icon = (Join-Path -Path $PWD -ChildPath "../images/会社アイコン.ico")

# ラベルを表示
$karibaraiLabel = New-Object System.Windows.Forms.Label
$karibaraiLabel.Location = New-Object System.Drawing.Point(10,10)
$karibaraiLabel.Size = New-Object System.Drawing.Size(200,30)
$karibaraiLabel.Text = "$targetYear 年 $targetMonth 月の仮払いがありますか？"
$karibaraiForm.Controls.Add($karibaraiLabel)

# ラベルを表示
$karibaraiLabel = New-Object System.Windows.Forms.Label
$karibaraiLabel.Location = New-Object System.Drawing.Point(10,51)
$karibaraiLabel.Size = New-Object System.Drawing.Size(70,20)
$karibaraiLabel.Text = "仮払い金額："
$karibaraiForm.Controls.Add($karibaraiLabel)

# テキストボックス
$karibaraiTextBox = New-Object System.Windows.Forms.TextBox 
$karibaraiTextBox.Location = New-Object System.Drawing.Point(80,50) 
$karibaraiTextBox.Size = New-Object System.Drawing.Size(100,100) 
$karibaraiForm.Controls.Add($karibaraiTextBox)

# OKボタンの設定
$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(40,100)
$OKButton.Size = New-Object System.Drawing.Size(75,30)
$OKButton.Text = "OK"
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$karibaraiForm.AcceptButton = $OKButton
$karibaraiForm.Controls.Add($OKButton)

# キャンセルボタンの設定
$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(130,100)
$CancelButton.Size = New-Object System.Drawing.Size(75,30)
$CancelButton.Text = "仮払いなし"
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$karibaraiForm.CancelButton = $CancelButton
$karibaraiForm.Controls.Add($CancelButton)

$yesNo_karibarai = $karibaraiForm.showDialog()

# キャンセルボタンか×ボタンが押されたら何もしない
if($yesNo_karibarai -eq 'Cancel'){
}
else{
    # OKボタンが押されたらテキストボックスの文字列を受け取る
    # テキストボックスが空なままOKボタンが押されたら
    for($yesNo_karibarai -eq 'OK'){

        # OKを押してしまったけれどやっぱり仮払いなかった場合
        if($yesNo_karibarai -eq 'Cancel'){
            break
        }
        elseif($karibaraiTextBox.text -eq ""){
            # エラー文上書きのためにサイズを0にする
            $errorLabel.size = New-Object System.Drawing.Size(0,0)
            # エラー文を表示
            $errorLabelKuuchi = New-Object System.Windows.Forms.Label
            $errorLabelKuuchi.Location = New-Object System.Drawing.Point(10,80)
            $errorLabelKuuchi.Size = New-Object System.Drawing.Size(270,50)
            $errorLabelKuuchi.Text = "仮払い金額を記入してください"
            $errorLabelKuuchi.ForeColor = "red"
            $errorLabelKuuchi.BringToFront()
            $karibaraiForm.Controls.Add($errorLabelKuuchi)
            $yesNo_karibarai = $karibaraiForm.showDialog()
        }
        # 半角数字かどうかの判定
        elseif(![int]::TryParse($karibaraiTextBox.text, [ref]$null)){
            write-host "半角数字じゃないよ"
            # エラー文上書きのためにサイズを0にする
            $errorLabelKuuchi.size = New-Object System.Drawing.Size(0,0)
            # エラー文を表示
            $errorLabel = New-Object System.Windows.Forms.Label
            $errorLabel.Location = New-Object System.Drawing.Point(10,80)
            $errorLabel.Size = New-Object System.Drawing.Size(270,50)
            $errorLabel.Text = "※半角数字で記入してください"
            $errorLabel.ForeColor = "blue"
            $karibaraiForm.Controls.Add($errorLabel)
            $yesNo_karibarai = $karibaraiForm.showDialog()
        }
        # 正常に半角数字が入力された場合はテキストボックスの文字列を取得してループを抜ける
        else{
            # ループを抜ける
            break
        }
    }
}

# --------------- 処理中のプログレスバーを表示 -------------

# プログレスバー用のフォームを用意
$formProgressBar = New-Object System.Windows.Forms.Form
$formProgressBar.Size = "300,200"
$formProgressBar.Startposition = "CenterScreen"
$formProgressBar.Text = "作成中…"
$formProgressBar.font = $Font
$formProgressBar.icon = (Join-Path -Path $PWD -ChildPath "../images/会社アイコン.ico")

# プログレスバー用のラベルを用意
$progressLabel = New-Object System.Windows.Forms.Label
$progressLabel.Location = New-Object System.Drawing.Point(10,40)
$progressLabel.Size = New-Object System.Drawing.Size(270,30)
$progressLabel.Text = "作成しています`r`nしばらくお待ちください"

# プログレスバーを用意
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = "10,100"
$progressBar.Size = "260,30"
$progressBar.Maximum = "10"
$progressBar.Minimum = "0"
$progressBar.Style = "Continuous"

# =========プログレスバーを進める2/10 =======
$progressBar.Value = 2
$formProgressBar.Controls.AddRange(@($progressBar,$progressLabel))
$formProgressBar.Topmost = $True
$formProgressBar.Show()


# ----------------------Excelを起動する--------------------------------
try {
    # 起動中のExcelプロセスを取得
    $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
}
catch {
    # Excelプロセスが起動してなければ新たに起動する
    $excel = New-Object -ComObject "Excel.Application" 
}

# Excelがメッセージダイアログを表示しないようにする
$excel.DisplayAlerts = $false
$excel.visible = $false

# =========プログレスバーを進める4/10 =======
$progressBar.Value += 2
$formProgressBar.Show()

# 勤務表のフルパス
$kinmuhyouFullPath = $kinmuhyou.FullName 

# 勤務表ブックを開く
$kinmuhyouBook = $excel.workbooks.open($kinmuhyouFullPath)
$kinmuhyouSheet = $kinmuhyouBook.worksheets.item([String]$targetMonth + '月')

# 小口ブックを開く
$koguchiBook = $excel.workbooks.open($koguchi)
$koguchiSheet = $koguchiBook.WorkSheets.item(1)


# =========プログレスバーを進める6/10 =======
$progressBar.Value += 2
$formProgressBar.Show()

# ------------- 仮払い金額を「仮払額」欄に入力する ----------------
# テキストボックスの文字列を取得し、小口の「仮払額」欄に入力する
$koguchiSheet.cells.item(6,2).formula = $karibaraiTextBox.text


# ------------- 勤務表の中身を小口にコピーする ----------------
# 「勤務内容」欄に書かれている勤務地を参考にして、勤務地情報リストテキストから該当情報を小口に記入する

# 小口の行カウンター
$koguchiRowCounter = 11

# 最終出勤日の行
$lastWorkDayRow = 14

# 勤務表の1日～月末まで1行ずつ繰り返す
for ($row = 14; $row -le 44; $row++) {
    # 勤務地判定のために「勤務内容」欄の文字列を取得
    $workPlace = $kinmuhyouSheet.cells.item($row, 26).formula
    
    # --- 出勤してるけど勤務内容に書いてない場合の処理 ---
    # 「勤務実績」欄の終了時刻の文字列を取得
    $kinmujisseki = $kinmuhyouSheet.cells.item($row, 7).text
    
    # 「作業場所」欄の文字列を取得
    $sagyoubasho = $kinmuhyouSheet.cells.item(7, 7).text
    
    # 出勤してるけど勤務内容に勤務地が書いてない場合
    if ($kinmujisseki -ne "" -and $workPlace -eq "") {
        # 作業場所に書いてある文字列を勤務地とする
        $workPlace = $sagyoubasho
    }
    # 在宅か休みの時以外の場合、小口に記入
    if ($workPlace -ne "" -and $workPlace -ne '在宅') {
        
        # ------------- 変数定義 ---------------
        # 適用(開始位置)
        $tekiyou = 6
        # 区間(開始位置)
        $kukan = 18
        # 交通機関(開始位置)
        $koutsukikan = 26
        # 金額(開始位置)
        $kingaku = 30
        # Substring()で取り除きたい、「勤務地_」の総文字数
        $workPlaceLength = [int]$workPlace.length + 1
        
        # ---------------勤務地情報リストを読み込む---------------------
        # 勤務地情報リストが書いてあるテキスト
        $infoTextFileName = "ツール用引数.txt"
        $infoTextFileFullpath = Join-Path -Path $PWD -ChildPath "..\user_info\$infoTextFileName"        
        # 勤務地情報リストテキストが存在したときの処理
        if(Test-Path $infoTextFileFullpath){
            
            $argumentText = (Get-Content $infoTextFileFullpath)
            
            # 「勤務内容」欄の文字列にマッチした勤務地の情報を、リストから取得 ( 配列の中身　[0]:適用　[1]:区間　[2]:交通機関　[3]:金額 )
            $workPlaceInfo = $argumentText | Select-String -Pattern ($workPlace + '_')
            
            # 「勤務内容」欄の内容が勤務の情報リストになかった場合、ポップアップを表示し終了する
            if($workPlaceInfo -eq $null){
                # ======= プログレスバーを閉じる =======
                $formProgressBar.Close()
                # ポップアップを表示
                $popup.popup("勤務地の情報が不足しています`r`n登録し直してください",0,"やり直してください",48) | Out-Null
                # 処理を中断し、終了
                breakExcel                
            }
            
            # 在宅フラグ(適用部分に1)が立っている場合、小口には記入しない
            elseif(([String]$workPlaceInfo[0]).Substring($workPlaceLength, ([String]$workPlaceInfo[0]).Length - $workPlaceLength) -eq '1'){
                # 小口に記入しない
            }
            
            # 上記以外の場合、小口に書き込む
            else{
                # 空白なら記入、埋まってたら下の段に移動する
                if($koguchiSheet.Cells.item($koguchiRowCounter,2).text -eq ""){
                    
                    # 「月」に記入
                    # B11、14、17...にユーザーが入力した対象月を入れる
                    $koguchiSheet.cells.item($koguchiRowCounter, 2) = $targetMonth
                    
                    # 「日」に記入
                    # 勤務表のC列をコピペ
                    $koguchiSheet.cells.item($koguchiRowCounter, 4) = $kinmuhyouSheet.cells.item($row, 3).text
                    
                    # 「適用（行先、要件）」に記入
                    $tekiyouText = ([String]$workPlaceInfo[0]).Substring($workPlaceLength, ([String]$workPlaceInfo[0]).Length - $workPlaceLength)
                    $koguchiSheet.Cells.item($koguchiRowCounter,$tekiyou) = $tekiyouText
                    
                    # 「区間」に記入
                    $kukanText = ([String]$workPlaceInfo[1]).Substring($workPlaceLength, ([String]$workPlaceInfo[1]).Length - $workPlaceLength)
                    $koguchiSheet.Cells.item($koguchiRowCounter,$kukan) = $kukanText
                    
                    # 「交通機関」に記入
                    
                    # 最初のお台場_を取り除いた文字列にする
                    # 小田急線`r`nJR山手線`r`nりんかい線　の状態
                    $koutsukikanText = ([String]$workPlaceInfo[2]).Substring($workPlaceLength, ([String]$workPlaceInfo[2]).Length - $workPlaceLength)
                    $koutsukikanArray = $koutsukikanText -split '`r`n'
                    
                    # 小口に記入する文字列を格納する変数を用意し、初期化する
                    $koutsukikanKaigyou = $null
                    
                    # 配列が1以下じゃない間、繰り返す
                    for ($i = 0; $i -lt $koutsukikanArray.Length; $i++) {
                        # 改行コードを足す
                        $koutsukikanKaigyou += $koutsukikanArray[$i] + "`r`n"
                    }
                    
                    # 最後の改行を削除する
                    $koutsukikanKaigyou = $koutsukikanKaigyou.Substring(0, $koutsukikanKaigyou.Length - 1)
                    $koguchiSheet.Cells.item($koguchiRowCounter,$koutsukikan) = $koutsukikanKaigyou
                    
                    # 4行以上なら交通機関の行幅を増やす(5行目までなら読める高さ)
                    if($koguchiSheet.Cells.item($koguchiRowCounter,$koutsukikan).text -match "^.+\n.+\n.+\n.+"){
                        $koguchiSheet.Range("Z$koguchiRowCounter").RowHeight = 40
                    }
                    
                    # 「金額」に記入
                    $kingakuText = ([String]$workPlaceInfo[3]).Substring($workPlaceLength, ([String]$workPlaceInfo[3]).Length - $workPlaceLength)
                    $koguchiSheet.Cells.item($koguchiRowCounter,$kingaku) = $kingakuText
                    
                }
                
                # 小口の行カウンターに3を追加し、次の行にする
                $koguchiRowCounter = $koguchiRowCounter + 3
                
            }
            
            # 勤務地情報リストテキストが存在したときの処理終了
        }else{
            # ======= プログレスバーを閉じる =======
            $formProgressBar.Close()
            # ポップアップを表示
            $popup.popup($infoTextFileName +"が見つかりません`r`nやり直してください",0,"やり直してください",48) | Out-Null
            # 処理を中断し、終了
            breakExcel
        }    
        # 「勤務内容」欄が空欄or在宅以外の処理 終了
    }
    # 実績
    if($kinmujisseki -ne ""){
        $lastWorkDayRow = $row
    }
    
}

# =========プログレスバーを進める8/10 =======
$progressBar.Value += 2
$formProgressBar.Show()

# 空行がある場合（小口の行カウンターが、記入可能行の終わり=74 未満のとき）は「適用（行先、要件）」に「以下余白」記入
if($koguchiRowCounter -lt 74){
    $koguchiSheet.Cells.item($koguchiRowCounter,6) = '以下余白'
}

# ------------- 個人情報欄の入力 --------------
# --- 年月日の入力 ---
$koguchiSheet.cells.item(78, 4) = $targetYear
$koguchiSheet.cells.item(78, 8) = $targetMonth

# 最終出勤日を日付欄に設定
$koguchiSheet.cells.item(78, 11) = $kinmuhyouSheet.cells.item($lastWorkDayRow, 3).text

# --- 名前の入力 ---
$targetPersonName = $kinmuhyouSheet.cells.range("W7").text
$koguchiSheet.cells.item(82, 21) = $targetPersonName
# 勤務表の名前が空白だった場合処理を中断する
if ($koguchiSheet.cells.item(82, 21).text -eq "") {
    $popup.popup($targetMonth + "月の勤務表に【名前】が記載されていません`r`n処理を中断します",0,"やり直してください",48) | Out-Null
    breakExcel
}

# --- 所属の入力 ---
$affiliation = $kinmuhyouSheet.cells.range("W6").text
# "部" を削除する
$affiliation -match "(?<affliationName>.+?)部" | Out-Null
$koguchiSheet.cells.item(80, 6) = $Matches.affliationName
# 勤務表の所属が空白だった場合処理を中断する
if ($koguchiSheet.cells.item(80, 6).text -eq "") {
    $popup.popup($targetMonth + "月の勤務表に【所属】が記載されていません`r`n処理を中断します",0,"やり直してください",48) | Out-Null
    breakExcel
}
# --- 印鑑の入力 ---
# 勤務表の該当シートの図形を取得
$allShapes = $kinmuhyouSheet.shapes
# コピペした時に図形のサイズを変更しないように設定する
# 2: セルに合わせて移動するがサイズ変更はしない
foreach ($shape in $allShapes) {
    $shape.placement = 2
}

# 印鑑をコピペしたいセルの位置
$targetStampCell = "AD82"

# 印鑑がないかもしれないフラグ
$haveStamp = $true
# 勤務表の印鑑のあるセルをクリップボードにコピー
$kinmuhyouSheet.range("AA7").copy() | Out-Null
# 小口シートに印鑑をペースト
$koguchiCell = $koguchiSheet.range($targetStampCell)
$koguchiSheet.paste($koguchiCell)
# ペースト先を編集
$koguchiSheet.range($targetStampCell).formula = ""
$koguchiSheet.range($targetStampCell).interior.colorindex = 0
# 罫線を編集するための宣言
$LineStyle = "microsoft.office.interop.excel.xlLineStyle" -as [type]
# 罫線をなしにする
$koguchiSheet.range($targetStampCell).borders.linestyle = $linestyle::xllinestylenone
# 印鑑（オブジェクト）が増えてなさそうなら、メッセージを表示する
$numberOfObject = 79
if ($koguchiSheet.shapes.count -eq $numberOfObject) {
    $haveStamp = $false
}

# 印鑑がないかもしれない場合注意喚起
if (!($haveStamp)) {
    # ======= プログレスバーを閉じる =======
    $formProgressBar.Close()
    # ポップアップを表示
    $popup.popup("印鑑が勤務表に入っていない`r`nまたは「印」から大幅にずれている可能性があります`r`n`r`n「印」の上に印鑑を貼り付けてやり直してください",0,"やり直してください",48) | Out-Null
    # 処理を中断し、終了
    breakExcel
}

# 文字色の変更（全部黒に）
$koguchiSheet.range("A1:BN90").font.colorindex = 1


# ---------------- 終了処理 ------------------

# 月が1桁 (ex 1月) の場合2桁 (ex 01) を用意する
$fileNameMonth = "{0:D2}" -f [int]$targetMonth

# 小口ブックの保存
$koguchiBook.save()

# 勤務表ブックと小口ブックを閉じる
$kinmuhyouBook.close()
$koguchiBook.close()

# --------新しい小口ファイル名を用意---------
# <社員番号>_小口交通費・出張旅費精算明細書_YYYYMM_<氏名>
$koguchiNewFileName = $kinmuhyou.name.Substring(0, 3) + "_小口交通費・出張旅費精算明細書_" + $targetYear + $fileNameMonth + "_" + $targetPersonName
# ファイル名に使えない文字が入っていたら削除する(氏名の間の空白など)
$koguchiNewFileName = remove-invalidFileNameChars $koguchiNewFileName
# 新しい小口ファイルのフルパス
$koguchiNewFullPath = join-path -Path $PWD -ChildPath ..\..\$koguchiKakunousakiName\$koguchiNewFileName

# ------------ファイル名を変更----------------

# すでに対象月の小口が作られているときの処理
# ※1桁まで対応
if (Test-Path ($koguchiNewfullPath + '_' + "[1-9]" + '.xlsx')) {

    
    # ------対象年月の小口が2つ以上存在してる場合--------

    # 同じ月の小口のファイル名を取得(_1など数字がついている)
    $onajiFileName = Get-ChildItem -Recurse -File -path ..\..\ | Where-Object name -CMatch "[0-9]{3}_小口交通費・出張旅費精算明細書_.+_.+_"

    # 同じ月の小口のファイル名を_で分ける
    # [0]: <社員番号>
    # [1]: 小口交通費・出張旅費精算明細書
    # [2]: <日付>
    # [3]: <氏名>
    # [4]:「1.xlsx」の数字部分 
    $splitBy_FileName = $onajiFileName -split "_"
    
    # -----------最大の数字を探す--------------
    for($i = 4; $i -lt (($onajiFileName.count)*5); $i = $i + 5){

        # 「1.xlsx」の数字部分を抜き出してインクリメントできるように数字にする
        $fileNameCountNumber = [int]($splitBy_FileName[$i].Substring(0,1))
        
        # もし今より大きかったら入れる
        if($fileNameCount -lt $fileNameCountNumber){
            $fileNameCount = $fileNameCountNumber
        }
        
    }

    # ファイル名の末尾の数字部分をインクリメント
    $fileNameCount = $fileNameCount + 1

    # ファイル名の変更に使用する文字列を用意
    $koguchiNewFileName = ($koguchiNewFileName + '_' + $fileNameCount + '.xlsx')

} elseif (Test-Path ($koguchiNewfullPath + '.xlsx')) {
    
    # ------対象年月の小口が1つ存在してる場合--------
    # <社員番号>_小口交通費・出張旅費精算明細書_YYYYMM_<氏名>_<numberOfFiles>.xlsxが存在しない

    # 「_1.xlsx」をファイル名に追加する
    $koguchiNewFileName = $koguchiNewFileName + '_1.xlsx'

}else{
    # 拡張子を追加
    $koguchiNewFileName = $koguchiNewFileName + '.xlsx'
}


# 小口ファイル名を変更
Rename-Item -path $koguchi -NewName $koguchiNewFileName -ErrorAction:Stop


# =======プログレスバーの終了8/10========
$progressBar.Value += 2
$formProgressBar.Show()
$formProgressBar.Close()

# ----------- 正常に終了したときポップアップを表示 ----------
# 氏名の空白をなくす
if($targetPersonName -match ' ' -or $targetPersonName -match '　'){
    $targetPersonName = $targetPersonName.replace('　', '  ')
    $targetPersonName = $targetPersonName.replace(' ', '')
}
$successEnd = $popup.popup($targetPersonName + "さんの`r`n" + $targetYear + "年" + $targetMonth + "月の小口が完成しました : )`r`n`r`nOKを押して不備がないか確認してください",0,"お待たせしました！",64)    

# ポップアップのOKが押されたら作成した小口が格納されているフォルダを開く
if($successEnd -eq '1'){
    $koguchiFilePath = join-path -Path $PWD -ChildPath ..\..\$koguchiKakunousakiName
    Start-Process $koguchiFilePath
}

# 使用したプロセスの解放
$kinmuhyouBook = $null
$kinmuhyouSheet = $null
$koguchiBook = $null
$koguchiSheet = $null
$koguchiCell = $null
[GC]::Collect()