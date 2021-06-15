# 
# 勤務表をもとに小口交通費請求書を作成するPowershell
# 
# 勤務表のファイル名：<3桁の社員番号>_勤務表_M月_<氏名>.xlsx
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
    # # 処理を終了する
    # exit
}

# 引数の空白を除きファイル名として使えない文字を消す関数
# fileName : ファイル名
function remove-invalidFileNameChars ($fileName) {
    $fileNameRemovedSpace = $fileName -replace "　", ""　-replace " ", ""
    $invalidChars = [IO.Path]::GetInvalidFileNameChars() -join ''
    $regex = "[{0}]" -f [RegEx]::Escape($invalidChars)
    return $fileNameRemovedSpace -replace $regex
}

# フォーム全体の設定をする関数
# formText : フォームの本文（文字列）
# formYoko : フォームの横幅
# formTate : フォームの縦幅
function makeForm ($formText, $formYoko, $formTate) {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $formText
    $form.Size = New-Object System.Drawing.Size($formYoko,$formTate)
    $form.StartPosition = "CenterScreen"
    $form.font = $Font
}

# ラベルを表示する関数
# $labelText : ラベルに書き込む文字列
# $form : フォームオブジェクト
function makeLabel ($labelText, $form) {
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,10)
    $label.Size = New-Object System.Drawing.Size(270,30)
    $label.Text = $labelText
    $form.Controls.Add($label)
    return $form
}

# -------------------- 主処理の準備 --------------------------

# 現在の年月日を取得する
$thisYear = (Get-Date).Year
$thisMonth = (Get-Date).Month
$today = (Get-Date).Day

# 現在日時から作成するべき勤務表の月次を判定
# 24日までは当月分を作る
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
$yesNo_yearMonthAreCorrect = [System.Windows.Forms.MessageBox]::Show("作成するのは 【 $thisYear 年 $targetMonth 月 】の小口でよろしいですか？`r`n`r`n「いいえ」で他の月を選択できます",'作成する小口の対象年月','YesNo','Question')

# 今年を小口作成の対象年とする
$targetYear = $thisYear

# ☆$yesNo_yearMonthAreCorrect -eq 'No'ループ開始☆
if($yesNo_yearMonthAreCorrect -eq 'No'){
    
    # フォントの指定
    $Font = New-Object System.Drawing.Font("メイリオ",8)

    # フォーム全体の設定
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "作成する小口の対象年月"
    $form.Size = New-Object System.Drawing.Size(265,200)
    $form.StartPosition = "CenterScreen"
    $form.font = $Font

    # ラベルを表示
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,10)
    $label.Size = New-Object System.Drawing.Size(270,30)
    $label.Text = "作成したい小口の年月を選択してください"
    $form.Controls.Add($label)

    # OKボタンの設定
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Point(40,100)
    $OKButton.Size = New-Object System.Drawing.Size(75,30)
    $OKButton.Text = "OK"
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $OKButton
    $form.Controls.Add($OKButton)

    # キャンセルボタンの設定
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Point(130,100)

    $CancelButton.Size = New-Object System.Drawing.Size(75,30)
    $CancelButton.Text = "Cancel"
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $CancelButton
    $form.Controls.Add($CancelButton)

    # コンボボックスを作成
    $Combo = New-Object System.Windows.Forms.Combobox
    $Combo.Location = New-Object System.Drawing.Point(50,50)
    $Combo.size = New-Object System.Drawing.Size(150,30)
    # リスト以外の入力を許可しない
    $Combo.DropDownStyle = "DropDownList"
    $Combo.FlatStyle = "standard"
    # $Combo.font = $Font
    $Combo.BackColor = "#005050"
    $Combo.ForeColor = "white"
        
    # -----------コンボボックスに項目を追加-----------
    for($counterForMove = (-6); $counterForMove -le 6; $counterForMove++){
        $date = get-date (get-date).AddMonths($counterForMove) -Format Y
        [void] $Combo.Items.Add("$date")
    }
    
    # フォームにコンボボックスを追加
    $form.Controls.Add($Combo)
    $Combo.SelectedIndex = 6
    
    # フォームを最前面に表示
    $form.Topmost = $True
    
    # フォームを表示＋選択結果を変数に格納
    $result = $form.ShowDialog()

    # 選択後、OKボタンが押された場合、選択項目を表示
    if ($result -eq "OK"){
        # ユーザーの回答を"年"で区切る
        $Combo.Text -match "(?<year>.+?)年(?<month>.+?)月" | out-null

        # ユーザー指定の年を小口作成の対象年として上書する
        $targetYear = $Matches.year

        # ユーザー指定の月を小口作成の対象月として上書きする
        $targetMonth = $Matches.month

    }else{
        # 処理を終了する
        exit
    }

# ☆$yesNo_yearMonthAreCorrect -eq 'No'ループ終了☆
}

Write-Host "$targetYear 年の"
Write-Host "$targetMonth 月の小口を作成します"

# ポップアップを作成
$popup = new-object -comobject wscript.shell

# ----------------------小口テンプレを取得------------------------
$koguchiTemplate = Get-ChildItem -Recurse -File | ? Name -Match "小口交通費・出張旅費精算明細書_テンプレ.xlsx"
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

# 小口格納フォルダが存在していない場合は作成する
if(!(Test-Path $PWD"\作成した小口交通費請求書")){
    New-Item -Path $PWD"\作成した小口交通費請求書" -ItemType Directory | Out-Null
}

$koguchi = Join-Path $PWD "作成した小口交通費請求書" | Join-Path -ChildPath "小口交通費・出張旅費精算明細書_コピー.xlsx"
Copy-Item -path $koguchiTemplate.FullName -Destination $koguchi

# ----------------テンプレートから小口交通費請求書を作成する---------------------

# ファイル名の勤務表_のあとの表記が「M月」表記の場合
$fileNameMonth = [string]("$targetMonth" + "月")

# もし「勤務表_YYYYMM」のような表記にするなら ↑ をコメントアウトして ↓ のコメントアウトをぬく
# $targetMonth00 = "{0:00}" -f [int]$targetMonth
# $fileNameMonth = ("$targetYear" + "$targetMonth00")

# 勤務表ファイルを取得
$kinmuhyou = Get-ChildItem -Recurse -File | ? Name -Match ("[0-9]{3}_勤務表_" + "$fileNameMonth" + "_.+")

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


# --------------- 処理中のプログレスバーを表示 -------------

# プログレスバー用のフォームを用意
$formProgressBar = New-Object System.Windows.Forms.Form
$formProgressBar.Size = "300,200"
$formProgressBar.Startposition = "CenterScreen"
$formProgressBar.Text = "作成中…"

# プログレスバーを用意
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = "10,100"
$progressBar.Size = "260,30"
$progressBar.Maximum = "10"
$progressBar.Minimum = "0"
$progressBar.Style = "Continuous"

# =========プログレスバーを進める2/10 =======
$progressBar.Value = 2
$formProgressBar.Controls.AddRange($progressBar)
$formProgressBar.Topmost = $True
$formProgressBar.Show()


# displaySharpMessage "White" ([string]$targetMonth + " 月の小口交通費請求書を作成します") "しばらくお待ちください。"

# ----------------------Excelを起動する--------------------------------
try {
    # 起動中のExcelプロセスを取得
    $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
}
catch {
    $excel = New-Object -ComObject "Excel.Application" 
}

# Excelがメッセージダイアログを表示しないようにする
$excel.DisplayAlerts = $false
$excel.visible = $true

# 勤務表のフルパス
$kinmuhyouFullPath = $kinmuhyou.FullName 

# 勤務表ブックを開く
$kinmuhyouBook = $excel.workbooks.open($kinmuhyouFullPath)
$kinmuhyouSheet = $kinmuhyouBook.worksheets.item([String]$targetMonth + '月')

# 小口ブックを開く
$koguchiBook = $excel.workbooks.open($koguchi)
$koguchiSheet = $koguchiBook.sheets(1)


# =========プログレスバーを進める4/10 =======
$progressBar.Value += 2
$formProgressBar.Show()

# ------------- 勤務表の中身を小口にコピーする ----------------
# 「勤務内容」欄に書かれている勤務地を参考にして、勤務地情報リストテキストから該当情報を小口に記入する

# 小口の行カウンター
$koguchiRowCounter = 11

# 勤務表の1日～月末まで1行ずつ繰り返す
for ($row = 14; $row -le 44; $row++) {
    # 勤務地判定のために「勤務内容」欄の文字列を取得
    $workPlace = $kinmuhyouSheet.cells.item($row, 26).formula
    Write-Host ("勤務地：" + $workPlace)
    $workPlaceLength = [int]$workPlace.length + 1
    write-host ('$workPlaceと＿の文字数：' + $workPlaceLength)
    
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
        
        # ---------------勤務地情報リストを読み込む---------------------
        # 勤務地情報リストが書いてあるテキスト
        $infoTextFileName = "ツール用引数.txt"
        $infoTextFileFullpath = "$PWD\$infoTextFileName"
        
        # 勤務地情報リストテキストが存在したときの処理
        if(Test-Path $infoTextFileFullpath){
            
            $argumentText = (Get-Content $infoTextFileFullpath)
            
            # 「勤務内容」欄の文字列にマッチした勤務地の情報を、リストから取得 ( 配列の中身　[0]:適用　[1]:区間　[2]:交通機関　[3]:金額 )
            $workPlaceInfo = $argumentText | Select-String -Pattern ($workPlace + '_')
            Write-Host ("勤務地list：" + $workPlaceInfo)
            
            # 「勤務内容」欄の内容が勤務の情報リストになかった場合、ポップアップを表示し終了する
            if($workPlaceInfo -eq $null){
                # ポップアップを表示
                $popup.popup("勤務地の情報が登録されていません`r`n初期設定もしくは上書きし、やり直してください",0,"やり直してください",48) | Out-Null
                
                # 処理を中断し、終了
                breakExcel
                exit
                
            }
            
            # 在宅フラグ(適用部分に1)が立っている場合、小口には記入しない
            elseif(([String]$workPlaceInfo[0]).Substring($workPlaceLength, ([String]$workPlaceInfo[0]).Length - $workPlaceLength) -eq '1'){
                # 小口に記入しない

                write-host "!!!!!!zaitaku!!!!!!"
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
            # ポップアップを表示
            $popup.popup("勤務地の情報リストが見つかりません`r`nやり直してください",0,"やり直してください",48) | Out-Null
        }
        
        # 「勤務内容」欄が空欄or在宅の処理終了
    }

}

# =========プログレスバーを進める6/10 =======
$progressBar.Value += 2
$formProgressBar.Show()

# ------------- 個人情報欄のコピー --------------
# --- 年月日のコピー ---
$koguchiSheet.cells.item(78, 4) = $targetYear
$koguchiSheet.cells.item(78, 8) = $targetMonth

# 月の最終日を日付欄に設定
$koguchiSheet.cells.item(78, 11) = [DateTime]::DaysInMonth($targetYear,$targetMonth)

# --- 名前のコピー ---
$targetPersonName = $kinmuhyouSheet.cells.range("W7").text
$koguchiSheet.cells.item(82, 21) = $targetPersonName
# 勤務表の名前が空白だった場合処理を中断する
if ($koguchiSheet.cells.item(82, 21).text -eq "") {
    $popup.popup($targetMonth + "月の勤務表に【名前】が記載されていません`r`n処理を中断します",0,"やり直してください",48) | Out-Null
    breakExcel
    exit
}

# --- 所属のコピー ---
$affiliation = $kinmuhyouSheet.cells.range("W6").text
# "部" を削除する
$affiliation -match "(?<affliationName>.+?)部" | Out-Null
$koguchiSheet.cells.item(80, 6) = $Matches.affliationName
# 勤務表の所属が空白だった場合処理を中断する
if ($koguchiSheet.cells.item(80, 6).text -eq "") {
    $popup.popup($targetMonth + "月の勤務表に【所属】が記載されていません`r`n処理を中断します",0,"やり直してください",48) | Out-Null
    breakExcel
    exit
}
# --- 印鑑のコピー ---
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

    $popup.popup("印鑑が勤務表に入っていない`r`nまたは印から大幅にずれている可能性があります`r`nやり直してください",0,"やり直してください",48) | Out-Null
    breakExcel
    exit

}

# 文字色の変更（全部黒に）
$koguchiSheet.range("A1:BN90").font.colorindex = 1

# ×ボタンを押したとき、処理途中のものを削除しよう


# =========プログレスバーを進める8/10 =======
$progressBar.Value += 2
$formProgressBar.Show()

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
$koguchiNewfullPath = Join-Path $PWD "作成した小口交通費請求書" | Join-Path -ChildPath $koguchiNewFileName

# ------------ファイル名を変更----------------

# すでに対象月の小口が作られているときの処理
# ※1桁まで対応
# if (Test-Path ($koguchiNewfullPath + "_$numberOfFiles.xlsx")) {
if (Test-Path ($koguchiNewfullPath + '_' + "[1-9]" + '.xlsx')) {

    
    # ------対象年月の小口が2つ以上存在してる場合--------
    # <社員番号>_小口交通費・出張旅費精算明細書_YYYYMM_<氏名>_<numberOfFiles>.xlsxが存在する

    # 同じ月の小口のファイル名を取得(_1など数字がついている)
    $onajiFileName = Get-ChildItem -Recurse | Where-Object name -CMatch "[0-9]{3}_小口交通費・出張旅費精算明細書_.+_.+_"

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
        $fileNameCountNumber
        
        # もし今より大きかったら入れる
        if($fileNameCount -lt $fileNameCountNumber){
            write-host "$fileNameCount を"
            $fileNameCount = $fileNameCountNumber
            write-host "$fileNameCount にしたよ"

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
$finish = $formProgressBar.Show()
$formProgressBar.Close()

# 正常に終了したときポップアップを表示
$popup.popup("お待たせしました！正常に終了しました`r`n仕上がりを確認してください",0,"正常終了",64) | Out-Null    

# 使用したプロセスの解放
$kinmuhyouBook = $null
$kinmuhyouSheet = $null
$koguchiBook = $null
$koguchiSheet = $null
$koguchiCell = $null
[GC]::Collect()


# 最後は「開く」「終了」の2択
# 開く→できあがったところのエクスプローラーを表示する