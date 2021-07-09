# ---------------アセンブリの読み込み---------------
Add-Type -Assembly System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# フォーム全体の設定をする関数
# formText : フォームの本文（文字列）
# formYoko : フォームの横幅
# formTate : フォームの縦幅
function makeForm ($formText, $formYoko, $formTate) {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $formText
    $form.Size = New-Object System.Drawing.Size($formYoko, $formTate)
    $form.StartPosition = "CenterScreen"
    $form.font = $font
}

# ラベルを表示する関数
# $labelText : ラベルに書き込む文字列
# $form : フォームオブジェクト
function makeLabel ($labelText, $form) {
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10, 10)
    $label.Size = New-Object System.Drawing.Size(270, 30)
    $label.Text = $labelText
    $form.Controls.Add($label)
    return $form
}

# 勤務表を保存せずに閉じて、Excelを中断する関数
function breakExcel {
    # Bookを閉じる
    $kinmuhyouBook.close()
    # 使用していたプロセスの解放
    $excel = $null
    $kinmuhyouBook = $null
    $kinmuhyouSheet = $null
    # ガベージコレクト
    [GC]::Collect()
    # 処理を終了する
    exit
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


# (現在日によって変わるので、get-date -Format Y にはしていない)
$yesNo_yearMonthAreCorrect = [System.Windows.Forms.MessageBox]::Show("【 $thisYear 年 $targetMonth 月 】の勤務表をもとに初期設定をしますか？`r`n`r`n「いいえ」で他の月を選択できます", '作成する小口の対象年月', 'YesNo', 'Question')

# 今年を小口作成の対象年とする
$targetYear = $thisYear

# ☆$yesNo_yearMonthAreCorrect -eq 'No'ループ開始☆
if ($yesNo_yearMonthAreCorrect -eq 'No') {
    
    # フォントの指定
    $font = New-Object System.Drawing.Font("メイリオ", 8)

    # フォーム全体の設定
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "作成する小口の対象年月"
    $form.Size = New-Object System.Drawing.Size(265, 200)
    $form.StartPosition = "CenterScreen"
    $form.font = $font

    # ラベルを表示
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10, 10)
    $label.Size = New-Object System.Drawing.Size(270, 30)
    $label.Text = "作成したい小口の年月を選択してください"
    $form.Controls.Add($label)

    # OKボタンの設定
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Point(40, 100)
    $OKButton.Size = New-Object System.Drawing.Size(75, 30)
    $OKButton.Text = "OK"
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $OKButton
    $form.Controls.Add($OKButton)

    # キャンセルボタンの設定
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Point(130, 100)

    $CancelButton.Size = New-Object System.Drawing.Size(75, 30)
    $CancelButton.Text = "Cancel"
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $CancelButton
    $form.Controls.Add($CancelButton)

    # コンボボックスを作成
    $Combo = New-Object System.Windows.Forms.Combobox
    $Combo.Location = New-Object System.Drawing.Point(50, 50)
    $Combo.size = New-Object System.Drawing.Size(150, 30)
    # リスト以外の入力を許可しない
    $Combo.DropDownStyle = "DropDownList"
    $Combo.FlatStyle = "standard"
    # $Combo.font = $font
    $Combo.BackColor = "#005050"
    $Combo.ForeColor = "white"
        
    # -----------コンボボックスに項目を追加-----------
    for ($counterForMove = (-6); $counterForMove -le 6; $counterForMove++) {
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
    if ($result -eq "OK") {
        # ユーザーの回答を"年"で区切る
        $Combo.Text -match "(?<year>.+?)年(?<month>.+?)月" | out-null

        # ユーザー指定の年を小口作成の対象年として上書する
        $targetYear = $Matches.year

        # ユーザー指定の月を小口作成の対象月として上書きする
        $targetMonth = $Matches.month

    }
    else {
        # 処理を終了する
        exit
    }

    # ☆$yesNo_yearMonthAreCorrect -eq 'No'ループ終了☆
}

# ----------- しばらくお待ちください画面 -----------

# フォントの指定
$font = New-Object System.Drawing.Font("メイリオ", 8)

# フォームの設定
$waitForm = New-Object System.Windows.Forms.Form
$waitForm.Text = "初期設定"
$waitForm.Size = New-Object System.Drawing.Size(265, 170)
$waitForm.StartPosition = "CenterScreen"
$waitForm.font = $font

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(70, 30)
$label.Size = New-Object System.Drawing.Size(270, 30)
$label.Text = "準備中です`r`nしばらくお待ちください"
$waitForm.Controls.Add($label)

###### 画像のフルパス変更してほしいです☆ #########
#PictureBox
$pic = New-Object System.Windows.Forms.PictureBox
$pic.Size = New-Object System.Drawing.Size(50, 50)
$pic.Image = [System.Drawing.Image]::FromFile($PWD.Path + "\resources\picutures\お待ちください猫.png")
$pic.Location = New-Object System.Drawing.Point(20,20) 
$pic.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
$waitForm.Controls.Add($pic)

# 可視化
$waitResult = $waitForm.Show()

# ------------------------------------------------------



Write-Host "$targetYear 年の"
Write-Host "$targetMonth 月の小口を作成します"

# ポップアップを作成
$popup = new-object -comobject wscript.shell

# ファイル名の勤務表_のあとの表記
$fileNameMonth = [string]$targetMonth + "月"
# 勤務表ファイルを取得
$kinmuhyou = Get-ChildItem -Recurse -File | ? Name -Match ("[0-9]{3}_勤務表_" + $fileNameMonth + "_.+") 
# 該当勤務表ファイルの個数確認
if ($kinmuhyou.Count -lt 1) {
    
    # ポップアップを表示
    $popup.popup("$targetMonth 月の勤務表ファイルが存在しません", 0, "やり直してください", 48) | Out-Null
    exit
}
elseif ($kinmuhyou.Count -gt 1) {
    # ポップアップを表示
    $popup.popup("$targetMonth 月の勤務表ファイルが多すぎます`r`n1つにしてください", 0, "やり直してください", 48) | Out-Null
    exit
}

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

# 勤務表のフルパス
$kinmuhyouFullPath = $kinmuhyou.FullName 

# 勤務表ブックを開く
$kinmuhyouBook = $excel.workbooks.open($kinmuhyouFullPath)
$kinmuhyouSheet = $kinmuhyouBook.worksheets.item([String]$targetMonth + '月')

# 入力内容をまとめて入れておくための配列
$inputContentsArray = @()

# 今回登録する勤務地を格納する配列
$workPlaceArray = @()

# すでに勤務地情報リストに書いてある勤務地を格納する配列
$registeredWorkPlaceArray = @()

# ---------------勤務地情報リストを読み込む---------------------
# 勤務地情報リストが書いてあるテキスト
$infoTextFileName = ".\resources\ツール用引数.txt"
$infoTextFileFullpath = "$PWD\$infoTextFileName"

# 勤務地情報リストテキストが存在したときの処理
if (Test-Path $infoTextFileFullpath) {

    # 勤務地情報リストテキストの内容を取得
    $argumentText = (Get-Content $infoTextFileFullpath)
    
    # 勤務地情報リストテキストにすでに書かれている情報を取得する
    for ($i = 0; $i -lt $argumentText.Length; $i++) {
        $argumentText[$i] -Match "(?<workplace>.+?)_" | Out-Null
        # すでに配列に入っている勤務地は追加しない
        if (!$registeredWorkPlaceArray.Contains($Matches.workplace)) {
            # 配列にない勤務地を配列に追加する
            $registeredWorkPlaceArray += $Matches.workplace
        }
    }
}

# 勤務表から勤務地一覧を取得する
# $kinmunaiyou : 勤務内容列セル
# $kinmujisseki : 勤務実績列セル
# $sagyoubasho : 作業場所セル
for ($Row = 14; $Row -le 44; $Row++) {
    # 「勤務内容」欄の文字列を取得
    $kinmunaiyou = $kinmuhyouSheet.cells.item($Row, 26).text
    # 「勤務実績」欄の終了時刻の文字列を取得
    $kinmujisseki = $kinmuhyouSheet.cells.item($Row, 7).text
    # 「作業場所」欄の文字列を取得
    $sagyoubasho = $kinmuhyouSheet.cells.item(7, 7).text

    # 勤務実績が空値でない＝出勤してる日
    if ($kinmujisseki -ne "") {

        # 勤務内容が空値でない＝勤務地などが書いてある
        if ($kinmunaiyou -ne "") {
            # 勤務内容から勤務地を持ってくる
            $workPlace = $kinmunaiyou        
        }
        else {
            # 出勤してるけど勤務内容に勤務地が書いてない場合
            # 作業場所から勤務地を持ってくる
            $workPlace = $sagyoubasho
        }   
    }

    # 今回登録する勤務地にまだ登録されてないかつ、ツール用引数.txtにまだ登録されていない場合は、今回登録する勤務地配列に追加する
    if (!$workPlaceArray.Contains($workPlace) -and !$registeredWorkPlaceArray.Contains($workPlace)) {
        $workPlaceArray += @($workPlace)
    }
}

# 今回登録するものがない場合はpopupを表示して終了
if ($workPlaceArray.Length -eq 0) {
    # ポップアップを表示
    $popup.popup("$targetmonth 月の勤務表の勤務地は既に登録されています。", 0, "登録済み", 64) | Out-Null
    breakExcel    
    exit
}


# =========================== 入力画面 ===========================

# ---------------- 変数定義 ----------------

# フォントを指定
$font = New-Object System.Drawing.Font("ＭＳ ゴシック", 11)

# フォームごとの要素を格納する配列
$forms = @()
# 適用
$outputTekiyous = @()
# 区間
$outputKukans = @()
# 交通機関
$outputKoutsukikans = @()
# 金額
$outputKingakus = @()

# 交通機関
$koutsukikan1 = @()
$koutsukikan2 = @()
$koutsukikan3 = @()
$koutsukikan4 = @()
$koutsukikan5 = @()
$koutsukikan6 = @()

# フォームを作りすぎないようにするためのフラグ
# $True : 新たにフォームを作る
# $False : 新たにフォームを作らない（上書きのみ）
# 最初のループは増やしたことにする
$isAdded = $True


# ---------------- 関数定義 ----------------

# フォームを作成する関数
# Args[0] : タイトルに表示する文字列
function drawForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "初期設定  【" + $Args[0] + "】"
    $form.Size = New-Object System.Drawing.Size(660, 700)
    $form.StartPosition = "CenterScreen"
    return $form
}


# ラベルを作成する関数
# Args[0] : フォーム内の設定座標（横の位置）
# Args[1] : フォーム内の設定座標（縦の位置。高さ）
# Args[2] : ラベルを表示する幅
# Args[3] : ラベルに表示する文字列
# Args[4] : ラベルを表示するフォーム
# Args[5] : ラベルのフォント
function drawLabel {
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point($Args[0], $Args[1])
    $label.Size = New-Object System.Drawing.Size($Args[2], 15)
    $label.Text = $Args[3]
    $label.forecolor = "black"
    $label.font = $Args[5]
    if ($Args[5] -ne $null) {
        $Args[5]
    }
    $Args[4].Controls.Add($label)
    return $label
}

# OKボタンを作成する関数
# Args[0] : フォーム内の設定座標（縦の位置。高さ）
# Args[1] : OKボタンに表示する文字列
# Args[2] : OKボタンを表示するフォーム
# result : OK
function drawOKButton {
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Point(20, $Args[0])
    $OKButton.Size = New-Object System.Drawing.Size(75, 30)
    $OKButton.Text = $Args[1]
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $Args[2].AcceptButton = $OKButton
    $Args[2].Controls.Add($OKButton)
}

# 在宅ボタンを作成する関数
# Args[0] : フォーム内の設定座標（縦の位置。高さ）
# Args[1] : 在宅ボタンに表示する文字列
# Args[2] : 在宅ボタンを表示するフォーム
# result : Yes
function drawAtHomeButton {
    $AtHomeButton = New-Object System.Windows.Forms.Button
    $AtHomeButton.Location = New-Object System.Drawing.Point(130, $Args[0])
    $AtHomeButton.Size = New-Object System.Drawing.Size(75, 30)
    $AtHomeButton.Text = $Args[1]
    $AtHomeButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $Args[2].Controls.Add($AtHomeButton)
}

# 戻るボタンを作成する関数
# Args[0] : フォーム内の設定座標（縦の位置。高さ）
# Args[1] : 戻るボタンに表示する文字列
# Args[2] : 戻るボタンを表示するフォーム
# result : Retry
function drawReturnButton {
    $ReturnButton = New-Object System.Windows.Forms.Button
    $ReturnButton.Location = New-Object System.Drawing.Point(240, $Args[0])
    $ReturnButton.Size = New-Object System.Drawing.Size(75, 30)
    $ReturnButton.Text = $Args[1]
    $ReturnButton.DialogResult = [System.Windows.Forms.DialogResult]::Retry
    if ($i -eq 0) {
        $ReturnButton.Enabled = $false; 
    }
    else {
        $ReturnButton.Enabled = $True;
    }
    $Args[2].Controls.Add($ReturnButton)
}

# 登録済み勤務地から選択ボタンを作成する関数
# result : No
function drawregisteredButton {
    $registeredButton = New-Object System.Windows.Forms.Button
    $registeredButton.Location = New-Object System.Drawing.Point(350, 610)
    $registeredButton.Size = New-Object System.Drawing.Size(155, 30)
    $registeredButton.Text = "登録済みの勤務地から選択する"
    $registeredButton.DialogResult = [System.Windows.Forms.DialogResult]::No
    # if ($i -eq 0) {
    #     $registeredButton.Enabled = $false; 
    # }else {
    #     $registeredButton.Enabled = $True;
    # }
    $forms[$i].Controls.Add($registeredButton)
}


# テキストボックスを作成する関数
# Args[0] : フォーム内の設定座標（横の位置）
# Args[1] : フォーム内の設定座標（縦の位置。高さ）
# Args[2] : テキストボックスの横幅
# Args[3] : テキストボックスの高さ
# Args[4] : テキストボックスを表示するフォーム
function drawTextBox {
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point($Args[0], $Args[1])
    $textBox.Size = New-Object System.Drawing.Size($Args[2], $Args[3])
    $textBox.BackColor = "white"
    $Args[4].Controls.Add($textBox)
    return $textBox
}


# しばらくお待ちください画面を閉じる
$waitForm.Close()


# 入力画面表示
# workPlaceArray : 勤務表から取得した、今回登録する勤務地一覧
:EMPTY for ($i = 0; $i -lt $workPlaceArray.Length; $i++) {

    # ---------------- Main ----------------- 

    # 戻るボタン押下後、エラーの場合 以外新しくフォームを作成する
    if ($isAdded) {
        # フォーム作成関数呼び出し
        $forms += drawForm $workPlaceArray[$i]   
    }

    # OKボタン関数呼び出し
    drawOKButton 610 "OK" $forms[$i]

    # 在宅ボタン関数呼び出し
    drawAtHomeButton 610 "在宅" $forms[$i]

    # 戻るボタン関数呼び出し
    drawReturnButton 610 "戻る" $forms[$i]

    # 登録済み勤務地から選択ボタン呼び出し
    drawregisteredButton


    # =============================== input ===============================

    # 在宅勤務時は「在宅ボタン」を押下させるメッセージラベルを作成
    $atHomeLabel = drawLabel 10 10 470 "★ 在宅勤務のときは【在宅】ボタンをクリック ★" $forms[$i]
    # 文字色指定
    $atHomeLabel.forecolor = "red" 
    # フォント指定（drawLabel関数の第5引数に入れると、文字色が赤くならない）
    $atHomeLabel.font = $font 


    # ---------------- 適用（行先、要件） ----------------- 
    # 適用ラベルのフォーム内の設定座標の高さ
    $tekiyouLabelLocate = 50
    # 適用テキストボックスのフォーム内の設定座標の高さ
    $tekiyouTextBoxLocate = 108

    # ラベル関数呼び出し
    drawLabel 10 $tekiyouLabelLocate 470 ("１．【 適用 】 勤務地 `"" + $workPlaceArray[$i] + "`" の時の適用を入力してください") $forms[$i] | Out-Null
    drawLabel 20 ($tekiyouLabelLocate + 20) 470 "ex.  自宅←→田町本社" $forms[$i] | Out-Null
    drawLabel 20 ($tekiyouLabelLocate + 40) 470 "      自宅→品川→東京テレポート→自宅 (勤務地複数の場合)" $forms[$i] | Out-Null

    # テキストボックス関数呼び出し
    $outputTekiyou = drawTextBox 20 $tekiyouTextBoxLocate 300 20  $forms[$i]

    # 戻るボタン押下後、エラーの場合 以外
    if ($isAdded) {
        # 適用テキストボックスを配列に追加
        $outputTekiyous += $outputTekiyou    
    }

    # ---------------- 区間 ----------------- 
    # 区間ラベルのフォーム内の設定座標の高さ
    $kukanLabelLocate = 150
    # 区間テキストボックスのフォーム内の設定座標の高さ
    $kukanTextBoxLocate = 208

    # ラベル関数呼び出し
    drawLabel  10 $kukanLabelLocate 550 ("２．【 区間 】 勤務地 `"" + $workPlaceArray[$i] + "`" の時の区間（自宅の最寄り駅←→勤務地の最寄り駅）を入力してください") $forms[$i] | Out-Null
    drawLabel 20 ($kukanLabelLocate + 20) 470 "ex.  <自宅の最寄り駅>←→田町 (往復の場合)" $forms[$i] | Out-Null
    drawLabel 20 ($kukanLabelLocate + 40) 670 "      <自宅の最寄り駅>→品川→東京テレポート→<自宅の最寄り駅> (勤務地複数の場合)" $forms[$i] | Out-Null

    # テキストボックス関数呼び出し
    $outputKukan = drawTextBox 20 $kukanTextBoxLocate 430 20 $forms[$i]

    # 戻るボタン押下後、エラーの場合 以外
    if ($isAdded) {
        # 区間テキストボックスを配列に追加
        $outputKukans += $outputKukan    
    }


    # ---------------- 交通機関 -----------------
    # 交通機関ラベルのフォーム内の設定座標の高さ
    $koutsukikanLabelLocate = 290
    # 交通機関テキストボックスのフォーム内の設定座標の高さ
    $koutsukikanTextBoxLocate = 288

    # ラベル関数呼び出し
    drawLabel 10 250 500 ("３．【 交通機関 】 勤務地 `"" + $workPlaceArray[$i] + "`" の時に利用する交通機関を入力してください") $forms[$i] | Out-Null
    drawLabel 20 270 500 "ex. JR山手線" $forms[$i] | Out-Null
    drawLabel 10 $koutsukikanLabelLocate 70 "交通機関１：" $forms[$i] | Out-Null
    drawLabel 10 ($koutsukikanLabelLocate + 40) 70 "交通機関２：" $forms[$i] | Out-Null
    drawLabel 10 ($koutsukikanLabelLocate + 80) 70 "交通機関３：" $forms[$i] | Out-Null
    drawLabel 10 ($koutsukikanLabelLocate + 120) 70 "交通機関４：" $forms[$i] | Out-Null
    drawLabel 10 ($koutsukikanLabelLocate + 160) 70 "交通機関５：" $forms[$i] | Out-Null
    drawLabel 10 ($koutsukikanLabelLocate + 200) 70 "交通機関６：" $forms[$i] | Out-Null

    # テキストボックス関数呼び出し
    $koutsukikan1 = drawTextBox 90 $koutsukikanTextBoxLocate 200 20 $forms[$i]
    $koutsukikan2 = drawTextBox 90 ($koutsukikanTextBoxLocate + 40) 200 20 $forms[$i]
    $koutsukikan3 = drawTextBox 90 ($koutsukikanTextBoxLocate + 80) 200 20 $forms[$i]
    $koutsukikan4 = drawTextBox 90 ($koutsukikanTextBoxLocate + 120) 200 20 $forms[$i]
    $koutsukikan5 = drawTextBox 90 ($koutsukikanTextBoxLocate + 160) 200 20 $forms[$i]
    $koutsukikan6 = drawTextBox 90 ($koutsukikanTextBoxLocate + 200) 200 20 $forms[$i]

    # 戻るボタン押下後、エラーの場合 以外
    if ($isAdded) {
        # 後の処理で使いやすくするため、各交通機関を配列に格納する
        $inputkoutsukikan = @($koutsukikan1, $koutsukikan2, $koutsukikan3, $koutsukikan4, $koutsukikan5, $koutsukikan6)
        $outputKoutsukikans+= , @($inputkoutsukikan)
    }
    

    # ---------------- 金額 -----------------
    # 金額ラベルのフォーム内の設定座標の高さ
    $kingakuLabelLocate = 530
    # 金額テキストボックスのフォーム内の設定座標の高さ
    $kingakuTextBoxLocate = 570

    # ラベル関数呼び出し
    drawLabel 10 $kingakuLabelLocate 500 ("４．【 金額 】 勤務地 `"" + $workPlaceArray[$i] + "`" の金額（往復代金）を入力してください") $forms[$i] | Out-Null
    drawLabel 20 ($kingakuLabelLocate + 20) 470 "ex.  750 （半角数字）" $forms[$i] | Out-Null

    # テキストボックス関数呼び出し
    $outputKingaku = drawTextBox 20 $kingakuTextBoxLocate 100 20 $forms[$i]

    # 戻るボタン押下後、エラーの場合 以外
    if ($isAdded) {
        # 金額テキストボックスを配列に追加
        $outputKingakus += $outputKingaku   
    }

    # 可視化
    $inputContentsResult = $forms[$i].ShowDialog()


    # =============================== output ===============================
    # --------------- OKボタンを押したら ---------------
    if ($inputContentsResult -eq "OK") {

        #  ---------------- 空白エラー判定 -----------------

        # 以下の変数をリセットする
        #
        # nullOrEmptyCount : 交通機関テキストボックスの空の個数
        # koutsukikans : 複数の交通機関テキストボックスを一つにまとめるための変数
        # outputKoutsukikan : 編集したkoutsukikansを代入する
        # isEmpty : 空白エラーを起こすためのフラグ
        #
        $nullOrEmptyCount = 0
        $koutsukikans = ""
        $outputKoutsukikan= ""
        $isEmpty = $false

        # テキストボックスの色を白に戻す
        $outputTekiyous[$i].BackColor = "white"
        $outputKukans[$i].BackColor = "white"
        $outputKingakus[$i].BackColor = "white"
        $outputKoutsukikans[$i][0].BackColor = "white"


        for ($l = 0; $l -lt $outputKoutsukikans[$i].length; $l++) {
            # 交通機関テキストボックスから空ではないものを抜き出す
            if ([string]::IsNullOrEmpty($outputKoutsukikans[$i][$l].text)) {
                # NULL や '' の場合
                $nullOrEmptyCount++
            }
            else {
                # 上記以外は設定された文字列を出力
                $koutsukikans += ($outputKoutsukikans[$i][$l].text + '`r`n')
            }
        }

        # 末尾の「`r`n」を消す
        $outputKoutsukikan+= $koutsukikans.Substring(0, $koutsukikans.Length - 4)

        while ($True) {
            # 交通機関が全て空だった場合の処理
            if ($nullOrEmptyCount -eq 6) {
                $outputKoutsukikans[$i][0].BackColor = "#ff99cc"
                $isEmpty = $True
            }
    
            # ユーザ入力に空白があった場合の処理
            $inputtedTextBoxes = @($outputTekiyous[$i], $outputKukans[$i], $outputKingakus[$i])
            # ユーザ入力に１つでも空白があった場合の処理
            foreach ($inputtedTextBox in $inputtedTextBoxes) {
                if ($inputtedTextBox.text -eq "") {
                    $inputtedTextBox.BackColor = "#ff99cc"
                    $isEmpty = $True
                }    
            }
            # 空白があった場合これ以降の処理をスキップする
            if ($isEmpty) {
                # 交通機関の空白カウントを初期化
                $nullOrEmptyCount = 0
                $i = $i - 1
                $isAdded = $false
                continue EMPTY
            }
            # エラーがない場合はループから抜ける
            break
        }
        

        # ---------------- 適用（行先、要件） -----------------

        $inputContentsArray += @($workPlaceArray[$i] + "_" + $outputTekiyous[$i].text)

        # ---------------- 区間 -----------------

        $inputContentsArray += @($workPlaceArray[$i] + "_" + $outputKukans[$i].text)

        # ---------------- 交通機関 -----------------

        $inputContentsArray += @($workPlaceArray[$i] + "_" + $outputKoutsukikan)

        # ---------------- 金額 -----------------

        $inputContentsArray += @($workPlaceArray[$i] + "_" + $outputKingakus[$i].text)

        # フォーム増やしたフラグ
        $isAdded = $True

    
        # --------------- 戻るボタンを押したら ---------------
    }
    elseif ($inputContentsResult -eq "Retry") {
        
        # 繰り返しの条件を2戻す
        # 例えば、1画面目が田町（$i = 1）2画面目がお台場（$i = 2）だったとき、田町の画面に戻りたいときは $i = 1 にしたい
        # for文の条件式？でインクリメントされているため、$iから2を引く必要がある
        $i = $i - 2
        # 配列になにも入っていない時（固定配列なので、最初の要素は空にするだけにした）
        if ($inputContentsArray.Length -le 4) {
            for ($j = 1; $j -lt 5; $j++) {
                $inputContentsArray[($inputContentsArray.Length - $j)] = ""
            }
            # 戻るボタンを押したら、テキストファイルに出力する要素を削除する    
        }
        else {
            $inputContentsArray = $inputContentsArray[0..($inputContentsArray.Length - 5)]
        }
        
        $isAdded = $false

        # --------------- 在宅ボタンを押したら ---------------
    }
    elseif ($inputContentsResult -eq "Yes") {
        $inputContentsArray += @($workPlaceArray[$i] + "_1")
        $inputContentsArray += @($workPlaceArray[$i] + "_1")
        $inputContentsArray += @($workPlaceArray[$i] + "_1")
        $inputContentsArray += @($workPlaceArray[$i] + "_1")
    }
    # 登録済み勤務地から選択する場合
    elseif ($inputContentsResult -eq "No") {
        # 登録済み勤務地選択用フォームを作成
        $selectForm = New-Object System.Windows.Forms.Form
        $selectForm.Text = "登録済みの勤務地から選択"
        $selectForm.Size = New-Object System.Drawing.Size(300, 200)
        $selectForm.StartPosition = "CenterScreen"
        
        # 可視化
        $selectResult = $selectForm.ShowDialog()
    
    }
    else {
        breakExcel
        exit
    }    
}

# 完了画面がほしいなあ

# 配列をテキストに出力する
foreach ($inputContent in $inputContentsArray) {
    $inputContent >> .\resources\ツール用引数.txt
}

# 勤務表ファイルを閉じる
breakExcel

# 変数の解放
$outputTekiyou = $null
$outputKukan = $null
$koutsukikan1 = $null
$koutsukikan2 = $null
$koutsukikan3 = $null
$koutsukikan4 = $null
$koutsukikan5 = $null
$koutsukikan6 = $null
$inputKoutsukikan = $null
$koutsukikans = $null
$outputKoutsukikans = $null
$outputKingaku = $null