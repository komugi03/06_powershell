# 構成図と合ってる

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
    $form.font = $Font
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

# 勤務表と小口を保存せずに閉じて、Excelを中断する関数
function endExcel {
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
    $Font = New-Object System.Drawing.Font("メイリオ", 8)

    # フォーム全体の設定
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "作成する小口の対象年月"
    $form.Size = New-Object System.Drawing.Size(265, 200)
    $form.StartPosition = "CenterScreen"
    $form.font = $Font

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
    # $Combo.font = $Font
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

Write-Host "$targetYear 年の"
Write-Host "$targetMonth 月の小口を作成します"

# ポップアップを作成
$popup = new-object -comobject wscript.shell

# ファイル名の勤務表_のあとの表記
$fileNameMonth = [string]$targetMonth+"月"
# 勤務表ファイルを取得
$kinmuhyou = Get-ChildItem -Recurse -File | ? Name -Match ("[0-9]{3}_勤務表_"+$fileNameMonth+"_.+") 
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

# ---------------勤務地情報リストを読み込む---------------------
# 勤務地情報リストが書いてあるテキスト
$infoTextFileName = "ツール用引数.txt"
$infoTextFileFullpath = "$PWD\$infoTextFileName"

# 勤務地情報リストテキストが存在したときの処理
if(Test-Path $infoTextFileFullpath){
    
    $argumentText = (Get-Content $infoTextFileFullpath)
    
    # 「勤務内容」欄の文字列にマッチした勤務地の情報を、リストから取得 ( 配列の中身　[0]:適用　[1]:区間　[2]:交通機関　[3]:金額 )
    $workPlaceInfo = $argumentText | Select-String -Pattern ($workPlace + '_')

}
# 入力内容をまとめて入れておく
$inputContentsArray = @()

# 勤務内容と作業場所から取ってきた勤務地が被らないように入ってる
$workPlaceArray = @()

# 勤務表から勤務地一覧を取得する
# $kinmunaiyou : 勤務内容列セル
# $kinmujisseki : 勤務実績列セル
# $sagyoubasho : 作業場所セル

for ($Row = 14; $Row -le 44; $Row++) {
    $kinmunaiyou = $kinmuhyouSheet.cells.item($Row, 26).text
    $kinmujisseki = $kinmuhyouSheet.cells.item($Row, 7).text
    $sagyoubasho = $kinmuhyouSheet.cells.item(7, 7).text

    if ($kinmujisseki -ne "") {

        if ($kinmunaiyou -ne "") {
            # 勤務内容から勤務地を持ってくる
            $workPlace = $kinmunaiyou        }
        else {
            # 作業場所から勤務地を持ってくる
            $workPlace = $sagyoubasho
        }   
    }

    if (!$workPlaceArray.Contains($workPlace)) {
        $workPlaceArray += @($workPlace)
        
    }
}


# =========================== 入力画面 ===========================

# ---------------- 変数定義 ----------------

# フォント
$Font = New-Object System.Drawing.Font("ＭＳ ゴシック",11)

# 戻るボタンを押した時用値保持
$tekiyouValue = ""
$kukanValue = ""
$koutsukikanValue = @()
$kingakuValue = ""

# フォームごとの要素一覧ためとく配列
$forms = @()
$outputTekiyous = @()
$outputKukans = @()
$outputkoutsukikans = @(@(),@())
$outputkingakus = @()
$koutsukikan1 = @()
$koutsukikan2 = @()
$koutsukikan3 = @()
$koutsukikan4 = @()
$koutsukikan5 = @()
$koutsukikan6 = @()

# フォームを増やしたかどうかのフラグ
# 最初のループは増やしたことにする
$isadded = $True


# ---------------- 関数定義 ----------------

# ラベルを表示
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

# OKボタン
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

# 在宅ボタン
# result : Yes
function drawAtHomeButton {
    $AtHomeButton = New-Object System.Windows.Forms.Button
    $AtHomeButton.Location = New-Object System.Drawing.Point(130, $Args[0])
    $AtHomeButton.Size = New-Object System.Drawing.Size(75, 30)
    $AtHomeButton.Text = $Args[1]
    $AtHomeButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $Args[2].Controls.Add($AtHomeButton)
}

# 戻るボタン
# result : Retry
function drawReturnButton {
    $ReturnButton = New-Object System.Windows.Forms.Button
    $ReturnButton.Location = New-Object System.Drawing.Point(240, $Args[0])
    $ReturnButton.Size = New-Object System.Drawing.Size(75, 30)
    $ReturnButton.Text = $Args[1]
    $ReturnButton.DialogResult = [System.Windows.Forms.DialogResult]::Retry
    if ($i -eq 0) {
        $ReturnButton.Enabled = $false; 
    }else {
        $ReturnButton.Enabled = $True;
    }
    $Args[2].Controls.Add($ReturnButton)
}

# テキストボックス
function drawTextBox {
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point($Args[0], $Args[1])
    $textBox.Size = New-Object System.Drawing.Size($Args[2], $Args[3])
    $textBox.BackColor = "#ff99cc"
    $Args[4].Controls.Add($textBox)
    return $textBox
}


# 入力画面表示
# workPlaceArray : 勤務表から取得した勤務地一覧
for ($i = 0; $i -lt $workPlaceArray.Length; $i++) {

    # ---------------- Main ----------------- 

    # フォーム全体の設定
    # ウィンドウ本体を作っている

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "初期設定  【"+$workPlaceArray[$i]+"】"
    $form.Size = New-Object System.Drawing.Size(660, 700)
    $form.StartPosition = "CenterScreen"

    # フォームを配列に追加
    if ($isadded) {
        $forms += $form    
    }

    # OKボタン関数呼び出し
    drawOKButton 610 "OK" $forms[$i]

    # 在宅ボタン関数呼び出し
    drawAtHomeButton 610 "在宅" $forms[$i]

    # 戻るボタン関数呼び出し
    drawReturnButton 610 "戻る" $forms[$i]


    # =============================== input ===============================

    # 勤務地が「在宅」の場合は「在宅ボタン」を押下する
    $atHomeLabel = drawLabel 10 10 470 "★ 在宅勤務のときは【在宅】ボタンをクリック ★" $forms[$i]
    $atHomeLabel.forecolor = "red" 
    $atHomeLabel.font = $Font 


    # ---------------- 適用（行先、要件） ----------------- 
    $tekiyouLabelLocate = 50
    $tekiyouTextBoxLocate = 108

    # ラベル関数呼び出し
    drawLabel 10 $tekiyouLabelLocate 470 ("１．【 適用 】 勤務地 `""+$workPlaceArray[$i]+"`" の時の適用を入力してください") $forms[$i] | Out-Null
    drawLabel 20 ($tekiyouLabelLocate + 20) 470 "ex.  自宅←→田町本社" $forms[$i] | Out-Null
    drawLabel 20 ($tekiyouLabelLocate + 40) 470 "      自宅→品川→東京テレポート→自宅 (勤務地複数の場合)" $forms[$i] | Out-Null


    # テキストボックス関数呼び出し
    $outputTekiyou = drawTextBox 20 $tekiyouTextBoxLocate 300 20  $forms[$i]

    # 適用テキストボックスを配列に追加
    if ($isadded) {
        $outputTekiyous += $outputTekiyou    
    }

    # ---------------- 区間 ----------------- 
    $kukanLabelLocate = 150
    $kukanTextBoxLocate = 208

    # ラベル関数呼び出し
    drawLabel  10 $kukanLabelLocate 550 ("２．【 区間 】 勤務地 `""+$workPlaceArray[$i]+"`" の時の区間（自宅の最寄り駅←→勤務地の最寄り駅）を入力してください") $forms[$i] | Out-Null
    drawLabel 20 ($kukanLabelLocate + 20) 470 "ex.  <自宅の最寄り駅>←→田町 (往復の場合)" $forms[$i] | Out-Null
    drawLabel 20 ($kukanLabelLocate + 40) 670 "      <自宅の最寄り駅>→品川→東京テレポート→<自宅の最寄り駅> (勤務地複数の場合)" $forms[$i] | Out-Null


    # テキストボックス関数呼び出し
    $outputKukan = drawTextBox 20 $kukanTextBoxLocate 430 20 $forms[$i]

    # 適用テキストボックスを配列に追加
    if ($isadded) {
        $outputKukans += $outputKukan    
    }


    # ---------------- 交通機関 ----------------- 
    $koutsukikanLabelLocate = 290
    $koutsukikanTextBoxLocate = 288

    # ラベル関数呼び出し
    drawLabel 10 250 500 ("３．【 交通機関 】 勤務地 `""+$workPlaceArray[$i]+"`" の時に利用する交通機関を入力してください") $forms[$i] | Out-Null
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

    # ---------------- 金額 -----------------
    $kingakuLabelLocate = 530
    $kingakuTextBoxLocate = 570

    # ラベル関数呼び出し
    drawLabel 10 $kingakuLabelLocate 500 ("４．【 金額 】 勤務地 `""+$workPlaceArray[$i]+"`" の金額（往復代金）を入力してください") $forms[$i] | Out-Null
    drawLabel 20 ($kingakuLabelLocate + 20) 470 "ex.  750 （半角数字）" $forms[$i] | Out-Null

    # テキストボックス関数呼び出し
    $outputkingaku = drawTextBox 20 $kingakuTextBoxLocate 100 20 $forms[$i]

    # 適用テキストボックスを配列に追加
    if ($isadded) {
        $outputkingakus += $outputkingaku    
    }


    # 可視化
    #$forms[$i].Add_Shown({$textBox.Select()})
    $inputContentsResult = $forms[$i].ShowDialog()




    # =============================== output ===============================

    if ($inputContentsResult -eq "OK") {

        # ---------------- 適用（行先、要件） ----------------- 

        $inputContentsArray += @($workPlaceArray[$i] + "_" + $outputTekiyous[$i].text)

        # ---------------- 区間 -----------------

        $inputContentsArray += @($workPlaceArray[$i] + "_" + $outputKukans[$i].text)

        # ---------------- 交通機関 ----------------- 
        $inputKoutsukikan = @($koutsukikan1.text, $koutsukikan2.text, $koutsukikan3.text, $koutsukikan4.text, $koutsukikan5.text, $koutsukikan6.text)

        foreach ($koutsukikan in $inputKoutsukikan) {
            if ([string]::IsNullOrEmpty($koutsukikan)) {
                # NULL や '' の場合
                Write-Host 'NULL or Empty'
            }
            else {
                # 上記以外は設定された文字列を出力
                $koutsukikans += $koutsukikan + '`r`n'
            }
        }

        $outputkoutsukikans += $koutsukikans.Substring(0, $koutsukikans.Length - 4)

        $inputContentsArray += @($workPlaceArray[$i] + "_" + $outputkoutsukikans[$i])

        # ---------------- 金額 -----------------

        $inputContentsArray += @($workPlaceArray[$i] + "_" + $outputkingakus[$i].text)

        # フォーム増やしたフラグ
        $isadded = $True


    # 戻るボタンを押したら
    }elseif ($inputContentsResult -eq "Retry") {
        
        $i = $i-2
        # 固定配列なので、最初の要素は空にするだけにした
        if ($inputContentsArray.Length -le 4) {
            for ($j = 1; $j -lt 5; $j++) {
                $inputContentsArray[($inputContentsArray.Length-$j)] = ""
            }
        # 戻るボタンを押したら、テキストファイルに出力する要素を削除する    
        } else {
            $inputContentsArray = $inputContentsArray[0..($inputContentsArray.Length-5)]
        }
        
        $isadded = $false

    # 在宅ボタンを押したら
    }elseif ($inputContentsResult -eq "Yes") {
        $inputContentsArray += @($workPlaceArray[$i] + "_1")
        $inputContentsArray += @($workPlaceArray[$i] + "_1")
        $inputContentsArray += @($workPlaceArray[$i] + "_1")
        $inputContentsArray += @($workPlaceArray[$i] + "_1")
    }else{
        endExcel
        exit
    }    
}

# 完了画面がほしいなあ

foreach($inputContent in $inputContentsArray){
    $inputContent >> ツール用引数.txt
}

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
$koutsukikan = $null
$koutsukikans = $null
$outputkoutsukikan = $null
$outputkingaku = $null

