# 
# 勤務表をもとに小口交通費請求書を作成するPowershell
# 
# 勤務表のファイル名：<3桁の社員番号>_勤務表_m月_<氏名>.xlsx
# 

# ----------------- 関数定義 ---------------------

# 勤務表と小口を保存せずに閉じて、Excelを中断する関数
function breakExcel {
    # Bookを閉じる
    $kinmuhyouBook.close()
    $koguchiBook.close()
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
# 最大文字数を基準にシャープの長さを決定する
# 引数1 : 文字色
# 引数2以降 : メッセージ
function displayMessagesSurroundedBySharp {
    # 変数の初期化
    $maxLengths = 0
    for ($i = 1; $i -lt $Args.length; $i++) {
        # メッセージの中で一番長い文字数を取得する
        if ( $maxLengths -lt $Args[$i].length) {
            $maxLengths = $Args[$i].length
        }
    }
    # メッセージの表示
    Write-Host ("`r`n" + '#' * ($maxLengths * 2 + 6) + "`r`n") -ForegroundColor $Args[0]
    for ($i = 1; $i -lt $Args.length; $i++) {
        Write-Host ('　　' + $Args[$i] + "　　`r`n") -ForegroundColor $Args[0]
    }
    Write-Host ('#' * ($maxLengths * 2 + 6) + "`r`n") -ForegroundColor $Args[0]
}

# 引数の空白を除きファイル名として使えない文字を消す関数
# fileName : ファイル名
function removeInvalidFileNameChars ($fileName) {
    $fileNameRemovedSpace = $fileName -replace "　", ""　-replace " ", ""
    $invalidChars = [IO.Path]::GetInvalidFileNameChars() -join ''
    $regex = "[{0}]" -f [RegEx]::Escape($invalidChars)
    return $fileNameRemovedSpace -replace $regex
}

# -------------------- 主処理 --------------------------

##### 注意書きを表示。問題ない場合にはEnterを押させる。#####

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

# ---------------アセンブリの読み込み---------------
Add-Type -Assembly System.Windows.Forms
Add-Type -AssemblyName System.Drawing
# # INPUTのために必要
# [void][System.Reflection.Assembly]::Load("Microsoft.VisualBasic, Version=8.0.0.0, Culture=Neutral, PublicKeyToken=b03f5f7f11d50a3a")

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
    $label.Text = "作成する小口の年月を選択してください`r`n※前月～翌月が選択できます※"
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
        $comboAnswer = $Combo.Text -split "年"

        # ユーザー指定の年を小口作成の対象年として上書する
        $targetYear = $comboAnswer[0]

        # ユーザー指定の月を小口作成の対象月として上書きする
        $targetMonth = $comboAnswer[1] -split "月"

    }else{
        # 処理を終了する
        exit
    }

# ☆$yesNo_yearMonthAreCorrect -eq 'No'ループ終了☆
}


echo "$targetYear 年の"
echo "$targetMonth 月の小口を作成します"



# -------（場所迷い中）---------------小口テンプレを取得------------------------
$koguchiTemplate = Get-ChildItem -Recurse -File | ? Name -Match "小口交通費・出張旅費精算明細書_テンプレ.xlsx"
# 該当小口ファイルの個数確認
if ($koguchiTemplate.Count -lt 1) {
    Write-Host "`r`n該当する小口ファイルが存在しません`r`n`r`nダウンロードし直してください`r`n" -ForegroundColor Red
    exit
}
elseif ($koguchiTemplate.Count -gt 1) {
    Write-Host "`r`n該当する小口ファイルが多すぎます`r`n`r`nダウンロードし直してください`r`n" -ForegroundColor Red
    exit
}

# ------（ユーザー指定の月が必要だから、コンボボックスより後）----------テンプレートから小口交通費請求書を作成する---------------------
# 作成した小口を格納するフォルダに、テンプレートをコピーする
# ※フォルダが存在していないとエラーが出る
$koguchi = Join-Path $PWD "作成した小口交通費請求書" | Join-Path -ChildPath "小口交通費・出張旅費精算明細書_コピー.xlsx"
Copy-Item -path $koguchiTemplate.FullName -Destination $koguchi

# 勤務表ファイルを取得
$kinmuhyou = Get-ChildItem -Recurse -File | ? Name -Match "[0-9]{3}_勤務表_($targetMonth)月_.+"

# 該当勤務表ファイルの個数確認
if ($kinmuhyou.Count -lt 1) {
    Write-Host "`r`n該当する勤務表ファイルが存在しません`r`n" -ForegroundColor Red
    exit
}
elseif ($kinmuhyou.Count -gt 1) {
    Write-Host "`r`n該当する勤務表ファイルが多すぎます`r`n" -ForegroundColor Red
    exit
}










# 勤務表からとってくる勤務地の情報は「勤務内容」の列からだけでOK