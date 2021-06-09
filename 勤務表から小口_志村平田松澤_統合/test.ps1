# ---------------アセンブリの読み込み---------------
# 
Add-Type -Assembly System.Windows.Forms
Add-Type -AssemblyName System.Drawing
# INPUTのために必要
[void][System.Reflection.Assembly]::Load("Microsoft.VisualBasic, Version=8.0.0.0, Culture=Neutral, PublicKeyToken=b03f5f7f11d50a3a")

# 現在の年月日を取得する
$thisYear = (Get-Date).Year
# $thisMonth = 12
$thisMonth = (Get-Date).Month
$today = (Get-Date).Day

# 現在日時から作成するべき勤務表の月次を判定
# 24日までは当月分を作る
if ($today -le 24) {
    $targetMonth = $thisMonth - 1
}
else {
    $targetMonth = $thisMonth
}

$yesNo_yearMonthAreCorrect = [System.Windows.Forms.MessageBox]::Show("作成するのは 【 $thisYear 年 $targetMonth 月 】の小口でよろしいですか？",'作成する小口の対象年月','YesNo','Question')

if($yesNo_yearMonthAreCorrect -eq 'No'){

    # フォントの指定
    $Font = New-Object System.Drawing.Font("メイリオ",8)

    # フォーム全体の設定
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "作成する小口の対象年月"
    $form.Size = New-Object System.Drawing.Size(300,200)
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
    
    # コンボボックスの選択肢に使う変数を準備
    # 去年
    $lastYear = $thisYear - 1
    # 来年
    $nextYear = $thisYear + 1
    
    # -----------コンボボックスに項目を追加-----------
    # 前の月
    if($thisMonth -eq '1'){
        # 去年の12月=1月に作成している
        $lastMonth = 12
        [void] $Combo.Items.Add("$lastYear 年 $lastMonth 月")
    }else{
        # 今年の先月
        $lastMonth = $thisMonth - 1
        [void] $Combo.Items.Add("$thisYear 年 $lastMonth 月")
    }
    # 今年の当月
    [void] $Combo.Items.Add("$thisYear 年 $thisMonth 月")
    
    # 次の月
    if($thisMonth -eq '12'){
        # 来年の1月=12月に作成している
        $nextMonth = 1
        [void] $Combo.Items.Add("$nextYear 年 $nextMonth 月")
    }else{
        # 今年の翌月
        $nextMonth = $thisMonth + 1
        [void] $Combo.Items.Add("$thisYear 年 $nextMonth 月")
    }
    
    # フォームにコンボボックスを追加
    $form.Controls.Add($Combo)
    $Combo.SelectedIndex = 1
    
    # フォームを最前面に表示
    $form.Topmost = $True
    
    # フォームを表示＋選択結果を変数に格納
    $result = $form.ShowDialog()

    # 選択後、OKボタンが押された場合、選択項目を表示
    if ($result -eq "OK"){
        $comboAnswer = $Combo.Text
    }else{
    exit
    }

    Write-Output $comboAnswer

    # 去年の12月なら
    if($comboAnswer -eq "$lastYear 年 $lastMonth 月"){
        echo "去年の12月だってさ"
    }

    # 今年の先月なら
    elseif($comboAnswer -eq "$thisYear 年 $lastMonth 月"){
        echo "先月だってよ"
    }

    # 当月なら
    elseif($comboAnswer -eq "$thisYear 年 $thisMonth 月"){
        echo "今月のだよ"
    }

    # 来年の1月なら
    elseif($comboAnswer -eq "$nextYear 年 $nextMonth 月") {
        echo "来年の1月！"
    }

    # 翌月なら
    elseif($comboAnswer -eq "$thisYear 年 $nextMonth 月") {
        echo "次の月！"
    }

}