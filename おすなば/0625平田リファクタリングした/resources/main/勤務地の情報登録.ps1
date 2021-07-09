#
#
#



# --------------- アセンブリの読み込み ---------------
Add-Type -Assembly System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---------------- 注意画面スクリプトの呼び出し -------------------
. (Join-Path -path $PWD -childpath "resources\scripts\注意画面.ps1")

# 注意画面.ps1の関数を実行する
$attentionForm = attentionThisTool

# フォームの可視化
$attentionResult = $attentionForm.ShowDialog()

if ($attentionResult -eq "Cancel") {
    exit 
}
# ---------------- 注意画面スクリプト終了 -------------------

# 現在の年月日を取得する
$targetYear = (Get-Date).Year
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
$yesNo_yearMonthAreCorrect = [System.Windows.Forms.MessageBox]::Show("【 $targetYear 年 $targetMonth 月 】の勤務表をもとに初期設定をしますか？`r`n`r`n「いいえ」で他の月を選択できます", '作成する小口の対象年月', 'YesNo', 'Question')

# ☆$yesNo_yearMonthAreCorrect -eq 'No'ループ開始☆
if ($yesNo_yearMonthAreCorrect -eq 'No') {

    # ---------------- 勤務表の月選択スクリプトを呼び出し -------------------
    . (Join-Path -Path $PWD -ChildPath ".\resources\scripts\勤務表の月選択.ps1")

    # 勤務表の月選択.ps1の関数を実行
    $choicedMonth = choiceMonth

    if ($choicedMonth -eq "OK") {
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
}
# ---------------- 勤務表の月選択スクリプト終了 -------------------

# ----------- しばらくお待ちくださいスクリプト呼び出し -----------
. (Join-Path -path $PWD -ChildPath "\resources\scripts\しばらくお待ちください.ps1")

# しばらくお待ちください.ps1の関数を実行
$waitCatForm = pleaseWait "resources/pictures/お待ちください猫.png"

# しばらくお待ちくださいフォームの可視化
$waitCatForm.show()
start-sleep -second 2
$waitCatForm.close()

# ---------------- 勤務表チェック処理 -------------------
# ポップアップを作成
$popup = new-object -comobject wscript.shell

# ファイル名の勤務表_のあとの表記
$targetMonth00 = "{0:00}" -f [int]$targetMonth
$fileNameMonth = ("$targetYear" + "$targetMonth00")
# 勤務表ファイルを取得
$kinmuhyou = Get-ChildItem -Recurse -File | ? Name -Match ("[0-9]{3}_勤務表_" + $fileNameMonth + "_.+") 
# 該当勤務表ファイルの個数確認
if ($kinmuhyou.Count -lt 1) {
    
    # しばらくお待ちください画面を閉じる
    $waitForm.Close()

    # ポップアップを表示
    $popup.popup("$targetMonth 月の勤務表ファイルが存在しません", 0, "やり直してください", 48) | Out-Null
    exit
}
elseif ($kinmuhyou.Count -gt 1) {
    
    # しばらくお待ちください画面を閉じる
    $waitForm.Close()

    # ポップアップを表示
    $popup.popup("$targetMonth 月の勤務表ファイルが多すぎます`r`n1つにしてください", 0, "やり直してください", 48) | Out-Null
    exit
}
# ---------------- 勤務表チェック処理終了 -------------------

# ----------------------Excelを起動処理--------------------------------
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

# ----------------------Excelを起動処理終了--------------------------------



