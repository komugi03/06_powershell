# $wsobj = new-object -comobject wscript.shell
# $result = $wsobj.popup("Hello!Project",0,"たいとる",0)

# function add ($x, $y) {
#     $result = $x + $y
#     $result    
# }

# add 1 2


# $koguchi = Join-Path $PWD "作成した小口交通費請求書" | Join-Path -ChildPath "小口交通費・出張旅費精算明細書_コピー.xlsx"
# Remove-Item -Path $koguchi


# if(Test-Path $PWD"\作成した小口交通費請求書"){
    #     echo "OK"
    # }
    # else{
        #     New-Item -Path $PWD"\作成した小口交通費請求書" -ItemType Directory | Out-Null
        # }
        
        # $i = 1
        # $i
        # # $koutsukikan1.textにしたい
        # $koi = ('$koutsukikan' + $i)
        # $koi
        
        # $targetMonth = (Get-Date).Month
        # $fileNameMonth = "$targetMonth 月"
        # $kinmuhyou = Get-ChildItem -Recurse -File | ? Name -Match "[0-9]{3}_勤務表_$fileNameMonth_.+"
        # $kinmuhyou
        
        
        # $workPlace = "お台場"
        # if ($workPlace -ne "" -and $workPlace -ne '在宅') {
        #     echo "OkKoguchi"
        # }else{
        #     echo "NoKoguchi"
        # }

# if(Test-Path $PWD"\ツール用引数.txt"){
#     $argumentText = (Get-Content $PWD"\ツール用引数.txt")[0..3]
#     $argumentText[0]
# }else{
#     Write-Output "ファイルはありません"
# }

# $workPlace = 'お'

# if(Test-Path $PWD"\ツール用引数.txt"){
#     $argumentText = (Get-Content $PWD"\ツール用引数.txt")
#     # 勤務地の情報をリストから取得 ( 配列の中身　[0]:適用　[1]:区間　[2]:交通機関　[3]:金額 )
#     $workPlaceInfo = $argumentText | Select-String -Pattern ($workPlace + '_')
#     if($workPlaceInfo -eq $null){
    
#         echo "catch!"
#     }

#     # $workPlaceInfo

#     # $workPlaceInfo[0]

# }else{
#     Write-Output "ファイルはありません"
# }

# $infomationTextFileName = "ツール用引数.txt"
# Test-Path $PWD"\"$infomationTextFileName

# $infoTextFileName = "ツール用引数.txt"
# $infoTextFileFullpath = "$PWD\$infoTextFileName"
# $infoTextFileFullpath
        
# if(Test-Path $infoTextFileFullpath){echo "OK"}


# $infoTextFileName = "ツール用引数.txt"
# $infoTextFileFullpath = "$PWD\$infoTextFileName"
# $argumentText = (Get-Content $infoTextFileFullpath)

# $workPlace = "お台場"
# $workPlaceInfo = $argumentText | Select-String -Pattern ($workPlace + '_')
# Write-Host ("勤務地list：" + $workPlaceInfo)

# Write-Host ("適用の行：" + $workPlaceInfo[0])


# $tekiyouText = [String]$workPlaceInfo[0]
# $tekiyouText

# $tekiyouText = $tekiyouText.Substring(4, $tekiyouText.Length - 4)
# Write-Host ("適用：" + $tekiyouText)

# $tekiyouText = ([String]$workPlaceInfo[0]).Substring(4, ([String]$workPlaceInfo[0]).Length - 4)
# Write-Host ("適用：" + $tekiyouText)


# プログレスバー
Add-Type -AssemblyName System.Windows.Forms

$formProgressBar = New-Object System.Windows.Forms.Form
$formProgressBar.Size = "300,200"
$formProgressBar.Startposition = "CenterScreen"
$formProgressBar.Text = "作成中…"

# $Button = New-Object System.Windows.Forms.Button
# $Button.Location = "110,20"
# $Button.Size = "80,30"
# $Button.Text = "開始"

# # ボタンのクリックイベント
# $Start = {
#     # For ( $i = 0 ; $i -lt 10 ; $i++ )
#     # {
#     #     $progressBar.Value = $i+1
#     #     start-sleep -s 1
#     # }
#     [System.Windows.Forms.MessageBox]::Show("お待たせしました！完成です！", "info")
# }
# $Button.Add_Click($Start)

# # プログレスバー
# $progressBar = New-Object System.Windows.Forms.ProgressBar
# $progressBar.Location = "10,100"
# $progressBar.Size = "260,30"
# $progressBar.Maximum = "10"
# $progressBar.Minimum = "0"
# $progressBar.Style = "Continuous"

# $progressBar.Value = 1


# # $formProgressBar.Controls.AddRange(@($progressBar,$Button))
# $formProgressBar.Controls.AddRange($progressBar)

# $formProgressBar.Topmost = $True

# $formProgressBar.Show()


# $progressBar.Value++

# $progressBar.Value++

# $formProgressBar.Show()
# # Start-Sleep -milliSeconds 300
# # $formProgressBar.Close()
# $formProgressBar
# $progressBar.Value++


# $progressBar.Value++
# $formProgressBar.Show()

# Start-Sleep -milliSeconds 500
# $progressBar.Value += 5

# # $formProgressBar.Visible = $false
# $finish = $formProgressBar.Show()
# $formProgressBar.Close()
# write-host $finish

# Write-Progress -Activity "処理中" -Status "現在の状態" -PercentComplete 20 -SecondsRemaining 5



# # プログレスバー　サンプル

# For($b = 1 ; $b -le 10000 ; $b++)
# {
    # $c = $b / 100
# write-progress -Activity "数値を加算しています" -Status "しばらくお待ちください"　-PercentComplete $c -CurrentOperation "$c % 完了"
# }



# Function Invoke-ProgressBar {
#     Param([Int]$Minimum, [Int]$Maximum, [ScriptBlock]$ScriptBlock)

#     Add-Type -AssemblyName System.Windows.Forms

#     $formProgressBar = New-Object System.Windows.Forms.Form
#     #$formProgressBar.Text = 'Progress Bar'
#     $formProgressBar.Width = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Width / 2
#     $formProgressBar.Height = $formProgressBar.Width / 5
#     $formProgressBar.StartPosition = 'CenterScreen'
#     $formProgressBar.TopMost = $true

#     $progressBar = New-Object System.Windows.Forms.ProgressBar
#     $progressBar.Width = $formProgressBar.ClientRectangle.Width * 0.9
#     $progressBar.Height = $formProgressBar.ClientRectangle.Height * 0.3
#     $progressBar.Left = ($formProgressBar.ClientRectangle.Width - $progressBar.Width) / 2
#     $progressBar.Top = ($formProgressBar.ClientRectangle.Height - $progressBar.Height) / 2
#     $progressBar.Visible = $true
#     $progressBar.Style = 'Continuous'
#     $progressBar.Minimum = $Minimum
#     $progressBar.Maximum = $Maximum
#     $progressBar.Value = $Minimum
#     $progressBar.Step = 1
#     $formProgressBar.Controls.Add($progressBar)

#     $formProgressBar.Add_Shown({
#         For ($i = $Minimum; $i -le $Maximum; $i++) {
#             & $ScriptBlock
#             $formProgressBar.Text = ($progressBar.Value / ($Maximum - $Minimum)).ToString('0%')
#             $progressBar.PerformStep()
#         }
#         $formProgressBar.Close()
#     })
#     [void]$formProgressBar.ShowDialog()
# }

# Invoke-ProgressBar -Minimum 0 -Maximum 42 -ScriptBlock {
#     # Some time-comsuming task ...
#     Start-Sleep -Milliseconds 100
#     Write-Host $i
# }


# $koutsu = 'お台場_小田急線`r`nJR山手線`r`nりんかい線'
# $koutsukikanText = $koutsu.Substring(4, $koutsu.length - 4)
# $koutsukikanText

# $koutsukikanArray = $koutsukikanText -split '`r`n'

# $koutsukikanArray.Length

# for ($i = 0; $i -lt $koutsukikanArray.Length; $i++) {
#     # 改行コードを足す
#     $koutsukikanKaigyou += $koutsukikanArray[$i] + "`r`n"
# }

# # 最後の改行を削除する
# $koutsukikanKaigyou = $koutsukikanKaigyou.Substring(0, $koutsukikanKaigyou.Length - 1)
# $koutsukikanKaigyou




# # --------新しい小口ファイル名を用意---------
# # <社員番号>_小口交通費・出張旅費精算明細書_YYYYMM_<氏名>
# # '116_小口交通費・出張旅費精算明細書_202105_志村瞳.xlsx'
# $koguchiNewFileName = '116' + "_小口交通費・出張旅費精算明細書_" + '202105' + "_" + '志村瞳'
# # ファイル名に使えない文字が入っていたら削除する(氏名の間の空白など)
# # $koguchiNewFileName = remove-invalidFileNameChars $koguchiNewFileName
# # 新しい小口ファイルのフルパス
# $koguchiNewfullPath = Join-Path $PWD "作成した小口交通費請求書" | Join-Path -ChildPath $koguchiNewFileName

# $koguchiNewfullPath

# # if (Test-Path ($koguchiNewfullPath + '_' + '[0-9][0-9]' + '.xlsx')) {
# #     echo "OK"
# # }

# Test-Path ($koguchiNewfullPath + "_.+" + '.xlsx')



# $targetMonth = "2"
# $fileNameMonth = "{0:D2}" -f [int]$targetMonth
# $fileNameMonth



# # ---------------アセンブリの読み込み---------------
# Add-Type -AssemblyName System.Windows.Forms
# Add-Type -AssemblyName System.Drawing

# # フォーム全体の設定をする関数
# # formText : フォームの本文（文字列）
# # formYoko : フォームの横幅
# # formTate : フォームの縦幅
# function makeForm ($formText, $formYoko, $formTate) {
#     $form = New-Object System.Windows.Forms.Form
#     $form.Text = $formText
#     $form.Size = New-Object System.Drawing.Size($formYoko,$formTate)
#     $form.StartPosition = "CenterScreen"
#     $form.font = $Font
#     return $form
# }

# # ラベルを表示する関数
# # $labelText : ラベルに書き込む文字列
# # $form : フォームオブジェクト
# function makeLabel ($labelText, $form) {
#     $label = New-Object System.Windows.Forms.Label
#     $label.Location = New-Object System.Drawing.Point(10,10)
#     $label.Size = New-Object System.Drawing.Size(270,30)
#     $label.Text = $labelText
#     $form.Controls.Add($label)
#     return $form
# }


# # フォーム全体の設定
# $form = makeForm "作成する小口の対象年月" 265 200

# # ラベルを表示
# $label = makeLabel "作成したい小口の年月を選択してください" $form

# # OKボタンの設定
# $OKButton = New-Object System.Windows.Forms.Button
# $OKButton.Location = New-Object System.Drawing.Point(40,100)
# $OKButton.Size = New-Object System.Drawing.Size(75,30)
# $OKButton.Text = "OK"
# $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
# $form.AcceptButton = $OKButton
# $form.Controls.Add($OKButton)

# # キャンセルボタンの設定
# $CancelButton = New-Object System.Windows.Forms.Button
# $CancelButton.Location = New-Object System.Drawing.Point(130,100)
# $CancelButton.Size = New-Object System.Drawing.Size(75,30)
# $CancelButton.Text = "Cancel"
# $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
# $form.CancelButton = $CancelButton
# $form.Controls.Add($CancelButton)

# # フォームを最前面に表示
# $form.Topmost = $True

# # フォームを表示＋選択結果を変数に格納
# $result = $form.ShowDialog()

# # 選択後、OKボタンが押された場合、選択項目を表示
# if ($result -eq "OK"){
#     # ユーザーの回答を"年"で区切る
#     $Combo.Text -match "(?<year>.+?)年(?<month>.+?)月" | out-null

#     # ユーザー指定の年を小口作成の対象年として上書する
#     $targetYear = $Matches.year

#     # ユーザー指定の月を小口作成の対象月として上書きする
#     $targetMonth = $Matches.month

# }else{
#     # 処理を終了する
#     exit
# }



# # ポップアップを作成
# $popup = new-object -comobject wscript.shell

# # 正常に終了したときポップアップを表示
# $successEnd = $popup.popup("お待たせしました！正常に終了しました`r`n仕上がりを確認してください",0,"正常終了",64)     

# if($successEnd -eq '1'){
#     Start-Process $PWD"\作成した小口交通費請求書"
#     # Invoke-Item $PWD"\作成した小口交通費請求書"
# }


# $wsobj = new-object -comobject wscript.shell
# $result = $wsobj.popup("Hello!Project",0,"wa",64)
# $result


# 作成する小口の年月が合っているか確認するダイアログを表示
# (現在日によって変わるので、get-date -Format Y にはしていない)
# $yesNo_yearMonthAreCorrect = [System.Windows.Forms.MessageBox]::Show("作成するのは 【 $thisYear 年 $targetMonth 月 】の小口でよろしいですか？`r`n`r`n「いいえ」で他の月を選択できます",'作成する小口の対象年月','YesNo','Question')
# $yesNo_yearMonthAreCorrect

# [System.Drawing.FontFamily]::Families


# ポップアップを作成
$popup = new-object -comobject wscript.shell

$targetPersonName = "志村瞳"
# $targetPersonName = '平田 隆'
# $targetPersonName = "松澤　夏海"

if($targetPersonName -match ' ' -or $targetPersonName -match '　'){
    # $targetPersonName = $targetPersonName -replace '　', '' -replace ' ', ''
    $targetPersonName = $targetPersonName.replace('　', '  ')
    $targetPersonName = $targetPersonName.replace(' ', '')
}

$successEnd = $popup.popup($targetPersonName + "さん : )`r`nOKを押して不備がないか確認してください",0,"お待たせしました！",64)    

