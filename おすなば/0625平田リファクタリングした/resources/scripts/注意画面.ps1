# 
# 実行前に当ツールについての注意画面を表示する
#

write-host "注意画面.ps1が呼び出されました"

# ラベル作成関数
function drawAttentionLabel {
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

# フォームを作成する関数
function attentionThisTool {
    . {
        # フォームの設定
        $attentionForm = New-Object System.Windows.Forms.Form
        $attentionForm.Text = "※ 注意事項 ※"
        $attentionForm.Size = New-Object System.Drawing.Size(550, 250)
        $attentionForm.StartPosition = "CenterScreen"
        $attentionForm.font = $font
        $attentionForm.Topmost = $True

        $tate = 10
        $yoko = 10
        
        drawAttentionLabel $yoko $tate 400 "【１】本ツールは選択した月の勤務表をもとに、未登録の勤務地を登録します。" $attentionForm | Out-Null
        drawAttentionLabel ($yoko + 30) ($tate + 20) 500 "「ダウンロードした勤務表」フォルダにSharePointから勤務表をダウンロードしてください。" $attentionForm | Out-Null
        drawAttentionLabel $yoko ($tate + 50) 500 "【２】本ツールは勤務表の「勤務内容」もしくは「作業場所」から勤務地を登録します。" $attentionForm | Out-Null
        drawAttentionLabel ($yoko + 30) ($tate + 70) 500 "また、「勤務実績」に記入がない場合は休日と判断し、登録しません。" $attentionForm | Out-Null
        drawAttentionLabel $yoko ($tate + 115) 500 "上記の事項が確認できましたら「OK」をクリックし、勤務地の登録を行ってください。" $attentionForm | Out-Null
        
        # OKボタンの設定
        $OKButton = New-Object System.Windows.Forms.Button
        $OKButton.Location = New-Object System.Drawing.Point(330, 160)
        $OKButton.Size = New-Object System.Drawing.Size(75, 30)
        $OKButton.Text = "OK"
        $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $attentionForm.AcceptButton = $OKButton
        $attentionForm.Controls.Add($OKButton)
        
        # キャンセルボタンの設定
        $CancelButton = New-Object System.Windows.Forms.Button
        $CancelButton.Location = New-Object System.Drawing.Point(430, 160)
        $CancelButton.Size = New-Object System.Drawing.Size(75, 30)
        $CancelButton.Text = "Cancel"
        $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $attentionForm.CancelButton = $CancelButton
        $attentionForm.Controls.Add($CancelButton)
        
        return
        
        # 可視化
        $attentionResult = $attentionForm.ShowDialog()
    } | Out-null
    
    # フォームを戻り値として返す
    return $attentionForm
}