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
        $attentionForm.Size = New-Object System.Drawing.Size(550, 270)
        $attentionForm.StartPosition = "CenterScreen"
        $attentionForm.formborderstyle = "FixedSingle"
        $attentionForm.font = $font
        $attentionForm.Topmost = $True
        $attentionForm.icon = (Join-Path -Path $PWD -ChildPath "../images/会社アイコン.ico")

        $tate = 15
        $yoko = 10
        
        drawAttentionLabel $yoko $tate 500 "下記の事項が確認できたら「OK」をクリックし、勤務地の登録を行ってください。" $attentionForm | Out-Null
        drawAttentionLabel $yoko ($tate + 35) 550 "【１】Excelファイルを開いている場合は閉じてください。データが破損する可能性があります。" $attentionForm | Out-Null
        drawAttentionLabel $yoko ($tate + 65) 500 "【２】勤務表の、勤務実績がある日のみ「勤務内容」欄から勤務地を取得し登録します。" $attentionForm | Out-Null
        drawAttentionLabel ($yoko + 25) ($tate + 85) 500 "「勤務内容」欄が空欄の場合は「作業場所」欄から勤務地を取得し登録します。" $attentionForm | Out-Null
        drawAttentionLabel $yoko ($tate + 115) 550 "【３】「01_ダウンロードした勤務表」フォルダに勤務表をダウンロードしてください。" $attentionForm | Out-Null
        drawAttentionLabel $yoko ($tate + 145) 550 "【４】本ツール使用中、開いた覚えのないExcelファイルが表示された場合は「×」をクリックしないでください。" $attentionForm | Out-Null
        
        # OKボタンの設定
        $OKButton = New-Object System.Windows.Forms.Button
        $OKButton.Location = New-Object System.Drawing.Point(330, 190)
        $OKButton.Size = New-Object System.Drawing.Size(75, 30)
        $OKButton.Text = "OK"
        $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $attentionForm.AcceptButton = $OKButton
        $attentionForm.Controls.Add($OKButton)
        
        # キャンセルボタンの設定
        $CancelButton = New-Object System.Windows.Forms.Button
        $CancelButton.Location = New-Object System.Drawing.Point(430, 190)
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