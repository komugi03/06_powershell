#
# コンボボックスを表示して、作成したい勤務表の月を選択させる
#

Write-Host "勤務表の月選択スクリプトが呼び出されました。"

function choiceMonth {
    $rtnVal = ""
    . {
        # フォントの指定
        $Font = New-Object System.Drawing.Font("Yu Gothic UI", 8)

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


    } | Out-Null
    return $form, $Combo
}
