#
# �R���{�{�b�N�X��\�����āA�쐬�������Ζ��\�̌���I��������
#

Write-Host "�Ζ��\�̌��I���X�N���v�g���Ăяo����܂����B"

function choiceMonth {
    $rtnVal = ""
    . {
        # �t�H���g�̎w��
        $Font = New-Object System.Drawing.Font("Yu Gothic UI", 8)

        # �t�H�[���S�̂̐ݒ�
        $form = New-Object System.Windows.Forms.Form
        $form.Text = "�쐬���鏬���̑Ώ۔N��"
        $form.Size = New-Object System.Drawing.Size(265, 200)
        $form.StartPosition = "CenterScreen"
        $form.font = $Font

        # ���x����\��
        $label = New-Object System.Windows.Forms.Label
        $label.Location = New-Object System.Drawing.Point(10, 10)
        $label.Size = New-Object System.Drawing.Size(270, 30)
        $label.Text = "�쐬�����������̔N����I�����Ă�������"
        $form.Controls.Add($label)

        # OK�{�^���̐ݒ�
        $OKButton = New-Object System.Windows.Forms.Button
        $OKButton.Location = New-Object System.Drawing.Point(40, 100)
        $OKButton.Size = New-Object System.Drawing.Size(75, 30)
        $OKButton.Text = "OK"
        $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.AcceptButton = $OKButton
        $form.Controls.Add($OKButton)

        # �L�����Z���{�^���̐ݒ�
        $CancelButton = New-Object System.Windows.Forms.Button
        $CancelButton.Location = New-Object System.Drawing.Point(130, 100)
        $CancelButton.Size = New-Object System.Drawing.Size(75, 30)
        $CancelButton.Text = "Cancel"
        $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form.CancelButton = $CancelButton
        $form.Controls.Add($CancelButton)

        # �R���{�{�b�N�X���쐬
        $Combo = New-Object System.Windows.Forms.Combobox
        $Combo.Location = New-Object System.Drawing.Point(50, 50)
        $Combo.size = New-Object System.Drawing.Size(150, 30)
        # ���X�g�ȊO�̓��͂������Ȃ�
        $Combo.DropDownStyle = "DropDownList"
        $Combo.FlatStyle = "standard"
        $Combo.BackColor = "#005050"
        $Combo.ForeColor = "white"
    
        # -----------�R���{�{�b�N�X�ɍ��ڂ�ǉ�-----------
        for ($counterForMove = (-6); $counterForMove -le 6; $counterForMove++) {
            $date = get-date (get-date).AddMonths($counterForMove) -Format Y
            [void] $Combo.Items.Add("$date")
        }

        # �t�H�[���ɃR���{�{�b�N�X��ǉ�
        $form.Controls.Add($Combo)
        $Combo.SelectedIndex = 6

        # �t�H�[�����őO�ʂɕ\��
        $form.Topmost = $True


    } | Out-Null
    return $form, $Combo
}
