# 
# ���s�O�ɓ��c�[���ɂ��Ă̒��Ӊ�ʂ�\������
#

write-host "���Ӊ��.ps1���Ăяo����܂���"

# ���x���쐬�֐�
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

# �t�H�[�����쐬����֐�
function attentionThisTool {
    . {
        # �t�H�[���̐ݒ�
        $attentionForm = New-Object System.Windows.Forms.Form
        $attentionForm.Text = "�� ���ӎ��� ��"
        $attentionForm.Size = New-Object System.Drawing.Size(550, 250)
        $attentionForm.StartPosition = "CenterScreen"
        $attentionForm.font = $font
        $attentionForm.Topmost = $True

        $tate = 10
        $yoko = 10
        
        drawAttentionLabel $yoko $tate 400 "�y�P�z�{�c�[���͑I���������̋Ζ��\�����ƂɁA���o�^�̋Ζ��n��o�^���܂��B" $attentionForm | Out-Null
        drawAttentionLabel ($yoko + 30) ($tate + 20) 500 "�u�_�E�����[�h�����Ζ��\�v�t�H���_��SharePoint����Ζ��\���_�E�����[�h���Ă��������B" $attentionForm | Out-Null
        drawAttentionLabel $yoko ($tate + 50) 500 "�y�Q�z�{�c�[���͋Ζ��\�́u�Ζ����e�v�������́u��Əꏊ�v����Ζ��n��o�^���܂��B" $attentionForm | Out-Null
        drawAttentionLabel ($yoko + 30) ($tate + 70) 500 "�܂��A�u�Ζ����сv�ɋL�����Ȃ��ꍇ�͋x���Ɣ��f���A�o�^���܂���B" $attentionForm | Out-Null
        drawAttentionLabel $yoko ($tate + 115) 500 "��L�̎������m�F�ł��܂�����uOK�v���N���b�N���A�Ζ��n�̓o�^���s���Ă��������B" $attentionForm | Out-Null
        
        # OK�{�^���̐ݒ�
        $OKButton = New-Object System.Windows.Forms.Button
        $OKButton.Location = New-Object System.Drawing.Point(330, 160)
        $OKButton.Size = New-Object System.Drawing.Size(75, 30)
        $OKButton.Text = "OK"
        $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $attentionForm.AcceptButton = $OKButton
        $attentionForm.Controls.Add($OKButton)
        
        # �L�����Z���{�^���̐ݒ�
        $CancelButton = New-Object System.Windows.Forms.Button
        $CancelButton.Location = New-Object System.Drawing.Point(430, 160)
        $CancelButton.Size = New-Object System.Drawing.Size(75, 30)
        $CancelButton.Text = "Cancel"
        $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $attentionForm.CancelButton = $CancelButton
        $attentionForm.Controls.Add($CancelButton)
        
        return
        
        # ����
        $attentionResult = $attentionForm.ShowDialog()
    } | Out-null
    
    # �t�H�[����߂�l�Ƃ��ĕԂ�
    return $attentionForm
}