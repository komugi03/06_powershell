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
        $attentionForm.Size = New-Object System.Drawing.Size(550, 270)
        $attentionForm.StartPosition = "CenterScreen"
        $attentionForm.formborderstyle = "FixedSingle"
        $attentionForm.font = $font
        $attentionForm.Topmost = $True
        $attentionForm.icon = (Join-Path -Path $PWD -ChildPath "../images/��ЃA�C�R��.ico")

        $tate = 15
        $yoko = 10
        
        drawAttentionLabel $yoko $tate 500 "���L�̎������m�F�ł�����uOK�v���N���b�N���A�Ζ��n�̓o�^���s���Ă��������B" $attentionForm | Out-Null
        drawAttentionLabel $yoko ($tate + 35) 550 "�y�P�zExcel�t�@�C�����J���Ă���ꍇ�͕��Ă��������B�f�[�^���j������\��������܂��B" $attentionForm | Out-Null
        drawAttentionLabel $yoko ($tate + 65) 500 "�y�Q�z�Ζ��\�́A�Ζ����т�������̂݁u�Ζ����e�v������Ζ��n���擾���o�^���܂��B" $attentionForm | Out-Null
        drawAttentionLabel ($yoko + 25) ($tate + 85) 500 "�u�Ζ����e�v�����󗓂̏ꍇ�́u��Əꏊ�v������Ζ��n���擾���o�^���܂��B" $attentionForm | Out-Null
        drawAttentionLabel $yoko ($tate + 115) 550 "�y�R�z�u01_�_�E�����[�h�����Ζ��\�v�t�H���_�ɋΖ��\���_�E�����[�h���Ă��������B" $attentionForm | Out-Null
        drawAttentionLabel $yoko ($tate + 145) 550 "�y�S�z�{�c�[���g�p���A�J�����o���̂Ȃ�Excel�t�@�C�����\�����ꂽ�ꍇ�́u�~�v���N���b�N���Ȃ��ł��������B" $attentionForm | Out-Null
        
        # OK�{�^���̐ݒ�
        $OKButton = New-Object System.Windows.Forms.Button
        $OKButton.Location = New-Object System.Drawing.Point(330, 190)
        $OKButton.Size = New-Object System.Drawing.Size(75, 30)
        $OKButton.Text = "OK"
        $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $attentionForm.AcceptButton = $OKButton
        $attentionForm.Controls.Add($OKButton)
        
        # �L�����Z���{�^���̐ݒ�
        $CancelButton = New-Object System.Windows.Forms.Button
        $CancelButton.Location = New-Object System.Drawing.Point(430, 190)
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