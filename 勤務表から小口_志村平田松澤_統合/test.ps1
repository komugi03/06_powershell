# ---------------�A�Z���u���̓ǂݍ���---------------
# 
Add-Type -Assembly System.Windows.Forms
Add-Type -AssemblyName System.Drawing
# INPUT�̂��߂ɕK�v
[void][System.Reflection.Assembly]::Load("Microsoft.VisualBasic, Version=8.0.0.0, Culture=Neutral, PublicKeyToken=b03f5f7f11d50a3a")

# ���݂̔N�������擾����
$thisYear = (Get-Date).Year
# $thisMonth = 12
$thisMonth = (Get-Date).Month
$today = (Get-Date).Day

# ���ݓ�������쐬����ׂ��Ζ��\�̌����𔻒�
# 24���܂ł͓����������
if ($today -le 24) {
    $targetMonth = $thisMonth - 1
}
else {
    $targetMonth = $thisMonth
}

$yesNo_yearMonthAreCorrect = [System.Windows.Forms.MessageBox]::Show("�쐬����̂� �y $thisYear �N $targetMonth �� �z�̏����ł�낵���ł����H",'�쐬���鏬���̑Ώ۔N��','YesNo','Question')

if($yesNo_yearMonthAreCorrect -eq 'No'){

    # �t�H���g�̎w��
    $Font = New-Object System.Drawing.Font("���C���I",8)

    # �t�H�[���S�̂̐ݒ�
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "�쐬���鏬���̑Ώ۔N��"
    $form.Size = New-Object System.Drawing.Size(300,200)
    $form.StartPosition = "CenterScreen"
    $form.font = $Font

    # ���x����\��
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,10)
    $label.Size = New-Object System.Drawing.Size(270,30)
    $label.Text = "�쐬���鏬���̔N����I�����Ă�������`r`n���O���`�������I���ł��܂���"
    $form.Controls.Add($label)

    # OK�{�^���̐ݒ�
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Point(40,100)
    $OKButton.Size = New-Object System.Drawing.Size(75,30)
    $OKButton.Text = "OK"
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $OKButton
    $form.Controls.Add($OKButton)

    # �L�����Z���{�^���̐ݒ�
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Point(130,100)

    $CancelButton.Size = New-Object System.Drawing.Size(75,30)
    $CancelButton.Text = "Cancel"
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $CancelButton
    $form.Controls.Add($CancelButton)

    # �R���{�{�b�N�X���쐬
    $Combo = New-Object System.Windows.Forms.Combobox
    $Combo.Location = New-Object System.Drawing.Point(50,50)
    $Combo.size = New-Object System.Drawing.Size(150,30)
    # ���X�g�ȊO�̓��͂������Ȃ�
    $Combo.DropDownStyle = "DropDownList"
    $Combo.FlatStyle = "standard"
    # $Combo.font = $Font
    $Combo.BackColor = "#005050"
    $Combo.ForeColor = "white"
    
    # �R���{�{�b�N�X�̑I�����Ɏg���ϐ�������
    # ���N
    $lastYear = $thisYear - 1
    # ���N
    $nextYear = $thisYear + 1
    
    # -----------�R���{�{�b�N�X�ɍ��ڂ�ǉ�-----------
    # �O�̌�
    if($thisMonth -eq '1'){
        # ���N��12��=1���ɍ쐬���Ă���
        $lastMonth = 12
        [void] $Combo.Items.Add("$lastYear �N $lastMonth ��")
    }else{
        # ���N�̐挎
        $lastMonth = $thisMonth - 1
        [void] $Combo.Items.Add("$thisYear �N $lastMonth ��")
    }
    # ���N�̓���
    [void] $Combo.Items.Add("$thisYear �N $thisMonth ��")
    
    # ���̌�
    if($thisMonth -eq '12'){
        # ���N��1��=12���ɍ쐬���Ă���
        $nextMonth = 1
        [void] $Combo.Items.Add("$nextYear �N $nextMonth ��")
    }else{
        # ���N�̗���
        $nextMonth = $thisMonth + 1
        [void] $Combo.Items.Add("$thisYear �N $nextMonth ��")
    }
    
    # �t�H�[���ɃR���{�{�b�N�X��ǉ�
    $form.Controls.Add($Combo)
    $Combo.SelectedIndex = 1
    
    # �t�H�[�����őO�ʂɕ\��
    $form.Topmost = $True
    
    # �t�H�[����\���{�I�����ʂ�ϐ��Ɋi�[
    $result = $form.ShowDialog()

    # �I����AOK�{�^���������ꂽ�ꍇ�A�I�����ڂ�\��
    if ($result -eq "OK"){
        $comboAnswer = $Combo.Text
    }else{
    exit
    }

    Write-Output $comboAnswer

    # ���N��12���Ȃ�
    if($comboAnswer -eq "$lastYear �N $lastMonth ��"){
        echo "���N��12�������Ă�"
    }

    # ���N�̐挎�Ȃ�
    elseif($comboAnswer -eq "$thisYear �N $lastMonth ��"){
        echo "�挎�����Ă�"
    }

    # �����Ȃ�
    elseif($comboAnswer -eq "$thisYear �N $thisMonth ��"){
        echo "�����̂���"
    }

    # ���N��1���Ȃ�
    elseif($comboAnswer -eq "$nextYear �N $nextMonth ��") {
        echo "���N��1���I"
    }

    # �����Ȃ�
    elseif($comboAnswer -eq "$thisYear �N $nextMonth ��") {
        echo "���̌��I"
    }

}