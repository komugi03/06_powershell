#
#
#



# --------------- �A�Z���u���̓ǂݍ��� ---------------
Add-Type -Assembly System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---------------- ���Ӊ�ʃX�N���v�g�̌Ăяo�� -------------------
. (Join-Path -path $PWD -childpath "resources\scripts\���Ӊ��.ps1")

# ���Ӊ��.ps1�̊֐������s����
$attentionForm = attentionThisTool

# �t�H�[���̉���
$attentionResult = $attentionForm.ShowDialog()

if ($attentionResult -eq "Cancel") {
    exit 
}
# ---------------- ���Ӊ�ʃX�N���v�g�I�� -------------------

# ���݂̔N�������擾����
$targetYear = (Get-Date).Year
$thisMonth = (Get-Date).Month
$today = (Get-Date).Day

# ���ݓ�������쐬����ׂ��Ζ��\�̌����𔻒�
# 24���܂ł͓����������
if ($today -le 24) {
    # �O�̌��������쐬�̑Ώی��Ƃ���
    $targetMonth = (Get-date).AddMonths(-1).month
}
else {
    # �����������쐬�̑Ώی��Ƃ���
    $targetMonth = $thisMonth
}

# (���ݓ��ɂ���ĕς��̂ŁAget-date -Format Y �ɂ͂��Ă��Ȃ�)
$yesNo_yearMonthAreCorrect = [System.Windows.Forms.MessageBox]::Show("�y $targetYear �N $targetMonth �� �z�̋Ζ��\�����Ƃɏ����ݒ�����܂����H`r`n`r`n�u�������v�ő��̌���I���ł��܂�", '�쐬���鏬���̑Ώ۔N��', 'YesNo', 'Question')

# ��$yesNo_yearMonthAreCorrect -eq 'No'���[�v�J�n��
if ($yesNo_yearMonthAreCorrect -eq 'No') {

    # ---------------- �Ζ��\�̌��I���X�N���v�g���Ăяo�� -------------------
    . (Join-Path -Path $PWD -ChildPath ".\resources\scripts\�Ζ��\�̌��I��.ps1")

    # �Ζ��\�̌��I��.ps1�̊֐������s
    $choicedMonth = choiceMonth

    if ($choicedMonth -eq "OK") {
        # ���[�U�[�̉񓚂�"�N"�ŋ�؂�
        $Combo.Text -match "(?<year>.+?)�N(?<month>.+?)��" | out-null

        # ���[�U�[�w��̔N�������쐬�̑Ώ۔N�Ƃ��ď㏑����
        $targetYear = $Matches.year

        # ���[�U�[�w��̌��������쐬�̑Ώی��Ƃ��ď㏑������
        $targetMonth = $Matches.month

    }
    else {   
        # �������I������
        exit
    }
}
# ---------------- �Ζ��\�̌��I���X�N���v�g�I�� -------------------

# ----------- ���΂炭���҂����������X�N���v�g�Ăяo�� -----------
. (Join-Path -path $PWD -ChildPath "\resources\scripts\���΂炭���҂���������.ps1")

# ���΂炭���҂���������.ps1�̊֐������s
$waitCatForm = pleaseWait "resources/pictures/���҂����������L.png"

# ���΂炭���҂����������t�H�[���̉���
$waitCatForm.show()
start-sleep -second 2
$waitCatForm.close()

# ---------------- �Ζ��\�`�F�b�N���� -------------------
# �|�b�v�A�b�v���쐬
$popup = new-object -comobject wscript.shell

# �t�@�C�����̋Ζ��\_�̂��Ƃ̕\�L
$targetMonth00 = "{0:00}" -f [int]$targetMonth
$fileNameMonth = ("$targetYear" + "$targetMonth00")
# �Ζ��\�t�@�C�����擾
$kinmuhyou = Get-ChildItem -Recurse -File | ? Name -Match ("[0-9]{3}_�Ζ��\_" + $fileNameMonth + "_.+") 
# �Y���Ζ��\�t�@�C���̌��m�F
if ($kinmuhyou.Count -lt 1) {
    
    # ���΂炭���҂�����������ʂ����
    $waitForm.Close()

    # �|�b�v�A�b�v��\��
    $popup.popup("$targetMonth ���̋Ζ��\�t�@�C�������݂��܂���", 0, "��蒼���Ă�������", 48) | Out-Null
    exit
}
elseif ($kinmuhyou.Count -gt 1) {
    
    # ���΂炭���҂�����������ʂ����
    $waitForm.Close()

    # �|�b�v�A�b�v��\��
    $popup.popup("$targetMonth ���̋Ζ��\�t�@�C�����������܂�`r`n1�ɂ��Ă�������", 0, "��蒼���Ă�������", 48) | Out-Null
    exit
}
# ---------------- �Ζ��\�`�F�b�N�����I�� -------------------

# ----------------------Excel���N������--------------------------------
try {
    # �N������Excel�v���Z�X���擾
    $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
}
catch {
    # Excel�v���Z�X���N�����ĂȂ���ΐV���ɋN������
    $excel = New-Object -ComObject "Excel.Application" 
}

# Excel�����b�Z�[�W�_�C�A���O��\�����Ȃ��悤�ɂ���
$excel.DisplayAlerts = $false
$excel.visible = $false

# �Ζ��\�̃t���p�X
$kinmuhyouFullPath = $kinmuhyou.FullName 

# �Ζ��\�u�b�N���J��
$kinmuhyouBook = $excel.workbooks.open($kinmuhyouFullPath)
$kinmuhyouSheet = $kinmuhyouBook.worksheets.item([String]$targetMonth + '��')

# ----------------------Excel���N�������I��--------------------------------



