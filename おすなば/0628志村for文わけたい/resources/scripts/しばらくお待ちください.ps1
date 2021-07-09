#
# 猫がしばらくお待ちくださいフォームを提供する
#

write-host "しばらくお待ちください.ps1が呼び出されました"

function pleaseWait($catPath) {
    .{
# フォントの指定
$font = New-Object System.Drawing.Font("メイリオ", 8)

# フォームの設定
$waitForm = New-Object System.Windows.Forms.Form
$waitForm.Text = "初期設定"
$waitForm.Size = New-Object System.Drawing.Size(265, 170)
$waitForm.StartPosition = "CenterScreen"
$waitForm.font = $font
# フォームを最前面に表示
$waitForm.Topmost = $True

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(70, 30)
$label.Size = New-Object System.Drawing.Size(270, 30)
$label.Text = "準備中です`r`nしばらくお待ちください"
$label.font = $font
$waitForm.Controls.Add($label)

###### 画像のフルパス変更してほしいです☆ #########
#PictureBox
$pic = New-Object System.Windows.Forms.PictureBox
$pic.Size = New-Object System.Drawing.Size(50, 50)
# おじぎ猫画像のフルパス
$catFullPath = join-path -path $PWD.path -childpath $catPath # "resources/pictures/お待ちください猫.png"
$pic.Image = [System.Drawing.Image]::FromFile($catFullPath)
$pic.Location = New-Object System.Drawing.Point(20,20) 
$pic.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
$waitForm.Controls.Add($pic)


return

} | Out-Null


# フォームをリターンする
return $waitForm

}