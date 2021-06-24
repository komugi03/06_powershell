#
# 勤務表をもとに小口を作成するPowerShell
# 
# 勤務表のファイル名:<社員番号>_勤務表_m月_<名前>.xlsx
#

# 今年
$thisYear = (Get-Date).Year

# 対話方式で進めていく

# 何月の小口を作成するかどうかの決定
$month = $input | Read-Host "何月の小口を作成しますか？[半角数字]"

if (-not([int]::TryParse($month,[ref]$null))) {
    Write-host "半角数字以外の値が入っているため、処理を終了します。"
    break
}

# 勤務表を受け取る
$kinmuhyou = gci -Recurse |? name -cmatch ("^[0-9]{2,3}_勤務表_")

# --------勤務表のファイル名から<名前>と<社員番号>取得--------
$fileName = $kinmuhyou.name.Split("_")
# <名前>
$name_Extension = $fileName[3].Split(".")
$lastName = $name_Extension[0]
# <社員番号>
$employeeNumber = $fileName[0]

$koguchiFileName = $employeeNumber+"_小口交通費・出張旅費精算明細書_"+$month+"月_"+$lastName+".xlsx"
#---------------------------------------------------------

if ($kinmuhyou -cmatch $month+"月") {
    $YorN = Read-Host `r`n$lastName"さんの"$thisYear"年"$month"月の小口を作成します。`r`nよろしいですか？[Y/N]"
}else {
    Write-host `r`n$month"月の勤務表が見つかりません。"
    break
}

if ($YorN -match "n") {
    $thisYear = Read-Host "`r`n何年の小口を作成しますか？[半角数字]"
    if (-not([int]::TryParse($month,[ref]$null))) {
        Write-host "半角数字以外の値が入っているため、処理を終了します。"
        break
    }
}
Write-host "`r`n*************************************************`r`n"`t$lastName"さんの"$thisYear"年"$month"月の小口を作成中...`r`n`tしばらくお待ちください`r`n*************************************************"


# ------------------------初期設定------------------------

#$moyoriStation = Read-Host "自宅の最寄り駅を入力してください。"

#$transportation = Read-Host "交通機関を入力してください。"

#$WorkPlace1_Fee = Read-Host "自宅-[$WorkPlace1]間の交通費(往復分)を入力してください。[半角数字]"

# --------------------------------------------------------

# Excelを開く
$excel = New-Object -ComObject Excel.Application

# Excelを見えるようにする
$excel.visible = $true

# 勤務表を開く
$kinmuhyouBook = $excel.workbooks.Open($kinmuhyou.fullname)
$kinmuhyouSheet = $kinmuhyouBook.sheets($month+'月')


# 小口を開く
$koguchi = gci -Recurse |? name -cmatch '小口交通費・出張旅費精算明細書'
$koguchiBook = $excel.workbooks.Open($koguchi.fullname)
$koguchiSheet = $koguchiBook.sheets('小口交通費')

# 勤務表と小口が必要。勤務表から必要な情報を小口に渡す。渡す値の加工が必要
# 「備考」から勤務地を取得

# 行カウント
$countRow = 11

for ($Row = 14; $Row -le 44; $Row++) {
        
    $workPlace = $kinmuhyouBook.sheets($month+'月').cells.item($Row,27).formula



    if (($workPlace -ne "") -And ($workPlace -ne "在宅")){

        # 適用セル
        $tekiyo = 6

        # 区間セル
        $kukan = 18

        # 交通機関セル
        $koutsukikan = 26

        # 金額セル
        $kingaku = 30

        $day = $kinmuhyouSheet.cells.item($Row,3).text


        switch -Regex ($workPlace) {
            "^お台場$" {
                # 月
                $koguchiSheet.cells.item($countRow,2).formula = $month
                # 日付
                $koguchiSheet.cells.item($countRow,4).formula = $day
                # 適用（行先、要件） 
                $koguchiSheet.cells.item($countRow,$tekiyo).formula = "自宅←→お台場"
                # 区間  ="$moyoriStation←→お台場"
                $koguchiSheet.cells.item($countRow,$kukan).formula = "町内会館前←→東京テレポート"
                # 交通機関 ="$transportation(多分複数になる)"←ユーザの入力からやる時どうしよう？
                $koguchiSheet.cells.item($countRow,$koutsukikan).formula = "神奈中バス`r`n小田急線`r`nJR埼京線-りんかい線" 
                # 金額 ="$WorkPlace1_Fee"
                $koguchiSheet.cells.item($countRow,$kingaku).formula = "2640"
            }
            "^田町$"{
                # 日付
                $koguchiSheet.cells.item($countRow,2).formula = $month
                $koguchiSheet.cells.item($countRow,4).formula = $day
                # 適用（行先、要件） 
                $koguchiSheet.cells.item($countRow,$tekiyo).formula = "自宅←→田町"
                # 区間  ="$moyoriStation←→田町"
                $koguchiSheet.cells.item($countRow,$kukan).formula = "町内会館前←→田町"
                # 交通機関 ="$transportation(多分複数になる)"←ユーザの入力からやる時どうしよう？
                $koguchiSheet.cells.item($countRow,$koutsukikan).formula = "神奈中バス`r`n小田急線`r`nJR山手線" 
                # 金額 ="$WorkPlace2_Fee"
                $koguchiSheet.cells.item($countRow,$kingaku).formula = "2030"
            }
            # どこにも該当しなかった場合
            Default {
                Write-Host "`r`n====================== 注意 ======================`r`n"     $day"日の勤務地：[ "$workPlace" ] は登録されていません`r`n登録してやり直してください`r`n`r`n又は現在作成された小口では"$day"日の行は空白であるため`r`n"$day"日のみご自身でご記入ください`r`n=================================================="
                $koguchiSheet.cells.item($countRow,2).formula = $month
                $koguchiSheet.cells.item($countRow,4).formula = $day
            }
        }
        # 行カウントのカウント
        $countRow = $countRow + 3
    }
}

# 記入が完了したら最後に「以下余白」を記入
$koguchiSheet.cells.item($countRow,$tekiyo).formula = "以下余白"

# 年記入
$koguchiSheet.cells.item(66,4).formula = [string]$thisYear
# 月記入
$koguchiSheet.cells.item(66,8).formula = $month
# 月末日記入
$lastDayOftheMonth = [DateTime]::DaysInMonth($thisYear,$month)
$koguchiSheet.cells.item(66,11).formula = [string]$lastDayOftheMonth


# 名前を付けて保存
# 小口のファイル名：<社員番号>_小口交通費・出張旅費精算明細書_m月_<名前>.xlsx
$koguchiBook.SaveAs("C:\Users\bvs20005\Documents\03_基盤講習\PowerShell-Lesson\勤務表から小口作成ツール\作成済み小口\"+$koguchiFileName)

# Excelを終了
$Excel.Quit()

# 変数の解放
$excel = $null
$koguchiBook = $null
$kinmuhyouBook = $null


