$inputLine = Read-Host "�L�����Ă�������"

$inputLine -replace "`n","" | Out-Null

$testFront = "
# �e�X�gPowershell
Write-Host start
Write-Host ("

$testEnd = ")
Write-Host end"

$testFront | Set-Content -encoding Default out.ps1
$inputLine | Add-Content out.ps1 -encoding Default
$testEnd | Add-Content out.ps1 -encoding Default

