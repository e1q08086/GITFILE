$MyCMD = ".\logOutput.ps1 "
Write-Host $MyCMD
Invoke-Expression $MyCMD

$MyCMD = "LogCreate " + $MyInvocation.MyCommand.Name + "_開始"
Write-Host $MyCMD
$Result = Invoke-Expression $MyCMD
$Result


#フォーマットチェック
try{
    cscript C:\script\wday.vbs >> $Result
    Write-Host "vbs " $LASTEXITCODE
    if($LASTEXITCODE -ne 0){
        throw
    }
}catch{
    Write-Output "異常" | Out-File $Result -Append -Encoding Default
#    Write-Host "異常" >> $Result
    Write-Error("エラー"+$_.Exception) >> $Result
}finally{
    $now = Get-Date
    Write-Host "処理終了時間: " $now

}


