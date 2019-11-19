##########################################################################
# ログ出力
# 引数１：　
#
#
##########################################################################
function LogCreate(
            $LogString
            ){

    # ログの出力先
    $LogPath = "C:\script\log"

    # ログファイル名
    $LogName = "ExecutLog"

    $Now = Get-Date

    # Log 出力文字列に時刻を付加(YYYY/MM/DD HH:MM:SS.MMM $LogString)
    $Log = $Now.ToString("yyyy/MM/dd HH:mm:ss.fff") + " "
    $Log += $LogString

    # ログファイル名が設定されていなかったらデフォルトのログファイル名をつける
    if( $LogName -eq $null ){
        $LogName = "LOG"
    }

    # ログファイル名(XXXX_YYYY-MM-DD.log)
    $LogFile = $LogName + "_" +$Now.ToString("yyyy-MM-dd") + ".log"

    # ログフォルダーがなかったら作成
    if( -not (Test-Path $LogPath) ) {
        New-Item $LogPath -Type Directory
    }

    # ログファイル名
    $LogFileName = Join-Path $LogPath $LogFile

    # ログ出力
    Write-Output $Log | Out-File -FilePath $LogFileName -Encoding Default -append

    # echo させるために出力したログを戻す
    Return $LogFileName
}

#LogCreate $MyInvocation.MyCommand.Name