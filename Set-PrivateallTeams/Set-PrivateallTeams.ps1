#############################################################
# 
# Title: Get-Teams.ps1
# Author: Shomi_Nijo
# Version: 1.0
# LastUpdate: 2019.12.18
# Description:
#
##############################################################

# 初期処理
$Error.Clear()
[String]$myPath = Split-Path $MyInvocation.MyCommand.Path -parent
Set-Location($myPath)
$ErrorActionPreference = "Stop"

# ********************************************
# Settings
# ********************************************
# 定数

## 出力ファイル
$outdir = $myPath + "\out"
$datestring = Get-Date -format "yyyyMMddHHmmss"

$TeamListFile = $outdir + "\TeamList_" + $datestring + ".csv"

$ErrGroupListFile = $outdir + "\ErrorTeamList_" + $datestring + ".csv"

# ログファイル
$logdir = $myPath + "\log"
$logFile = $logdir + "\log_getteam_" + $datestring + ".log" 
# ログステータス
$logStatus_Inf = "情報"
$logStatus_War = "警告"
$logStatus_Err = "エラー"

[String]$EncodingStr = "UTF8"

# 変数
$credential = $NULL
$TeamList = $NULL
$catchError = $false
$TeamCount = 0
$ErrorTeamCount = 0
$Team = ""

Function WriteLog($status, $message)
{
    $logmessage = [DateTime]::Now.ToString() + "`t" + $status + "`t" + $message
    $logmessage | Out-File $logFile -Encoding $EncodingStr -append
    
    # コンソールで出力する場合は、下記のコメント化をはずす
    Write-Host $logmessage
}

# ********************************************
# MAIN
# ********************************************

try
{
    # 出力フォルダ作成
    if(![System.IO.Directory]::Exists($outdir))
    {
        [System.IO.Directory]::CreateDirectory($outdir)
    }
    if(![System.IO.Directory]::Exists($logdir))
    {
        [System.IO.Directory]::CreateDirectory($logdir)
    }
    
    # パラメータ
    WriteLog $logStatus_Inf ">>> パラメータ"
    if($IsWithGroupMember -eq $NULL)
    {
        $IsWithGroupMember = $False
    }
    WriteLog $logStatus_Inf ("IsWithGroupMember:" + $IsWithGroupMember)

    WriteLog $logStatus_Inf "スクリプトを開始します。"



    # 接続アカウント/パスワード入力（管理者アカウント）
    #WriteLog $logStatus_Inf "管理者アカウント/パスワードを入力してください。"
    $credential = Get-Credential

    # Teams接続
    try
    {
        #WriteLog $logStatus_Inf "Teamsへの接続を開始します。"
        Import-Module MicrosoftTeams
        Connect-MicrosoftTeams -Credential $credential
    }
    catch
    {
        #WriteLog $logStatus_err ("Teamsに接続できませんでした。" + $error[0].Exception.Message)
        break
    }

    WriteLog $logStatus_Inf "publicチームを取得します。"

    # publicチームの取得
    $TeamList = Get-Team | Where-Object {$_.Visibility -eq "Public"}
    if(($TeamList -eq $NULL) -or ($TeamList.count -lt 0))
    {
        WriteLog $logStatus_War "チームがありません。処理を終了します。"
        exit
    }
    else
    {
        WriteLog $logStatus_Inf ("publicチーム数：" + $TeamList.count)
        #Add-Content $Teamlistfile -Value "GroupId,DisplayName,Visibility,MemberCount" -Encoding $EncodingStr
    }

    foreach($Team in $TeamList)
    {
        try
        {
            
            Set-Team -GroupId $Team.GroupId -Visibility "private"
            WriteLog $logStatus_Inf ("チーム名：" + $Team.DisplayName + "をprivateチームに変更しました")
            $TeamCount++

            # チーム内の出力/エラーメンバー数の初期化
            #$MemInCount = 0
            #$ErrorMemInTeamCount = 0
        }
        catch
        {
            WriteLog $logStatus_War ("チームを出力できません。チーム名：" + $Team.DisplayName)
            WriteLog $logStatus_Err ($error[0].Exception.Message + "`n" + $error[0].ScriptStackTrace)
            $group.DisplayName | Out-File $errGroupListFile -Encoding $EncodingStr -append
            $ErrorTeamCount++
            $catchError = $true
            continue
        }
    }

}
catch
{
    WriteLog $logStatus_Err ($error[0].Exception.Message + "`n" + $error[0].ScriptStackTrace)
}
finally
{
    WriteLog $logStatus_Inf "スクリプトを終了します。"
}
