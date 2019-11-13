# 関数

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


# publicチームを取得
$TeamList = Get-Team| Where-Object {$_.Visibility -eq "Public"}

# 公開範囲をPrivateに変更する
Set-Team -GroupId **** -Visibility Private

# 終了処理
