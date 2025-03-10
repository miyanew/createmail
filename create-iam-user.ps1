param(
    [Alias("i")]
    [Parameter(Mandatory=$false)]
    [string]$NewIamUser = "",

    [Alias("p")]
    [Parameter(Mandatory=$false)]
    [string]$AdmProfile = "scm_tmp_role",

    [Alias("e")]
    [Parameter(Mandatory=$false)]
    [ValidateSet("prd", "mgt", "stg1", "stg2", "stg3", "stg4", "stg5", "stg6", "stg7", "stg8", "stg9")]
    [string]$EnvScm = "prd",

    [Alias("g")]
    [Parameter(Mandatory=$false)]
    [ValidateSet("user", "ope", "adm")]
    [string]$Grade = "user",

    [Parameter(Mandatory=$false)]
    [switch]$Help = $false
)

$HelpText = @"
概要: 
  AWSにIAMユーザーを作成し、適切なグループに追加するスクリプト

使用方法:
  .\create-iam-user.ps1 [-NewIamUser <新規ユーザー名>] [-AdmProfile <管理者プロファイル>] 
                        [-EnvScm <環境>] [-Grade <権限グレード>] [-Help]

パラメータ:
  -p, -AdmProfile  オプション: 管理者権限を持つAWSプロファイル名（デフォルト: "scm_tmp_role"）
  -i, -NewIamUser  オプション: 作成する新しいIAMユーザーの名前（指定なしの場合は入力を求められます）
  -e, -EnvScm      オプション: 環境を指定（prd, mgt, stg1-9から選択、デフォルト: "prd"）
  -g, -Grade       オプション: ユーザーの権限グレード（user, ope, admから選択、デフォルト: "user"）
  -h, -Help        オプション: このヘルプ情報を表示します

出力:
  ユーザー情報がJSON形式で "AddAwsUser_{ユーザー名}_to_{グループ名}.json" という名前のファイルに保存されます。
  このJSONファイルは、後でメール送信スクリプトで使用できます。

例:
  .\create-iam-user.ps1 -p "admin-profile" -i "new-user" -e "prd" -g "user"
  .\create-iam-user.ps1 -h
"@

$gradeConfig = @{
    "user" = @{
        GroupName = "$($EnvScm)-scm-user-group"
    }
    "ope" = @{
        GroupName = "$($EnvScm)-scm-operator-group"
    }
    "adm" = @{
        GroupName = "$($EnvScm)-scm-administrator-group"
    }
}

function Main {
    Read-RequiredParameters -NewUserRef ([ref]$NewIamUser) -AdmProfileRef ([ref]$AdmProfile)
    
    $groupName = $gradeConfig[$Grade].GroupName
    
    Show-StartMessage -NewIamUser $NewIamUser -AdmProfile $AdmProfile -EnvScm $EnvScm -Grade $Grade -GroupName $groupName
    
    $userCreation = Create-IamUser -NewIamUser $NewIamUser -AdmProfile $AdmProfile
    Add-UserToGroup -NewIamUser $NewIamUser -GroupName $groupName -AdmProfile $AdmProfile
    $accessKey = Create-AccessKey -NewIamUser $NewIamUser -AdmProfile $AdmProfile

    $userParams = @{
        Created = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        UserArn = $userCreation.User.Arn
        Username = $NewIamUser
        AccessKeyId = $accessKey.AccessKeyId
        SecretAccessKey = $accessKey.SecretAccessKey
        EmailAddress = "$($NewIamUser)@sample.com"  # デフォルトのメールアドレス形式
        GroupName = $groupName
        Environment = $EnvScm
        Grade = $Grade
    }
    
    if ($Grade -ne "user") {
        $initialPassword = Create-LoginProfile -NewIamUser $NewIamUser -AdmProfile $AdmProfile
        $userParams["InitialPassword"] = $initialPassword
    }
    
    Write-Output "ユーザー情報:"
    Write-Output $userParams
    
    Save-UserParams -UserParams $userParams -NewIamUser $NewIamUser -GroupName $groupName
    
    Write-Host "===================================================================" -ForegroundColor Green
    Write-Host "処理が完了しました。ユーザー情報はJSONファイルに保存されました。" -ForegroundColor Green
    Write-Host "メール送信のためにsend-user-email.ps1を使用してください。" -ForegroundColor Green
    Write-Host "===================================================================" -ForegroundColor Green
}

function Read-RequiredParameters {
    param (
        [ref]$NewIamUserRef,
        [ref]$AdmProfileRef
    )
    
    if (-not $NewIamUserRef.Value) {
        $NewIamUserRef.Value = Read-Host "新規IAMユーザー名を入力してください"
    }

    if (-not $AdmProfileRef.Value) {
        $AdmProfileRef.Value = Read-Host "管理者権限を持つAWS CLIプロファイル名を入力してください"
    }
}

function Show-StartMessage {
    param (
        [string]$NewIamUser,
        [string]$AdmProfile,
        [string]$EnvScm,
        [string]$Grade,
        [string]$GroupName
    )

    Write-Host "===========================================================" -ForegroundColor Green
    Write-Host "IAMユーザー作成プロセスを開始しています" -ForegroundColor Green
    Write-Host "===========================================================" -ForegroundColor Green
    Write-Host "設定情報:" -ForegroundColor Green
    Write-Host " - 新規ユーザー名:       $NewIamUser" -ForegroundColor White
    Write-Host " - 管理者プロファイル:   $AdmProfile" -ForegroundColor White
    Write-Host " - 環境:                 $EnvScm" -ForegroundColor White
    Write-Host " - 権限グレード:         $Grade" -ForegroundColor White
    Write-Host " - グループ名:           $GroupName" -ForegroundColor White
    Write-Host "===========================================================" -ForegroundColor Green
    Write-Host "このスクリプトは以下を実行します:" -ForegroundColor Green
    Write-Host " 1. 新規IAMユーザーの作成" -ForegroundColor White
    Write-Host " 2. 指定されたグループへのユーザーの追加" -ForegroundColor White
    Write-Host " 3. アクセスキーの生成" -ForegroundColor White
    if ($Grade -ne "user") {
        Write-Host " 4. ログインプロファイルの作成（ランダムパスワード付き）" -ForegroundColor White
    }
    Write-Host " 5. ユーザー情報のJSONファイルへの保存" -ForegroundColor White
    Write-Host "===========================================================" -ForegroundColor Green
}

function Create-IamUser {
    param (
        [string]$NewIamUser,
        [string]$AdmProfile
    )

    try {
        $userCreation = aws iam create-user --user-name $NewIamUser --profile $AdmProfile --no-verify | ConvertFrom-Json
        Write-Host "ユーザー作成成功: $($userCreation.User.Arn)" -ForegroundColor Green
        return $userCreation
    }
    catch {
        Write-Host "エラー: ユーザー作成に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}

function Add-UserToGroup {
    param (
        [string]$NewIamUser,
        [string]$GroupName,
        [string]$AdmProfile
    )

    try {
        aws iam add-user-to-group --user-name $NewIamUser --group-name $GroupName --profile $AdmProfile --no-verify
        Write-Host "グループ追加成功: ユーザー '$NewIamUser' をグループ '$GroupName' に追加しました" -ForegroundColor Green
    }
    catch {
        Write-Host "エラー: グループへの追加に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Create-AccessKey {
    param (
        [string]$NewIamUser,
        [string]$AdmProfile
    )

    try {
        $accessKey = aws iam create-access-key --user-name $NewIamUser --profile $AdmProfile --no-verify | ConvertFrom-Json
        Write-Host "アクセスキー作成成功: $($accessKey.AccessKey.AccessKeyId)" -ForegroundColor Green
        return $accessKey.AccessKey
    }
    catch {
        Write-Host "エラー: アクセスキー作成に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}

function Create-LoginProfile {
    param (
        [string]$NewIamUser,
        [string]$AdmProfile
    )

    try {
        # ランダムパスワードの生成
        $randomPassword = (aws secretsmanager get-random-password --password-length 16 --exclude-punctuation --require-each-included-type --region ap-northeast-3 --profile $AdmProfile --no-verify | ConvertFrom-Json).RandomPassword
        
        aws iam create-login-profile `
            --user-name $NewIamUser `
            --password $randomPassword `
            --password-reset-required $true `
            --profile $AdmProfile
            
        Write-Host "ログインプロファイル作成成功: 初期パスワードが設定されました" -ForegroundColor Green
        return $randomPassword
    }
    catch {
        Write-Host "エラー: ログインプロファイル作成に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}

function Save-UserParams {
    param (
        [hashtable]$UserParams,
        [string]$NewIamUser,
        [string]$GroupName
    )

    try {
        $fileName = "AddAwsUser_$($NewIamUser)_to_$($GroupName).json"
        $UserParams | ConvertTo-Json -Depth 10 | Out-File -FilePath $fileName -Encoding utf8
        Write-Host "ユーザー情報をJSONファイルに保存しました: $fileName" -ForegroundColor Green
        return $fileName
    }
    catch {
        Write-Host "警告: ユーザー情報の保存に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

if ($Help) {
    Write-Host $HelpText -ForegroundColor Cyan
    exit 0
}

Main