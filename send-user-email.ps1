param(
    [Alias("f")]
    [Parameter(Mandatory=$false)]
    [string]$UserParamsFile = "",

    [Alias("t")]
    [Parameter(Mandatory=$false)]
    [string]$ToAddress = "",

    [Alias("s")]
    [Parameter(Mandatory=$false)]
    [string]$FromAddress = "",

    [Alias("b")]
    [Parameter(Mandatory=$false)]
    [string]$Subject = "【AWS】新規IAMユーザー作成のお知らせ",

    [Parameter(Mandatory=$false)]
    [switch]$Help = $false
)

$HelpText = @"
概要: 
  IAMユーザー情報を含むJSONファイルを読み込み、Outlookでメール下書きを作成するスクリプト

使用方法:
  .\send-user-email.ps1 [-UserParamsFile <JSONファイル>] [-ToAddress <宛先メールアドレス>] 
                         [-FromAddress <送信元メールアドレス>] [-Subject <件名>] [-Help]

パラメータ:
  -f, -UserParamsFile  オプション: IAMユーザー情報を含むJSONファイル（create-iam-user.ps1で作成されたもの）
  -t, -ToAddress       オプション: 宛先メールアドレス（JSONファイル内にEmailAddressが含まれている場合はそれを優先）
  -s, -FromAddress     オプション: 送信元メールアドレス
  -b, -Subject         オプション: メールの件名（デフォルト: "【AWS】新規IAMユーザー作成のお知らせ"）
  -h, -Help            オプション: このヘルプ情報を表示します

例:
  .\send-user-email.ps1 -f "AddAwsUser_new-user_to_prd-scm-user-group.json"
  .\send-user-email.ps1 -f "AddAwsUser_new-user_to_prd-scm-user-group.json" -t "user@example.com"
  .\send-user-email.ps1 -h
"@

function Main {
    # JSONファイルの読み込み
    if (-not $UserParamsFile) {
        # ファイルが指定されていない場合は検索してリストから選択する
        $jsonFiles = Get-ChildItem -Filter "AddAwsUser_*.json" -File
        
        if ($jsonFiles.Count -eq 0) {
            Write-Host "エラー: ユーザー情報を含むJSONファイルが見つかりません。" -ForegroundColor Red
            Write-Host "create-iam-user.ps1を実行して、先にユーザーを作成してください。" -ForegroundColor Yellow
            exit 1
        }
        
        Write-Host "利用可能なユーザー情報ファイル:" -ForegroundColor Cyan
        for ($i = 0; $i -lt $jsonFiles.Count; $i++) {
            Write-Host "[$i] $($jsonFiles[$i].Name)" -ForegroundColor White
        }
        
        $selection = Read-Host "使用するファイル番号を入力してください [0-$($jsonFiles.Count - 1)]"
        if ($selection -ge 0 -and $selection -lt $jsonFiles.Count) {
            $UserParamsFile = $jsonFiles[$selection].Name
        }
        else {
            Write-Host "エラー: 無効な選択です。" -ForegroundColor Red
            exit 1
        }
    }
    
    # JSONファイルの存在チェック
    if (-not (Test-Path $UserParamsFile)) {
        Write-Host "エラー: 指定されたファイル '$UserParamsFile' が見つかりません。" -ForegroundColor Red
        exit 1
    }
    
    try {
        # JSONファイルを読み込む
        $userParams = Get-Content $UserParamsFile -Raw | ConvertFrom-Json
        
        # デフォルトのメールアドレスを設定
        if (-not $ToAddress -and $userParams.EmailAddress) {
            $ToAddress = $userParams.EmailAddress
        }
        
        # パラメータをHashtableに変換（辞書のようにアクセスできるように）
        $userParamsHash = @{}
        $userParams.PSObject.Properties | ForEach-Object {
            $userParamsHash[$_.Name] = $_.Value
        }
        
        # メール送信処理
        Show-StartMessage -UserParamsFile $UserParamsFile -ToAddress $ToAddress -Subject $Subject
        
        Prepare-OutlookMail -FromAddress $FromAddress -ToAddress $ToAddress -Subject $Subject -UserParams $userParamsHash
        
        Write-Host "Outlookでメールの下書きを作成しました。内容を確認して送信してください。" -ForegroundColor Green
    }
    catch {
        Write-Host "エラー: 処理中にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}

function Show-StartMessage {
    param (
        [string]$UserParamsFile,
        [string]$ToAddress,
        [string]$Subject
    )

    Write-Host "===========================================================" -ForegroundColor Green
    Write-Host "IAMユーザー情報のメール準備プロセスを開始しています" -ForegroundColor Green
    Write-Host "===========================================================" -ForegroundColor Green
    Write-Host "設定情報:" -ForegroundColor Green
    Write-Host " - ユーザー情報ファイル: $UserParamsFile" -ForegroundColor White
    Write-Host " - メール宛先:         $ToAddress" -ForegroundColor White
    Write-Host " - メール件名:         $Subject" -ForegroundColor White
    Write-Host "===========================================================" -ForegroundColor Green
}

function Build-MailBody {
    param (
        [hashtable]$UserParams
    )
    
    $mailBody = @"
お疲れ様です。

AWSアカウントに新しいIAMユーザーを作成しましたのでお知らせします。

■ ユーザー情報
ユーザー名: $($UserParams.Username)
アクセスキーID: $($UserParams.AccessKeyId)
シークレットアクセスキー: $($UserParams.SecretAccessKey)
作成日時: $($UserParams.Created)

"@

    # ログインプロファイルがある場合は追加
    if ($UserParams.ContainsKey("InitialPassword")) {
        $mailBody += @"
■ コンソールログイン情報
初期パスワード: $($UserParams.InitialPassword)
※ 初回ログイン時にパスワードの変更が必要です

コンソールURL: https://console.aws.amazon.com/

"@
    }

    $mailBody += @"
■ 所属グループ情報
グループ名: $($UserParams.GroupName)
環境: $($UserParams.Environment)
権限グレード: $($UserParams.Grade)

■ 注意事項
・アクセスキーとシークレットキーは安全に保管してください。
・これらの認証情報を共有したり、公開リポジトリにコミットしないでください。
・AWS CLIの設定方法については以下を参照してください。
  https://docs.aws.amazon.com/ja_jp/cli/latest/userguide/cli-configure-files.html

ご不明点がございましたら、お気軽にお問い合わせください。

"@

    return $mailBody
}

function Prepare-OutlookMail {
    param (
        [string]$FromAddress,
        [string]$ToAddress,
        [string]$Subject,
        [hashtable]$UserParams
    )

    try {
        # メール本文の作成
        $mailBody = Build-MailBody -UserParams $UserParams
        
        # COMオブジェクトとしてOutlookを起動
        $outlook = New-Object -ComObject Outlook.Application
        $mail = $outlook.CreateItem(0) # 0 = olMailItem

        # メールの各項目を設定
        if ($ToAddress) {
            $mail.To = $ToAddress
        }
        if ($FromAddress) {
            $mail.SentOnBehalfOfName = $FromAddress
        }
        $mail.Subject = $Subject
        $mail.Body = $mailBody

        # 下書きとして表示（送信はしない）
        $mail.Display()
        
        return $mail
    }
    catch {
        Write-Host "エラー: Outlookでメールを作成できませんでした: $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
}

if ($Help) {
    Write-Host $HelpText -ForegroundColor Cyan
    exit 0
}

Main
