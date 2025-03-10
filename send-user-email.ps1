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
    [string]$Subject = "�yAWS�z�V�KIAM���[�U�[�쐬�̂��m�点",

    [Parameter(Mandatory=$false)]
    [switch]$Help = $false
)

$HelpText = @"
�T�v: 
  IAM���[�U�[�����܂�JSON�t�@�C����ǂݍ��݁AOutlook�Ń��[�����������쐬����X�N���v�g

�g�p���@:
  .\send-user-email.ps1 [-UserParamsFile <JSON�t�@�C��>] [-ToAddress <���惁�[���A�h���X>] 
                         [-FromAddress <���M�����[���A�h���X>] [-Subject <����>] [-Help]

�p�����[�^:
  -f, -UserParamsFile  �I�v�V����: IAM���[�U�[�����܂�JSON�t�@�C���icreate-iam-user.ps1�ō쐬���ꂽ���́j
  -t, -ToAddress       �I�v�V����: ���惁�[���A�h���X�iJSON�t�@�C������EmailAddress���܂܂�Ă���ꍇ�͂����D��j
  -s, -FromAddress     �I�v�V����: ���M�����[���A�h���X
  -b, -Subject         �I�v�V����: ���[���̌����i�f�t�H���g: "�yAWS�z�V�KIAM���[�U�[�쐬�̂��m�点"�j
  -h, -Help            �I�v�V����: ���̃w���v����\�����܂�

��:
  .\send-user-email.ps1 -f "AddAwsUser_new-user_to_prd-scm-user-group.json"
  .\send-user-email.ps1 -f "AddAwsUser_new-user_to_prd-scm-user-group.json" -t "user@example.com"
  .\send-user-email.ps1 -h
"@

function Main {
    # JSON�t�@�C���̓ǂݍ���
    if (-not $UserParamsFile) {
        # �t�@�C�����w�肳��Ă��Ȃ��ꍇ�͌������ă��X�g����I������
        $jsonFiles = Get-ChildItem -Filter "AddAwsUser_*.json" -File
        
        if ($jsonFiles.Count -eq 0) {
            Write-Host "�G���[: ���[�U�[�����܂�JSON�t�@�C����������܂���B" -ForegroundColor Red
            Write-Host "create-iam-user.ps1�����s���āA��Ƀ��[�U�[���쐬���Ă��������B" -ForegroundColor Yellow
            exit 1
        }
        
        Write-Host "���p�\�ȃ��[�U�[���t�@�C��:" -ForegroundColor Cyan
        for ($i = 0; $i -lt $jsonFiles.Count; $i++) {
            Write-Host "[$i] $($jsonFiles[$i].Name)" -ForegroundColor White
        }
        
        $selection = Read-Host "�g�p����t�@�C���ԍ�����͂��Ă������� [0-$($jsonFiles.Count - 1)]"
        if ($selection -ge 0 -and $selection -lt $jsonFiles.Count) {
            $UserParamsFile = $jsonFiles[$selection].Name
        }
        else {
            Write-Host "�G���[: �����ȑI���ł��B" -ForegroundColor Red
            exit 1
        }
    }
    
    # JSON�t�@�C���̑��݃`�F�b�N
    if (-not (Test-Path $UserParamsFile)) {
        Write-Host "�G���[: �w�肳�ꂽ�t�@�C�� '$UserParamsFile' ��������܂���B" -ForegroundColor Red
        exit 1
    }
    
    try {
        # JSON�t�@�C����ǂݍ���
        $userParams = Get-Content $UserParamsFile -Raw | ConvertFrom-Json
        
        # �f�t�H���g�̃��[���A�h���X��ݒ�
        if (-not $ToAddress -and $userParams.EmailAddress) {
            $ToAddress = $userParams.EmailAddress
        }
        
        # �p�����[�^��Hashtable�ɕϊ��i�����̂悤�ɃA�N�Z�X�ł���悤�Ɂj
        $userParamsHash = @{}
        $userParams.PSObject.Properties | ForEach-Object {
            $userParamsHash[$_.Name] = $_.Value
        }
        
        # ���[�����M����
        Show-StartMessage -UserParamsFile $UserParamsFile -ToAddress $ToAddress -Subject $Subject
        
        Prepare-OutlookMail -FromAddress $FromAddress -ToAddress $ToAddress -Subject $Subject -UserParams $userParamsHash
        
        Write-Host "Outlook�Ń��[���̉��������쐬���܂����B���e���m�F���đ��M���Ă��������B" -ForegroundColor Green
    }
    catch {
        Write-Host "�G���[: �������ɃG���[���������܂���: $($_.Exception.Message)" -ForegroundColor Red
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
    Write-Host "IAM���[�U�[���̃��[�������v���Z�X���J�n���Ă��܂�" -ForegroundColor Green
    Write-Host "===========================================================" -ForegroundColor Green
    Write-Host "�ݒ���:" -ForegroundColor Green
    Write-Host " - ���[�U�[���t�@�C��: $UserParamsFile" -ForegroundColor White
    Write-Host " - ���[������:         $ToAddress" -ForegroundColor White
    Write-Host " - ���[������:         $Subject" -ForegroundColor White
    Write-Host "===========================================================" -ForegroundColor Green
}

function Build-MailBody {
    param (
        [hashtable]$UserParams
    )
    
    $mailBody = @"
�����l�ł��B

AWS�A�J�E���g�ɐV����IAM���[�U�[���쐬���܂����̂ł��m�点���܂��B

�� ���[�U�[���
���[�U�[��: $($UserParams.Username)
�A�N�Z�X�L�[ID: $($UserParams.AccessKeyId)
�V�[�N���b�g�A�N�Z�X�L�[: $($UserParams.SecretAccessKey)
�쐬����: $($UserParams.Created)

"@

    # ���O�C���v���t�@�C��������ꍇ�͒ǉ�
    if ($UserParams.ContainsKey("InitialPassword")) {
        $mailBody += @"
�� �R���\�[�����O�C�����
�����p�X���[�h: $($UserParams.InitialPassword)
�� ���񃍃O�C�����Ƀp�X���[�h�̕ύX���K�v�ł�

�R���\�[��URL: https://console.aws.amazon.com/

"@
    }

    $mailBody += @"
�� �����O���[�v���
�O���[�v��: $($UserParams.GroupName)
��: $($UserParams.Environment)
�����O���[�h: $($UserParams.Grade)

�� ���ӎ���
�E�A�N�Z�X�L�[�ƃV�[�N���b�g�L�[�͈��S�ɕۊǂ��Ă��������B
�E�����̔F�؏������L������A���J���|�W�g���ɃR�~�b�g���Ȃ��ł��������B
�EAWS CLI�̐ݒ���@�ɂ��Ă͈ȉ����Q�Ƃ��Ă��������B
  https://docs.aws.amazon.com/ja_jp/cli/latest/userguide/cli-configure-files.html

���s���_���������܂�����A���C�y�ɂ��₢���킹���������B

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
        # ���[���{���̍쐬
        $mailBody = Build-MailBody -UserParams $UserParams
        
        # COM�I�u�W�F�N�g�Ƃ���Outlook���N��
        $outlook = New-Object -ComObject Outlook.Application
        $mail = $outlook.CreateItem(0) # 0 = olMailItem

        # ���[���̊e���ڂ�ݒ�
        if ($ToAddress) {
            $mail.To = $ToAddress
        }
        if ($FromAddress) {
            $mail.SentOnBehalfOfName = $FromAddress
        }
        $mail.Subject = $Subject
        $mail.Body = $mailBody

        # �������Ƃ��ĕ\���i���M�͂��Ȃ��j
        $mail.Display()
        
        return $mail
    }
    catch {
        Write-Host "�G���[: Outlook�Ń��[�����쐬�ł��܂���ł���: $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
}

if ($Help) {
    Write-Host $HelpText -ForegroundColor Cyan
    exit 0
}

Main
