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
�T�v: 
  AWS��IAM���[�U�[���쐬���A�K�؂ȃO���[�v�ɒǉ�����X�N���v�g

�g�p���@:
  .\create-iam-user.ps1 [-NewIamUser <�V�K���[�U�[��>] [-AdmProfile <�Ǘ��҃v���t�@�C��>] 
                        [-EnvScm <��>] [-Grade <�����O���[�h>] [-Help]

�p�����[�^:
  -p, -AdmProfile  �I�v�V����: �Ǘ��Ҍ���������AWS�v���t�@�C�����i�f�t�H���g: "scm_tmp_role"�j
  -i, -NewIamUser  �I�v�V����: �쐬����V����IAM���[�U�[�̖��O�i�w��Ȃ��̏ꍇ�͓��͂����߂��܂��j
  -e, -EnvScm      �I�v�V����: �����w��iprd, mgt, stg1-9����I���A�f�t�H���g: "prd"�j
  -g, -Grade       �I�v�V����: ���[�U�[�̌����O���[�h�iuser, ope, adm����I���A�f�t�H���g: "user"�j
  -h, -Help        �I�v�V����: ���̃w���v����\�����܂�

�o��:
  ���[�U�[���JSON�`���� "AddAwsUser_{���[�U�[��}_to_{�O���[�v��}.json" �Ƃ������O�̃t�@�C���ɕۑ�����܂��B
  ����JSON�t�@�C���́A��Ń��[�����M�X�N���v�g�Ŏg�p�ł��܂��B

��:
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
        EmailAddress = "$($NewIamUser)@sample.com"  # �f�t�H���g�̃��[���A�h���X�`��
        GroupName = $groupName
        Environment = $EnvScm
        Grade = $Grade
    }
    
    if ($Grade -ne "user") {
        $initialPassword = Create-LoginProfile -NewIamUser $NewIamUser -AdmProfile $AdmProfile
        $userParams["InitialPassword"] = $initialPassword
    }
    
    Write-Output "���[�U�[���:"
    Write-Output $userParams
    
    Save-UserParams -UserParams $userParams -NewIamUser $NewIamUser -GroupName $groupName
    
    Write-Host "===================================================================" -ForegroundColor Green
    Write-Host "�������������܂����B���[�U�[����JSON�t�@�C���ɕۑ�����܂����B" -ForegroundColor Green
    Write-Host "���[�����M�̂��߂�send-user-email.ps1���g�p���Ă��������B" -ForegroundColor Green
    Write-Host "===================================================================" -ForegroundColor Green
}

function Read-RequiredParameters {
    param (
        [ref]$NewIamUserRef,
        [ref]$AdmProfileRef
    )
    
    if (-not $NewIamUserRef.Value) {
        $NewIamUserRef.Value = Read-Host "�V�KIAM���[�U�[������͂��Ă�������"
    }

    if (-not $AdmProfileRef.Value) {
        $AdmProfileRef.Value = Read-Host "�Ǘ��Ҍ���������AWS CLI�v���t�@�C��������͂��Ă�������"
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
    Write-Host "IAM���[�U�[�쐬�v���Z�X���J�n���Ă��܂�" -ForegroundColor Green
    Write-Host "===========================================================" -ForegroundColor Green
    Write-Host "�ݒ���:" -ForegroundColor Green
    Write-Host " - �V�K���[�U�[��:       $NewIamUser" -ForegroundColor White
    Write-Host " - �Ǘ��҃v���t�@�C��:   $AdmProfile" -ForegroundColor White
    Write-Host " - ��:                 $EnvScm" -ForegroundColor White
    Write-Host " - �����O���[�h:         $Grade" -ForegroundColor White
    Write-Host " - �O���[�v��:           $GroupName" -ForegroundColor White
    Write-Host "===========================================================" -ForegroundColor Green
    Write-Host "���̃X�N���v�g�͈ȉ������s���܂�:" -ForegroundColor Green
    Write-Host " 1. �V�KIAM���[�U�[�̍쐬" -ForegroundColor White
    Write-Host " 2. �w�肳�ꂽ�O���[�v�ւ̃��[�U�[�̒ǉ�" -ForegroundColor White
    Write-Host " 3. �A�N�Z�X�L�[�̐���" -ForegroundColor White
    if ($Grade -ne "user") {
        Write-Host " 4. ���O�C���v���t�@�C���̍쐬�i�����_���p�X���[�h�t���j" -ForegroundColor White
    }
    Write-Host " 5. ���[�U�[����JSON�t�@�C���ւ̕ۑ�" -ForegroundColor White
    Write-Host "===========================================================" -ForegroundColor Green
}

function Create-IamUser {
    param (
        [string]$NewIamUser,
        [string]$AdmProfile
    )

    try {
        $userCreation = aws iam create-user --user-name $NewIamUser --profile $AdmProfile --no-verify | ConvertFrom-Json
        Write-Host "���[�U�[�쐬����: $($userCreation.User.Arn)" -ForegroundColor Green
        return $userCreation
    }
    catch {
        Write-Host "�G���[: ���[�U�[�쐬�Ɏ��s���܂���: $($_.Exception.Message)" -ForegroundColor Red
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
        Write-Host "�O���[�v�ǉ�����: ���[�U�[ '$NewIamUser' ���O���[�v '$GroupName' �ɒǉ����܂���" -ForegroundColor Green
    }
    catch {
        Write-Host "�G���[: �O���[�v�ւ̒ǉ��Ɏ��s���܂���: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Create-AccessKey {
    param (
        [string]$NewIamUser,
        [string]$AdmProfile
    )

    try {
        $accessKey = aws iam create-access-key --user-name $NewIamUser --profile $AdmProfile --no-verify | ConvertFrom-Json
        Write-Host "�A�N�Z�X�L�[�쐬����: $($accessKey.AccessKey.AccessKeyId)" -ForegroundColor Green
        return $accessKey.AccessKey
    }
    catch {
        Write-Host "�G���[: �A�N�Z�X�L�[�쐬�Ɏ��s���܂���: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}

function Create-LoginProfile {
    param (
        [string]$NewIamUser,
        [string]$AdmProfile
    )

    try {
        # �����_���p�X���[�h�̐���
        $randomPassword = (aws secretsmanager get-random-password --password-length 16 --exclude-punctuation --require-each-included-type --region ap-northeast-3 --profile $AdmProfile --no-verify | ConvertFrom-Json).RandomPassword
        
        aws iam create-login-profile `
            --user-name $NewIamUser `
            --password $randomPassword `
            --password-reset-required $true `
            --profile $AdmProfile
            
        Write-Host "���O�C���v���t�@�C���쐬����: �����p�X���[�h���ݒ肳��܂���" -ForegroundColor Green
        return $randomPassword
    }
    catch {
        Write-Host "�G���[: ���O�C���v���t�@�C���쐬�Ɏ��s���܂���: $($_.Exception.Message)" -ForegroundColor Red
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
        Write-Host "���[�U�[����JSON�t�@�C���ɕۑ����܂���: $fileName" -ForegroundColor Green
        return $fileName
    }
    catch {
        Write-Host "�x��: ���[�U�[���̕ۑ��Ɏ��s���܂���: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

if ($Help) {
    Write-Host $HelpText -ForegroundColor Cyan
    exit 0
}

Main