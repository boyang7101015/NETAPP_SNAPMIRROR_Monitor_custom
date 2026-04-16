<#
.SYNOPSIS
Interactive generator for Config\settings.json

.DESCRIPTION
1. Prompt for ONTAP, Mail, History, and Log settings in the terminal
2. Encrypt ONTAP and SMTP passwords using Windows DPAPI for the current user
3. Save the generated JSON configuration to Config\settings.json

.NOTES
Encrypted passwords created by ConvertFrom-SecureString without a custom key
can usually only be decrypted by the same Windows user on the same machine.
#>

[CmdletBinding()]
param(
    [string]$OutputPath
)

Set-StrictMode -Version 2
$ErrorActionPreference = 'Stop'

function Get-ProjectRoot {
    $scriptRoot = Split-Path -Parent $MyInvocation.PSCommandPath
    if (-not [string]::IsNullOrWhiteSpace($scriptRoot)) {
        return (Split-Path -Parent $scriptRoot)
    }

    return (Get-Location).Path
}

function Get-DefaultOutputPath {
    return (Join-Path (Get-ProjectRoot) 'Config\settings.json')
}

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Get-DefaultOutputPath
}

function Read-RequiredValue {
    param(
        [Parameter(Mandatory = $true)][string]$Prompt,
        [Parameter(Mandatory = $false)][string]$DefaultValue
    )

    while ($true) {
        $displayPrompt = if ([string]::IsNullOrWhiteSpace($DefaultValue)) {
            $Prompt
        }
        else {
            '{0} [{1}]' -f $Prompt, $DefaultValue
        }

        $value = Read-Host -Prompt $displayPrompt
        if ([string]::IsNullOrWhiteSpace($value)) {
            $value = $DefaultValue
        }

        if (-not [string]::IsNullOrWhiteSpace($value)) {
            return $value.Trim()
        }

        Write-Host 'This value is required. Please try again.' -ForegroundColor Yellow
    }
}

function Read-OptionalValue {
    param(
        [Parameter(Mandatory = $true)][string]$Prompt,
        [Parameter(Mandatory = $false)][string]$DefaultValue
    )

    $displayPrompt = if ([string]::IsNullOrWhiteSpace($DefaultValue)) {
        $Prompt
    }
    else {
        '{0} [{1}]' -f $Prompt, $DefaultValue
    }

    $value = Read-Host -Prompt $displayPrompt
    if ([string]::IsNullOrWhiteSpace($value)) {
        return $DefaultValue
    }

    return $value.Trim()
}

function Read-BoolValue {
    param(
        [Parameter(Mandatory = $true)][string]$Prompt,
        [Parameter(Mandatory = $true)][bool]$DefaultValue
    )

    $defaultLabel = if ($DefaultValue) { 'Y' } else { 'N' }

    while ($true) {
        $value = Read-Host -Prompt ('{0} [Y/N] (default: {1})' -f $Prompt, $defaultLabel)
        if ([string]::IsNullOrWhiteSpace($value)) {
            return $DefaultValue
        }

        switch ($value.Trim().ToUpperInvariant()) {
            'Y' { return $true }
            'YES' { return $true }
            'N' { return $false }
            'NO' { return $false }
            default {
                Write-Host 'Please enter Y or N.' -ForegroundColor Yellow
            }
        }
    }
}

function Read-IntValue {
    param(
        [Parameter(Mandatory = $true)][string]$Prompt,
        [Parameter(Mandatory = $true)][int]$DefaultValue
    )

    while ($true) {
        $value = Read-Host -Prompt ('{0} [{1}]' -f $Prompt, $DefaultValue)
        if ([string]::IsNullOrWhiteSpace($value)) {
            return $DefaultValue
        }

        $parsed = 0
        if ([int]::TryParse($value.Trim(), [ref]$parsed)) {
            return $parsed
        }

        Write-Host 'Please enter a valid integer.' -ForegroundColor Yellow
    }
}

function Read-ArrayValue {
    param(
        [Parameter(Mandatory = $true)][string]$Prompt,
        [Parameter(Mandatory = $false)][string[]]$DefaultValue = @()
    )

    $defaultText = if (@($DefaultValue).Count -gt 0) {
        $DefaultValue -join ', '
    }
    else {
        ''
    }

    $raw = Read-OptionalValue -Prompt ('{0} (comma-separated)' -f $Prompt) -DefaultValue $defaultText
    if ([string]::IsNullOrWhiteSpace($raw)) {
        return @()
    }

    return @(
        $raw -split ',' |
        ForEach-Object { $_.Trim() } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )
}

function Read-SecureValue {
    param(
        [Parameter(Mandatory = $true)][string]$Prompt
    )

    while ($true) {
        $secureValue = Read-Host -Prompt $Prompt -AsSecureString
        if ($secureValue.Length -gt 0) {
            return $secureValue
        }

        Write-Host 'Password cannot be empty. Please try again.' -ForegroundColor Yellow
    }
}

function Convert-SecureValueToEncryptedString {
    param(
        [Parameter(Mandatory = $true)][System.Security.SecureString]$SecureValue
    )

    return ConvertFrom-SecureString -SecureString $SecureValue
}

function New-ConfigObject {
    $projectRoot = Get-ProjectRoot
    $defaultHistoryFolder = Join-Path $projectRoot 'History'
    $defaultLogFolder = Join-Path $projectRoot 'Logs'

    Write-Host ''
    Write-Host '=== ONTAP Settings ===' -ForegroundColor Cyan
    $clusterUrl = Read-RequiredValue -Prompt 'ONTAP Cluster URL' -DefaultValue 'https://cluster2.demo.netapp.com'
    $apiPath = Read-RequiredValue -Prompt 'ONTAP API Path' -DefaultValue '/api/private/cli/snapmirror?fields=source-path,source-vserver,source-volume,destination-path,destination-vserver,destination-volume,last-transfer-type,last-transfer-error,last-transfer-size,last-transfer-duration,last-transfer-end-timestamp'
    $ontapUsername = Read-RequiredValue -Prompt 'ONTAP Username' -DefaultValue 'admin'
    $ontapPasswordEncrypted = Convert-SecureValueToEncryptedString -SecureValue (Read-SecureValue -Prompt 'ONTAP Password')
    $ignoreCertificate = Read-BoolValue -Prompt 'Ignore ONTAP TLS certificate errors' -DefaultValue $true

    Write-Host ''
    Write-Host '=== History Settings ===' -ForegroundColor Cyan
    $historyFolder = Read-RequiredValue -Prompt 'History folder' -DefaultValue $defaultHistoryFolder

    while ($true) {
        $dedupMode = Read-RequiredValue -Prompt 'History DedupMode (None / ByCollectTime / ByTransferResult)' -DefaultValue 'ByTransferResult'
        if ($dedupMode -in @('None', 'ByCollectTime', 'ByTransferResult')) {
            break
        }

        Write-Host 'DedupMode must be None, ByCollectTime, or ByTransferResult.' -ForegroundColor Yellow
    }

    Write-Host ''
    Write-Host '=== Mail Settings ===' -ForegroundColor Cyan
    $smtpServer = Read-RequiredValue -Prompt 'SMTP server' -DefaultValue 'smtp.demo.com'
    $smtpPort = Read-IntValue -Prompt 'SMTP port' -DefaultValue 587
    $useSsl = Read-BoolValue -Prompt 'Use SSL/TLS for SMTP' -DefaultValue $true
    $useAuthentication = Read-BoolValue -Prompt 'Use SMTP authentication' -DefaultValue $true
    $sender = Read-RequiredValue -Prompt 'Mail sender address' -DefaultValue 'netapp-report@demo.com'

    $senderPasswordEncrypted = ''
    if ($useAuthentication) {
        $senderPasswordEncrypted = Convert-SecureValueToEncryptedString -SecureValue (Read-SecureValue -Prompt 'SMTP sender password')
    }

    $to = Read-ArrayValue -Prompt 'Mail To recipients' -DefaultValue @('user1@demo.com', 'user2@demo.com')
    if (@($to).Count -lt 1) {
        throw 'Mail.To must contain at least one recipient.'
    }

    $cc = Read-ArrayValue -Prompt 'Mail Cc recipients' -DefaultValue @()
    $bcc = Read-ArrayValue -Prompt 'Mail Bcc recipients' -DefaultValue @()

    Write-Host ''
    Write-Host '=== Log Settings ===' -ForegroundColor Cyan
    $logFolder = Read-RequiredValue -Prompt 'Log folder' -DefaultValue $defaultLogFolder

    return [ordered]@{
        Ontap = [ordered]@{
            ClusterUrl         = $clusterUrl
            ApiPath            = $apiPath
            Username           = $ontapUsername
            PasswordEncrypted  = $ontapPasswordEncrypted
            IgnoreCertificate  = $ignoreCertificate
        }
        History = [ordered]@{
            Folder    = $historyFolder
            DedupMode = $dedupMode
        }
        Mail = [ordered]@{
            SmtpServer              = $smtpServer
            Port                    = $smtpPort
            UseSsl                  = $useSsl
            UseAuthentication       = $useAuthentication
            Sender                  = $sender
            SenderPasswordEncrypted = $senderPasswordEncrypted
            To                      = @($to)
            Cc                      = @($cc)
            Bcc                     = @($bcc)
        }
        Log = [ordered]@{
            Folder = $logFolder
        }
    }
}

function Show-ConfigSummary {
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)][string]$Path
    )

    Write-Host ''
    Write-Host '=== Summary ===' -ForegroundColor Cyan
    Write-Host ('Output Path           : {0}' -f $Path)
    Write-Host ('ONTAP Cluster URL     : {0}' -f $Config.Ontap.ClusterUrl)
    Write-Host ('ONTAP API Path        : {0}' -f $Config.Ontap.ApiPath)
    Write-Host ('ONTAP Username        : {0}' -f $Config.Ontap.Username)
    Write-Host ('ONTAP Password        : <encrypted>')
    Write-Host ('Ignore Certificate    : {0}' -f $Config.Ontap.IgnoreCertificate)
    Write-Host ('History Folder        : {0}' -f $Config.History.Folder)
    Write-Host ('History DedupMode     : {0}' -f $Config.History.DedupMode)
    Write-Host ('SMTP Server           : {0}' -f $Config.Mail.SmtpServer)
    Write-Host ('SMTP Port             : {0}' -f $Config.Mail.Port)
    Write-Host ('SMTP UseSsl           : {0}' -f $Config.Mail.UseSsl)
    Write-Host ('SMTP Auth             : {0}' -f $Config.Mail.UseAuthentication)
    Write-Host ('Mail Sender           : {0}' -f $Config.Mail.Sender)
    Write-Host ('Sender Password       : {0}' -f $(if ($Config.Mail.UseAuthentication) { '<encrypted>' } else { '<not used>' }))
    Write-Host ('Mail To               : {0}' -f (@($Config.Mail.To) -join ', '))
    Write-Host ('Mail Cc               : {0}' -f $(if (@($Config.Mail.Cc).Count -gt 0) { @($Config.Mail.Cc) -join ', ' } else { '<empty>' }))
    Write-Host ('Mail Bcc              : {0}' -f $(if (@($Config.Mail.Bcc).Count -gt 0) { @($Config.Mail.Bcc) -join ', ' } else { '<empty>' }))
    Write-Host ('Log Folder            : {0}' -f $Config.Log.Folder)
}

function Save-ConfigFile {
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)][string]$Path
    )

    $folder = Split-Path -Path $Path -Parent
    if (-not (Test-Path -LiteralPath $folder)) {
        New-Item -Path $folder -ItemType Directory -Force | Out-Null
    }

    if (Test-Path -LiteralPath $Path) {
        $overwrite = Read-BoolValue -Prompt ('File already exists. Overwrite {0}' -f $Path) -DefaultValue $false
        if (-not $overwrite) {
            throw 'Operation cancelled. Existing settings.json was not overwritten.'
        }
    }

    $json = $Config | ConvertTo-Json -Depth 5
    [System.IO.File]::WriteAllText($Path, $json, (New-Object System.Text.UTF8Encoding($false)))
}

function Main {
    Write-Host ('Generating settings file: {0}' -f $OutputPath) -ForegroundColor Green
    Write-Host 'Passwords will be encrypted for the current Windows user on this machine.' -ForegroundColor Green

    $config = New-ConfigObject
    Show-ConfigSummary -Config $config -Path $OutputPath

    $confirm = Read-BoolValue -Prompt 'Write settings.json now' -DefaultValue $true
    if (-not $confirm) {
        throw 'Operation cancelled before writing settings.json.'
    }

    Save-ConfigFile -Config $config -Path $OutputPath
    Write-Host ''
    Write-Host ('settings.json saved successfully: {0}' -f $OutputPath) -ForegroundColor Green
    Write-Host 'Use the same Windows account to run Send-SnapMirrorReport.ps1 so encrypted passwords can be decrypted.' -ForegroundColor Green
}

Main
