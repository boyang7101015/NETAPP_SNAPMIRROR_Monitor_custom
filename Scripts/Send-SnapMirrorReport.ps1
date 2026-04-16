<#
.SYNOPSIS
NetApp ONTAP SnapMirror daily reporting script (single-file edition)

.DESCRIPTION
1. Load ONTAP, History, Mail, and Log settings from settings.json
2. Query ONTAP REST API for the latest transfer result of each SnapMirror relationship
3. Build the current snapshot using destination_volume as the report subject
4. Append results into a monthly CSV history file
5. Compare each destination_volume with the previous record from the current month or previous month
6. Generate an HTML email detail report and send it
7. Write execution details to a log file
#>

[CmdletBinding()]
param(
    [string]$ConfigPath
)

Set-StrictMode -Version 2
$ErrorActionPreference = 'Stop'

$script:LogFilePath = $null
$script:LastCollectedRecords = @()
$script:LastSkippedDuplicateRecords = @()

function Get-ProjectRoot {
    try {
        $scriptRoot = Split-Path -Parent $MyInvocation.PSCommandPath
        if (-not [string]::IsNullOrWhiteSpace($scriptRoot)) {
            return (Split-Path -Parent $scriptRoot)
        }
    }
    catch {
    }

    return (Get-Location).Path
}

$script:DefaultLogFolder = Join-Path (Get-ProjectRoot) 'Logs'

function Get-DefaultConfigPath {
    try {
        $scriptRoot = Split-Path -Parent $MyInvocation.PSCommandPath
        if (-not [string]::IsNullOrWhiteSpace($scriptRoot)) {
            $candidate = Join-Path (Split-Path -Parent $scriptRoot) 'Config\settings.json'
            return $candidate
        }
    }
    catch {
    }

    return (Join-Path (Get-Location).Path 'Config\settings.json')
}

if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
    $ConfigPath = Get-DefaultConfigPath
}

function Resolve-ConfiguredPath {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$BasePath
    )

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return $Path
    }

    return [System.IO.Path]::GetFullPath((Join-Path $BasePath $Path))
}

function Write-Log {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR')][string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = '[{0}] [{1}] {2}' -f $timestamp, $Level, $Message
    Write-Host $line

    try {
        $path = $script:LogFilePath
        if ([string]::IsNullOrWhiteSpace($path)) {
            if (-not (Test-Path -LiteralPath $script:DefaultLogFolder)) {
                New-Item -Path $script:DefaultLogFolder -ItemType Directory -Force | Out-Null
            }
            $path = Join-Path $script:DefaultLogFolder ('SnapMirrorReport_{0}.log' -f (Get-Date -Format 'yyyy-MM-dd'))
        }

        $folder = Split-Path -Path $path -Parent
        if (-not (Test-Path -LiteralPath $folder)) {
            New-Item -Path $folder -ItemType Directory -Force | Out-Null
        }

        Add-Content -LiteralPath $path -Value $line -Encoding UTF8
    }
    catch {
        Write-Host ('[{0}] [ERROR] Failed to write log file: {1}' -f $timestamp, $_.Exception.Message)
    }
}

function Show-TerminalSummary {
    param(
        [Parameter(Mandatory = $false)][array]$Records,
        [Parameter(Mandatory = $false)][array]$SkippedRecords
    )

    $safeRecords = @($Records)
    $safeSkippedRecords = @($SkippedRecords)

    if ($safeRecords.Count -eq 0) {
        Write-Host ''
        Write-Host '==== SnapMirror Current Result ===='
        Write-Host 'No records collected.'
        Write-Host '==================================='
        return
    }

    $displayRows = Get-SummaryDisplayRows -Records $safeRecords

    Write-Host ''
    Write-Host '==== SnapMirror Current Result ===='
    $displayRows | Sort-Object DestinationVolume | Format-Table -AutoSize | Out-Host
    Write-Host '==================================='

    if ($safeSkippedRecords.Count -gt 0) {
        $skippedRows = Get-SummaryDisplayRows -Records $safeSkippedRecords
        Write-Host ''
        Write-Host '==== Skipped Duplicate Records ===='
        $skippedRows | Sort-Object DestinationVolume | Format-Table -AutoSize | Out-Host
        Write-Host '==================================='
    }
}

function Get-SummaryDisplayRows {
    param(
        [Parameter(Mandatory = $false)][array]$Records
    )

    $safeRecords = @($Records)

    foreach ($record in $safeRecords) {
        [PSCustomObject]@{
            DestinationVolume = $record.DestinationVolume
            SourceVolume      = $record.SourceVolume
            TransferType      = $record.LastTransferType
            TransferSize      = $record.LastTransferSizeDisplay
            Duration          = $record.LastTransferDurationDisplay
            EndTime           = if ([string]::IsNullOrWhiteSpace([string]$record.LastTransferEndTimestamp)) { 'N/A' } else { Format-DateTimeDisplay -DateTimeValue $record.LastTransferEndTimestamp }
            Status            = $record.Status
            ErrorMessage      = if ([string]::IsNullOrWhiteSpace([string]$record.LastTransferError)) { '' } else { [string]$record.LastTransferError }
        }
    }
}

function Save-SummaryReport {
    param(
        [Parameter(Mandatory = $false)][array]$Records,
        [Parameter(Mandatory = $false)][array]$SkippedRecords,
        [Parameter(Mandatory = $true)][datetime]$CollectTime
    )

    try {
        $safeRecords = @($Records)
        $safeSkippedRecords = @($SkippedRecords)
        $targetFolder = if (-not [string]::IsNullOrWhiteSpace($script:LogFilePath)) {
            Split-Path -Path $script:LogFilePath -Parent
        }
        else {
            $script:DefaultLogFolder
        }

        if (-not (Test-Path -LiteralPath $targetFolder)) {
            New-Item -Path $targetFolder -ItemType Directory -Force | Out-Null
        }

        $summaryPath = Join-Path $targetFolder ('SnapMirrorSummary_{0}.txt' -f $CollectTime.ToString('yyyy-MM-dd_HHmmss'))

        $lines = New-Object System.Collections.Generic.List[string]
        $lines.Add('==== SnapMirror Current Result ====')
        $lines.Add(('Collect Time: {0}' -f (Format-DateTimeDisplay -DateTimeValue $CollectTime)))

        if ($safeRecords.Count -eq 0) {
            $lines.Add('No records collected.')
        }
        else {
            $displayRows = Get-SummaryDisplayRows -Records $safeRecords
            $tableText = ($displayRows | Sort-Object DestinationVolume | Format-Table -AutoSize | Out-String)
            foreach ($line in ($tableText -split "(`r`n|`n|`r)")) {
                if ($null -ne $line) {
                    $lines.Add($line.TrimEnd())
                }
            }
        }

        $lines.Add('===================================')

        if ($safeSkippedRecords.Count -gt 0) {
            $lines.Add('')
            $lines.Add('==== Skipped Duplicate Records ====')
            $skippedTableText = ((Get-SummaryDisplayRows -Records $safeSkippedRecords) | Sort-Object DestinationVolume | Format-Table -AutoSize | Out-String)
            foreach ($line in ($skippedTableText -split "(`r`n|`n|`r)")) {
                if ($null -ne $line) {
                    $lines.Add($line.TrimEnd())
                }
            }
            $lines.Add('===================================')
        }

        [System.IO.File]::WriteAllLines($summaryPath, $lines, (New-Object System.Text.UTF8Encoding($true)))
        Write-Log -Message ('Summary report saved: {0}' -f $summaryPath)
    }
    catch {
        Write-Log -Message ('Failed to save summary report: {0}' -f $_.Exception.Message) -Level 'ERROR'
    }
}

function Initialize-LogFile {
    param(
        [Parameter(Mandatory = $false)]$Config
    )

    $logFolder = $script:DefaultLogFolder

    try {
        if ($Config -and $Config.Log -and -not [string]::IsNullOrWhiteSpace([string]$Config.Log.Folder)) {
            $logFolder = [string]$Config.Log.Folder
        }

        if (-not (Test-Path -LiteralPath $logFolder)) {
            New-Item -Path $logFolder -ItemType Directory -Force | Out-Null
        }

        $script:LogFilePath = Join-Path $logFolder ('SnapMirrorReport_{0}.log' -f (Get-Date -Format 'yyyy-MM-dd'))
    }
    catch {
        $script:LogFilePath = Join-Path $script:DefaultLogFolder ('SnapMirrorReport_{0}.log' -f (Get-Date -Format 'yyyy-MM-dd'))
        Write-Log -Message ('Failed to initialize log file path. Falling back to default path. Error: {0}' -f $_.Exception.Message) -Level 'WARN'
    }
}

function Read-ConfigFile {
    param(
        [Parameter(Mandatory = $true)][string]$Path
    )

    try {
        if (-not (Test-Path -LiteralPath $Path)) {
            throw "Configuration file does not exist: $Path"
        }

        $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($raw)) {
            throw "Configuration file is empty: $Path"
        }

        $config = $raw | ConvertFrom-Json
        $configFolder = Split-Path -Path $Path -Parent

        if ($config.History -and -not [string]::IsNullOrWhiteSpace([string]$config.History.Folder)) {
            $config.History.Folder = Resolve-ConfiguredPath -Path ([string]$config.History.Folder) -BasePath $configFolder
        }

        if ($config.Log -and -not [string]::IsNullOrWhiteSpace([string]$config.Log.Folder)) {
            $config.Log.Folder = Resolve-ConfiguredPath -Path ([string]$config.Log.Folder) -BasePath $configFolder
        }

        Write-Log -Message ('Configuration file loaded successfully: {0}' -f $Path)
        return $config
    }
    catch {
        Write-Log -Message ('Failed to load configuration file: {0}' -f $_.Exception.Message) -Level 'ERROR'
        throw
    }
}

function Validate-Config {
    param(
        [Parameter(Mandatory = $true)]$Config
    )

    $errors = New-Object System.Collections.Generic.List[string]

    if (-not $Config.Ontap) {
        $errors.Add('Missing Ontap configuration section.')
    }
    else {
        $ontapHasPlainPassword = -not [string]::IsNullOrWhiteSpace((Get-ConfigValue -Object $Config.Ontap -PropertyName 'Password'))
        $ontapHasEncryptedPassword = -not [string]::IsNullOrWhiteSpace((Get-ConfigValue -Object $Config.Ontap -PropertyName 'PasswordEncrypted'))

        if ([string]::IsNullOrWhiteSpace([string]$Config.Ontap.ClusterUrl)) { $errors.Add('Ontap.ClusterUrl cannot be empty.') }
        if ([string]::IsNullOrWhiteSpace([string]$Config.Ontap.ApiPath)) { $errors.Add('Ontap.ApiPath cannot be empty.') }
        if ([string]::IsNullOrWhiteSpace([string]$Config.Ontap.Username)) { $errors.Add('Ontap.Username cannot be empty.') }
        if (-not $ontapHasPlainPassword -and -not $ontapHasEncryptedPassword) {
            $errors.Add('Ontap.Password or Ontap.PasswordEncrypted cannot be empty.')
        }
    }

    if (-not $Config.History) {
        $errors.Add('Missing History configuration section.')
    }
    elseif ([string]::IsNullOrWhiteSpace([string]$Config.History.Folder)) {
        $errors.Add('History.Folder cannot be empty.')
    }
    elseif ($Config.History.DedupMode) {
        $dedupMode = [string]$Config.History.DedupMode
        if ($dedupMode -notin @('None', 'ByCollectTime', 'ByTransferResult')) {
            $errors.Add('History.DedupMode must be one of: None, ByCollectTime, ByTransferResult.')
        }
    }

    if (-not $Config.Log) {
        $errors.Add('Missing Log configuration section.')
    }
    elseif ([string]::IsNullOrWhiteSpace([string]$Config.Log.Folder)) {
        $errors.Add('Log.Folder cannot be empty.')
    }

    if (-not $Config.Mail) {
        $errors.Add('Missing Mail configuration section.')
    }
    else {
        $mailHasPlainPassword = -not [string]::IsNullOrWhiteSpace((Get-ConfigValue -Object $Config.Mail -PropertyName 'SenderPassword'))
        $mailHasEncryptedPassword = -not [string]::IsNullOrWhiteSpace((Get-ConfigValue -Object $Config.Mail -PropertyName 'SenderPasswordEncrypted'))

        if ([string]::IsNullOrWhiteSpace([string]$Config.Mail.SmtpServer)) { $errors.Add('Mail.SmtpServer cannot be empty.') }
        if (-not $Config.Mail.Port) { $errors.Add('Mail.Port cannot be empty.') }
        if ([string]::IsNullOrWhiteSpace([string]$Config.Mail.Sender)) { $errors.Add('Mail.Sender cannot be empty.') }
        if (-not $Config.Mail.To -or @($Config.Mail.To).Count -lt 1) { $errors.Add('Mail.To must contain at least one recipient.') }

        if ($Config.Mail.UseAuthentication -eq $true) {
            if ([string]::IsNullOrWhiteSpace([string]$Config.Mail.Sender)) { $errors.Add('Mail.Sender cannot be empty when Mail.UseAuthentication is true.') }
            if (-not $mailHasPlainPassword -and -not $mailHasEncryptedPassword) {
                $errors.Add('Mail.SenderPassword or Mail.SenderPasswordEncrypted cannot be empty when Mail.UseAuthentication is true.')
            }
        }
    }

    if ($errors.Count -gt 0) {
        foreach ($err in $errors) {
            Write-Log -Message $err -Level 'ERROR'
        }
        throw ('Configuration validation failed with {0} error(s).' -f $errors.Count)
    }

    Write-Log -Message 'Configuration validation completed successfully.'
}

function Get-ConfigValue {
    param(
        [Parameter(Mandatory = $true)]$Object,
        [Parameter(Mandatory = $true)][string]$PropertyName
    )

    if ($null -eq $Object) {
        return $null
    }

    $property = $Object.PSObject.Properties[$PropertyName]
    if ($null -eq $property) {
        return $null
    }

    return $property.Value
}

function Convert-SecureStringToPlainText {
    param(
        [Parameter(Mandatory = $true)][System.Security.SecureString]$SecureString
    )

    $bstr = [IntPtr]::Zero

    try {
        $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
        return [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    }
    finally {
        if ($bstr -ne [IntPtr]::Zero) {
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
        }
    }
}

function Get-ConfigPassword {
    param(
        [Parameter(Mandatory = $true)]$Section,
        [Parameter(Mandatory = $true)][string]$PlainPropertyName,
        [Parameter(Mandatory = $true)][string]$EncryptedPropertyName,
        [Parameter(Mandatory = $true)][string]$SectionName
    )

    try {
        $plainValue = [string](Get-ConfigValue -Object $Section -PropertyName $PlainPropertyName)
        if (-not [string]::IsNullOrWhiteSpace($plainValue)) {
            return $plainValue
        }

        $encryptedValue = [string](Get-ConfigValue -Object $Section -PropertyName $EncryptedPropertyName)
        if ([string]::IsNullOrWhiteSpace($encryptedValue)) {
            throw ('{0}.{1} and {0}.{2} are both empty.' -f $SectionName, $PlainPropertyName, $EncryptedPropertyName)
        }

        $secureValue = ConvertTo-SecureString -String $encryptedValue
        return Convert-SecureStringToPlainText -SecureString $secureValue
    }
    catch {
        Write-Log -Message ('Failed to resolve password for {0}: {1}' -f $SectionName, $_.Exception.Message) -Level 'ERROR'
        throw
    }
}

function Ensure-Folder {
    param(
        [Parameter(Mandatory = $true)][string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
        Write-Log -Message ('Created folder: {0}' -f $Path)
    }
}

function Build-RestApiUrl {
    param(
        [Parameter(Mandatory = $true)][string]$ClusterUrl,
        [Parameter(Mandatory = $true)][string]$ApiPath
    )

    try {
        if ([string]::IsNullOrWhiteSpace($ClusterUrl)) { throw 'ClusterUrl is missing.' }
        if ([string]::IsNullOrWhiteSpace($ApiPath)) { throw 'ApiPath is missing.' }

        $base = $ClusterUrl.TrimEnd('/')
        $path = $ApiPath.Trim()
        if (-not $path.StartsWith('/')) {
            $path = '/' + $path
        }

        $url = $base + $path
        Write-Log -Message ('Built API URL: {0}' -f $url)
        return $url
    }
    catch {
        Write-Log -Message ('Failed to build API URL: {0}' -f $_.Exception.Message) -Level 'ERROR'
        throw
    }
}

function Get-BasicAuthHeader {
    param(
        [Parameter(Mandatory = $true)][string]$Username,
        [Parameter(Mandatory = $true)][string]$Password
    )

    try {
        $plainText = '{0}:{1}' -f $Username, $Password
        $bytes = [System.Text.Encoding]::ASCII.GetBytes($plainText)
        $encoded = [System.Convert]::ToBase64String($bytes)

        return @{
            Authorization = "Basic $encoded"
            Accept        = 'application/json'
        }
    }
    catch {
        Write-Log -Message ('Failed to build Basic Authentication header: {0}' -f $_.Exception.Message) -Level 'ERROR'
        throw
    }
}

function Set-IgnoreCertificatePolicy {
    param(
        [Parameter(Mandatory = $true)][bool]$IgnoreCertificate
    )

    try {
        $protocols = 0
        $protocolNames = New-Object System.Collections.Generic.List[string]

        foreach ($name in @('Tls', 'Tls11', 'Tls12')) {
            if ([Enum]::IsDefined([System.Net.SecurityProtocolType], $name)) {
                $protocols = $protocols -bor [int][System.Net.SecurityProtocolType]::$name
                $protocolNames.Add($name)
            }
        }

        if ($protocols -ne 0) {
            [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]$protocols
        }

        [System.Net.ServicePointManager]::Expect100Continue = $false

        if ($PSVersionTable.PSVersion.Major -lt 6) {
            if (-not ([System.Management.Automation.PSTypeName]'TrustAllCertsPolicy').Type) {
                Add-Type @"
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;

public static class TrustAllCertsPolicy {
    private static readonly RemoteCertificateValidationCallback Callback = IgnoreValidation;

    public static bool IgnoreValidation(
        object sender,
        X509Certificate certificate,
        X509Chain chain,
        SslPolicyErrors sslPolicyErrors) {
        return true;
    }

    public static void Enable() {
        ServicePointManager.ServerCertificateValidationCallback = Callback;
    }

    public static void Disable() {
        ServicePointManager.ServerCertificateValidationCallback = null;
    }
}
"@ -ErrorAction Stop
            }

            if ($IgnoreCertificate) {
                [TrustAllCertsPolicy]::Enable()
            }
            else {
                [TrustAllCertsPolicy]::Disable()
            }
        }

        if ($IgnoreCertificate) {
            Write-Log -Message 'TLS certificate validation bypass is enabled.'
        }

        if ($protocolNames.Count -gt 0) {
            Write-Log -Message ('Enabled security protocols: {0}' -f ($protocolNames -join ', '))
        }
    }
    catch {
        Write-Log -Message ('Failed to configure TLS certificate bypass: {0}' -f $_.Exception.Message) -Level 'ERROR'
        throw
    }
}

function Invoke-OntapRestApi {
    param(
        [Parameter(Mandatory = $true)][string]$Url,
        [Parameter(Mandatory = $true)][hashtable]$Headers,
        [Parameter(Mandatory = $false)][bool]$IgnoreCertificate = $false
    )

    try {
        Set-IgnoreCertificatePolicy -IgnoreCertificate $IgnoreCertificate

        $invokeParams = @{
            Uri         = $Url
            Method      = 'GET'
            Headers     = $Headers
            ErrorAction = 'Stop'
        }

        if ($PSVersionTable.PSVersion.Major -ge 6 -and $IgnoreCertificate) {
            $invokeParams.SkipCertificateCheck = $true
        }

        if ($PSVersionTable.PSVersion.Major -lt 6) {
            $invokeParams.UseBasicParsing = $true
        }

        $webResponse = Invoke-WebRequest @invokeParams
        Write-Log -Message ('API request completed successfully. HTTP Status: {0}' -f $webResponse.StatusCode)

        if ($null -eq $webResponse -or [string]::IsNullOrWhiteSpace([string]$webResponse.Content)) {
            throw 'API response content is empty.'
        }

        return ($webResponse.Content | ConvertFrom-Json)
    }
    catch {
        Write-Log -Message ('API request failed: {0}' -f $_.Exception.Message) -Level 'ERROR'
        if ($_.Exception.InnerException) {
            Write-Log -Message ('API request inner exception: {0}' -f $_.Exception.InnerException.Message) -Level 'ERROR'
        }
        throw
    }
}

function Convert-IsoDurationToTimeSpan {
    param(
        [Parameter(Mandatory = $false)][string]$IsoDuration
    )

    try {
        if ([string]::IsNullOrWhiteSpace($IsoDuration)) {
            return [TimeSpan]::Zero
        }

        return [System.Xml.XmlConvert]::ToTimeSpan($IsoDuration)
    }
    catch {
        Write-Log -Message ('Failed to parse ISO 8601 duration. Raw value: {0}. Error: {1}' -f $IsoDuration, $_.Exception.Message) -Level 'WARN'
        return [TimeSpan]::Zero
    }
}

function Convert-BytesToReadableSize {
    param(
        [Parameter(Mandatory = $false)]$Bytes
    )

    try {
        if ($null -eq $Bytes -or $Bytes -eq '') {
            return 'N/A'
        }

        [double]$size = $Bytes
        $units = @('Bytes', 'KB', 'MB', 'GB', 'TB', 'PB')
        $index = 0

        while ($size -ge 1024 -and $index -lt ($units.Count - 1)) {
            $size = $size / 1024
            $index++
        }

        if ($index -eq 0) {
            return ('{0} {1}' -f [math]::Round($size, 0), $units[$index])
        }

        return ('{0:N2} {1}' -f $size, $units[$index])
    }
    catch {
        Write-Log -Message ('Failed to convert byte size to readable format: {0}' -f $_.Exception.Message) -Level 'WARN'
        return 'N/A'
    }
}

function Format-DurationDisplay {
    param(
        [Parameter(Mandatory = $true)][TimeSpan]$Duration
    )

    if ($null -eq $Duration) {
        return 'N/A'
    }

    $parts = @()
    if ($Duration.Days -gt 0) { $parts += ('{0}d' -f $Duration.Days) }
    if ($Duration.Hours -gt 0) { $parts += ('{0}h' -f $Duration.Hours) }
    if ($Duration.Minutes -gt 0) { $parts += ('{0}m' -f $Duration.Minutes) }
    if ($Duration.Seconds -gt 0 -or $parts.Count -eq 0) { $parts += ('{0}s' -f $Duration.Seconds) }

    return ($parts -join ' ')
}

function Format-DateTimeDisplay {
    param(
        [Parameter(Mandatory = $false)]$DateTimeValue
    )

    try {
        if ($null -eq $DateTimeValue -or [string]::IsNullOrWhiteSpace([string]$DateTimeValue)) {
            return 'N/A'
        }

        if ($DateTimeValue -is [DateTimeOffset]) {
            return $DateTimeValue.ToString('yyyy-MM-dd HH:mm:ss zzz')
        }

        if ($DateTimeValue -is [DateTime]) {
            return $DateTimeValue.ToString('yyyy-MM-dd HH:mm:ss')
        }

        $parsed = [DateTimeOffset]::Parse([string]$DateTimeValue)
        return $parsed.ToString('yyyy-MM-dd HH:mm:ss zzz')
    }
    catch {
        return [string]$DateTimeValue
    }
}

function Get-StatusFromTransferError {
    param(
        [Parameter(Mandatory = $false)][string]$LastTransferError
    )

    if ([string]::IsNullOrWhiteSpace($LastTransferError)) {
        return 'Success'
    }

    return 'Failed'
}

function Build-CurrentRecordObject {
    param(
        [Parameter(Mandatory = $true)]$ApiRecord,
        [Parameter(Mandatory = $true)][datetime]$CollectTime
    )

    try {
        $durationRaw = $null
        $sizeBytes = $null
        $endTimestamp = $null
        $startTimestamp = $null
        $duration = [TimeSpan]::Zero

        if ($ApiRecord.PSObject.Properties.Name -contains 'last_transfer_duration') {
            $durationRaw = [string]$ApiRecord.last_transfer_duration
            $duration = Convert-IsoDurationToTimeSpan -IsoDuration $durationRaw
        }

        if ($ApiRecord.PSObject.Properties.Name -contains 'last_transfer_size') {
            $sizeBytes = $ApiRecord.last_transfer_size
        }

        if ($ApiRecord.PSObject.Properties.Name -contains 'last_transfer_end_timestamp' -and -not [string]::IsNullOrWhiteSpace([string]$ApiRecord.last_transfer_end_timestamp)) {
            $endTimestamp = [DateTimeOffset]::Parse([string]$ApiRecord.last_transfer_end_timestamp)
            $startTimestamp = $endTimestamp.Subtract($duration)
        }

        $lastTransferError = $null
        if ($ApiRecord.PSObject.Properties.Name -contains 'last_transfer_error') {
            $lastTransferError = [string]$ApiRecord.last_transfer_error
        }

        return [PSCustomObject]@{
            CollectTime                 = $CollectTime.ToString('o')
            SourcePath                  = [string]$ApiRecord.source_path
            SourceVserver               = [string]$ApiRecord.source_vserver
            SourceVolume                = [string]$ApiRecord.source_volume
            DestinationPath             = [string]$ApiRecord.destination_path
            DestinationVserver          = [string]$ApiRecord.destination_vserver
            DestinationVolume           = [string]$ApiRecord.destination_volume
            LastTransferType            = [string]$ApiRecord.last_transfer_type
            LastTransferError           = $lastTransferError
            LastTransferSizeBytes       = if ($null -ne $sizeBytes -and $sizeBytes -ne '') { [string]$sizeBytes } else { '' }
            LastTransferSizeDisplay     = Convert-BytesToReadableSize -Bytes $sizeBytes
            LastTransferDurationRaw     = $durationRaw
            LastTransferDurationDisplay = Format-DurationDisplay -Duration $duration
            LastTransferEndTimestamp    = if ($endTimestamp) { $endTimestamp.ToString('o') } else { '' }
            CalculatedStartTimestamp    = if ($startTimestamp) { $startTimestamp.ToString('o') } else { '' }
            Status                      = Get-StatusFromTransferError -LastTransferError $lastTransferError
        }
    }
    catch {
        Write-Log -Message ('Failed to build current record object. DestinationVolume={0}. Error: {1}' -f [string]$ApiRecord.destination_volume, $_.Exception.Message) -Level 'ERROR'
        throw
    }
}

function Get-MonthlyCsvPath {
    param(
        [Parameter(Mandatory = $true)][string]$Folder,
        [Parameter(Mandatory = $true)][datetime]$ReferenceDate
    )

    return (Join-Path $Folder ('SnapMirror_{0}.csv' -f $ReferenceDate.ToString('yyyy-MM')))
}

function Import-HistoryCsv {
    param(
        [Parameter(Mandatory = $true)][string]$CsvPath
    )

    try {
        if (-not (Test-Path -LiteralPath $CsvPath)) {
            Write-Log -Message ('History CSV does not exist, skipping import: {0}' -f $CsvPath) -Level 'WARN'
            return @()
        }

        $rows = Import-Csv -LiteralPath $CsvPath
        if ($null -eq $rows) {
            return @()
        }

        Write-Log -Message ('History CSV imported successfully: {0}. Rows: {1}' -f $CsvPath, @($rows).Count)
        return @($rows)
    }
    catch {
        Write-Log -Message ('Failed to import history CSV: {0}. Error: {1}' -f $CsvPath, $_.Exception.Message) -Level 'ERROR'
        throw
    }
}

function Find-PreviousRecord {
    param(
        [Parameter(Mandatory = $true)]$CurrentRecord,
        [Parameter(Mandatory = $false)][array]$CurrentMonthHistory = @(),
        [Parameter(Mandatory = $false)][array]$PreviousMonthHistory = @()
    )

    try {
        $targetVolume = [string]$CurrentRecord.DestinationVolume
        $targetCollect = [datetime]::Parse([string]$CurrentRecord.CollectTime)

        $candidate = $CurrentMonthHistory |
            Where-Object {
                $_.DestinationVolume -eq $targetVolume -and
                -not [string]::IsNullOrWhiteSpace($_.CollectTime) -and
                ([datetime]::Parse($_.CollectTime) -lt $targetCollect)
            } |
            Sort-Object { [datetime]::Parse($_.CollectTime) } -Descending |
            Select-Object -First 1

        if ($candidate) {
            return $candidate
        }

        return $PreviousMonthHistory |
            Where-Object {
                $_.DestinationVolume -eq $targetVolume -and
                -not [string]::IsNullOrWhiteSpace($_.CollectTime)
            } |
            Sort-Object { [datetime]::Parse($_.CollectTime) } -Descending |
            Select-Object -First 1
    }
    catch {
        Write-Log -Message ('Failed to find previous record. DestinationVolume={0}. Error: {1}' -f $CurrentRecord.DestinationVolume, $_.Exception.Message) -Level 'ERROR'
        return $null
    }
}

function Append-HistoryCsv {
    param(
        [Parameter(Mandatory = $true)][string]$CsvPath,
        [Parameter(Mandatory = $true)][array]$Records
    )

    try {
        $folder = Split-Path -Path $CsvPath -Parent
        Ensure-Folder -Path $folder

        if (-not (Test-Path -LiteralPath $CsvPath)) {
            $Records | Export-Csv -LiteralPath $CsvPath -NoTypeInformation -Encoding UTF8
        }
        else {
            $Records | Export-Csv -LiteralPath $CsvPath -NoTypeInformation -Encoding UTF8 -Append
        }

        Write-Log -Message ('CSV write completed successfully: {0}. Added rows: {1}' -f $CsvPath, @($Records).Count)
    }
    catch {
        Write-Log -Message ('Failed to write CSV: {0}. Error: {1}' -f $CsvPath, $_.Exception.Message) -Level 'ERROR'
        throw
    }
}

function New-ReportAttachmentZip {
    param(
        [Parameter(Mandatory = $true)][datetime]$CollectTime,
        [Parameter(Mandatory = $true)][string[]]$SourcePaths,
        [Parameter(Mandatory = $true)][string]$OutputFolder
    )

    try {
        Ensure-Folder -Path $OutputFolder

        $existingPaths = @(
            $SourcePaths |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and (Test-Path -LiteralPath $_) } |
            Select-Object -Unique
        )

        if ($existingPaths.Count -eq 0) {
            Write-Log -Message 'No CSV or log files were found for zip attachment; email will be sent without attachment.' -Level 'WARN'
            return $null
        }

        $zipPath = Join-Path $OutputFolder ('SnapMirrorReport_{0}.zip' -f $CollectTime.ToString('yyyy-MM-dd_HHmmss'))

        if (Test-Path -LiteralPath $zipPath) {
            Remove-Item -LiteralPath $zipPath -Force
        }

        Compress-Archive -LiteralPath $existingPaths -DestinationPath $zipPath -CompressionLevel Optimal -Force
        Write-Log -Message ('Created email attachment zip: {0}' -f $zipPath)
        return $zipPath
    }
    catch {
        Write-Log -Message ('Failed to create email attachment zip: {0}' -f $_.Exception.Message) -Level 'ERROR'
        throw
    }
}

function Get-HistoryDedupMode {
    param(
        [Parameter(Mandatory = $false)]$Config
    )

    $mode = 'ByTransferResult'

    try {
        if ($Config -and $Config.History -and -not [string]::IsNullOrWhiteSpace([string]$Config.History.DedupMode)) {
            $mode = [string]$Config.History.DedupMode
        }
    }
    catch {
    }

    return $mode
}

function Get-RecordDedupKey {
    param(
        [Parameter(Mandatory = $true)]$Record,
        [Parameter(Mandatory = $true)][string]$Mode
    )

    $destinationVolume = [string]$Record.DestinationVolume

    if ($Mode -eq 'ByCollectTime') {
        return '{0}|{1}' -f $destinationVolume, ([string]$Record.CollectTime)
    }

    if ($Mode -eq 'ByTransferResult') {
        return '{0}|{1}|{2}|{3}|{4}|{5}|{6}' -f `
            $destinationVolume, `
            ([string]$Record.LastTransferEndTimestamp), `
            ([string]$Record.LastTransferType), `
            ([string]$Record.LastTransferSizeBytes), `
            ([string]$Record.LastTransferDurationRaw), `
            ([string]$Record.LastTransferError), `
            ([string]$Record.Status)
    }

    return [guid]::NewGuid().ToString()
}

function Remove-DuplicateRecords {
    param(
        [Parameter(Mandatory = $true)][array]$CurrentRecords,
        [Parameter(Mandatory = $false)][array]$ExistingHistory = @(),
        [Parameter(Mandatory = $true)][string]$Mode
    )

    $safeCurrentRecords = @($CurrentRecords)
    $safeExistingHistory = @($ExistingHistory)

    if ($Mode -eq 'None') {
        return @{
            RecordsToWrite = $safeCurrentRecords
            SkippedCount   = 0
            SkippedRecords = @()
        }
    }

    $existingKeys = @{}

    foreach ($historyRecord in $safeExistingHistory) {
        $key = Get-RecordDedupKey -Record $historyRecord -Mode $Mode
        if (-not $existingKeys.ContainsKey($key)) {
            $existingKeys[$key] = $true
        }
    }

    $recordsToWrite = @()
    $skippedRecords = @()
    $skippedCount = 0

    foreach ($record in $safeCurrentRecords) {
        $key = Get-RecordDedupKey -Record $record -Mode $Mode
        if ($existingKeys.ContainsKey($key)) {
            $skippedCount++
            $skippedRecords += $record
            continue
        }

        $existingKeys[$key] = $true
        $recordsToWrite += $record
    }

    return @{
        RecordsToWrite = $recordsToWrite
        SkippedCount   = $skippedCount
        SkippedRecords = $skippedRecords
    }
}

function Convert-RecordToReportValues {
    param(
        [Parameter(Mandatory = $false)]$Record
    )

    if (-not $Record) {
        return @{
            StartTime    = 'N/A'
            EndTime      = 'N/A'
            Duration     = 'N/A'
            Size         = 'N/A'
            Type         = 'N/A'
            Status       = 'N/A'
            ErrorMessage = 'N/A'
        }
    }

    return @{
        StartTime    = if ([string]::IsNullOrWhiteSpace([string]$Record.CalculatedStartTimestamp)) { 'N/A' } else { Format-DateTimeDisplay -DateTimeValue $Record.CalculatedStartTimestamp }
        EndTime      = if ([string]::IsNullOrWhiteSpace([string]$Record.LastTransferEndTimestamp)) { 'N/A' } else { Format-DateTimeDisplay -DateTimeValue $Record.LastTransferEndTimestamp }
        Duration     = if ([string]::IsNullOrWhiteSpace([string]$Record.LastTransferDurationDisplay)) { 'N/A' } else { [string]$Record.LastTransferDurationDisplay }
        Size         = if ([string]::IsNullOrWhiteSpace([string]$Record.LastTransferSizeDisplay)) { 'N/A' } else { [string]$Record.LastTransferSizeDisplay }
        Type         = if ([string]::IsNullOrWhiteSpace([string]$Record.LastTransferType)) { 'N/A' } else { [string]$Record.LastTransferType }
        Status       = if ([string]::IsNullOrWhiteSpace([string]$Record.Status)) { 'N/A' } else { [string]$Record.Status }
        ErrorMessage = if ([string]::IsNullOrWhiteSpace([string]$Record.LastTransferError)) { 'N/A' } else { [string]$Record.LastTransferError }
    }
}

function Build-ComparisonReportHtml {
    param(
        [Parameter(Mandatory = $true)][array]$CurrentRecords,
        [Parameter(Mandatory = $true)][hashtable]$PreviousRecordMap,
        [Parameter(Mandatory = $true)][datetime]$CollectTime
    )

    try {
        $sb = New-Object System.Text.StringBuilder

        [void]$sb.AppendLine('<html>')
        [void]$sb.AppendLine('<head>')
        [void]$sb.AppendLine("<meta http-equiv='Content-Type' content='text/html; charset=utf-8' />")
        [void]$sb.AppendLine('<style>')
        [void]$sb.AppendLine("body { font-family: Arial, sans-serif; font-size: 14px; color: #222; background-color: #f7f7f7; }")
        [void]$sb.AppendLine('.container { padding: 16px; }')
        [void]$sb.AppendLine('.header { font-size: 18px; font-weight: bold; margin-bottom: 16px; }')
        [void]$sb.AppendLine('.meta { margin-bottom: 20px; color: #555; }')
        [void]$sb.AppendLine('.volume-block { background: #ffffff; border: 1px solid #d9d9d9; margin-bottom: 18px; padding: 12px; }')
        [void]$sb.AppendLine('.volume-title { font-size: 16px; font-weight: bold; margin-bottom: 10px; }')
        [void]$sb.AppendLine('table { border-collapse: collapse; width: 100%; background-color: #fafafa; }')
        [void]$sb.AppendLine('th, td { border: 1px solid #cfcfcf; padding: 8px 10px; text-align: left; vertical-align: top; }')
        [void]$sb.AppendLine('th { background-color: #eeeeee; font-weight: bold; }')
        [void]$sb.AppendLine('.status-success { color: #1f6f43; font-weight: bold; }')
        [void]$sb.AppendLine('.status-failed { color: #b00020; font-weight: bold; background-color: #fdeaea; }')
        [void]$sb.AppendLine('</style>')
        [void]$sb.AppendLine('</head>')
        [void]$sb.AppendLine('<body>')
        [void]$sb.AppendLine("<div class='container'>")
        [void]$sb.AppendLine("<div class='header'>NetApp SnapMirror Daily Report</div>")
        [void]$sb.AppendLine(("<div class='meta'>Collect Time: {0}</div>" -f ([System.Net.WebUtility]::HtmlEncode((Format-DateTimeDisplay -DateTimeValue $CollectTime)))))

        foreach ($current in ($CurrentRecords | Sort-Object DestinationVolume)) {
            $volume = [string]$current.DestinationVolume
            $previous = $null
            if ($PreviousRecordMap.ContainsKey($volume)) {
                $previous = $PreviousRecordMap[$volume]
            }

            $currentDisplay = Convert-RecordToReportValues -Record $current
            $previousDisplay = Convert-RecordToReportValues -Record $previous
            $currentStatusClass = if ($currentDisplay.Status -eq 'Failed') { 'status-failed' } elseif ($currentDisplay.Status -eq 'Success') { 'status-success' } else { '' }
            $previousStatusClass = if ($previousDisplay.Status -eq 'Failed') { 'status-failed' } elseif ($previousDisplay.Status -eq 'Success') { 'status-success' } else { '' }

            [void]$sb.AppendLine("<div class='volume-block'>")
            [void]$sb.AppendLine(("<div class='volume-title'>Volume: {0}</div>" -f [System.Net.WebUtility]::HtmlEncode($volume)))
            [void]$sb.AppendLine('<table>')
            [void]$sb.AppendLine('<tr><th>Item</th><th>Current</th><th>Previous</th></tr>')
            [void]$sb.AppendLine(("<tr><td>Start Time</td><td>{0}</td><td>{1}</td></tr>" -f [System.Net.WebUtility]::HtmlEncode($currentDisplay.StartTime), [System.Net.WebUtility]::HtmlEncode($previousDisplay.StartTime)))
            [void]$sb.AppendLine(("<tr><td>End Time</td><td>{0}</td><td>{1}</td></tr>" -f [System.Net.WebUtility]::HtmlEncode($currentDisplay.EndTime), [System.Net.WebUtility]::HtmlEncode($previousDisplay.EndTime)))
            [void]$sb.AppendLine(("<tr><td>Transfer Duration</td><td>{0}</td><td>{1}</td></tr>" -f [System.Net.WebUtility]::HtmlEncode($currentDisplay.Duration), [System.Net.WebUtility]::HtmlEncode($previousDisplay.Duration)))
            [void]$sb.AppendLine(("<tr><td>Transfer Size</td><td>{0}</td><td>{1}</td></tr>" -f [System.Net.WebUtility]::HtmlEncode($currentDisplay.Size), [System.Net.WebUtility]::HtmlEncode($previousDisplay.Size)))
            [void]$sb.AppendLine(("<tr><td>Transfer Type</td><td>{0}</td><td>{1}</td></tr>" -f [System.Net.WebUtility]::HtmlEncode($currentDisplay.Type), [System.Net.WebUtility]::HtmlEncode($previousDisplay.Type)))
            [void]$sb.AppendLine(("<tr><td>Status</td><td class='{0}'>{1}</td><td class='{2}'>{3}</td></tr>" -f $currentStatusClass, [System.Net.WebUtility]::HtmlEncode($currentDisplay.Status), $previousStatusClass, [System.Net.WebUtility]::HtmlEncode($previousDisplay.Status)))
            [void]$sb.AppendLine(("<tr><td>Error Message</td><td>{0}</td><td>{1}</td></tr>" -f [System.Net.WebUtility]::HtmlEncode($currentDisplay.ErrorMessage), [System.Net.WebUtility]::HtmlEncode($previousDisplay.ErrorMessage)))
            [void]$sb.AppendLine('</table>')
            [void]$sb.AppendLine('</div>')
        }

        [void]$sb.AppendLine('</div>')
        [void]$sb.AppendLine('</body>')
        [void]$sb.AppendLine('</html>')

        Write-Log -Message 'HTML report built successfully.'
        return $sb.ToString()
    }
    catch {
        Write-Log -Message ('Failed to build HTML report: {0}' -f $_.Exception.Message) -Level 'ERROR'
        throw
    }
}

function Get-SmtpCredential {
    param(
        [Parameter(Mandatory = $true)]$MailConfig
    )

    try {
        if ($MailConfig.UseAuthentication -ne $true) {
            return $null
        }

        $plainPassword = Get-ConfigPassword -Section $MailConfig -PlainPropertyName 'SenderPassword' -EncryptedPropertyName 'SenderPasswordEncrypted' -SectionName 'Mail'
        $securePassword = ConvertTo-SecureString -String $plainPassword -AsPlainText -Force
        return New-Object System.Management.Automation.PSCredential($MailConfig.Sender, $securePassword)
    }
    catch {
        Write-Log -Message ('Failed to create SMTP credential: {0}' -f $_.Exception.Message) -Level 'ERROR'
        throw
    }
}

function Send-ReportMail {
    param(
        [Parameter(Mandatory = $true)]$MailConfig,
        [Parameter(Mandatory = $true)][string]$Subject,
        [Parameter(Mandatory = $true)][string]$BodyHtml,
        [Parameter(Mandatory = $false)][string[]]$AttachmentPaths = @()
    )

    $mailMessage = $null
    $smtpClient = $null
    $mailAttachments = @()

    try {
        $mailMessage = New-Object System.Net.Mail.MailMessage
        $mailMessage.From = $MailConfig.Sender

        foreach ($address in @($MailConfig.To)) {
            if (-not [string]::IsNullOrWhiteSpace([string]$address)) {
                [void]$mailMessage.To.Add([string]$address)
            }
        }

        foreach ($address in @($MailConfig.Cc)) {
            if (-not [string]::IsNullOrWhiteSpace([string]$address)) {
                [void]$mailMessage.CC.Add([string]$address)
            }
        }

        foreach ($address in @($MailConfig.Bcc)) {
            if (-not [string]::IsNullOrWhiteSpace([string]$address)) {
                [void]$mailMessage.Bcc.Add([string]$address)
            }
        }

        $mailMessage.Subject = $Subject
        $mailMessage.SubjectEncoding = [System.Text.Encoding]::UTF8
        $mailMessage.BodyEncoding = [System.Text.Encoding]::UTF8
        $mailMessage.IsBodyHtml = $true
        $mailMessage.Body = $BodyHtml

        foreach ($attachmentPath in @($AttachmentPaths)) {
            if ([string]::IsNullOrWhiteSpace([string]$attachmentPath)) {
                continue
            }

            if (-not (Test-Path -LiteralPath $attachmentPath)) {
                Write-Log -Message ('Attachment file not found, skipping: {0}' -f $attachmentPath) -Level 'WARN'
                continue
            }

            $attachment = New-Object System.Net.Mail.Attachment($attachmentPath)
            $mailAttachments += $attachment
            [void]$mailMessage.Attachments.Add($attachment)
        }

        $smtpClient = New-Object System.Net.Mail.SmtpClient($MailConfig.SmtpServer, [int]$MailConfig.Port)
        $smtpClient.EnableSsl = [bool]$MailConfig.UseSsl
        $smtpClient.DeliveryMethod = [System.Net.Mail.SmtpDeliveryMethod]::Network
        $smtpClient.Timeout = 30000

        if ($MailConfig.UseAuthentication -eq $true) {
            $smtpClient.Credentials = (Get-SmtpCredential -MailConfig $MailConfig).GetNetworkCredential()
        }
        else {
            $smtpClient.UseDefaultCredentials = $false
            $smtpClient.Credentials = $null
        }

        $smtpClient.Send($mailMessage)
        Write-Log -Message ('Email sent successfully. Subject: {0}' -f $Subject)
    }
    catch {
        Write-Log -Message ('Failed to send email: {0}' -f $_.Exception.Message) -Level 'ERROR'
        throw
    }
    finally {
        foreach ($attachment in $mailAttachments) {
            if ($attachment) { $attachment.Dispose() }
        }
        if ($mailMessage) { $mailMessage.Dispose() }
        if ($smtpClient) { $smtpClient.Dispose() }
    }
}

function Main {
    Write-Log -Message 'Script started.'

    $config = Read-ConfigFile -Path $ConfigPath
    Initialize-LogFile -Config $config
    Write-Log -Message ('Config path in use: {0}' -f $ConfigPath)

    Validate-Config -Config $config
    Ensure-Folder -Path $config.History.Folder
    Ensure-Folder -Path $config.Log.Folder

    $collectTime = Get-Date
    $apiUrl = Build-RestApiUrl -ClusterUrl $config.Ontap.ClusterUrl -ApiPath $config.Ontap.ApiPath
    $ontapPassword = Get-ConfigPassword -Section $config.Ontap -PlainPropertyName 'Password' -EncryptedPropertyName 'PasswordEncrypted' -SectionName 'Ontap'
    $headers = Get-BasicAuthHeader -Username $config.Ontap.Username -Password $ontapPassword
    $response = Invoke-OntapRestApi -Url $apiUrl -Headers $headers -IgnoreCertificate ([bool]$config.Ontap.IgnoreCertificate)

    if (-not $response.records) {
        Write-Log -Message 'No records were returned by the API.' -Level 'WARN'
        $apiRecords = @()
    }
    else {
        $apiRecords = @($response.records)
    }

    Write-Log -Message ('Record count returned by API: {0}' -f $apiRecords.Count)

    $currentRecords = @()
    foreach ($apiRecord in $apiRecords) {
        $currentRecords += Build-CurrentRecordObject -ApiRecord $apiRecord -CollectTime $collectTime
    }
    $script:LastCollectedRecords = @($currentRecords)

    $currentCsvPath = Get-MonthlyCsvPath -Folder $config.History.Folder -ReferenceDate $collectTime
    $previousCsvPath = Get-MonthlyCsvPath -Folder $config.History.Folder -ReferenceDate ($collectTime.AddMonths(-1))
    $currentMonthHistory = @(Import-HistoryCsv -CsvPath $currentCsvPath)
    $previousMonthHistory = @(Import-HistoryCsv -CsvPath $previousCsvPath)
    $dedupMode = Get-HistoryDedupMode -Config $config

    Write-Log -Message ('History dedup mode: {0}' -f $dedupMode)

    $previousRecordMap = @{}
    foreach ($record in $currentRecords) {
        $previousRecordMap[[string]$record.DestinationVolume] = Find-PreviousRecord -CurrentRecord $record -CurrentMonthHistory $currentMonthHistory -PreviousMonthHistory $previousMonthHistory
    }

    $dedupResult = Remove-DuplicateRecords -CurrentRecords $currentRecords -ExistingHistory $currentMonthHistory -Mode $dedupMode
    $recordsToWrite = @($dedupResult.RecordsToWrite)
    $script:LastSkippedDuplicateRecords = @($dedupResult.SkippedRecords)

    if ($dedupResult.SkippedCount -gt 0) {
        Write-Log -Message ('Skipped duplicate record count: {0}' -f $dedupResult.SkippedCount) -Level 'WARN'
    }

    if ($recordsToWrite.Count -gt 0) {
        Append-HistoryCsv -CsvPath $currentCsvPath -Records $recordsToWrite
    }
    else {
        Write-Log -Message 'No new records to append after deduplication.' -Level 'WARN'
    }
    $html = Build-ComparisonReportHtml -CurrentRecords $currentRecords -PreviousRecordMap $previousRecordMap -CollectTime $collectTime

    $hasFailed = ($currentRecords | Where-Object { $_.Status -eq 'Failed' } | Measure-Object).Count -gt 0
    $subjectDate = $collectTime.ToString('yyyy-MM-dd')
    $subject = if ($hasFailed) { "[NetApp SnapMirror Daily Report][FAILED] $subjectDate" } else { "[NetApp SnapMirror Daily Report] $subjectDate" }
    $attachmentZipPath = New-ReportAttachmentZip -CollectTime $collectTime -SourcePaths @($currentCsvPath, $script:LogFilePath) -OutputFolder $config.Log.Folder

    Send-ReportMail -MailConfig $config.Mail -Subject $subject -BodyHtml $html -AttachmentPaths @($attachmentZipPath)
    Save-SummaryReport -Records $currentRecords -SkippedRecords $script:LastSkippedDuplicateRecords -CollectTime $collectTime
    Show-TerminalSummary -Records $currentRecords -SkippedRecords $script:LastSkippedDuplicateRecords
    Write-Log -Message 'Script finished successfully.'
}

try {
    Initialize-LogFile
    Main
    exit 0
}
catch {
    if ($script:LastCollectedRecords -and @($script:LastCollectedRecords).Count -gt 0) {
        Show-TerminalSummary -Records $script:LastCollectedRecords -SkippedRecords $script:LastSkippedDuplicateRecords
    }
    Write-Log -Message ('Script failed: {0}' -f $_.Exception.Message) -Level 'ERROR'
    Write-Log -Message ('Exception details: {0}' -f $_.ToString()) -Level 'ERROR'
    throw
}
