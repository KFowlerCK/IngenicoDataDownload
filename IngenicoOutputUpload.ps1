$ProgressPreference = 'SilentlyContinue'
$InformationPreference = 'SilentlyContinue'
$WarningPreference = 'SilentlyContinue'
$ErrorActionPreference = 'Stop'

function Get-StableHostname {
    $name = $null

    try { $name = (& cmd.exe /c echo %COMPUTERNAME%).Trim() } catch {}
    if ([string]::IsNullOrWhiteSpace($name)) {
        try { $name = [System.Environment]::GetEnvironmentVariable('COMPUTERNAME', 'Machine') } catch {}
    }
    if ([string]::IsNullOrWhiteSpace($name)) {
        try { $name = [System.Environment]::MachineName } catch {}
    }
    if ([string]::IsNullOrWhiteSpace($name)) {
        try { $name = (& hostname).Trim() } catch {}
    }
    if ([string]::IsNullOrWhiteSpace($name)) {
        $name = 'UNKNOWN-HOST'
    }

    return $name
}

$TenantId       = $env:INGENICO_TENANT_ID
$ClientId       = $env:INGENICO_CLIENT_ID
$ClientSecret   = $env:INGENICO_CLIENT_SECRET
$SiteId         = $env:INGENICO_SITE_ID
$LibraryFolder  = $env:INGENICO_LIBRARY_FOLDER

if ([string]::IsNullOrWhiteSpace($TenantId))      { throw 'Missing environment variable: INGENICO_TENANT_ID' }
if ([string]::IsNullOrWhiteSpace($ClientId))      { throw 'Missing environment variable: INGENICO_CLIENT_ID' }
if ([string]::IsNullOrWhiteSpace($ClientSecret))  { throw 'Missing environment variable: INGENICO_CLIENT_SECRET' }
if ([string]::IsNullOrWhiteSpace($SiteId))        { throw 'Missing environment variable: INGENICO_SITE_ID' }
if ([string]::IsNullOrWhiteSpace($LibraryFolder)) { throw 'Missing environment variable: INGENICO_LIBRARY_FOLDER' }

$regPath = 'HKLM:\SOFTWARE\WOW6432Node\Synchronics\CounterPoint\8.0'

$DeviceName = Get-StableHostname
$Now = Get-Date
$NowString = $Now.ToString('yyyy-MM-dd HH:mm:ss')
$SafeTime = $Now.ToString('yyyyMMdd_HHmmss')
$SafeHost = ($DeviceName -replace '[^A-Za-z0-9._-]', '_')

$InjectedSN = 'NOT FOUND'
$Model = 'NOT FOUND'
$Application = 'NOT FOUND'
$KSN = 'NOT FOUND'
$Status = 'FAILED'
$ErrorMessage = ''

try {
    $installPath = (Get-ItemProperty -Path $regPath -ErrorAction Stop).InstallPath
    if (-not $installPath) {
        throw 'InstallPath was empty.'
    }

    $counterpointBin = Join-Path $installPath 'BIN'

    $device = Get-PnpDevice -PresentOnly |
        Where-Object {
            $_.InstanceId -match 'VID_0B00&PID_0081' -or
            (($_.InstanceId -match '^USBVCOM') -and ($_.FriendlyName -match '^Ingenico'))
        } |
        Select-Object -First 1

    if (-not $device) {
        throw 'No matching Ingenico device found.'
    }

    if ($device.FriendlyName -match '\(COM(\d+)\)') {
        $comPort = 'COM' + $matches[1]
    }
    else {
        throw 'Could not extract COM port.'
    }

    $exePath = Join-Path $counterpointBin 'IngenicoConsoleUtility.exe'
    if (-not (Test-Path $exePath)) {
        throw 'Utility not found.'
    }

    $output = & $exePath ('/Port:' + $comPort) '/INFO' 2>&1

    $snLine = $output | Where-Object { $_ -match 'Injected Serial Number\s*[:=]\s*(.+)' } | Select-Object -First 1
    if ($snLine) {
        $InjectedSN = ($snLine -replace '.*[:=]\s*', '').Trim()
    }

    $modelLine = $output | Where-Object { $_ -match 'Model\s*[:=]\s*(.+)' } | Select-Object -First 1
    if ($modelLine) {
        $Model = ($modelLine -replace '.*[:=]\s*', '').Trim()
    }

    $appLine = $output | Where-Object { $_ -match 'Application\s*[:=]\s*(.+)' } | Select-Object -First 1
    if ($appLine) {
        $Application = ($appLine -replace '.*[:=]\s*', '').Trim()
    }

    $ksnLines = $output | Where-Object { $_ -match '^\s*KSN(?:_\d+)?\s*:\s*(.+)$' }
    if ($ksnLines) {
        $KSN = ($ksnLines | ForEach-Object { ($_ -replace '^\s*', '').Trim() }) -join ' | '
    }

    $Status = 'SUCCESS'
}
catch {
    $ErrorMessage = $_.Exception.Message
}

$RowObject = [PSCustomObject]@{
    Hostname     = $DeviceName
    Time         = $NowString
    InjectedSN   = $InjectedSN
    Model        = $Model
    Application  = $Application
    KSN          = $KSN
    Status       = $Status
    ErrorMessage = $ErrorMessage
}

$FileName = "${SafeHost}_${SafeTime}.csv"
$TempFile = Join-Path $env:TEMP $FileName

$RowObject | Export-Csv -Path $TempFile -NoTypeInformation -Encoding UTF8

try {
    $TokenBody = @{
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = 'https://graph.microsoft.com/.default'
        grant_type    = 'client_credentials'
    }

    $TokenResponse = Invoke-RestMethod `
        -Method Post `
        -Uri ('https://login.microsoftonline.com/' + $TenantId + '/oauth2/v2.0/token') `
        -Body $TokenBody `
        -ContentType 'application/x-www-form-urlencoded'

    $Headers = @{
        Authorization = 'Bearer ' + $TokenResponse.access_token
    }

    $Url = 'https://graph.microsoft.com/v1.0/sites/' + $SiteId + '/drive/root:/' + $LibraryFolder + '/' + $FileName + ':/content'

    $UploadResult = Invoke-RestMethod `
        -Method Put `
        -Uri $Url `
        -Headers $Headers `
        -InFile $TempFile `
        -ContentType 'text/csv'

    Write-Output ('Resolved Hostname: ' + $DeviceName)
    Write-Output ('Uploaded File: ' + $FileName)
    Write-Output ('Updated: ' + $UploadResult.webUrl)
}
catch {
    Write-Output ('Resolved Hostname: ' + $DeviceName)
    Write-Output ('Upload File: ' + $FileName)
    Write-Output ('UPLOAD ERROR: ' + $_.Exception.Message)
    if ($_.ErrorDetails.Message) {
        Write-Output $_.ErrorDetails.Message
    }
}