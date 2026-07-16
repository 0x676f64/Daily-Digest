<#
  Test-DigestDiagnostics.ps1 -- throwaway. Import, publish, Start, read Output.
  Dumps what the engine actually sees. Sends nothing. Reads no mailboxes.
#>

Write-Output "===== 1. IDENTITY ====="
Write-Output "IDENTITY_ENDPOINT present : $([bool]$env:IDENTITY_ENDPOINT)"
Write-Output "PowerShell version        : $($PSVersionTable.PSVersion)"

Write-Output "`n===== 2. AUTOMATION VARIABLES ====="
foreach ($n in 'StorageAccount','ConfigContainer','ConfigBlob') {
    try {
        $v = Get-AutomationVariable -Name $n -ErrorAction Stop
        Write-Output ("{0,-16} = '{1}'  (type: {2})" -f $n, $v, $v.GetType().Name)
    } catch {
        Write-Output ("{0,-16} = <<FAILED: {1}>>" -f $n, $_.Exception.Message)
    }
}

Write-Output "`n===== 3. RAW BLOB FETCH ====="
$acct = Get-AutomationVariable -Name 'StorageAccount'
$cont = Get-AutomationVariable -Name 'ConfigContainer'
$blob = Get-AutomationVariable -Name 'ConfigBlob'
$url  = "https://$acct.blob.core.windows.net/$cont/$blob"
Write-Output "URL: $url"

$stTok = (Invoke-RestMethod -Method Get `
    -Uri "$($env:IDENTITY_ENDPOINT)?resource=https://storage.azure.com&api-version=2019-08-01" `
    -Headers @{ 'X-IDENTITY-HEADER' = $env:IDENTITY_HEADER }).access_token
Write-Output "Storage token acquired    : $([bool]$stTok)"

$raw = Invoke-RestMethod -Method Get -Uri $url `
        -Headers @{ Authorization = "Bearer $stTok"; 'x-ms-version' = '2021-08-06' }
Write-Output "Returned .NET type        : $($raw.GetType().FullName)"
Write-Output "Is string?                : $($raw -is [string])"

Write-Output "`n===== 4. PARSED CONFIG ====="
$cfg = if ($raw -is [string]) { $raw | ConvertFrom-Json } else { $raw }
Write-Output "cfg type                  : $($cfg.GetType().FullName)"
Write-Output "Top-level properties      : $(($cfg.PSObject.Properties.Name) -join ', ')"
Write-Output "timeZoneId                : '$($cfg.timeZoneId)'"
Write-Output "sender                    : '$($cfg.sender)'"
Write-Output "orgName                   : '$($cfg.orgName)'"
Write-Output "recipients count          : $(@($cfg.recipients).Count)"
foreach ($r in @($cfg.recipients)) {
    Write-Output ("  - '{0}'  enabled={1} (type {2})" -f $r.email, $r.enabled, $(if($null -ne $r.enabled){$r.enabled.GetType().Name}else{'NULL'}))
}
Write-Output "enabled recipients        : $(@($cfg.recipients | Where-Object { $_.enabled }).Count)"

Write-Output "`n===== 5. TIME WINDOW ====="
try {
    $tz = [System.TimeZoneInfo]::FindSystemTimeZoneById($cfg.timeZoneId)
    Write-Output "Timezone resolved         : $($tz.Id)"
    $nowLocal   = [System.TimeZoneInfo]::ConvertTimeFromUtc([DateTime]::UtcNow, $tz)
    $startLocal = $nowLocal.Date.AddDays(-1)
    Write-Output "Now (local)               : $nowLocal"
    Write-Output "Yesterday start (local)   : $startLocal"
    Write-Output "Label                     : '$($startLocal.ToString('dddd, MMMM d'))'"
    Write-Output "Current culture           : $([System.Globalization.CultureInfo]::CurrentCulture.Name) / UI $([System.Globalization.CultureInfo]::CurrentUICulture.Name)"
} catch {
    Write-Output "TIMEZONE FAILED           : $($_.Exception.Message)"
}

Write-Output "`n===== 6. GRAPH ACCESS PER RECIPIENT ====="
Write-Output "(Tests the ACTUAL endpoints the engine uses. A directory lookup like /users/{id}"
Write-Output " would need User.Read.All, which we intentionally did NOT grant.)"

# PS 5.1 hides the Graph error body behind a generic message -- dig it out.
function Get-GraphErr($err) {
    try {
        $resp = $err.Exception.Response
        if ($resp) {
            $stream = $resp.GetResponseStream()
            $stream.Position = 0
            $body = (New-Object System.IO.StreamReader($stream)).ReadToEnd()
            if ($body) { return $body }
        }
    } catch {}
    return $err.Exception.Message
}

$gTok = (Invoke-RestMethod -Method Get `
    -Uri "$($env:IDENTITY_ENDPOINT)?resource=https://graph.microsoft.com&api-version=2019-08-01" `
    -Headers @{ 'X-IDENTITY-HEADER' = $env:IDENTITY_HEADER }).access_token
Write-Output "Graph token acquired      : $([bool]$gTok)"
$h = @{ Authorization = "Bearer $gTok" }

# Show what roles the token itself claims to carry -- decodes the JWT payload.
try {
    $payload = $gTok.Split('.')[1].Replace('-','+').Replace('_','/')
    while ($payload.Length % 4) { $payload += '=' }
    $claims = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($payload)) | ConvertFrom-Json
    Write-Output "Token app id (appid)      : $($claims.appid)"
    Write-Output "Token roles claim         : $(($claims.roles) -join ', ')"
    if (-not $claims.roles) {
        Write-Output "  >> NO ROLES IN TOKEN. The app role assignments were never granted."
        Write-Output "  >> Run Grant-DigestPermissions.ps1, then wait a few minutes."
    }
} catch { Write-Output "Could not decode token claims: $($_.Exception.Message)" }

foreach ($r in @($cfg.recipients)) {
    $e = $r.email
    try {
        $null = Invoke-RestMethod -Method Get -ErrorAction Stop -Headers $h `
            -Uri "https://graph.microsoft.com/v1.0/users/$e/mailFolders/inbox/messages?`$top=1&`$select=subject"
        Write-Output "  MAIL  OK   $e"
    } catch { Write-Output "  MAIL  FAIL $e : $(Get-GraphErr $_)" }

    try {
        $s = [DateTime]::UtcNow.AddDays(-1).ToString('yyyy-MM-ddTHH:mm:ssZ')
        $t = [DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ')
        $null = Invoke-RestMethod -Method Get -ErrorAction Stop -Headers $h `
            -Uri "https://graph.microsoft.com/v1.0/users/$e/calendarView?startDateTime=$s&endDateTime=$t&`$top=1&`$select=subject"
        Write-Output "  CAL   OK   $e"
    } catch { Write-Output "  CAL   FAIL $e : $(Get-GraphErr $_)" }
}

try {
    $null = Invoke-RestMethod -Method Get -ErrorAction Stop -Headers $h `
        -Uri "https://graph.microsoft.com/v1.0/users/$($cfg.sender)/mailFolders/inbox?`$select=id"
    Write-Output "  SENDER OK  $($cfg.sender)"
} catch { Write-Output "  SENDER FAIL $($cfg.sender) : $(Get-GraphErr $_)" }

Write-Output "`nHOW TO READ A 403:"
Write-Output "  'Access is denied' / Authorization_RequestDenied  -> app roles missing or not propagated"
Write-Output "  'Access to OData is disabled'                     -> Application Access Policy is blocking"

Write-Output "`n===== DONE ====="