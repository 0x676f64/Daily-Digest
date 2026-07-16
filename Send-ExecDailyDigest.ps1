<#
  Executive Daily Digest - ENGINE  (dual-mode)
  --------------------------------------------
  Runs headless on a schedule. Two ways to authenticate, auto-detected:

    * In Azure (Automation runbook): uses the account's MANAGED IDENTITY.
      No secret stored anywhere. Config is read from the storage blob.

    * On your desktop (testing): uses a client secret from environment
      variables, and reads digest-config.json from disk.

  Graph application permissions required (granted to whichever identity is used,
  and admin-consented):  Mail.Read, Calendars.Read, Mail.Send
  Scope the identity to the leadership mailboxes with an Exchange Online
  Application Access Policy. (Both done once by Grant-DigestPermissions.ps1.)
#>

$MaxItems = 50
$UseManagedIdentity = [bool]$env:IDENTITY_ENDPOINT   # present in the Azure Automation sandbox

# --- read a setting from Automation variable (in Azure) or env var (local) ----
function Get-Setting {
    param([string]$AutoName, [string]$EnvName, [string]$Default = $null)
    if (Get-Command Get-AutomationVariable -ErrorAction SilentlyContinue) {
        try { $v = Get-AutomationVariable -Name $AutoName; if ($v) { return $v } } catch {}
    }
    if ($EnvName -and (Test-Path "env:$EnvName")) { return (Get-Item "env:$EnvName").Value }
    $Default
}

# --- token from the managed identity (Azure) ----------------------------------
function Get-MiToken {
    param([string]$Resource)
    (Invoke-RestMethod -Method Get `
        -Uri "$($env:IDENTITY_ENDPOINT)?resource=$Resource&api-version=2019-08-01" `
        -Headers @{ 'X-IDENTITY-HEADER' = $env:IDENTITY_HEADER }).access_token
}

# --- Graph token (either mode) ------------------------------------------------
function Get-GraphToken {
    if ($UseManagedIdentity) { return Get-MiToken -Resource 'https://graph.microsoft.com' }
    $body = @{
        client_id     = $env:DIGEST_CLIENT_ID
        client_secret = $env:DIGEST_CLIENT_SECRET
        scope         = 'https://graph.microsoft.com/.default'
        grant_type    = 'client_credentials'
    }
    (Invoke-RestMethod -Method Post `
        -Uri "https://login.microsoftonline.com/$($env:DIGEST_TENANT_ID)/oauth2/v2.0/token" `
        -ContentType 'application/x-www-form-urlencoded' -Body $body).access_token
}

# --- config (blob in Azure, file locally) -------------------------------------
function Get-Config {
    if ($UseManagedIdentity) {
        $acct = Get-Setting 'StorageAccount'  'DIGEST_STORAGE_ACCOUNT'
        $cont = Get-Setting 'ConfigContainer' 'DIGEST_CONFIG_CONTAINER' 'config'
        $blob = Get-Setting 'ConfigBlob'      'DIGEST_CONFIG_BLOB'      'digest-config.json'
        $tok  = Get-MiToken -Resource 'https://storage.azure.com'
        $url  = "https://$acct.blob.core.windows.net/$cont/$blob"
        try {
            $raw = Invoke-RestMethod -Method Get -Uri $url -ErrorAction Stop `
                       -Headers @{ Authorization = "Bearer $tok"; 'x-ms-version' = '2021-08-06' }
        } catch {
            throw "Could not read config from $url : $($_.Exception.Message). " +
                  "A 404 means the container or blob name is wrong (names are case-sensitive) " +
                  "or the file was never uploaded. A 403 means the managed identity is missing " +
                  "the Storage Blob Data Contributor role."
        }
        $parsed = if ($raw -is [string]) { $raw | ConvertFrom-Json } else { $raw }
        if (-not $parsed.timeZoneId -or -not $parsed.recipients) {
            throw "Config at $url parsed but is missing required fields (timeZoneId / recipients). Re-download it from the Digest Console."
        }
        return $parsed
    }
    $path = if ($env:DIGEST_CONFIG_PATH) { $env:DIGEST_CONFIG_PATH } else { Join-Path $PSScriptRoot 'digest-config.json' }
    if (-not (Test-Path $path)) { throw "Config not found at $path. Save it from the Digest Console first." }
    Get-Content $path -Raw | ConvertFrom-Json
}

$cfg = Get-Config

# --- Graph helpers ------------------------------------------------------------
function Invoke-GraphGet {
    param([string]$Uri, [hashtable]$ExtraHeaders = @{})
    $headers = @{ Authorization = "Bearer $script:Token" }
    foreach ($k in $ExtraHeaders.Keys) { $headers[$k] = $ExtraHeaders[$k] }
    $items = @()
    do {
        $page = Invoke-RestMethod -Method Get -Uri $Uri -Headers $headers
        if ($page.value) { $items += $page.value }
        $Uri = $page.'@odata.nextLink'
    } while ($Uri)
    $items
}

function Get-YesterdayWindowUtc {
    $tz         = [System.TimeZoneInfo]::FindSystemTimeZoneById($cfg.timeZoneId)
    $nowLocal   = [System.TimeZoneInfo]::ConvertTimeFromUtc([DateTime]::UtcNow, $tz)
    $startLocal = $nowLocal.Date.AddDays(-1)
    $endLocal   = $nowLocal.Date
    [pscustomobject]@{
        StartUtc = [System.TimeZoneInfo]::ConvertTimeToUtc($startLocal, $tz).ToString("yyyy-MM-ddTHH:mm:ssZ")
        EndUtc   = [System.TimeZoneInfo]::ConvertTimeToUtc($endLocal,   $tz).ToString("yyyy-MM-ddTHH:mm:ssZ")
        Label    = $startLocal.ToString('dddd, MMMM d')
    }
}

function Get-UnreadYesterday {
    param([string]$Upn, [string]$StartUtc, [string]$EndUtc)
    $filter = "isRead eq false and receivedDateTime ge $StartUtc and receivedDateTime lt $EndUtc"
    $uri = "https://graph.microsoft.com/v1.0/users/$Upn/mailFolders/inbox/messages" +
           "?`$filter=$([uri]::EscapeDataString($filter))" +
           "&`$select=subject,from,receivedDateTime,webLink&`$top=$MaxItems"
    Invoke-GraphGet -Uri $uri | Sort-Object receivedDateTime -Descending
}

function Get-EventsYesterday {
    param([string]$Upn, [string]$StartUtc, [string]$EndUtc)
    $uri = "https://graph.microsoft.com/v1.0/users/$Upn/calendarView" +
           "?startDateTime=$StartUtc&endDateTime=$EndUtc" +
           "&`$select=subject,start,end,organizer,responseStatus,isCancelled,isAllDay,showAs,webLink&`$top=100"
    Invoke-GraphGet -Uri $uri -ExtraHeaders @{ 'Prefer' = "outlook.timezone=""$($cfg.timeZoneId)""" }
}

function Split-Events {
    param([array]$Events)
    $needs = @(); $gone = @(); $cal = @()
    foreach ($e in $Events) {
        $r = $e.responseStatus.response
        if     ($e.isCancelled)                   { $gone  += $e }
        elseif ($r -eq 'declined')                { $gone  += $e }
        elseif ($r -in @('none','notResponded'))  { $needs += $e }
        else                                      { $cal   += $e }
    }
    [pscustomobject]@{ NeedsResponse = $needs; DeclinedCancelled = $gone; OnCalendar = $cal }
}

function Get-AiIntro {
    param([string]$Facts)
    if (-not $cfg.useAiNarrative) { return $null }
    $body = @{
        messages = @(
            @{ role='system'; content='Write a 2-sentence, plain, no-fluff intro for an executive daily digest. No emojis, no greeting.' }
            @{ role='user';   content=$Facts }
        )
        max_tokens = 120; temperature = 0.3
    } | ConvertTo-Json -Depth 6
    try {
        $uri = "$($env:AOAI_ENDPOINT)/openai/deployments/$($env:AOAI_DEPLOYMENT)/chat/completions?api-version=2024-06-01"
        (Invoke-RestMethod -Method Post -Uri $uri -ContentType 'application/json' `
            -Headers @{ 'api-key'=$env:AOAI_API_KEY } -Body $body).choices[0].message.content
    } catch { $null }
}

# --- HTML (eye-catching, Outlook-safe: tables + inline styles) ----------------
function New-DigestHtml {
    param([string]$Label, [array]$Unread, $Meetings, [string]$Intro)

    $navy = $cfg.headerColor
    $red  = $cfg.accentColor
    $ink  = '#1a2733'; $muted='#5b6b7b'; $line='#e3e8ee'
    $enc  = { param($s) [System.Net.WebUtility]::HtmlEncode($s) }

    function Pill($text,$bg){ "<span style='background:$bg;color:#ffffff;font-size:11px;font-weight:700;padding:2px 9px;border-radius:20px'>$text</span>" }

    function Metric($num,$label,$fg,$lblc){
        "<td width='33%' style='padding:16px 10px;text-align:center;border-right:1px solid #e8edf3'>" +
        "<div style='font-size:30px;font-weight:700;color:$fg;line-height:1'>$num</div>" +
        "<div style='font-size:11px;font-weight:700;color:$lblc;text-transform:uppercase;letter-spacing:.5px;margin-top:5px'>$label</div></td>"
    }
    function EmailRow($m){
        $from = if ($m.from.emailAddress.name){$m.from.emailAddress.name}else{$m.from.emailAddress.address}
        $t = ([datetime]$m.receivedDateTime).ToLocalTime().ToString('h:mm tt')
        "<tr><td style='padding:10px 14px;border-bottom:1px solid $line'>" +
        "<div style='font-size:14px;font-weight:600;color:$navy'>$(& $enc $m.subject)</div>" +
        "<div style='font-size:12px;color:$muted;margin-top:3px'>$(& $enc $from) &nbsp;&middot;&nbsp; $t</div></td></tr>"
    }
    function EventRow($e){
        $t = ([datetime]$e.start.dateTime).ToString('h:mm tt')
        "<tr><td style='padding:10px 14px;border-bottom:1px solid $line'>" +
        "<div style='font-size:14px;font-weight:600;color:$navy'>$(& $enc $e.subject)</div>" +
        "<div style='font-size:12px;color:$muted;margin-top:3px'>$t &nbsp;&middot;&nbsp; from $(& $enc $e.organizer.emailAddress.name)</div></td></tr>"
    }
    function Section($title,$rows,$count,$accent,$tint){
        if ($count -eq 0){ return "" }
        "<div style='margin:0 0 20px;border:1px solid $line;border-left:4px solid $accent;border-radius:6px;overflow:hidden'>" +
        "<div style='background:$tint;padding:11px 14px'>" +
        "<span style='font-size:13px;font-weight:700;color:$ink;text-transform:uppercase;letter-spacing:.4px'>$title</span> " +
        "$(Pill $count $accent)</div>" +
        "<table role='presentation' width='100%' cellpadding='0' cellspacing='0' style='border-collapse:collapse'>$rows</table></div>"
    }

    $emailRows = ($Unread | Select-Object -First $MaxItems | ForEach-Object { EmailRow $_ }) -join ''
    $needRows  = ($Meetings.NeedsResponse     | ForEach-Object { EventRow $_ }) -join ''
    $goneRows  = ($Meetings.DeclinedCancelled | ForEach-Object { EventRow $_ }) -join ''
    $calRows   = ($Meetings.OnCalendar        | ForEach-Object { EventRow $_ }) -join ''
    $introBlock = if ($Intro){ "<div style='background:#f7f9fc;border:1px solid $line;border-radius:6px;padding:14px 16px;margin:0 0 20px;font-size:14px;color:$ink;line-height:1.5'>$(& $enc $Intro)</div>" } else { "" }

    $sec = ""
    if ($cfg.sections.needsResponse)     { $sec += Section 'Needs your response'  $needRows $Meetings.NeedsResponse.Count     $red      '#fdecee' }
    if ($cfg.sections.unread)            { $sec += Section 'Unread from yesterday' $emailRows $Unread.Count                   $navy     '#eaf1f9' }
    if ($cfg.sections.declinedCancelled) { $sec += Section 'Declined or cancelled' $goneRows $Meetings.DeclinedCancelled.Count '#8a97a6' '#f1f4f7' }
    if ($cfg.sections.onCalendar)        { $sec += Section 'On your calendar'      $calRows  $Meetings.OnCalendar.Count       '#8a97a6' '#f1f4f7' }

    $logo = if ($cfg.logoUrl){ "<img src='$($cfg.logoUrl)' height='24' style='display:block;margin-bottom:8px' alt='$($cfg.orgName)'>" } else { "" }

    @"
<div style="background:#eef1f5;padding:24px 12px">
<div style="max-width:600px;margin:0 auto;font-family:'Segoe UI',Arial,sans-serif;background:#ffffff;border:1px solid $line;border-radius:10px;overflow:hidden">
  <div style="background:$navy;padding:20px 24px">
    $logo
    <div style="color:#ffffff;font-size:17px;font-weight:600">Daily executive digest</div>
    <div style="color:#9db4cc;font-size:12px;margin-top:3px">$($cfg.orgName) &nbsp;&middot;&nbsp; Recap of $Label</div>
  </div>
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;border-bottom:1px solid $line">
    <tr>
      $(Metric $Meetings.NeedsResponse.Count 'Need response' $red '#8a3a44')
      $(Metric $Unread.Count 'Unread' $navy '#3c5a78')
      $(Metric $Meetings.DeclinedCancelled.Count 'Cancelled' '#5f6b78' '#5f6b78')
    </tr>
  </table>
  <div style="padding:22px 24px">$introBlock$sec</div>
  <div style="padding:14px 24px;background:#fafbfc;border-top:1px solid $line;font-size:11px;color:#9aa7b4">Automated recap from $($cfg.orgName). Times shown in local time.</div>
</div>
</div>
"@
}

function Send-Digest {
    param([string]$ToUpn, [string]$Subject, [string]$Html)
    $payload = @{
        message = @{
            subject      = $Subject
            body         = @{ contentType='HTML'; content=$Html }
            toRecipients = @(@{ emailAddress = @{ address = $ToUpn } })
        }
        saveToSentItems = $false
    } | ConvertTo-Json -Depth 8
    Invoke-RestMethod -Method Post `
        -Uri "https://graph.microsoft.com/v1.0/users/$($cfg.sender)/sendMail" `
        -Headers @{ Authorization = "Bearer $script:Token" } `
        -ContentType 'application/json' -Body $payload
}

# --- Main ---------------------------------------------------------------------
$script:Token = Get-GraphToken
$win = Get-YesterdayWindowUtc
Write-Output "Digest run for $($win.Label) (managed identity: $UseManagedIdentity)"

foreach ($r in ($cfg.recipients | Where-Object { $_.enabled })) {
    $upn = $r.email
    try {
        $unread   = Get-UnreadYesterday -Upn $upn -StartUtc $win.StartUtc -EndUtc $win.EndUtc
        $events   = Get-EventsYesterday -Upn $upn -StartUtc $win.StartUtc -EndUtc $win.EndUtc
        $meetings = Split-Events -Events $events

        if ($unread.Count -eq 0 -and $meetings.NeedsResponse.Count -eq 0 -and $meetings.DeclinedCancelled.Count -eq 0) {
            Write-Output "$upn : nothing to report, skipped."; continue
        }
        $facts = "Unread: $($unread.Count). Never-responded meetings: $($meetings.NeedsResponse.Count). Declined/cancelled: $($meetings.DeclinedCancelled.Count)."
        $intro = Get-AiIntro -Facts $facts
        $html  = New-DigestHtml -Label $win.Label -Unread $unread -Meetings $meetings -Intro $intro
        Send-Digest -ToUpn $upn -Subject "$($cfg.subjectPrefix) - $($win.Label)" -Html $html
        Write-Output "$upn : sent ($($unread.Count) unread, $($meetings.NeedsResponse.Count) unanswered invites)."
    }
    catch { Write-Warning "$upn : $($_.Exception.Message)" }
}