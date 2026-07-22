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
  Mailbox scoping is enforced by RBAC for Applications (Grant-DigestPermissionsRBAC.ps1).
  Application Access Policies are legacy and should not be used.

  PARAMETERS (leave both blank for normal scheduled operation):
    TargetDate      - yyyy-MM-dd. Report on this day instead of yesterday.
                      Useful for testing against a day you know had activity.
    SendEvenIfEmpty - send the digest even when there is nothing to report,
                      so you can see the layout land in your inbox.
#>

param(
    [string] $TargetDate = '',
    [bool]   $SendEvenIfEmpty = $false
)

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

# Returns two windows plus the local reference day:
#   Email   : looks BACKWARD  (the previous day) -- what you may have missed.
#   Calendar: looks FORWARD   (from local midnight of the reference day
#             through the end of the NEXT day) -- what is coming up.
# TargetDate overrides the reference "today" for testing.
function Get-ReportWindowUtc {
    $tz = [System.TimeZoneInfo]::FindSystemTimeZoneById($cfg.timeZoneId)
    $nowLocal = [System.TimeZoneInfo]::ConvertTimeFromUtc([DateTime]::UtcNow, $tz)

    if ($TargetDate) {
        $parsed = [datetime]::MinValue
        if (-not [datetime]::TryParseExact($TargetDate, 'yyyy-MM-dd',
                [Globalization.CultureInfo]::InvariantCulture,
                [Globalization.DateTimeStyles]::None, [ref]$parsed)) {
            throw "TargetDate '$TargetDate' is not valid. Use yyyy-MM-dd, e.g. 2026-07-14."
        }
        $todayLocal = $parsed.Date
    } else {
        $todayLocal = $nowLocal.Date
    }

    # Email: the full previous calendar day.
    $mailStartLocal = $todayLocal.AddDays(-1)
    $mailEndLocal   = $todayLocal

    # Calendar: today + tomorrow (through end of tomorrow).
    $calStartLocal  = $todayLocal
    $calEndLocal    = $todayLocal.AddDays(2)

    $toUtc = { param($d) [System.TimeZoneInfo]::ConvertTimeToUtc($d, $tz).ToString("yyyy-MM-ddTHH:mm:ssZ") }

    [pscustomobject]@{
        # email window
        MailStartUtc = & $toUtc $mailStartLocal
        MailEndUtc   = & $toUtc $mailEndLocal
        MailLabel    = $mailStartLocal.ToString('dddd, MMMM d')
        # calendar window
        CalStartUtc  = & $toUtc $calStartLocal
        CalEndUtc    = & $toUtc $calEndLocal
        # local anchors for grouping today vs tomorrow
        TodayLocal   = $todayLocal
        TomorrowLocal= $todayLocal.AddDays(1)
        TodayLabel   = $todayLocal.ToString('dddd, MMMM d')
        TomorrowLabel= $todayLocal.AddDays(1).ToString('dddd, MMMM d')
    }
}

function Get-UnreadEmail {
    param([string]$Upn, [string]$StartUtc, [string]$EndUtc)
    $filter = "isRead eq false and receivedDateTime ge $StartUtc and receivedDateTime lt $EndUtc"
    $uri = "https://graph.microsoft.com/v1.0/users/$Upn/mailFolders/inbox/messages" +
           "?`$filter=$([uri]::EscapeDataString($filter))" +
           "&`$select=subject,from,receivedDateTime,webLink&`$top=$MaxItems"
    Invoke-GraphGet -Uri $uri | Sort-Object receivedDateTime -Descending
}

function Get-UpcomingEvents {
    param([string]$Upn, [string]$StartUtc, [string]$EndUtc)
    $uri = "https://graph.microsoft.com/v1.0/users/$Upn/calendarView" +
           "?startDateTime=$StartUtc&endDateTime=$EndUtc" +
           "&`$select=subject,start,end,organizer,responseStatus,isCancelled,isAllDay,showAs,webLink&`$top=100"
    Invoke-GraphGet -Uri $uri -ExtraHeaders @{ 'Prefer' = "outlook.timezone=""$($cfg.timeZoneId)""" }
}

# Forward-looking: bucket upcoming events into Today and Tomorrow, skip cancelled
# ones, and mark whether each still needs the recipient to respond.
function Split-Events {
    param([array]$Events, $Win)
    $tz = [System.TimeZoneInfo]::FindSystemTimeZoneById($cfg.timeZoneId)
    $today = @(); $tomorrow = @(); $needsCount = 0

    foreach ($e in $Events) {
        if ($e.isCancelled) { continue }               # not relevant looking forward
        if ($e.isAllDay)    { continue }               # skip all-day/OOO blocks

        # Event start, in local time, for day grouping.
        $startLocal = [datetime]$e.start.dateTime
        $r = $e.responseStatus.response
        $needsResponse = $r -in @('none','notResponded')
        if ($needsResponse) { $needsCount++ }

        $row = [pscustomobject]@{
            Subject   = $e.subject
            StartLocal= $startLocal
            Organizer = $e.organizer.emailAddress.name
            Needs     = $needsResponse
            WebLink   = $e.webLink
        }

        if     ($startLocal.Date -eq $Win.TodayLocal)    { $today    += $row }
        elseif ($startLocal.Date -eq $Win.TomorrowLocal) { $tomorrow += $row }
    }

    $today    = @($today    | Sort-Object StartLocal)
    $tomorrow = @($tomorrow | Sort-Object StartLocal)

    [pscustomobject]@{
        Today            = $today
        Tomorrow         = $tomorrow
        NeedsResponseCnt = $needsCount
    }
}

# Builds a compact, readable summary of the day for the model to reason over.
# Feeding real subjects/senders (not just counts) is what lets it point at a priority.
function Build-AiFacts {
    param([array]$Unread, $Meetings)
    $lines = @()

    if ($Unread.Count) {
        $lines += "UNREAD EMAIL FROM YESTERDAY (may have been missed):"
        foreach ($m in ($Unread | Select-Object -First 15)) {
            $from = if ($m.from.emailAddress.name) { $m.from.emailAddress.name } else { $m.from.emailAddress.address }
            $lines += "  - '$($m.subject)' from $from"
        }
    }
    if ($Meetings.Today.Count) {
        $lines += "MEETINGS TODAY:"
        foreach ($e in $Meetings.Today) {
            $tag = if ($e.Needs) { " [NOT YET ACCEPTED]" } else { "" }
            $lines += "  - $($e.StartLocal.ToString('h:mm tt')) '$($e.Subject)' from $($e.Organizer)$tag"
        }
    }
    if ($Meetings.Tomorrow.Count) {
        $lines += "MEETINGS TOMORROW:"
        foreach ($e in $Meetings.Tomorrow) {
            $tag = if ($e.Needs) { " [NOT YET ACCEPTED]" } else { "" }
            $lines += "  - $($e.StartLocal.ToString('h:mm tt')) '$($e.Subject)' from $($e.Organizer)$tag"
        }
    }
    if (-not $lines) { $lines += "No unread email, and no meetings scheduled for today or tomorrow." }
    $lines -join "`n"
}

function Get-AiIntro {
    param([string]$Facts)
    if (-not $cfg.useAiNarrative) { return $null }

    $system = @'
You write the opening line of an executive's daily email digest. Below the line you write,
the reader sees the full itemized lists, so DO NOT re-list everything.

The digest has two parts: email that may have been MISSED yesterday, and the meetings
COMING UP today and tomorrow.

Your job: in ONE to TWO sentences, orient them to the day ahead and flag anything needing action.
- Lead with the day ahead: how busy today looks, and call out the first or most important
  meeting (especially any marked NOT YET ACCEPTED -- those need an RSVP).
- Then, if there is time-sensitive unread email (contracts, legal, signatures, renewals,
  named customers, money, deadlines), point to the most important one.
- If today is clear and nothing was missed, say so plainly in one sentence.

Rules: plain professional English. No greeting, no name, no emojis, no bullet points,
no sign-off. Use only plain ASCII punctuation -- a hyphen "-" instead of an em-dash, and
straight quotes. Do not invent anything not present in the data. Never exceed two sentences.
'@
    $body = @{
        messages = @(
            @{ role='system'; content=$system }
            @{ role='user';   content=$Facts }
        )
        max_completion_tokens = 400
    } | ConvertTo-Json -Depth 6
    $endpoint   = Get-Setting 'AOAI_ENDPOINT'   'AOAI_ENDPOINT'
    $deployment = Get-Setting 'AOAI_DEPLOYMENT' 'AOAI_DEPLOYMENT'
    $apiKey     = Get-Setting 'AOAI_API_KEY'    'AOAI_API_KEY'
    if (-not ($endpoint -and $deployment -and $apiKey)) {
        Write-Warning "AI intro on, but AOAI settings missing (AOAI_ENDPOINT/AOAI_DEPLOYMENT/AOAI_API_KEY). Sending without intro."
        return $null
    }
    try {
        $uri = "$($endpoint.TrimEnd('/'))/openai/deployments/$deployment/chat/completions?api-version=2025-01-01-preview"
        $text = (Invoke-RestMethod -Method Post -Uri $uri -ContentType 'application/json' `
            -Headers @{ 'api-key'=$apiKey } -Body $body).choices[0].message.content
        if ($text) { return $text.Trim() }
        return $null
    } catch {
        Write-Warning "AI intro skipped: $($_.Exception.Message)"
        return $null   # never let the AI layer break the digest
    }
}

# --- HTML (eye-catching, Outlook-safe: tables + inline styles) ----------------
function New-DigestHtml {
    param($Win, [array]$Unread, $Meetings, [string]$Intro)

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
        $t = $e.StartLocal.ToString('h:mm tt')
        $flag = if ($e.Needs) {
            " &nbsp; <span style='background:$red;color:#fff;font-size:10px;font-weight:700;padding:1px 7px;border-radius:20px'>RSVP</span>"
        } else { "" }
        "<tr><td style='padding:10px 14px;border-bottom:1px solid $line'>" +
        "<div style='font-size:14px;font-weight:600;color:$navy'>$(& $enc $e.Subject)$flag</div>" +
        "<div style='font-size:12px;color:$muted;margin-top:3px'>$t &nbsp;&middot;&nbsp; from $(& $enc $e.Organizer)</div></td></tr>"
    }
    function Section($title,$rows,$count,$accent,$tint){
        if ($count -eq 0){ return "" }
        "<div style='margin:0 0 20px;border:1px solid $line;border-left:4px solid $accent;border-radius:6px;overflow:hidden'>" +
        "<div style='background:$tint;padding:11px 14px'>" +
        "<span style='font-size:13px;font-weight:700;color:$ink;text-transform:uppercase;letter-spacing:.4px'>$title</span> " +
        "$(Pill $count $accent)</div>" +
        "<table role='presentation' width='100%' cellpadding='0' cellspacing='0' style='border-collapse:collapse'>$rows</table></div>"
    }

    $emailRows    = ($Unread | Select-Object -First $MaxItems | ForEach-Object { EmailRow $_ }) -join ''
    $todayRows    = ($Meetings.Today    | ForEach-Object { EventRow $_ }) -join ''
    $tomorrowRows = ($Meetings.Tomorrow | ForEach-Object { EventRow $_ }) -join ''
    $introBlock = if ($Intro){ "<div style='background:#f7f9fc;border:1px solid $line;border-radius:6px;padding:14px 16px;margin:0 0 20px;font-size:14px;color:$ink;line-height:1.5'>$(& $enc $Intro)</div>" } else { "" }

    # Forward-looking calendar first (what is coming), then missed email.
    $sec = ""
    $sec += Section "Today - $($Win.TodayLabel)"       $todayRows    $Meetings.Today.Count    $navy '#eaf1f9'
    $sec += Section "Tomorrow - $($Win.TomorrowLabel)" $tomorrowRows $Meetings.Tomorrow.Count  $navy '#eaf1f9'
    if ($cfg.sections.unread) {
        $sec += Section 'Unread from yesterday' $emailRows $Unread.Count $red '#fdecee'
    }

    $logo = if ($cfg.logoUrl){ "<img src='$($cfg.logoUrl)' height='24' style='display:block;margin-bottom:8px' alt='$($cfg.orgName)'>" } else { "" }

    @"
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<div style="background:#eef1f5;padding:24px 12px">
<div style="max-width:600px;margin:0 auto;font-family:'Segoe UI',Arial,sans-serif;background:#ffffff;border:1px solid $line;border-radius:10px;overflow:hidden">
  <div style="background:$navy;padding:20px 24px">
    $logo
    <div style="color:#ffffff;font-size:17px;font-weight:600">Daily executive digest</div>
    <div style="color:#9db4cc;font-size:12px;margin-top:3px">$($cfg.orgName) &nbsp;&middot;&nbsp; $($Win.TodayLabel)</div>
  </div>
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;border-bottom:1px solid $line">
    <tr>
      $(Metric $Meetings.Today.Count 'Today' $navy '#3c5a78')
      $(Metric $Meetings.Tomorrow.Count 'Tomorrow' $navy '#3c5a78')
      $(Metric $Meetings.NeedsResponseCnt 'Need RSVP' $red '#8a3a44')
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
    # Encode the body as UTF-8 bytes explicitly. Without this, PowerShell 5.1 sends the
    # JSON as Latin-1 and any non-ASCII the AI produced (em-dash, curly quotes) arrives
    # mojibaked as "a-hat" sequences.
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($payload)
    Invoke-RestMethod -Method Post `
        -Uri "https://graph.microsoft.com/v1.0/users/$($cfg.sender)/sendMail" `
        -Headers @{ Authorization = "Bearer $script:Token" } `
        -ContentType 'application/json; charset=utf-8' -Body $bytes
}

# --- Main ---------------------------------------------------------------------
$script:Token = Get-GraphToken
$win = Get-ReportWindowUtc
Write-Output "Digest run: email recap $($win.MailLabel); calendar $($win.TodayLabel) + $($win.TomorrowLabel) (managed identity: $UseManagedIdentity)"
Write-Output "Email window : $($win.MailStartUtc) -> $($win.MailEndUtc)"
Write-Output "Cal window   : $($win.CalStartUtc) -> $($win.CalEndUtc)"
if ($TargetDate)      { Write-Output "TargetDate override active: $TargetDate" }
if ($SendEvenIfEmpty) { Write-Output "SendEvenIfEmpty active: will send even with nothing to report" }

foreach ($r in ($cfg.recipients | Where-Object { $_.enabled })) {
    $upn = $r.email
    try {
        $unread   = Get-UnreadEmail    -Upn $upn -StartUtc $win.MailStartUtc -EndUtc $win.MailEndUtc
        $events   = Get-UpcomingEvents -Upn $upn -StartUtc $win.CalStartUtc  -EndUtc $win.CalEndUtc
        $meetings = Split-Events -Events $events -Win $win

        Write-Output ("$upn : {0} unread, {1} today, {2} tomorrow, {3} need RSVP." -f `
            $unread.Count, $meetings.Today.Count, $meetings.Tomorrow.Count, $meetings.NeedsResponseCnt)

        if ($unread.Count -eq 0 -and $meetings.Today.Count -eq 0 -and $meetings.Tomorrow.Count -eq 0) {
            if (-not $SendEvenIfEmpty) {
                Write-Output "$upn : nothing to report, skipped. (Use SendEvenIfEmpty to send anyway.)"
                continue
            }
            Write-Output "$upn : nothing to report, but SendEvenIfEmpty is set -- sending."
        }
        $facts = Build-AiFacts -Unread $unread -Meetings $meetings
        $intro = Get-AiIntro -Facts $facts
        $html  = New-DigestHtml -Win $win -Unread $unread -Meetings $meetings -Intro $intro
        Send-Digest -ToUpn $upn -Subject "$($cfg.subjectPrefix) - $($win.TodayLabel)" -Html $html
        Write-Output "$upn : sent."
    }
    catch { Write-Warning "$upn : $($_.Exception.Message)" }
}