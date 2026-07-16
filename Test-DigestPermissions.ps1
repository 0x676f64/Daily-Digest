<#
  Test-DigestPermissions.ps1 -- run LOCALLY from your PC (not in Azure).
  Verifies the two things that cause a 403:
    1. Are the Graph app roles actually assigned to the managed identity?
    2. Is the Application Access Policy scoping the identity to the mailbox?

  Example:
    .\Test-DigestPermissions.ps1 `
        -AutomationPrincipalId <managedIdentityObjectId> `
        -TestMailbox jrodriguez@ceasusa.com
#>

param(
    [Parameter(Mandatory)] [string] $AutomationPrincipalId,
    [Parameter(Mandatory)] [string] $TestMailbox
)

$GraphAppId = '00000003-0000-0000-c000-000000000000'
$Expected   = @('Mail.Read','Calendars.Read','Mail.Send')

Write-Host "=== 1. APP ROLE ASSIGNMENTS ===" -ForegroundColor Cyan
Connect-MgGraph -Scopes 'Application.Read.All' -NoWelcome

$mi = Get-MgServicePrincipal -ServicePrincipalId $AutomationPrincipalId
Write-Host "Managed identity : $($mi.DisplayName)"
Write-Host "App ID           : $($mi.AppId)"

$graph = Get-MgServicePrincipal -Filter "appId eq '$GraphAppId'"
$assignments = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $mi.Id

$granted = @()
foreach ($a in $assignments) {
    if ($a.ResourceId -eq $graph.Id) {
        $role = $graph.AppRoles | Where-Object { $_.Id -eq $a.AppRoleId }
        if ($role) { $granted += $role.Value }
    }
}

if ($granted.Count -eq 0) {
    Write-Host "  NO Graph app roles assigned!" -ForegroundColor Red
    Write-Host "  -> Grant-DigestPermissions.ps1 did not run successfully. Run it." -ForegroundColor Yellow
} else {
    foreach ($e in $Expected) {
        if ($granted -contains $e) { Write-Host "  OK      $e" -ForegroundColor Green }
        else                       { Write-Host "  MISSING $e" -ForegroundColor Red }
    }
    $extra = $granted | Where-Object { $Expected -notcontains $_ }
    if ($extra) { Write-Host "  (also granted: $($extra -join ', '))" -ForegroundColor DarkGray }
}

Write-Host "`n=== 2. APPLICATION ACCESS POLICY ===" -ForegroundColor Cyan
Connect-ExchangeOnline -ShowBanner:$false

$policies = Get-ApplicationAccessPolicy -ErrorAction SilentlyContinue | Where-Object { $_.AppId -eq $mi.AppId }
if (-not $policies) {
    Write-Host "  No access policy found for this app." -ForegroundColor Yellow
    Write-Host "  -> Without one, the app can read EVERY mailbox in the tenant." -ForegroundColor Yellow
    Write-Host "  -> Note: section 3 will report 'Granted' for ANY mailbox in this state." -ForegroundColor Yellow
    Write-Host "  -> This is NOT the cause of a 403." -ForegroundColor DarkGray
} else {
    foreach ($p in $policies) {
        Write-Host "  Policy      : $($p.AccessRight)"
        Write-Host "  Scope group : $($p.ScopeIdentity)"
    }
}

Write-Host "`n=== 3. EFFECTIVE ACCESS TO $TestMailbox ===" -ForegroundColor Cyan
$result = Test-ApplicationAccessPolicy -AppId $mi.AppId -Identity $TestMailbox
Write-Host "  Result : $($result.AccessCheckResult)" -ForegroundColor $(if ($result.AccessCheckResult -eq 'Granted') { 'Green' } else { 'Red' })

if ($result.AccessCheckResult -ne 'Granted') {
    Write-Host "  -> The mailbox is not in the policy's scope group. Add it to the group." -ForegroundColor Yellow
}

Write-Host "`n=== VERDICT ===" -ForegroundColor Cyan
if ($granted.Count -eq 0) {
    Write-Host "Cause: app roles never granted. Run Grant-DigestPermissions.ps1." -ForegroundColor Yellow
} elseif ($result.AccessCheckResult -ne 'Granted') {
    Write-Host "Cause: Application Access Policy is blocking. Fix the group membership." -ForegroundColor Yellow
} else {
    Write-Host "Both look correct. If Graph still 403s, consent may still be propagating (wait ~10 min)." -ForegroundColor Green
}