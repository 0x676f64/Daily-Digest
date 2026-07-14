<#
  Grant-DigestPermissions.ps1  —  run ONCE after the Deploy to Azure button.

  Does the two things ARM/Bicep cannot:
    1. Grants the digest's managed identity its Microsoft Graph permissions
       (Mail.Read, Calendars.Read, Mail.Send).
    2. Scopes that identity to ONLY the leadership mailboxes, via an Exchange
       Online Application Access Policy.

  You need to be a Global/Privileged Role Admin to consent Graph app roles,
  and have an Exchange admin role for the access policy.

  Prereqs (install once):
    Install-Module Microsoft.Graph -Scope CurrentUser
    Install-Module ExchangeOnlineManagement -Scope CurrentUser

  Example:
    .\Grant-DigestPermissions.ps1 `
        -AutomationPrincipalId <managedIdentityObjectId from the deployment output> `
        -LeadershipGroup leadership-digest@ceasusa.com
#>

param(
    [Parameter(Mandatory)] [string] $AutomationPrincipalId,   # managed identity object id (Bicep output)
    [Parameter(Mandatory)] [string] $LeadershipGroup          # mail-enabled security group (email or GUID)
)

$GraphAppId   = '00000003-0000-0000-c000-000000000000'
$Permissions  = @('Mail.Read', 'Calendars.Read', 'Mail.Send')

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes 'Application.Read.All', 'AppRoleAssignment.ReadWrite.All' -NoWelcome

$mi    = Get-MgServicePrincipal -ServicePrincipalId $AutomationPrincipalId
$graph = Get-MgServicePrincipal -Filter "appId eq '$GraphAppId'"
Write-Host "Managed identity: $($mi.DisplayName)  (appId $($mi.AppId))" -ForegroundColor Green

foreach ($perm in $Permissions) {
    $role = $graph.AppRoles | Where-Object { $_.Value -eq $perm -and $_.AllowedMemberTypes -contains 'Application' }
    if (-not $role) { Write-Warning "Role $perm not found; skipping."; continue }
    try {
        New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $mi.Id `
            -PrincipalId $mi.Id -ResourceId $graph.Id -AppRoleId $role.Id -ErrorAction Stop | Out-Null
        Write-Host "  granted $perm" -ForegroundColor Green
    } catch {
        if ($_.Exception.Message -match 'already exists|Permission being assigned already') {
            Write-Host "  $perm already granted" -ForegroundColor DarkGray
        } else { Write-Warning "  $perm : $($_.Exception.Message)" }
    }
}

Write-Host "`nScoping the identity to the leadership mailboxes..." -ForegroundColor Cyan
Connect-ExchangeOnline -ShowBanner:$false
try {
    New-ApplicationAccessPolicy -AppId $mi.AppId `
        -PolicyScopeGroupId $LeadershipGroup -AccessRight RestrictAccess `
        -Description 'Executive Daily Digest - read/send limited to leadership' -ErrorAction Stop | Out-Null
    Write-Host "  access policy created — identity can only touch $LeadershipGroup" -ForegroundColor Green
} catch {
    if ($_.Exception.Message -match 'already') { Write-Host "  access policy already exists" -ForegroundColor DarkGray }
    else { Write-Warning "  $($_.Exception.Message)" }
}

Write-Host "`nDone. Graph consent can take a few minutes to propagate before the first run." -ForegroundColor Cyan
Write-Host "Test now with:  Test-ApplicationAccessPolicy -AppId $($mi.AppId) -Identity ceo@ceasusa.com" -ForegroundColor DarkGray
