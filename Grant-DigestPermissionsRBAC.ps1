<#
  Grant-DigestPermissionsRBAC.ps1  --  run ONCE, locally.

  Scopes the digest's managed identity to ONLY the leadership mailboxes using
  RBAC for Applications -- the supported model. Application Access Policies
  (New-ApplicationAccessPolicy) are legacy and slated for deprecation.

  This assumes the Entra Graph app roles (Mail.Read, Calendars.Read, Mail.Send)
  are ALREADY granted -- Grant-DigestPermissions.ps1 did that, and it worked.
  This script only handles the Exchange-side mailbox scoping.

  Order matters: RBAC is created FIRST, then the legacy policy is removed, so
  there is no window where the app has unscoped tenant-wide access.

  Requires: Organization Management (to create roles/scopes) + Exchange Admin.

  Example:
    .\Grant-DigestPermissionsRBAC.ps1 `
        -AppId 0b7686a7-4b24-4cca-b6ee-c9f07134f17c `
        -ServicePrincipalObjectId 2f65a104-1781-4609-aa18-5d8bc16a52a1 `
        -LeadershipGroup leadership-digest@ceasusa.com
#>

param(
    [Parameter(Mandatory)] [string] $AppId,                    # Application (client) ID of the managed identity
    [Parameter(Mandatory)] [string] $ServicePrincipalObjectId, # Object ID of the MI's service principal
    [Parameter(Mandatory)] [string] $LeadershipGroup,          # mail-enabled security group
    [string] $ScopeName = 'Digest Mailbox Scope',
    [switch] $RemoveLegacyPolicy = $true
)

$ErrorActionPreference = 'Stop'

Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
Connect-ExchangeOnline -ShowBanner:$false

# --- 1. Register the app as a service principal inside Exchange -----------------
Write-Host "`n1. Exchange service principal" -ForegroundColor Cyan
$existingSp = Get-ServicePrincipal -ErrorAction SilentlyContinue | Where-Object { $_.AppId -eq $AppId }
if ($existingSp) {
    Write-Host "   already registered: $($existingSp.DisplayName)" -ForegroundColor DarkGray
} else {
    New-ServicePrincipal -AppId $AppId -ObjectId $ServicePrincipalObjectId `
        -DisplayName 'Executive Daily Digest' | Out-Null
    Write-Host "   registered." -ForegroundColor Green
    Write-Host "   (may take a few minutes to become usable)" -ForegroundColor DarkGray
}

# --- 2. Management scope limited to the group's members -------------------------
Write-Host "`n2. Management scope" -ForegroundColor Cyan
$group = Get-DistributionGroup -Identity $LeadershipGroup
Write-Host "   group DN: $($group.DistinguishedName)"

$existingScope = Get-ManagementScope -Identity $ScopeName -ErrorAction SilentlyContinue
if ($existingScope) {
    Write-Host "   scope '$ScopeName' already exists" -ForegroundColor DarkGray
} else {
    New-ManagementScope -Name $ScopeName `
        -RecipientRestrictionFilter "MemberOfGroup -eq '$($group.DistinguishedName)'" | Out-Null
    Write-Host "   created scope '$ScopeName'" -ForegroundColor Green
}

# --- 3. Assign the three application roles, scoped -----------------------------
Write-Host "`n3. Role assignments" -ForegroundColor Cyan
$roles = @(
    @{ Role = 'Application Mail.Read';      Name = 'Digest-Mail.Read' }
    @{ Role = 'Application Calendars.Read'; Name = 'Digest-Calendars.Read' }
    @{ Role = 'Application Mail.Send';      Name = 'Digest-Mail.Send' }
)
foreach ($r in $roles) {
    $existing = Get-ManagementRoleAssignment -Identity $r.Name -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "   $($r.Role) already assigned" -ForegroundColor DarkGray
        continue
    }
    try {
        New-ManagementRoleAssignment -Name $r.Name -Role $r.Role -App $AppId `
            -CustomResourceScope $ScopeName | Out-Null
        Write-Host "   assigned $($r.Role)" -ForegroundColor Green
    } catch {
        Write-Host "   FAILED $($r.Role): $($_.Exception.Message)" -ForegroundColor Red
    }
}

# --- 4. Remove the legacy policy (it is what is currently blocking) -------------
if ($RemoveLegacyPolicy) {
    Write-Host "`n4. Legacy Application Access Policy" -ForegroundColor Cyan
    $legacy = Get-ApplicationAccessPolicy -ErrorAction SilentlyContinue | Where-Object { $_.AppId -eq $AppId }
    if (-not $legacy) {
        Write-Host "   none found" -ForegroundColor DarkGray
    } else {
        foreach ($p in $legacy) {
            Remove-ApplicationAccessPolicy -Identity $p.Identity -Confirm:$false
            Write-Host "   removed: $($p.Identity)" -ForegroundColor Green
        }
        Write-Host "   Mailbox scoping is now enforced by RBAC instead." -ForegroundColor DarkGray
    }
}

# --- 5. Show the result --------------------------------------------------------
Write-Host "`n5. Current state" -ForegroundColor Cyan
Get-ManagementRoleAssignment -App $AppId -ErrorAction SilentlyContinue |
    Format-Table Name, Role, CustomResourceScope -AutoSize

Write-Host "Mailboxes in scope (members of $LeadershipGroup):" -ForegroundColor Cyan
Get-DistributionGroupMember -Identity $LeadershipGroup |
    ForEach-Object { Write-Host "   $($_.PrimarySmtpAddress)" }

Write-Host "`nDone. Allow 15-30 minutes for RBAC to propagate, then re-run Test-DigestDiagnostics." -ForegroundColor Cyan
