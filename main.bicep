// Executive Daily Digest - one-click infrastructure
// Deploys: storage + config container, Automation account with a managed
// identity, the runbook (pulled from your repo), a weekday-morning schedule,
// and the role grant so the engine can read its config.
//
// After this deploys, run Grant-DigestPermissions.ps1 once to give the managed
// identity its Graph permissions and scope it to the leadership mailboxes.

@description('Region for all resources.')
param location string = resourceGroup().location

@description('Short prefix for resource names (lowercase letters/numbers).')
param namePrefix string = 'execdigest'

@description('Raw URL to Send-ExecDailyDigest.ps1 (e.g. a GitHub raw link).')
param runbookUri string

@description('Timezone the daily job fires in.')
param scheduleTimeZone string = 'America/Chicago'

@description('First run. Defaults to 7:00 AM tomorrow, Central. Adjust the offset for other zones.')
param scheduleStartTime string = '${substring(dateTimeAdd(utcNow(), 'P1D'), 0, 10)}T07:00:00-06:00'

var storageName    = toLower('${namePrefix}${uniqueString(resourceGroup().id)}')
var automationName = '${namePrefix}-aa-${uniqueString(resourceGroup().id)}'
var blobDataContributor = subscriptionResourceId('Microsoft.Authorization/roleDefinitions', 'ba92f5b4-2d11-453d-a403-e96b0029c9fe')

// ---- storage + config container ----
resource sa 'Microsoft.Storage/storageAccounts@2023-05-01' = {
  name: storageName
  location: location
  sku: { name: 'Standard_LRS' }
  kind: 'StorageV2'
  properties: {
    minimumTlsVersion: 'TLS1_2'
    allowBlobPublicAccess: false
  }
}
resource blobService 'Microsoft.Storage/storageAccounts/blobServices@2023-05-01' = {
  parent: sa
  name: 'default'
}
resource configContainer 'Microsoft.Storage/storageAccounts/blobServices/containers@2023-05-01' = {
  parent: blobService
  name: 'config'
}

// ---- Automation account with a system-assigned managed identity ----
resource automation 'Microsoft.Automation/automationAccounts@2023-11-01' = {
  name: automationName
  location: location
  identity: { type: 'SystemAssigned' }
  properties: {
    sku: { name: 'Basic' }
  }
}

// The runbook, imported and published straight from your repo.
resource runbook 'Microsoft.Automation/automationAccounts/runbooks@2023-11-01' = {
  parent: automation
  name: 'Send-ExecDailyDigest'
  location: location
  properties: {
    runbookType: 'PowerShell'
    logProgress: false
    logVerbose: false
    publishContentLink: {
      uri: runbookUri
    }
  }
}

// Non-secret pointers the engine reads to find its config blob.
resource vStorage 'Microsoft.Automation/automationAccounts/variables@2023-11-01' = {
  parent: automation
  name: 'StorageAccount'
  properties: { value: '"${sa.name}"', isEncrypted: false }
}
resource vContainer 'Microsoft.Automation/automationAccounts/variables@2023-11-01' = {
  parent: automation
  name: 'ConfigContainer'
  properties: { value: '"config"', isEncrypted: false }
}
resource vBlob 'Microsoft.Automation/automationAccounts/variables@2023-11-01' = {
  parent: automation
  name: 'ConfigBlob'
  properties: { value: '"digest-config.json"', isEncrypted: false }
}

// ---- weekday-morning schedule, linked to the runbook ----
resource schedule 'Microsoft.Automation/automationAccounts/schedules@2023-11-01' = {
  parent: automation
  name: 'DailyMorning'
  properties: {
    frequency: 'Week'
    interval: 1
    startTime: scheduleStartTime
    timeZone: scheduleTimeZone
    advancedSchedule: {
      weekDays: [ 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday' ]
    }
  }
}
resource jobSchedule 'Microsoft.Automation/automationAccounts/jobSchedules@2023-11-01' = {
  parent: automation
  name: guid(automation.id, 'DailyMorning', 'Send-ExecDailyDigest')
  properties: {
    runbook: { name: runbook.name }
    schedule: { name: schedule.name }
  }
  dependsOn: [ runbook, schedule ]
}

// ---- let the engine's identity read the config blob ----
resource blobRole 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  scope: sa
  name: guid(sa.id, automation.id, blobDataContributor)
  properties: {
    principalId: automation.identity.principalId
    roleDefinitionId: blobDataContributor
    principalType: 'ServicePrincipal'
  }
}

// ---- what you need for the one post-deploy step ----
output automationAccountName string = automation.name
output storageAccountName string = sa.name
@description('Feed this to Grant-DigestPermissions.ps1 as -AutomationPrincipalId.')
output managedIdentityObjectId string = automation.identity.principalId
output uploadConfigHint string = 'Upload digest-config.json into the "config" container of ${sa.name}.'
