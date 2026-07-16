// Executive Daily Digest - infrastructure (Option B: portal runbook import)
//
// Deploys: storage + config container, Automation account with a system-assigned
// managed identity, a weekday-morning schedule, and the role grant so the engine
// can read its config blob.
//
// The runbook itself is NOT deployed here -- you import Send-ExecDailyDigest.ps1
// through the portal, so the code never leaves your machine except into your own
// tenant. After importing, link it to the 'DailyMorning' schedule (2 clicks).
//
// Then run Grant-DigestPermissions.ps1 once to give the managed identity its
// Graph permissions and scope it to the leadership mailboxes.

@description('Region for all resources.')
param location string = resourceGroup().location

@description('Short prefix for resource names (lowercase letters/numbers).')
param namePrefix string = 'execdigest'

@description('Timezone the daily job fires in. DST is handled automatically.')
param scheduleTimeZone string = 'America/Chicago'

@description('First run: 7:00 AM tomorrow. The offset must match your CURRENT clock (-05:00 = Central Daylight, summer; -06:00 = Central Standard, winter). Only the first run uses this; the timeZone above governs the rest.')
param scheduleStartTime string = '${substring(dateTimeAdd(utcNow(), 'P1D'), 0, 10)}T07:00:00-05:00'

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

// ---- weekday-morning schedule (link it to the runbook in the portal after import) ----
resource schedule 'Microsoft.Automation/automationAccounts/schedules@2023-11-01' = {
  parent: automation
  name: 'DailyMorning'
  properties: {
    description: 'Executive Daily Digest - weekday morning send'
    frequency: 'Week'
    interval: 1
    startTime: scheduleStartTime
    timeZone: scheduleTimeZone
    advancedSchedule: {
      weekDays: [ 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday' ]
    }
  }
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

// ---- what you need for the next steps ----
output automationAccountName string = automation.name
output storageAccountName string = sa.name
@description('Feed this to Grant-DigestPermissions.ps1 as -AutomationPrincipalId.')
output managedIdentityObjectId string = automation.identity.principalId
output nextSteps string = 'Import Send-ExecDailyDigest.ps1 into ${automation.name} > Runbooks > Import, publish it, link it to the DailyMorning schedule, then upload digest-config.json to the "config" container in ${sa.name}.'
