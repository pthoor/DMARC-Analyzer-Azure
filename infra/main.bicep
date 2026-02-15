// ──────────────────────────────────────────────────────────────
// DMARC-to-Sentinel Pipeline — Infrastructure
// Deploys: Log Analytics, Custom Table, DCR, Storage, App Insights,
//          Function App (PowerShell 7.4) with Managed Identity,
//          and RBAC role assignment for Logs Ingestion.
// ──────────────────────────────────────────────────────────────

targetScope = 'resourceGroup'

// ── Parameters ──

@description('Azure region for all resources.')
param location string = resourceGroup().location

@description('Base name used to derive resource names (e.g., "dmarc").')
@minLength(3)
@maxLength(16)
param baseName string = 'dmarc'

@description('Object ID of the shared mailbox user in Entra ID.')
param mailboxUserId string

@description('Random secret used to validate Graph change notifications.')
@secure()
param graphClientState string

@description('Retention period in days for the Log Analytics workspace.')
@minValue(30)
@maxValue(730)
param retentionInDays int = 90

@description('Use an existing Log Analytics workspace. Leave empty to create a new one.')
param existingWorkspaceId string = ''

@description('Optional: Resource ID of an existing Application Insights instance.')
param existingAppInsightsId string = ''

// ── Variables ──

var uniqueSuffix = uniqueString(resourceGroup().id, baseName)
var workspaceName = '${baseName}-law-${uniqueSuffix}'
var dcrName = '${baseName}-dcr-${uniqueSuffix}'
var storageName = toLower('${baseName}st${uniqueSuffix}')
var appInsightsName = '${baseName}-ai-${uniqueSuffix}'
var hostingPlanName = '${baseName}-plan-${uniqueSuffix}'
var functionAppName = '${baseName}-func-${uniqueSuffix}'
var customTableName = 'DMARCReports_CL'
var streamName = 'Custom-${customTableName}'

// Monitoring Metrics Publisher role definition ID
var monitoringMetricsPublisherRoleId = '3913510d-42f4-4e42-8a64-420c390055eb'

// ── Log Analytics Workspace ──

resource workspace 'Microsoft.OperationalInsights/workspaces@2023-09-01' = if (empty(existingWorkspaceId)) {
  name: workspaceName
  location: location
  properties: {
    sku: {
      name: 'PerGB2018'
    }
    retentionInDays: retentionInDays
  }
}

var workspaceId = empty(existingWorkspaceId) ? workspace.id : existingWorkspaceId

// ── Custom Table ──

resource customTable 'Microsoft.OperationalInsights/workspaces/tables@2022-10-01' = {
  name: '${empty(existingWorkspaceId) ? workspaceName : last(split(existingWorkspaceId, '/'))}/${customTableName}'
  properties: {
    schema: {
      name: customTableName
      columns: [
        { name: 'TimeGenerated', type: 'dateTime' }
        { name: 'ReportOrgName', type: 'string' }
        { name: 'ReportEmail', type: 'string' }
        { name: 'ReportExtraContactInfo', type: 'string' }
        { name: 'ReportId', type: 'string' }
        { name: 'ReportDateRangeBegin', type: 'dateTime' }
        { name: 'ReportDateRangeEnd', type: 'dateTime' }
        { name: 'Domain', type: 'string' }
        { name: 'PolicyPublished_p', type: 'string' }
        { name: 'PolicyPublished_sp', type: 'string' }
        { name: 'PolicyPublished_pct', type: 'int' }
        { name: 'PolicyPublished_adkim', type: 'string' }
        { name: 'PolicyPublished_aspf', type: 'string' }
        { name: 'PolicyPublished_fo', type: 'string' }
        { name: 'SourceIP', type: 'string' }
        { name: 'MessageCount', type: 'int' }
        { name: 'PolicyEvaluated_disposition', type: 'string' }
        { name: 'PolicyEvaluated_dkim', type: 'string' }
        { name: 'PolicyEvaluated_spf', type: 'string' }
        { name: 'PolicyEvaluated_reason_type', type: 'string' }
        { name: 'PolicyEvaluated_reason_comment', type: 'string' }
        { name: 'HeaderFrom', type: 'string' }
        { name: 'EnvelopeFrom', type: 'string' }
        { name: 'EnvelopeTo', type: 'string' }
        { name: 'DkimResult', type: 'string' }
        { name: 'DkimDomain', type: 'string' }
        { name: 'DkimSelector', type: 'string' }
        { name: 'SpfResult', type: 'string' }
        { name: 'SpfDomain', type: 'string' }
        { name: 'SpfScope', type: 'string' }
        { name: 'DkimAuthResults', type: 'string' }
        { name: 'SpfAuthResults', type: 'string' }
      ]
    }
    retentionInDays: retentionInDays
  }
  dependsOn: empty(existingWorkspaceId) ? [workspace] : []
}

// ── Data Collection Rule (kind: Direct — built-in logsIngestion endpoint) ──

resource dcr 'Microsoft.Insights/dataCollectionRules@2023-03-11' = {
  name: dcrName
  location: location
  kind: 'Direct'
  properties: {
    streamDeclarations: {
      '${streamName}': {
        columns: [
          { name: 'TimeGenerated', type: 'datetime' }
          { name: 'ReportOrgName', type: 'string' }
          { name: 'ReportEmail', type: 'string' }
          { name: 'ReportExtraContactInfo', type: 'string' }
          { name: 'ReportId', type: 'string' }
          { name: 'ReportDateRangeBegin', type: 'datetime' }
          { name: 'ReportDateRangeEnd', type: 'datetime' }
          { name: 'Domain', type: 'string' }
          { name: 'PolicyPublished_p', type: 'string' }
          { name: 'PolicyPublished_sp', type: 'string' }
          { name: 'PolicyPublished_pct', type: 'int' }
          { name: 'PolicyPublished_adkim', type: 'string' }
          { name: 'PolicyPublished_aspf', type: 'string' }
          { name: 'PolicyPublished_fo', type: 'string' }
          { name: 'SourceIP', type: 'string' }
          { name: 'MessageCount', type: 'int' }
          { name: 'PolicyEvaluated_disposition', type: 'string' }
          { name: 'PolicyEvaluated_dkim', type: 'string' }
          { name: 'PolicyEvaluated_spf', type: 'string' }
          { name: 'PolicyEvaluated_reason_type', type: 'string' }
          { name: 'PolicyEvaluated_reason_comment', type: 'string' }
          { name: 'HeaderFrom', type: 'string' }
          { name: 'EnvelopeFrom', type: 'string' }
          { name: 'EnvelopeTo', type: 'string' }
          { name: 'DkimResult', type: 'string' }
          { name: 'DkimDomain', type: 'string' }
          { name: 'DkimSelector', type: 'string' }
          { name: 'SpfResult', type: 'string' }
          { name: 'SpfDomain', type: 'string' }
          { name: 'SpfScope', type: 'string' }
          { name: 'DkimAuthResults', type: 'string' }
          { name: 'SpfAuthResults', type: 'string' }
        ]
      }
    }
    destinations: {
      logAnalytics: [
        {
          workspaceResourceId: workspaceId
          name: 'dmarcWorkspace'
        }
      ]
    }
    dataFlows: [
      {
        streams: [streamName]
        destinations: ['dmarcWorkspace']
        transformKql: 'source'
        outputStream: streamName
      }
    ]
  }
  dependsOn: [customTable]
}

// ── Storage Account ──

resource storageAccount 'Microsoft.Storage/storageAccounts@2023-05-01' = {
  name: storageName
  location: location
  sku: {
    name: 'Standard_LRS'
  }
  kind: 'StorageV2'
  properties: {
    supportsHttpsTrafficOnly: true
    minimumTlsVersion: 'TLS1_2'
  }
}

// ── Application Insights ──

resource appInsights 'Microsoft.Insights/components@2020-02-02' = if (empty(existingAppInsightsId)) {
  name: appInsightsName
  location: location
  kind: 'web'
  properties: {
    Application_Type: 'web'
    WorkspaceResourceId: workspaceId
  }
  dependsOn: empty(existingWorkspaceId) ? [workspace] : []
}

var appInsightsConnectionString = empty(existingAppInsightsId)
  ? appInsights.properties.ConnectionString
  : reference(existingAppInsightsId, '2020-02-02').ConnectionString

// ── Hosting Plan (Consumption) ──

resource hostingPlan 'Microsoft.Web/serverfarms@2023-12-01' = {
  name: hostingPlanName
  location: location
  sku: {
    name: 'Y1'
    tier: 'Dynamic'
  }
  kind: 'functionapp'
  properties: {
    reserved: true // Linux
  }
}

// ── Function App ──

resource functionApp 'Microsoft.Web/sites@2023-12-01' = {
  name: functionAppName
  location: location
  kind: 'functionapp,linux'
  identity: {
    type: 'SystemAssigned'
  }
  properties: {
    serverFarmId: hostingPlan.id
    httpsOnly: true
    siteConfig: {
      linuxFxVersion: 'PowerShell|7.4'
      appSettings: [
        { name: 'AzureWebJobsStorage', value: 'DefaultEndpointsProtocol=https;AccountName=${storageAccount.name};AccountKey=${storageAccount.listKeys().keys[0].value}' }
        { name: 'WEBSITE_CONTENTAZUREFILECONNECTIONSTRING', value: 'DefaultEndpointsProtocol=https;AccountName=${storageAccount.name};AccountKey=${storageAccount.listKeys().keys[0].value}' }
        { name: 'WEBSITE_CONTENTSHARE', value: functionAppName }
        { name: 'FUNCTIONS_EXTENSION_VERSION', value: '~4' }
        { name: 'FUNCTIONS_WORKER_RUNTIME', value: 'powershell' }
        { name: 'APPLICATIONINSIGHTS_CONNECTION_STRING', value: appInsightsConnectionString }
        { name: 'MAILBOX_USER_ID', value: mailboxUserId }
        { name: 'DCR_ENDPOINT', value: dcr.properties.logsIngestion.endpoint }
        { name: 'DCR_IMMUTABLE_ID', value: dcr.properties.immutableId }
        { name: 'DCR_STREAM_NAME', value: streamName }
        { name: 'GRAPH_CLIENT_STATE', value: graphClientState }
        // GRAPH_SUBSCRIPTION_ID is set after running New-GraphSubscription.ps1
      ]
      ftpsState: 'Disabled'
      minTlsVersion: '1.2'
    }
  }
}

// ── Role Assignment: Function App → Monitoring Metrics Publisher on DCR ──

resource dcrRoleAssignment 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  name: guid(dcr.id, functionApp.id, monitoringMetricsPublisherRoleId)
  scope: dcr
  properties: {
    roleDefinitionId: subscriptionResourceId(
      'Microsoft.Authorization/roleDefinitions',
      monitoringMetricsPublisherRoleId
    )
    principalId: functionApp.identity.principalId
    principalType: 'ServicePrincipal'
  }
}

// ── Outputs ──

output functionAppName string = functionApp.name
output functionAppPrincipalId string = functionApp.identity.principalId
output functionAppDefaultHostName string = functionApp.properties.defaultHostName
output dcrEndpoint string = dcr.properties.logsIngestion.endpoint
output dcrImmutableId string = dcr.properties.immutableId
output dcrStreamName string = streamName
output workspaceId string = workspaceId
output workspaceName string = empty(existingWorkspaceId) ? workspace.name : last(split(existingWorkspaceId, '/'))
