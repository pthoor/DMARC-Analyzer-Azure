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
var storageName = toLower(take('${baseName}st${uniqueSuffix}', 24))
var appInsightsName = '${baseName}-ai-${uniqueSuffix}'
var hostingPlanName = '${baseName}-plan-${uniqueSuffix}'
var functionAppName = '${baseName}-func-${uniqueSuffix}'
// Key Vault names must be 3-24 characters, all lowercase, start with a letter, and contain only letters, numbers, and '-'.
var keyVaultBase = toLower(take(baseName, 24 - 2 - length(uniqueSuffix)))
var keyVaultName = !regex('^[a-z][a-z0-9-]*$', keyVaultBase) ? error('Parameter "baseName" must be suitable for Key Vault naming: start with a letter, contain only letters, numbers, or hyphens, and produce a valid Key Vault name.') : '${keyVaultBase}kv${uniqueSuffix}'
var customTableName = 'DMARCReports_CL'
var streamName = 'Custom-${customTableName}'

// Monitoring Metrics Publisher role definition ID
var monitoringMetricsPublisherRoleId = '3913510d-42f4-4e42-8a64-420c390055eb'

// Key Vault Secrets User role definition ID
var keyVaultSecretsUserRoleId = '4633458b-17de-408a-b874-0445c86b69e6'

// Storage RBAC role definition IDs (for MI-based storage auth)
var storageBlobDataOwnerRoleId = 'b7e6dc6d-f1e8-4753-8033-0f276bb0955b'
var storageQueueDataContributorRoleId = '974c5e8b-45b9-4653-ba55-5f855dd0fb88'
var storageTableDataContributorRoleId = '0a9a7e1f-b9d0-4cc4-a60d-0319b160aaa3'

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

// Validates existing workspace resource ID format when provided.
// Expected format:
//   /subscriptions/{subId}/resourceGroups/{rg}/providers/Microsoft.OperationalInsights/workspaces/{name}
var isExistingWorkspaceIdValid = empty(existingWorkspaceId) || !empty(regex('^/subscriptions/[^/]+/resourceGroups/[^/]+/providers/Microsoft\\.OperationalInsights/workspaces/[^/]+$', existingWorkspaceId))

var resolvedWorkspaceName = empty(existingWorkspaceId)
  ? workspaceName
  : (isExistingWorkspaceIdValid
      ? last(split(existingWorkspaceId, '/'))
      : error('Parameter existingWorkspaceId must be a full resource ID of a Log Analytics workspace, e.g. /subscriptions/{subId}/resourceGroups/{rg}/providers/Microsoft.OperationalInsights/workspaces/{name}.'))

// ── Custom Table ──

resource customTable 'Microsoft.OperationalInsights/workspaces/tables@2022-10-01' = {
  name: '${resolvedWorkspaceName}/${customTableName}'
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
    allowBlobPublicAccess: false
    minimumTlsVersion: 'TLS1_2'
    defaultToOAuthAuthentication: true
    // allowSharedKeyAccess cannot be set to false on Consumption plan —
    // the platform mounts the content share using shared key internally.
    // Slice 2 (Flex Consumption + private endpoints) will remove this limitation.
  }
}

// ── Key Vault ──

resource keyVault 'Microsoft.KeyVault/vaults@2023-07-01' = {
  name: keyVaultName
  location: location
  properties: {
    sku: {
      family: 'A'
      name: 'standard'
    }
    tenantId: subscription().tenantId
    enableRbacAuthorization: true
    enabledForDeployment: false
    enabledForDiskEncryption: false
    enabledForTemplateDeployment: false
    enableSoftDelete: true
    softDeleteRetentionInDays: 90
    // Public network access is enabled to avoid the complexity and cost of private endpoints.
    // Access is controlled via RBAC - only the Function App's managed identity can read secrets.
    publicNetworkAccess: 'Enabled'
  }
}

// Store the Graph client state secret in Key Vault
resource graphClientStateSecret 'Microsoft.KeyVault/vaults/secrets@2023-07-01' = {
  parent: keyVault
  name: 'graph-client-state'
  properties: {
    value: graphClientState
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

// Validates existing App Insights resource ID format when provided.
// Expected format:
//   /subscriptions/{subId}/resourceGroups/{rg}/providers/Microsoft.Insights/components/{name}
var isExistingAppInsightsIdValid = empty(existingAppInsightsId) || !empty(regex('^/subscriptions/[^/]+/resourceGroups/[^/]+/providers/Microsoft\\.Insights/components/[^/]+$', existingAppInsightsId))

var appInsightsConnectionString = empty(existingAppInsightsId)
  ? appInsights.properties.ConnectionString
  : (isExistingAppInsightsIdValid
      ? reference(existingAppInsightsId, '2020-02-02').ConnectionString
      : error('Parameter existingAppInsightsId must be a full resource ID of an Application Insights component, e.g. /subscriptions/{subId}/resourceGroups/{rg}/providers/Microsoft.Insights/components/{name}.'))

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
// Implicit dependency on Storage Account via storageAccount.name (MI-based auth)
// and the remaining storageAccount.listKeys() call for the content share.
// Bicep resolves these references automatically — no explicit dependsOn needed.

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
        // MI-based storage auth — no account key in this setting.
        // Requires Storage Blob Data Owner, Queue Data Contributor, and Table Data Contributor
        // RBAC roles on the storage account (assigned below).
        { name: 'AzureWebJobsStorage__accountName', value: storageAccount.name }
        // Content share still requires a connection string on Consumption (Y1) plan.
        // This is a platform limitation; moving to Flex Consumption (Slice 2) removes
        // this last key dependency and allows allowSharedKeyAccess = false.
        { name: 'WEBSITE_CONTENTAZUREFILECONNECTIONSTRING', value: 'DefaultEndpointsProtocol=https;AccountName=${storageAccount.name};AccountKey=${storageAccount.listKeys().keys[0].value}' }
        { name: 'WEBSITE_CONTENTSHARE', value: functionAppName }
        { name: 'FUNCTIONS_EXTENSION_VERSION', value: '~4' }
        { name: 'FUNCTIONS_WORKER_RUNTIME', value: 'powershell' }
        { name: 'APPLICATIONINSIGHTS_CONNECTION_STRING', value: appInsightsConnectionString }
        { name: 'MAILBOX_USER_ID', value: mailboxUserId }
        { name: 'DCR_ENDPOINT', value: dcr.properties.logsIngestion.endpoint }
        { name: 'DCR_IMMUTABLE_ID', value: dcr.properties.immutableId }
        { name: 'DCR_STREAM_NAME', value: streamName }
        { name: 'GRAPH_CLIENT_STATE', value: '@Microsoft.KeyVault(SecretUri=${graphClientStateSecret.properties.secretUri})' }
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

// ── Role Assignment: Function App → Key Vault Secrets User on Key Vault ──

resource keyVaultRoleAssignment 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  name: guid(keyVault.id, functionApp.id, keyVaultSecretsUserRoleId)
  scope: keyVault
  properties: {
    roleDefinitionId: subscriptionResourceId(
      'Microsoft.Authorization/roleDefinitions',
      keyVaultSecretsUserRoleId
    )
    principalId: functionApp.identity.principalId
    principalType: 'ServicePrincipal'
  }
}

// ── Role Assignments: Function App → Storage Account (MI-based auth) ──
// Required for AzureWebJobsStorage__accountName identity-based connection.

resource storageBlobDataOwner 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  name: guid(storageAccount.id, functionApp.id, storageBlobDataOwnerRoleId)
  scope: storageAccount
  properties: {
    roleDefinitionId: subscriptionResourceId(
      'Microsoft.Authorization/roleDefinitions',
      storageBlobDataOwnerRoleId
    )
    principalId: functionApp.identity.principalId
    principalType: 'ServicePrincipal'
  }
}

resource storageQueueDataContributor 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  name: guid(storageAccount.id, functionApp.id, storageQueueDataContributorRoleId)
  scope: storageAccount
  properties: {
    roleDefinitionId: subscriptionResourceId(
      'Microsoft.Authorization/roleDefinitions',
      storageQueueDataContributorRoleId
    )
    principalId: functionApp.identity.principalId
    principalType: 'ServicePrincipal'
  }
}

resource storageTableDataContributor 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  name: guid(storageAccount.id, functionApp.id, storageTableDataContributorRoleId)
  scope: storageAccount
  properties: {
    roleDefinitionId: subscriptionResourceId(
      'Microsoft.Authorization/roleDefinitions',
      storageTableDataContributorRoleId
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
output workspaceName string = resolvedWorkspaceName
output keyVaultName string = keyVault.name
output keyVaultUri string = keyVault.properties.vaultUri
