using 'main.bicep'

param baseName = 'dmarc'
param mailboxUserId = '13a15f32-dd72-4341-a6a3-8448452d4fc7'
param graphClientState = readEnvironmentVariable('GRAPH_CLIENT_STATE', '')
param retentionInDays = 90

// Optional: use an existing Log Analytics workspace
// param existingWorkspaceId = '/subscriptions/<sub>/resourceGroups/<rg>/providers/Microsoft.OperationalInsights/workspaces/<name>'
