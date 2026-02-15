using 'main.bicep'

param baseName = 'dmarc'
param mailboxUserId = '<object-id-of-shared-mailbox-user>'
param graphClientState = '<generate-a-random-secret>'
param retentionInDays = 90

// Optional: use an existing Log Analytics workspace
// param existingWorkspaceId = '/subscriptions/<sub>/resourceGroups/<rg>/providers/Microsoft.OperationalInsights/workspaces/<name>'
