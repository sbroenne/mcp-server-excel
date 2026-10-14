// Application Insights resources module
// Called by appinsights.bicep - deploys into an existing resource group

param location string
param logAnalyticsName string
param appInsightsName string
param retentionInDays int
param tags object

// Name of the workspace-transform DCR that drops noisy or unsafe telemetry at ingestion time.
var telemetrySanitizationDcrName = 'dcr-excelmcp-drop-noisy-metrics'

// Log Analytics Workspace (required backend for Application Insights)
resource logAnalytics 'Microsoft.OperationalInsights/workspaces@2023-09-01' = {
  name: logAnalyticsName
  location: location
  tags: tags
  properties: {
    sku: {
      name: 'PerGB2018'
    }
    retentionInDays: retentionInDays
    features: {
      enableLogAccessUsingOnlyResourcePermissions: true
    }
    workspaceCapping: {
      dailyQuotaGb: 2 // Cap at 2 GB/day to prevent runaway costs
    }
    publicNetworkAccessForIngestion: 'Enabled'
    publicNetworkAccessForQuery: 'Enabled'
  }
}

// Usage is exported explicitly as native OpenTelemetry requests and custom events.
// No metrics, counters, framework logs, or raw exceptions are part of the contract.
// Reject them at ingestion as defense in depth, including from older clients.
resource telemetrySanitizationDcr 'Microsoft.Insights/dataCollectionRules@2023-03-11' = {
  name: telemetrySanitizationDcrName
  location: location
  tags: tags
  kind: 'WorkspaceTransforms'
  properties: {
    dataFlows: [
      {
        streams: [
          'Microsoft-Table-AppMetrics'
        ]
        destinations: [
          'excelmcpLogs'
        ]
        transformKql: 'source | where false'
      }
      {
        streams: [
          'Microsoft-Table-AppPerformanceCounters'
        ]
        destinations: [
          'excelmcpLogs'
        ]
        // Drop all rows - performance counters are not useful telemetry for a short-lived CLI/MCP server.
        transformKql: 'source | where false'
      }
      {
        streams: [
          'Microsoft-Table-AppExceptions'
        ]
        destinations: [
          'excelmcpLogs'
        ]
        // Unhandled failures are now sanitized custom events, never raw exceptions.
        transformKql: 'source | where false'
      }
      {
        streams: [
          'Microsoft-Table-AppEvents'
        ]
        destinations: [
          'excelmcpLogs'
        ]
        transformKql: 'source | where Name != "UnhandledException" or tostring(Properties["Sanitized"]) == "true"'
      }
      {
        streams: [
          'Microsoft-Table-AppTraces'
        ]
        destinations: [
          'excelmcpLogs'
        ]
        // Framework logs are not part of the approved telemetry contract.
        transformKql: 'source | where false'
      }
    ]
    destinations: {
      logAnalytics: [
        {
          name: 'excelmcpLogs'
          workspaceResourceId: logAnalytics.id
        }
      ]
    }
  }
}

// Associate only after both resources exist. Referencing the DCR from the initial
// workspace PUT makes a fresh deployment fail because the DCR depends on the workspace.
resource telemetrySanitizationAssociation 'Microsoft.Insights/dataCollectionRuleAssociations@2023-03-11' = {
  name: 'default'
  scope: logAnalytics
  properties: {
    dataCollectionRuleId: telemetrySanitizationDcr.id
  }
}

// Application Insights (workspace-based)
resource appInsights 'Microsoft.Insights/components@2020-02-02' = {
  name: appInsightsName
  location: location
  tags: tags
  kind: 'other' // 'other' for non-web applications like console apps
  properties: {
    Application_Type: 'other'
    WorkspaceResourceId: logAnalytics.id
    IngestionMode: 'LogAnalytics'
    publicNetworkAccessForIngestion: 'Enabled'
    publicNetworkAccessForQuery: 'Enabled'
    RetentionInDays: retentionInDays
  }
}

// Outputs
output logAnalyticsWorkspaceId string = logAnalytics.id
output appInsightsName string = appInsights.name
output appInsightsConnectionString string = appInsights.properties.ConnectionString
output appInsightsInstrumentationKey string = appInsights.properties.InstrumentationKey
