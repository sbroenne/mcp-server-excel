// Application Insights resources module
// Called by appinsights.bicep - deploys into an existing resource group

param location string
param logAnalyticsName string
param appInsightsName string
param retentionInDays int
param tags object

// Name of the workspace-transform DCR that drops noisy AppMetrics rows at ingestion time
// (see dropNoisyMetricsDcr below). Referenced via resourceId() on the workspace to avoid
// a circular symbolic dependency between the workspace and the DCR.
var dropNoisyMetricsDcrName = 'dcr-excelmcp-drop-noisy-metrics'

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
    // Applies dropNoisyMetricsDcr's ingestion-time transform to tables (e.g. AppMetrics)
    // that don't have a table-specific DCR of their own.
    defaultDataCollectionRuleResourceId: resourceId('Microsoft.Insights/dataCollectionRules', dropNoisyMetricsDcrName)
    publicNetworkAccessForIngestion: 'Enabled'
    publicNetworkAccessForQuery: 'Enabled'
  }
}

// Workspace-transform DCR: drops noisy .NET HttpClient meter instruments and HeartbeatState rows
// from AppMetrics, and drops the entire AppPerformanceCounters table, before billing/ingestion.
// These were driving the vast majority of billed ingestion for this short-lived CLI/MCP server:
// - http.client.* gauges/histograms come from the SDK's own outbound HTTP calls (AI ingestion
//   endpoint) - not useful telemetry (see PR #661, PR #725). The http.client.* portion of this
//   drop-list must stay in sync with the identical list in
//   src/ExcelMcp.McpServer/Program.cs (DroppedHttpClientMetricNames) - update both when changed.
// - HeartbeatState and AppPerformanceCounters (Requests/Sec, Private Bytes, % Processor Time, ...)
//   are emitted by Microsoft.ApplicationInsights.WorkerService 3.1.2 with no corresponding
//   ApplicationInsightsServiceOptions flag or ITelemetryModule to disable them in-process (that
//   SDK version has no PerformanceCollectorModule/heartbeat feature at all - see the NOTE in
//   Program.cs's ConfigureTelemetry). This DCR transform is therefore the only mechanism
//   suppressing them, not a defense-in-depth backstop for an in-process disable.
resource dropNoisyMetricsDcr 'Microsoft.Insights/dataCollectionRules@2023-03-11' = {
  name: dropNoisyMetricsDcrName
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
        transformKql: 'source | where Name !in (\'http.client.open_connections\',\'http.client.active_requests\',\'http.client.connection.duration\',\'http.client.request.time_in_queue\',\'http.client.request.duration\',\'HeartbeatState\')'
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
