// Application Insights infrastructure for ExcelMcp INTEGRATION TESTS
// This is a separate instance from production for testing telemetry functionality
//
// Deploy with: az deployment sub create --location <location> --template-file appinsights-test.bicep --parameters appinsights-test.parameters.json

targetScope = 'subscription'

@description('Name of the resource group to create')
param resourceGroupName string = 'excelmcp-test-observability'

@description('Azure region for all resources')
param location string = 'westeurope'

@description('Name of the Log Analytics workspace')
param logAnalyticsName string = 'excelmcp-test-logs'

@description('Name of the Application Insights resource')
param appInsightsName string = 'excelmcp-test-appinsights'

@description('Data retention in days - shorter for test instance to reduce costs')
@minValue(30)
@maxValue(730)
param retentionInDays int = 30

@description('Tags to apply to all resources')
param tags object = {
  project: 'ExcelMcp'
  purpose: 'TelemetryTesting'
  environment: 'test'
  managedBy: 'Bicep'
}

// Resource Group
resource rg 'Microsoft.Resources/resourceGroups@2024-03-01' = {
  name: resourceGroupName
  location: location
  tags: tags
}

// Deploy resources into the resource group
module observability 'appinsights-resources.bicep' = {
  name: 'test-observability-deployment'
  scope: rg
  params: {
    location: location
    logAnalyticsName: logAnalyticsName
    appInsightsName: appInsightsName
    retentionInDays: retentionInDays
    tags: tags
  }
}

// Outputs
output resourceGroupName string = rg.name
output logAnalyticsWorkspaceId string = observability.outputs.logAnalyticsWorkspaceId
output appInsightsName string = observability.outputs.appInsightsName
output appInsightsConnectionString string = observability.outputs.appInsightsConnectionString
output appInsightsInstrumentationKey string = observability.outputs.appInsightsInstrumentationKey
