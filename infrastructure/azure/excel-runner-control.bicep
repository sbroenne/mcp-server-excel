targetScope = 'subscription'

@description('Dedicated Excel runner resource group; never the mcp-windows group.')
param resourceGroupName string = 'rg-excel-copilot-runner'

@description('GitHub control application service principal; never the VM managed identity.')
param principalId string

resource group 'Microsoft.Resources/resourceGroups@2025-04-01' existing = {
  name: resourceGroupName
}

resource controlRole 'Microsoft.Authorization/roleDefinitions@2022-04-01' = {
  name: guid(subscription().id, group.id, 'excel-runner-control')
  properties: {
    roleName: 'Excel runner control ${resourceGroupName}'
    description: 'Dedicated-group VM lifecycle and settings required by shutdown scheduling, guest commands and vault policy inspection.'
    type: 'CustomRole'
    assignableScopes: [group.id]
    permissions: [{
      actions: [
        'Microsoft.Resources/subscriptions/resourceGroups/read'
        'Microsoft.Compute/virtualMachines/read'
        'Microsoft.Compute/virtualMachines/write'
        'Microsoft.Compute/virtualMachines/instanceView/read'
        'Microsoft.Compute/virtualMachines/start/action'
        'Microsoft.Compute/virtualMachines/restart/action'
        'Microsoft.Compute/virtualMachines/deallocate/action'
        'Microsoft.Compute/virtualMachines/runCommand/action'
        'Microsoft.KeyVault/vaults/read'
        'Microsoft.DevTestLab/schedules/read'
        'Microsoft.DevTestLab/schedules/write'
      ]
      notActions: []
      dataActions: []
      notDataActions: []
    }]
  }
}

resource auditRole 'Microsoft.Authorization/roleDefinitions@2022-04-01' = {
  name: guid(subscription().id, 'excel-runner-identity-audit')
  properties: {
    roleName: 'Excel runner identity role-assignment audit'
    description: 'Read role-assignment metadata so the control workflow can reject a privileged VM identity.'
    type: 'CustomRole'
    assignableScopes: [subscription().id]
    permissions: [{
      actions: [
        'Microsoft.Authorization/roleAssignments/read'
        'Microsoft.Authorization/roleDefinitions/read'
      ]
      notActions: []
      dataActions: []
      notDataActions: []
    }]
  }
}

module controlAssignment 'excel-runner-control-assignment.bicep' = {
  scope: group
  name: 'excel-runner-control-assignment'
  params: {
    principalId: principalId
    roleDefinitionId: controlRole.id
  }
}

resource auditAssignment 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  name: guid(subscription().id, principalId, auditRole.id)
  properties: {
    principalId: principalId
    principalType: 'ServicePrincipal'
    roleDefinitionId: auditRole.id
  }
}
