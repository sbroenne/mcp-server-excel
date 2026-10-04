@description('Temporary no-cost browser access for user sign-in/Excel activation.')
param location string = resourceGroup().location

resource vnet 'Microsoft.Network/virtualNetworks@2023-05-01' existing = {
  name: 'vnet-excel-copilot-runner'
}

resource nsg 'Microsoft.Network/networkSecurityGroups@2023-05-01' existing = {
  name: 'vm-excel-copilot-runner-nsg'
}

resource privateRdp 'Microsoft.Network/networkSecurityGroups/securityRules@2023-05-01' = {
  parent: nsg
  name: 'AllowPrivateActivation'
  properties: {
    priority: 900
    protocol: 'Tcp'
    access: 'Allow'
    direction: 'Inbound'
    sourceAddressPrefix: 'VirtualNetwork'
    sourcePortRange: '*'
    destinationAddressPrefix: '*'
    destinationPortRange: '3389'
  }
}

// Developer is free and uses the existing VNet; never silently use a paid SKU.
resource bastion 'Microsoft.Network/bastionHosts@2024-05-01' = {
  name: 'excel-activation-desktop'
  location: location
  tags: {
    project: 'mcp-server-excel'
    purpose: 'excel-cloud-agent-devtest'
  }
  sku: {
    name: 'Developer'
  }
  properties: {
    virtualNetwork: {
      id: vnet.id
    }
  }
}

output bastionName string = bastion.name
