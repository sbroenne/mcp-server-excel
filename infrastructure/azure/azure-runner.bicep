// Azure Bicep Template for Excel Integration Test Runner
// Automates provisioning of Windows VM with GitHub Actions self-hosted runner

@description('Location for all resources')
param location string = 'swedencentral'

@description('VM size - B2ms provides 8GB RAM needed for Excel automation')
param vmSize string = 'Standard_B2ms'

@description('Admin username for the VM')
param adminUsername string = 'azureuser'

@description('Admin password for the VM')
@secure()
param adminPassword string

var vmName = 'vm-excel-runner'
var nicName = '${vmName}-nic'
var nsgName = '${vmName}-nsg'
var bastionPublicIpName = 'bastion-ip'
var bastionName = 'bastion-excel-runner'
var vnetName = 'vnet-excel-runner'
var subnetName = 'subnet-default'
var bastionSubnetName = 'AzureBastionSubnet' // Required name for Bastion
var osDiskName = '${vmName}-osdisk'

// Network Security Group - Allow outbound HTTPS for GitHub runner
resource nsg 'Microsoft.Network/networkSecurityGroups@2023-05-01' = {
  name: nsgName
  location: location
  properties: {
    securityRules: [
      {
        name: 'AllowHTTPS'
        properties: {
          priority: 1001
          protocol: 'Tcp'
          access: 'Allow'
          direction: 'Outbound'
          sourceAddressPrefix: '*'
          sourcePortRange: '*'
          destinationAddressPrefix: 'Internet'
          destinationPortRange: '443'
        }
      }
    ]
  }
}

// Virtual Network with Bastion subnet
resource vnet 'Microsoft.Network/virtualNetworks@2023-05-01' = {
  name: vnetName
  location: location
  properties: {
    addressSpace: {
      addressPrefixes: [
        '10.0.0.0/16'
      ]
    }
    subnets: [
      {
        name: subnetName
        properties: {
          addressPrefix: '10.0.0.0/24'
          networkSecurityGroup: {
            id: nsg.id
          }
        }
      }
      {
        name: bastionSubnetName // Must be named 'AzureBastionSubnet'
        properties: {
          addressPrefix: '10.0.1.0/26' // Minimum /26 required for Bastion
        }
      }
    ]
  }
}

// Public IP Address for Bastion (Standard SKU required)
resource bastionPublicIp 'Microsoft.Network/publicIPAddresses@2023-05-01' = {
  name: bastionPublicIpName
  location: location
  sku: {
    name: 'Standard'
  }
  properties: {
    publicIPAllocationMethod: 'Static'
  }
}

// Network Interface - No public IP (using Bastion)
resource nic 'Microsoft.Network/networkInterfaces@2023-05-01' = {
  name: nicName
  location: location
  properties: {
    ipConfigurations: [
      {
        name: 'ipconfig1'
        properties: {
          privateIPAllocationMethod: 'Dynamic'
          subnet: {
            id: vnet.properties.subnets[0].id
          }
        }
      }
    ]
  }
}

// Virtual Machine
resource vm 'Microsoft.Compute/virtualMachines@2023-07-01' = {
  name: vmName
  location: location
  properties: {
    hardwareProfile: {
      vmSize: vmSize
    }
    osProfile: {
      computerName: vmName
      adminUsername: adminUsername
      adminPassword: adminPassword
      windowsConfiguration: {
        enableAutomaticUpdates: true
        provisionVMAgent: true
        timeZone: 'UTC'
      }
    }
    storageProfile: {
      imageReference: {
        publisher: 'MicrosoftWindowsServer'
        offer: 'WindowsServer'
        sku: '2022-datacenter'
        version: 'latest'
      }
      osDisk: {
        name: osDiskName
        createOption: 'FromImage'
        managedDisk: {
          storageAccountType: 'Premium_LRS'
        }
        diskSizeGB: 128
      }
    }
    networkProfile: {
      networkInterfaces: [
        {
          id: nic.id
        }
      ]
    }
  }
}

// Azure Bastion (Developer SKU)
resource bastion 'Microsoft.Network/bastionHosts@2023-05-01' = {
  name: bastionName
  location: location
  sku: {
    name: 'Developer'
  }
  properties: {
    ipConfigurations: [
      {
        name: 'bastionIpConfig'
        properties: {
          subnet: {
            id: '${vnet.id}/subnets/${bastionSubnetName}'
          }
          publicIPAddress: {
            id: bastionPublicIp.id
          }
        }
      }
    ]
  }
}

// Outputs
output vmPrivateIP string = nic.properties.ipConfigurations[0].properties.privateIPAddress
output bastionName string = bastionName
output vmResourceId string = vm.id
output vmName string = vmName
output nextSteps string = 'Connect via Azure Portal → VM → Connect → Bastion. Then install Excel, .NET SDK, and GitHub runner manually'
output monthlyCost string = 'Estimated ~$200/month (VM $61 + Bastion Developer $140) in Sweden Central'
output manualSetup string = 'Install: 1) Office 365 Excel, 2) .NET 10 SDK, 3) GitHub Actions Runner from https://github.com/sbroenne/mcp-server-excel/settings/actions/runners'
