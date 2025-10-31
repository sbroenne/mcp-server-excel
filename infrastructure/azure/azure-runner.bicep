# Azure Bicep Template for Excel Integration Test Runner
# Automates provisioning of Windows VM with GitHub Actions self-hosted runner

@description('Location for all resources (East US is cheapest)')
param location string = 'eastus'

@description('VM size - B2s is cheapest burstable VM suitable for testing workload')
param vmSize string = 'Standard_B2s'

@description('Admin username for the VM')
param adminUsername string = 'azureuser'

@description('Admin password for the VM')
@secure()
param adminPassword string

@description('GitHub repository URL (e.g., https://github.com/sbroenne/mcp-server-excel)')
param githubRepoUrl string

@description('GitHub runner registration token (generate from Settings > Actions > Runners > New self-hosted runner)')
@secure()
param githubRunnerToken string

var vmName = 'vm-excel-runner'
var nicName = '${vmName}-nic'
var nsgName = '${vmName}-nsg'
var publicIpName = '${vmName}-ip'
var vnetName = 'vnet-excel-runner'
var subnetName = 'subnet-default'
var osDiskName = '${vmName}-osdisk'

// Network Security Group - Restrict RDP access
resource nsg 'Microsoft.Network/networkSecurityGroups@2023-05-01' = {
  name: nsgName
  location: location
  properties: {
    securityRules: [
      {
        name: 'AllowRDP'
        properties: {
          priority: 1000
          protocol: 'Tcp'
          access: 'Allow'
          direction: 'Inbound'
          sourceAddressPrefix: '*' // Configure to your IP after deployment
          sourcePortRange: '*'
          destinationAddressPrefix: '*'
          destinationPortRange: '3389'
        }
      }
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

// Virtual Network
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
    ]
  }
}

// Public IP Address
resource publicIp 'Microsoft.Network/publicIPAddresses@2023-05-01' = {
  name: publicIpName
  location: location
  sku: {
    name: 'Basic'
  }
  properties: {
    publicIPAllocationMethod: 'Dynamic'
    dnsSettings: {
      domainNameLabel: toLower(vmName)
    }
  }
}

// Network Interface
resource nic 'Microsoft.Network/networkInterfaces@2023-05-01' = {
  name: nicName
  location: location
  properties: {
    ipConfigurations: [
      {
        name: 'ipconfig1'
        properties: {
          privateIPAllocationMethod: 'Dynamic'
          publicIPAddress: {
            id: publicIp.id
          }
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

// Auto-shutdown schedule to save costs (7 PM UTC daily)
resource autoShutdown 'Microsoft.DevTestLab/schedules@2018-09-15' = {
  name: 'shutdown-computevm-${vmName}'
  location: location
  properties: {
    status: 'Enabled'
    taskType: 'ComputeVmShutdownTask'
    dailyRecurrence: {
      time: '1900'
    }
    timeZoneId: 'UTC'
    notificationSettings: {
      status: 'Disabled'
    }
    targetResourceId: vm.id
  }
}

// VM Extension - Install .NET SDK and GitHub runner
resource vmExtension 'Microsoft.Compute/virtualMachines/extensions@2023-07-01' = {
  parent: vm
  name: 'SetupGitHubRunner'
  location: location
  properties: {
    publisher: 'Microsoft.Compute'
    type: 'CustomScriptExtension'
    typeHandlerVersion: '1.10'
    autoUpgradeMinorVersion: true
    protectedSettings: {
      commandToExecute: 'powershell -ExecutionPolicy Unrestricted -Command "Invoke-WebRequest -Uri https://aka.ms/dotnet/8.0/dotnet-sdk-win-x64.exe -OutFile C:\\dotnet-sdk.exe; Start-Process C:\\dotnet-sdk.exe -ArgumentList \'/quiet\' -Wait; New-Item -Path C:\\actions-runner -ItemType Directory -Force; Set-Location C:\\actions-runner; Invoke-WebRequest -Uri https://github.com/actions/runner/releases/download/v2.321.0/actions-runner-win-x64-2.321.0.zip -OutFile actions-runner.zip; Expand-Archive actions-runner.zip -Force; .\\config.cmd --url ${githubRepoUrl} --token ${githubRunnerToken} --name azure-excel-runner --labels self-hosted,windows,excel --runnergroup Default --work _work --unattended; .\\svc.cmd install; .\\svc.cmd start"'
    }
  }
}

// Outputs
output vmPublicIP string = publicIp.properties.dnsSettings.fqdn
output vmResourceId string = vm.id
output vmName string = vmName
output nextSteps string = 'RDP to VM using output vmPublicIP and install Office 365 Excel manually'
output monthlyCost string = 'Estimated ~$30/month with auto-shutdown (12h/day)'
