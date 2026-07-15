// Azure Bicep Template for Excel Integration Test Runner (cost-optimized)
//
// Design goals (see infrastructure/azure/README.md):
//   * Cheapest sustainable footprint. The VM is DEALLOCATED when idle and started
//     only for the nightly integration run (.github/workflows/integration-tests.yml),
//     so you pay ~a couple of compute-hours/night instead of 24/7.
//   * StandardSSD_LRS OS disk (not Premium).
//   * RDP over a Standard public IP, locked by NSG to a single admin IP — used once to
//     install Office/Excel. No Azure Bastion (that alone was ~$140/mo).
//   * Daily auto-shutdown as a runaway-cost backstop in case a deallocate step fails.
//
// Rough cost (D2as v5, StandardSSD 128 GB, Standard IP): ~$15-25/month with nightly runs.

@description('Location for all resources. Cheapest compute regions: eastus2 / eastus.')
param location string = 'eastus2'

@description('VM size - D2as v5 provides sustained CPU and 8 GB RAM for long Excel integration runs.')
param vmSize string = 'Standard_D2as_v5'

@description('Admin username for the VM.')
param adminUsername string = 'azureuser'

@description('Admin password for the VM.')
@secure()
param adminPassword string

@description('Source IP/CIDR allowed to RDP to the VM (your workstation public IP). Used once to install Office.')
param rdpSourceAddressPrefix string

@description('OS disk size in GB. 128 leaves headroom for Windows Updates + Office.')
param osDiskSizeGB int = 128

@description('Daily auto-shutdown time (HHmm, 24h) in the auto-shutdown time zone. Backstop only.')
param autoShutdownTime string = '0400'

@description('Time zone id for the auto-shutdown schedule.')
param autoShutdownTimeZone string = 'UTC'

var vmName = 'vm-excel-runner'
var nicName = '${vmName}-nic'
var nsgName = '${vmName}-nsg'
var publicIpName = '${vmName}-ip'
var vnetName = 'vnet-excel-runner'
var subnetName = 'subnet-default'
var osDiskName = '${vmName}-osdisk'

// Network Security Group: outbound HTTPS for the GitHub runner, inbound RDP from admin IP only.
resource nsg 'Microsoft.Network/networkSecurityGroups@2023-05-01' = {
  name: nsgName
  location: location
  properties: {
    securityRules: [
      {
        name: 'AllowRdpFromAdmin'
        properties: {
          priority: 1000
          protocol: 'Tcp'
          access: 'Allow'
          direction: 'Inbound'
          sourceAddressPrefix: rdpSourceAddressPrefix
          sourcePortRange: '*'
          destinationAddressPrefix: '*'
          destinationPortRange: '3389'
        }
      }
      {
        name: 'AllowHTTPSOutbound'
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

// Standard public IP (Static). Standard SKU only supports Static allocation; the standing
// charge is small (~$3/mo) and it gives a stable address for the RDP NSG rule.
resource publicIp 'Microsoft.Network/publicIPAddresses@2023-05-01' = {
  name: publicIpName
  location: location
  sku: {
    name: 'Standard'
  }
  properties: {
    publicIPAllocationMethod: 'Static'
  }
}

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
          publicIPAddress: {
            id: publicIp.id
          }
        }
      }
    ]
  }
}

resource vm 'Microsoft.Compute/virtualMachines@2023-07-01' = {
  name: vmName
  location: location
  properties: {
    licenseType: 'Windows_Client'
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
        publisher: 'MicrosoftWindowsDesktop'
        offer: 'windows-11'
        sku: 'win11-24h2-pro'
        version: 'latest'
      }
      osDisk: {
        name: osDiskName
        createOption: 'FromImage'
        managedDisk: {
          storageAccountType: 'StandardSSD_LRS'
        }
        diskSizeGB: osDiskSizeGB
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

// Daily auto-shutdown backstop (DevTest Labs schedule). Only guards against a failed
// deallocate; the nightly workflow normally deallocates the VM as soon as tests finish.
resource autoShutdown 'Microsoft.DevTestLab/schedules@2018-09-15' = {
  name: 'shutdown-computevm-${vmName}'
  location: location
  properties: {
    status: 'Enabled'
    taskType: 'ComputeVmShutdownTask'
    dailyRecurrence: {
      time: autoShutdownTime
    }
    timeZoneId: autoShutdownTimeZone
    notificationSettings: {
      status: 'Disabled'
      timeInMinutes: 30
    }
    targetResourceId: vm.id
  }
}

output vmName string = vmName
output vmResourceId string = vm.id
output publicIp string = publicIp.properties.ipAddress
output rdpHint string = 'RDP to ${publicIp.properties.ipAddress} as ${adminUsername}. Allowed source: ${rdpSourceAddressPrefix}'
output nextSteps string = 'RDP in, install Office/Excel, then register the runner (see infrastructure/azure/setup-runner.ps1 / README.md). The nightly workflow starts and deallocates this VM automatically.'
