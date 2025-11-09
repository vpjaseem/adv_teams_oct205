<# 
Build an Azure Windows VM using existing network (no NSG on NIC)

What it does
- Prompts for Azure login (Connect-AzAccount)
- **Selects subscription by name**: 'Azure Subscription CAN PAYG'
- Reuses RGs: rg-vms (VM), rg-generic (network)
- Reuses vNet/subnet: vn01-east-us / sn-internet-east-us
- Creates/reuses Public IP 'win01-ip' and NIC 'win01-nic' in rg-vms
- Tries to set SecurityType = Standard; on subscription error, retries without SecurityType
- Uses Windows Server 2025 Datacenter Gen2 image; falls back to 2022 if needed
- **Boot diagnostics disabled** (no storage accounts created)
- OS disk: Standard HDD (Standard_LRS), delete with VM
- NIC: Accelerated Networking ON, no NSG
#>

# -------------------- Settings --------------------
$SubscriptionName = 'Azure Subscription CAN PAYG'   # <<< select by Name

$VmRg          = 'rg-vms'
$Location      = 'eastus'
$VmName        = 'win01'
$VmSize        = 'Standard_D2s_v3'

$PipName       = 'win01-ip'
$NicName       = 'win01-nic'

$VnetRg        = 'rg-generic'
$VnetName      = 'vn01-east-us'
$SubnetName    = 'sn-internet-east-us'

# Local admin
$LocalAdminUser = 'user01'
$LocalAdminPass = 'MyVMPass@aj91'   # ensure your policy allows this complexity

# Primary (preferred) image: Windows Server 2025 Datacenter Gen2
$ImagePublisher = 'MicrosoftWindowsServer'
$ImageOffer     = 'WindowsServer'
$ImageSkuPrimary = '2025-datacenter-g2'
$ImageSkuFallback = '2022-datacenter'    # fallback if 2025 not available
$ImageVersion   = 'latest'

# -------------------- Modules & Auth --------------------
$ErrorActionPreference = 'Stop'
Import-Module Az.Accounts -ErrorAction Stop
Import-Module Az.Resources -ErrorAction Stop
Import-Module Az.Network   -ErrorAction Stop
Import-Module Az.Compute   -ErrorAction Stop

Write-Host "Authenticating to Azure..." -ForegroundColor Cyan
Connect-AzAccount | Out-Null

# ---- Select subscription by NAME (Option B) ----
$sub = Get-AzSubscription -SubscriptionName $SubscriptionName -ErrorAction SilentlyContinue
if (-not $sub) { throw "Subscription '$SubscriptionName' not found or not accessible to the signed-in account." }
Select-AzSubscription -SubscriptionName $SubscriptionName | Out-Null
Write-Host "Using subscription: $($sub.Name) [$($sub.Id)]" -ForegroundColor Green

# -------------------- Basic Checks --------------------
$vmRgObj = Get-AzResourceGroup -Name $VmRg -ErrorAction SilentlyContinue
if (-not $vmRgObj) { throw "Resource group '$VmRg' not found." }

$vnet = Get-AzVirtualNetwork -ResourceGroupName $VnetRg -Name $VnetName -ErrorAction Stop
$subnet = Get-AzVirtualNetworkSubnetConfig -Name $SubnetName -VirtualNetwork $vnet
if (-not $subnet) { throw "Subnet '$SubnetName' not found in vNet '$VnetName' (RG: $VnetRg)." }

# Avoid accidental overwrite
$existingVm = Get-AzVM -Name $VmName -ResourceGroupName $VmRg -ErrorAction SilentlyContinue
if ($existingVm) { throw "A VM named '$VmName' already exists in '$VmRg'." }

# -------------------- Public IP (create or reuse) --------------------
$pip = Get-AzPublicIpAddress -Name $PipName -ResourceGroupName $VmRg -ErrorAction SilentlyContinue
if ($pip) {
    Write-Host "Using existing Public IP: $PipName" -ForegroundColor Yellow
} else {
    Write-Host "Creating Public IP: $PipName" -ForegroundColor Cyan
    $pip = New-AzPublicIpAddress `
        -Name $PipName `
        -ResourceGroupName $VmRg `
        -Location $Location `
        -AllocationMethod Static `
        -Sku Standard `
        -IdleTimeoutInMinutes 4
}

# -------------------- NIC (create/reuse; Accelerated + no NSG) --------------------
$nic = Get-AzNetworkInterface -Name $NicName -ResourceGroupName $VmRg -ErrorAction SilentlyContinue
if ($nic) {
    Write-Host "Using existing NIC: $NicName" -ForegroundColor Yellow
    $ipconfig = $nic.IpConfigurations | Select-Object -First 1
    if (-not $ipconfig) { throw "Existing NIC '$NicName' has no IP configuration." }

    $needsSubnetFix = ($ipconfig.Subnet.Id -ne $subnet.Id)
    $needsPipFix    = (-not $ipconfig.PublicIpAddress) -or ($ipconfig.PublicIpAddress.Id -ne $pip.Id)
    $hasNsg         = $nic.NetworkSecurityGroup -ne $null
    $needsAccel     = -not $nic.EnableAcceleratedNetworking

    if ($hasNsg) {
        Write-Host "Detaching NSG from NIC (per requirement)..." -ForegroundColor Yellow
        $nic.NetworkSecurityGroup = $null
    }
    if ($needsSubnetFix -or $needsPipFix) {
        Write-Host "Updating NIC IP configuration..." -ForegroundColor Yellow
        $ipconfig.Subnet = New-Object Microsoft.Azure.Commands.Network.Models.PSSubnet -Property @{ Id = $subnet.Id }
        $ipconfig.PublicIpAddress = $pip
    }
    if ($needsAccel) {
        Write-Host "Enabling Accelerated Networking on NIC..." -ForegroundColor Yellow
        $nic.EnableAcceleratedNetworking = $true
    }
    if ($hasNsg -or $needsSubnetFix -or $needsPipFix -or $needsAccel) {
        Set-AzNetworkInterface -NetworkInterface $nic | Out-Null
        $nic = Get-AzNetworkInterface -Name $NicName -ResourceGroupName $VmRg
    }
} else {
    Write-Host "Creating NIC: $NicName (Accelerated ON, no NSG)" -ForegroundColor Cyan
    $nic = New-AzNetworkInterface `
        -Name $NicName `
        -ResourceGroupName $VmRg `
        -Location $Location `
        -SubnetId $subnet.Id `
        -PublicIpAddressId $pip.Id `
        -IpConfigurationName 'ipconfig1' `
        -EnableAcceleratedNetworking
    # No -NetworkSecurityGroup parameter => no NSG attached
}

# -------------------- Local admin credential --------------------
$securePass = ConvertTo-SecureString -String $LocalAdminPass -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential ($LocalAdminUser, $securePass)

# -------------------- Helper: build VM config --------------------
function New-VMConfig {
    param([string]$SkuToUse)

    $cfg = New-AzVMConfig -VMName $VmName -VMSize $VmSize

    # Try to set SecurityType = Standard (some subs require a feature flag; we catch/handle during create)
    try {
        $cfg = Set-AzVMSecurityProfile -VM $cfg -SecurityType 'Standard'
        $global:SecurityTypeSet = $true
    } catch {
        $global:SecurityTypeSet = $false
    }

    $cfg = Set-AzVMOperatingSystem `
        -VM $cfg `
        -Windows `
        -ComputerName $VmName `
        -Credential $cred `
        -ProvisionVMAgent `
        -EnableAutoUpdate

    # Primary image attempt
    $cfg = Set-AzVMSourceImage -VM $cfg -PublisherName $ImagePublisher -Offer $ImageOffer -Skus $SkuToUse -Version $ImageVersion

    # OS disk: Standard HDD LRS, delete with VM
    $cfg = Set-AzVMOSDisk -VM $cfg -CreateOption FromImage -StorageAccountType 'Standard_LRS'
    $cfg.StorageProfile.OsDisk.DeleteOption = 'Delete'

    # Disable boot diagnostics (prevents storage account creation)
    $cfg = Set-AzVMBootDiagnostic -VM $cfg -Disable

    # Attach NIC (primary); set delete option where supported
    try {
        $cfg = Add-AzVMNetworkInterface -VM $cfg -Id $nic.Id -Primary -DeleteOption 'Delete'
    } catch {
        $cfg = Add-AzVMNetworkInterface -VM $cfg -Id $nic.Id -Primary
    }

    return $cfg
}

function New-VMConfig-NoSecurity {
    param([string]$SkuToUse)

    $cfg = New-AzVMConfig -VMName $VmName -VMSize $VmSize

    $cfg = Set-AzVMOperatingSystem -VM $cfg -Windows -ComputerName $VmName -Credential $cred -ProvisionVMAgent -EnableAutoUpdate
    $cfg = Set-AzVMSourceImage  -VM $cfg -PublisherName $ImagePublisher -Offer $ImageOffer -Skus $SkuToUse -Version $ImageVersion

    # OS disk: Standard HDD LRS, delete with VM
    $cfg = Set-AzVMOSDisk -VM $cfg -CreateOption FromImage -StorageAccountType 'Standard_LRS'
    $cfg.StorageProfile.OsDisk.DeleteOption = 'Delete'

    # Disable boot diagnostics (prevents storage account creation)
    $cfg = Set-AzVMBootDiagnostic -VM $cfg -Disable

    try {
        $cfg = Add-AzVMNetworkInterface -VM $cfg -Id $nic.Id -Primary -DeleteOption 'Delete'
    } catch {
        $cfg = Add-AzVMNetworkInterface -VM $cfg -Id $nic.Id -Primary
    }
    return $cfg
}

# -------------------- Build config (try 2025, fallback to 2022) --------------------
$vmConfig = $null
try {
    $vmConfig = New-VMConfig -SkuToUse $ImageSkuPrimary
} catch {
    Write-Warning "Windows Server 2025 Gen2 image may be unavailable in '$Location'. Falling back to 2022 Datacenter."
    $vmConfig = New-VMConfig -SkuToUse $ImageSkuFallback
}

# -------------------- Create VM (retry if SecurityType=Standard blocked) --------------------
Write-Host "Creating VM '$VmName' in '$VmRg' ($Location)..." -ForegroundColor Cyan
$retryWithoutSecurityType = $false
try {
    $null = New-AzVM -ResourceGroupName $VmRg -Location $Location -VM $vmConfig -ErrorAction Stop
} catch {
    $msg = $_.Exception.Message
    $needsStdRetry = $msg -match 'UseStandardSecurityType' -or $msg -match 'StandardSecurityTypeAsFirstClassEnum' -or $msg -match "value 'Standard' is not available for property 'securityType'"
    if ($needsStdRetry) {
        Write-Warning "Subscription not enabled for explicit SecurityType='Standard'. Retrying without setting SecurityType..."
        $retryWithoutSecurityType = $true
    } else {
        Write-Host "`n--- VM creation failed ---" -ForegroundColor Red
        Write-Host $msg -ForegroundColor Red
        if ($_.Exception.InnerException) { Write-Host ("Inner: " + $_.Exception.InnerException.Message) -ForegroundColor Red }
        throw
    }
}

if ($retryWithoutSecurityType) {
    try {
        $vmConfig2 = New-VMConfig-NoSecurity -SkuToUse $ImageSkuPrimary
    } catch {
        $vmConfig2 = New-VMConfig-NoSecurity -SkuToUse $ImageSkuFallback
    }
    $null = New-AzVM -ResourceGroupName $VmRg -Location $Location -VM $vmConfig2 -ErrorAction Stop
}

# -------------------- Output sumsmary --------------------
$pipNow = Get-AzPublicIpAddress -Name $PipName -ResourceGroupName $VmRg
$pubIp  = if ($pipNow) { $pipNow.IpAddress } else { '' }

Write-Host ""
Write-Host "===== VM Created Successfully =====" -ForegroundColor Green
Write-Host ("Subscription : {0}s" -f $sub.Name)
Write-Host ("Name         : {0}" -f $VmName)
Write-Host ("ResourceGroup: {0}" -f $VmRg)
Write-Host ("Location     : {0}" -f $Location)
Write-Host ("Size         : {0}" -f $VmSize)
Write-Host ("Image SKU    : {0}" -f ($vmConfig.StorageProfile.ImageReference.Sku))
Write-Host ("NIC          : {0}" -f $nic.Name)
Write-Host ("Accel Net    : {0}" -f $nic.EnableAcceleratedNetworking)
Write-Host ("Subnet       : {0}" -f $SubnetName)
Write-Host ("vNet         : {0}" -f $VnetName)
Write-Host ("Public IP    : {0}" -f $pip.Name)
Write-Host ("Public IP IP : {0}" -f $pubIp)
Write-Host ("Local Admin  : {0}" -f $LocalAdminUser)
