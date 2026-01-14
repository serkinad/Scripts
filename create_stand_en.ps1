<#
.SYNOPSIS
    Script for automated deployment of test environments in Hyper-V.

.DESCRIPTION
    This script creates virtual machines for test environments with flexible configuration.
    The user specifies the environment name and the number of servers of each type.
    All virtual machines are created based on a parent differencing disk.

.PARAMETER No parameters
    The script works in interactive mode, requesting data from the user.

.EXAMPLE
    PS C:\> .\Create-TestStand.ps1
    
    === Creating Test Environment ===
    Enter 'X' at any stage to exit the script
    
    === STEP 1: Environment Name ===
    Enter X to exit the script
    Enter environment name (e.g., mail): mail
    
    === STEP 2: Domain Controllers ===
    Available options:
      1 - create 1 server(s)
      2 - create 2 server(s)
      X - EXIT the script
    How many domain controllers are needed?: 1
    ...

.NOTES
    Author: Aleksandr Serkin
    Version: 1.0
    Creation Date: 14.01.2026
    
    REQUIREMENTS:
    1. Hyper-V must be installed and enabled
    2. PowerShell must be run as Administrator
    3. Parent VHD must exist at: C:\vm\parent\parent.vhdx
    4. Sufficient free disk space
    
    NOTES:
    - All virtual machines are created as Generation 2
    - Differencing disks are used to save space
    - You can exit the script at any stage by entering 'X'
    
.LINK
    [Hyper-V Documentation]: https://docs.microsoft.com/en-us/windows-server/virtualization/hyper-v/hyper-v-technology-overview
#>

# Script for automating test environment deployment

# Function for numeric input with exit option
function Get-NumberInput {
    param(
        [string]$Prompt,
        [int]$Min,
        [int]$Max
    )
    
    Write-Host "`nAvailable options:" -ForegroundColor Yellow
    for ($i = $Min; $i -le $Max; $i++) {
        Write-Host "  $i" -ForegroundColor White -NoNewline
        Write-Host " - create $i server(s)" -ForegroundColor Gray
    }
    Write-Host "  X" -ForegroundColor White -NoNewline
    Write-Host " - EXIT the script" -ForegroundColor Red
    
    do {
        $input = Read-Host $Prompt
        
        # Check for exit
        if ($input -eq "X" -or $input -eq "x") {
            Write-Host "Terminating script..." -ForegroundColor Red
            exit
        }
        
        $number = $input -as [int]
        
        if ($number -eq $null -or $number -lt $Min -or $number -gt $Max) {
            Write-Host "Error: enter a number from $Min to $Max or X to exit the script" -ForegroundColor Red
        }
    } while ($number -eq $null -or $number -lt $Min -or $number -gt $Max)
    
    return $number
}

# Function for creating virtual machines
function Create-VirtualMachine {
    <#
    .SYNOPSIS
        Creates a virtual machine with specified parameters.
    
    .DESCRIPTION
        Creates a differencing disk based on a parent VHD and configures the virtual machine.
    
    .PARAMETER vmName
        Virtual machine name.
    
    .PARAMETER vmType
        Virtual machine type (used in VHD file name).
    
    .PARAMETER cpuCount
        Number of virtual processors.
    
    .PARAMETER memoryGB
        Amount of RAM in gigabytes.
    
    .EXAMPLE
        Create-VirtualMachine -vmName "mail.dc1" -vmType "dc1" -cpuCount 2 -memoryGB 8
    #>
    
    param(
        [string]$vmName,
        [string]$vmType,
        [int]$cpuCount,
        [int]$memoryGB
    )
    
    $vmPath = "c:\vm\$vmName"
    $vhdPath = "$vmPath\vhd\$vmType.vhdx"
    
    Write-Host "Creating $vmName..." -ForegroundColor Green
    
    # Create directories
    New-Item -Path "$vmPath\vhd" -ItemType Directory -Force | Out-Null
    
    # Create VHD
    try {
        New-VHD -Path $vhdPath -ParentPath "C:\vm\parent\parent.vhdx" -Differencing -ErrorAction Stop
        Write-Host "  VHD created: $vhdPath" -ForegroundColor Green
    }
    catch {
        Write-Host "  Error creating VHD: $_" -ForegroundColor Red
        return $false
    }
    
    # Create VM
    try {
        New-VM -Name $vmName -Path $vmPath -VHDPath $vhdPath -Generation 2 -ErrorAction Stop
        Write-Host "  Virtual machine created: $vmName" -ForegroundColor Green
    }
    catch {
        Write-Host "  Error creating VM: $_" -ForegroundColor Red
        return $false
    }
    
    # Configure processor and memory
    try {
        Set-VMProcessor -VMName $vmName -Count $cpuCount -ErrorAction Stop
        Set-VMMemory -VMName $vmName -StartupBytes ($memoryGB * 1GB) -ErrorAction Stop
        Write-Host "  Processor ($cpuCount cores) and memory (${memoryGB}GB) settings applied" -ForegroundColor Green
    }
    catch {
        Write-Host "  Error configuring VM: $_" -ForegroundColor Red
        return $false
    }
    
    return $true
}

# Main script
Write-Host "`n=== Creating Test Environment ===" -ForegroundColor Cyan
Write-Host "Enter 'X' at any stage to exit the script" -ForegroundColor Red

# 1. Request environment name
Write-Host "`n=== STEP 1: Environment Name ===" -ForegroundColor Yellow
Write-Host "Enter X to exit the script" -ForegroundColor Red
$name = Read-Host "Enter environment name (e.g., mail)"

if ($name -eq "X" -or $name -eq "x") {
    Write-Host "Script terminated by user." -ForegroundColor Red
    exit
}

Write-Host "Machines will be named: $name.dc1, $name.exchange1, etc." -ForegroundColor Yellow

# 2. Request number of Domain Controllers
Write-Host "`n=== STEP 2: Domain Controllers ===" -ForegroundColor Yellow
$dcCount = Get-NumberInput -Prompt "How many domain controllers are needed?" -Min 1 -Max 2

# Create Domain Controllers
if ($dcCount -gt 0) {
    Write-Host "`n=== Creating Domain Controllers ===" -ForegroundColor Cyan
    $successCount = 0
    for ($i = 1; $i -le $dcCount; $i++) {
        if (Create-VirtualMachine -vmName "$name.dc$i" -vmType "dc$i" -cpuCount 2 -memoryGB 8) {
            $successCount++
        }
    }
    Write-Host "Successfully created domain controllers: $successCount out of $dcCount" -ForegroundColor $(if ($successCount -eq $dcCount) {"Green"} else {"Yellow"})
}

# 3. Request number of Exchange servers
Write-Host "`n=== STEP 3: Exchange Servers ===" -ForegroundColor Yellow
$exchangeCount = Get-NumberInput -Prompt "How many Exchange servers are needed?" -Min 0 -Max 2

# Create Exchange servers
if ($exchangeCount -gt 0) {
    Write-Host "`n=== Creating Exchange Servers ===" -ForegroundColor Cyan
    $successCount = 0
    for ($i = 1; $i -le $exchangeCount; $i++) {
        if (Create-VirtualMachine -vmName "$name.exchange$i" -vmType "exchange$i" -cpuCount 8 -memoryGB 32) {
            $successCount++
        }
    }
    Write-Host "Successfully created Exchange servers: $successCount out of $exchangeCount" -ForegroundColor $(if ($successCount -eq $exchangeCount) {"Green"} else {"Yellow"})
}

# 4. Request number of SQL servers
Write-Host "`n=== STEP 4: SQL Servers ===" -ForegroundColor Yellow
$sqlCount = Get-NumberInput -Prompt "How many SQL servers are needed?" -Min 0 -Max 2

# Create SQL servers
if ($sqlCount -gt 0) {
    Write-Host "`n=== Creating SQL Servers ===" -ForegroundColor Cyan
    $successCount = 0
    for ($i = 1; $i -le $sqlCount; $i++) {
        if (Create-VirtualMachine -vmName "$name.sql$i" -vmType "sql$i" -cpuCount 8 -memoryGB 32) {
            $successCount++
        }
    }
    Write-Host "Successfully created SQL servers: $successCount out of $sqlCount" -ForegroundColor $(if ($successCount -eq $sqlCount) {"Green"} else {"Yellow"})
}

# 5. Request number of regular servers
Write-Host "`n=== STEP 5: Regular Servers ===" -ForegroundColor Yellow
$serverCount = Get-NumberInput -Prompt "How many regular servers are needed?" -Min 0 -Max 2

# Create regular servers
if ($serverCount -gt 0) {
    Write-Host "`n=== Creating Regular Servers ===" -ForegroundColor Cyan
    $successCount = 0
    for ($i = 1; $i -le $serverCount; $i++) {
        if (Create-VirtualMachine -vmName "$name.server$i" -vmType "server$i" -cpuCount 8 -memoryGB 32) {
            $successCount++
        }
    }
    Write-Host "Successfully created regular servers: $successCount out of $serverCount" -ForegroundColor $(if ($successCount -eq $serverCount) {"Green"} else {"Yellow"})
}

# Final report
Write-Host "`n=== FINAL SUMMARY ===" -ForegroundColor Cyan
Write-Host "Environment name: $name" -ForegroundColor Yellow
Write-Host "Domain Controllers: $dcCount" -ForegroundColor Green
Write-Host "Exchange Servers: $exchangeCount" -ForegroundColor Green
Write-Host "SQL Servers: $sqlCount" -ForegroundColor Green
Write-Host "Regular Servers: $serverCount" -ForegroundColor Green

$totalVMs = $dcCount + $exchangeCount + $sqlCount + $serverCount
Write-Host "`nTotal virtual machines to create: $totalVMs" -ForegroundColor Magenta

Write-Host "`nEnvironment creation completed!" -ForegroundColor Green
Write-Host "Script finished." -ForegroundColor Cyan