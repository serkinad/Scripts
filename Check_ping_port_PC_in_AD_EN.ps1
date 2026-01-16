<#
.SYNOPSIS
    Checks computer availability and port status in the network with results export to CSV.

.DESCRIPTION
    This script performs network connectivity checks (ping) and verifies the status of specified ports.
    Supports multiple methods for obtaining computer lists: Active Directory search, OU-based search, or manual input.
    Results are saved to a CSV file with detailed statistics.

.PARAMETER None
    The script operates in interactive mode without parameters.

.EXAMPLE
    .\PortChecker.ps1
    Launches the interactive menu for checking computer and port availability.

.NOTES
    Author: Administrator
    Version: 2.0
    Created: 2024
    Requirements: PowerShell 5.1+, ActiveDirectory module, AD read permissions

    Features:
    1. Multiple computer list sources support
    2. Ping check (always performed)
    3. Single or multiple port checks
    4. Parallel processing of up to 10 computers simultaneously
    5. Results export to CSV with ";" delimiter
    6. Automatic directory creation for saving
    7. Path variable support ($PSScriptRoot, $home, $desktop)
    8. Detailed statistics on check results

.LINK
    https://learn.microsoft.com/powershell/
    https://learn.microsoft.com/windows-server/identity/ad-ds/get-started/ad-ds-introduction
#>

[CmdletBinding()]
param()

function Show-Menu {
    <#
    .SYNOPSIS
        Displays the main script menu.
    
    .DESCRIPTION
        Clears the console and displays the script title.
    #>
    Clear-Host
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host "     PORT AND PING AVAILABILITY CHECKER" -ForegroundColor Yellow
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host ""
}

function Get-ComputersByCriteria {
    <#
    .SYNOPSIS
        Retrieves a list of computers based on selected criteria.
    
    .DESCRIPTION
        Depending on the selected criterion, retrieves a list of computers:
        - By partial name from Active Directory
        - By Organizational Unit (OU)
        - Manual input separated by commas
    
    .PARAMETER Choice
        Computer search criteria:
        - "NamePart" - search by partial name
        - "OU" - search by OU
        - "Manual" - manual input
    
    .OUTPUTS
        System.Array
        Array of strings containing computer names.
    
    .EXAMPLE
        Get-ComputersByCriteria -Choice "NamePart"
        Requests a partial name and returns a list of computers from AD.
    #>
    param(
        [ValidateSet("NamePart", "OU", "Manual", "Back")]
        [string]$Choice
    )
    
    $computers = @()
    
    switch ($Choice) {
        "NamePart" {
            $partName = Read-Host "Enter partial computer name"
            if (-not [string]::IsNullOrWhiteSpace($partName)) {
                try {
                    $computers = Get-ADComputer -LDAPFilter "(cn=*$partName*)" -ErrorAction Stop | 
                                Select-Object -ExpandProperty Name
                }
                catch {
                    Write-Host "Error searching for computers: $_" -ForegroundColor Red
                    Read-Host "Press Enter to continue"
                }
            }
        }
        
        "OU" {
            $OUPath = Read-Host "Enter OU path (e.g., ou=Computers,dc=domain,dc=local)"
            if (-not [string]::IsNullOrWhiteSpace($OUPath)) {
                try {
                    $computers = Get-ADComputer -SearchBase $OUPath -Filter * -ErrorAction Stop | 
                                Select-Object -ExpandProperty Name
                }
                catch {
                    Write-Host "Error searching for computers: $_" -ForegroundColor Red
                    Read-Host "Press Enter to continue"
                }
            }
        }
        
        "Manual" {
            $manualInput = Read-Host "Enter computer names separated by commas"
            if (-not [string]::IsNullOrWhiteSpace($manualInput)) {
                $computers = $manualInput.Split(',').Trim() | Where-Object { $_ -ne '' }
            }
        }
    }
    
    return $computers
}

function Test-Port {
    <#
    .SYNOPSIS
        Checks TCP port availability on the specified computer.
    
    .DESCRIPTION
        Attempts to establish a TCP connection to the specified port on the computer.
        Uses asynchronous connection with timeout.
    
    .PARAMETER comp
        Computer name or IP address to check.
    
    .PARAMETER port
        TCP port number to check.
    
    .PARAMETER timeout
        Connection timeout in milliseconds (default: 200 ms).
    
    .OUTPUTS
        System.Boolean
        $true - port is open, $false - port is closed or unavailable.
    
    .EXAMPLE
        Test-Port -comp "server01" -port 3389
        Checks availability of port 3389 on server01.
    #>
    param(
        [string]$comp,
        [int]$port,
        [int]$timeout = 200
    )
    
    $client = $null
    try {
        $client = New-Object System.Net.Sockets.TcpClient
        $task = $client.ConnectAsync($comp, $port)
        $completed = $task.Wait($timeout)
        return $completed -and $client.Connected
    }
    catch { 
        return $false 
    }
    finally { 
        if ($client) { 
            $client.Dispose() 
        } 
    }
}

function Test-Ping {
    <#
    .SYNOPSIS
        Checks computer availability using ICMP ping.
    
    .DESCRIPTION
        Sends an ICMP echo request to the specified computer and checks for a response.
    
    .PARAMETER comp
        Computer name or IP address to check.
    
    .PARAMETER count
        Number of ping requests (default: 1).
    
    .PARAMETER timeout
        Response timeout in milliseconds (default: 1000 ms).
    
    .OUTPUTS
        System.Boolean
        $true - computer responds to ping, $false - no response.
    
    .EXAMPLE
        Test-Ping -comp "server01"
        Checks availability of server01 via ping.
    #>
    param(
        [string]$comp,
        [int]$count = 1,
        [int]$timeout = 1000
    )
    
    try {
        $ping = New-Object System.Net.NetworkInformation.Ping
        $reply = $ping.Send($comp, $timeout)
        return $reply.Status -eq [System.Net.NetworkInformation.IPStatus]::Success
    }
    catch {
        return $false
    }
}

function Read-PathWithTabCompletion {
    <#
    .SYNOPSIS
        Prompts for a path with autocompletion and variable support.
    
    .DESCRIPTION
        Allows the user to enter a path with Tab autocompletion support
        and predefined variable usage.
    
    .PARAMETER Prompt
        Input prompt text.
    
    .PARAMETER DefaultPath
        Default path to use if input is empty.
    
    .OUTPUTS
        System.String
        Full file path with .csv extension.
    
    .EXAMPLE
        Read-PathWithTabCompletion -Prompt "Enter path" -DefaultPath "C:\default.csv"
        Prompts the user for a path with a hint and default value.
    
    .NOTES
        Supported variables:
        - $PSScriptRoot - script directory
        - $home - user's home directory
        - $desktop - user's desktop
    #>
    param(
        [string]$Prompt = "Enter path",
        [string]$DefaultPath
    )
    
    # Use ReadLine for autocompletion support
    # Check if PSReadLine module is available
    if (Get-Module -Name PSReadLine -ErrorAction SilentlyContinue) {
        Write-Host $Prompt -ForegroundColor Yellow
        if (-not [string]::IsNullOrWhiteSpace($DefaultPath)) {
            Write-Host "Default: $DefaultPath" -ForegroundColor Gray
            Write-Host "Press Enter to use default path" -ForegroundColor Gray
        }
        Write-Host "Use Tab for autocompletion" -ForegroundColor Gray
        Write-Host "Path: " -NoNewline -ForegroundColor White
        $userInput = [Microsoft.PowerShell.PSConsoleReadLine]::ReadLine()
        
        # If input is empty, return default path
        if ([string]::IsNullOrWhiteSpace($userInput)) {
            return $DefaultPath
        }
        
        return $userInput
    }
    else {
        # Alternative without PSReadLine
        Write-Host $Prompt -ForegroundColor Yellow
        if (-not [string]::IsNullOrWhiteSpace($DefaultPath)) {
            Write-Host "Default: $DefaultPath" -ForegroundColor Gray
        }
        Write-Host "Available variables: `$PSScriptRoot, `$home, `$desktop" -ForegroundColor Gray
        Write-Host "Example: `$home\Documents\results.csv" -ForegroundColor Gray
        $userInput = Read-Host "Enter path (or press Enter to use default path)"
        
        # If input is empty, return default path
        if ([string]::IsNullOrWhiteSpace($userInput)) {
            return $DefaultPath
        }
        
        # Process special variables
        $userInput = $userInput.Replace('$PSScriptRoot', $PSScriptRoot)
        $userInput = $userInput.Replace('$home', $HOME)
        $userInput = $userInput.Replace('$desktop', [Environment]::GetFolderPath('Desktop'))
        $userInput = $userInput.Replace('`$PSScriptRoot', $PSScriptRoot)
        $userInput = $userInput.Replace('`$home', $HOME)
        $userInput = $userInput.Replace('`$desktop', [Environment]::GetFolderPath('Desktop'))
        
        # If path is entered but no .csv extension, add it
        if (-not $userInput.EndsWith('.csv') -and -not [string]::IsNullOrWhiteSpace($userInput)) {
            $userInput += ".csv"
        }
        
        return $userInput
    }
}

# Main menu
do {
    Show-Menu
    
    # Get computer list
    $allComputers = @()
    
    # Show selection menu once
    Show-Menu
    Write-Host "SELECT COMPUTER SEARCH CRITERIA:" -ForegroundColor Green
    Write-Host "1. Full or partial computer name" -ForegroundColor White
    Write-Host "2. By OU" -ForegroundColor White
    Write-Host "3. Computer names separated by commas" -ForegroundColor White
    Write-Host "0. Exit" -ForegroundColor Red
    Write-Host ""
    
    $searchChoice = Read-Host "Select option (1-3 or 0)"
    
    switch ($searchChoice) {
        "1" {
            $foundComputers = Get-ComputersByCriteria -Choice "NamePart"
            if ($foundComputers.Count -gt 0) {
                $allComputers += $foundComputers
            }
        }
        
        "2" {
            $foundComputers = Get-ComputersByCriteria -Choice "OU"
            if ($foundComputers.Count -gt 0) {
                $allComputers += $foundComputers
            }
        }
        
        "3" {
            $foundComputers = Get-ComputersByCriteria -Choice "Manual"
            if ($foundComputers.Count -gt 0) {
                $allComputers += $foundComputers
            }
        }
        
        "0" {
            exit
        }
        
        default {
            Write-Host "Invalid choice!" -ForegroundColor Red
            Read-Host "Press Enter to continue"
            continue
        }
    }
    
    # Check that computer list is not empty
    if ($allComputers.Count -eq 0) {
        Write-Host "Computer list is empty! Need to add at least one computer." -ForegroundColor Red
        Write-Host ""
        
        do {
            Write-Host "Select action:" -ForegroundColor Yellow
            Write-Host "1. Return to main menu" -ForegroundColor White
            Write-Host "2. Exit" -ForegroundColor White
            Write-Host ""
            
            $emptyListChoice = Read-Host "Enter 1 or 2"
            
            switch ($emptyListChoice) {
                "1" {
                    # Return to start of main loop
                    $continueMain = $true
                    break
                }
                "2" {
                    exit
                }
                default {
                    Write-Host "Invalid choice! Enter 1 or 2." -ForegroundColor Red
                    $continueMain = $false
                }
            }
        } while (-not $continueMain)
        
        if ($continueMain) {
            continue
        }
    }
    
    # If we got here, there are computers to check
    Write-Host "Found computers: $($allComputers.Count)" -ForegroundColor Green
    Read-Host "Press Enter to continue"
    
    # Get ports to check (no default ports)
    Show-Menu
    Write-Host "Computers found for checking: $($allComputers.Count)" -ForegroundColor Green
    Write-Host ""
    $portsInput = Read-Host "Enter ports to check separated by commas (or leave empty for ping only)"
    
    $ports = @()
    if (-not [string]::IsNullOrWhiteSpace($portsInput)) {
        $ports = $portsInput.Split(',').Trim() | ForEach-Object { 
            if ($_ -match '^\d+$') { [int]$_ } 
        } | Where-Object { $_ -gt 0 -and $_ -lt 65536 }
        
        if ($ports.Count -gt 0) {
            Write-Host "Ports to be checked: $($ports -join ', ')" -ForegroundColor Yellow
        }
    }
    
    # Ping is always performed by default
    Write-Host "Ping will be performed for all computers" -ForegroundColor Yellow
    
    # Request save directory
    Write-Host ""
    $defaultPath = Join-Path $PSScriptRoot "PortCheck_Results_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    
    # Use enhanced function for path input
    $savePath = Read-PathWithTabCompletion -Prompt "Enter path to save CSV file" -DefaultPath $defaultPath
    
    # Check that path is not empty
    if ([string]::IsNullOrWhiteSpace($savePath)) {
        Write-Host "Path not specified. Using default path." -ForegroundColor Yellow
        $savePath = $defaultPath
    }
    
    Write-Host "File will be saved as: $savePath" -ForegroundColor Green
    
    # Check and create directory if needed
    $directory = Split-Path $savePath -Parent
    if (-not [string]::IsNullOrWhiteSpace($directory) -and -not (Test-Path $directory)) {
        try {
            New-Item -ItemType Directory -Path $directory -Force | Out-Null
            Write-Host "Directory created: $directory" -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to create directory: $_" -ForegroundColor Red
            Write-Host "Saving to default path: $defaultPath" -ForegroundColor Yellow
            $savePath = $defaultPath
        }
    }
    
    # Perform checks
    $results = @()
    $computerCount = $allComputers.Count
    $current = 0
    
    Write-Host ""
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host "     CHECKING STARTED" -ForegroundColor Yellow
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host "Starting check of $computerCount computers..." -ForegroundColor Yellow
    Write-Host "Ping is being performed for all computers" -ForegroundColor Yellow
    if ($ports.Count -gt 0) {
        Write-Host "Checking ports: $($ports -join ', ')" -ForegroundColor Yellow
    }
    Write-Host ""
    
    # Define functions inside ForEach-Object for visibility in parallel context
    $allComputers | ForEach-Object -Parallel {
        $comp = $_
        $ports = $using:ports
        
        # Define functions inside parallel block
        function Local-TestPing {
            param(
                [string]$comp,
                [int]$count = 1,
                [int]$timeout = 1000
            )
            
            try {
                $ping = New-Object System.Net.NetworkInformation.Ping
                $reply = $ping.Send($comp, $timeout)
                return $reply.Status -eq [System.Net.NetworkInformation.IPStatus]::Success
            }
            catch {
                return $false
            }
        }
        
        function Local-TestPort {
            param(
                [string]$comp,
                [int]$port,
                [int]$timeout = 200
            )
            
            $client = $null
            try {
                $client = New-Object System.Net.Sockets.TcpClient
                $task = $client.ConnectAsync($comp, $port)
                $completed = $task.Wait($timeout)
                return $completed -and $client.Connected
            }
            catch { 
                return $false 
            }
            finally { 
                if ($client) { 
                    $client.Dispose() 
                } 
            }
        }
        
        # Create result object
        $result = [PSCustomObject]@{
            ComputerName = $comp
            Ping = "NotTested"
            TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }
        
        # Add properties for ports if specified
        foreach ($port in $ports) {
            $result | Add-Member -NotePropertyName "Port_$port" -NotePropertyValue "NotTested" -Force
        }
        
        # Ping check (always performed)
        $pingResult = Local-TestPing -comp $comp
        $result.Ping = if ($pingResult) { "Success" } else { "Failed" }
        
        # Port check (only if specified)
        foreach ($port in $ports) {
            $portResult = Local-TestPort -comp $comp -port $port
            $result."Port_$port" = if ($portResult) { "Open" } else { "Closed" }
        }
        
        # Return result
        $result
        
    } -ThrottleLimit 10 | ForEach-Object {
        $results += $_
        $current++
        Write-Progress -Activity "Checking computers" -Status "Processed: $current of $computerCount" `
                      -PercentComplete (($current / $computerCount) * 100)
    }
    
    Write-Progress -Activity "Checking computers" -Completed
    
    # Save results
    try {
        # Check that path is not empty before saving
        if ([string]::IsNullOrWhiteSpace($savePath)) {
            $savePath = $defaultPath
            Write-Host "Path not specified. Using default path: $savePath" -ForegroundColor Yellow
        }
        
        $results | Export-Csv -Path $savePath -Encoding UTF8 -NoTypeInformation -Delimiter ';'
        Write-Host ""
        Write-Host "=========================================" -ForegroundColor Green
        Write-Host "     CHECK RESULTS" -ForegroundColor Yellow
        Write-Host "=========================================" -ForegroundColor Green
        Write-Host ""
        Write-Host "Results saved to: $savePath" -ForegroundColor Green
        Write-Host "Computers processed: $($results.Count)" -ForegroundColor Green
        
        # Show brief statistics
        $successPing = ($results | Where-Object { $_.Ping -eq "Success" }).Count
        $failedPing = $computerCount - $successPing
        Write-Host "Successful ping: $successPing of $($results.Count)" -ForegroundColor Green
        Write-Host "Failed ping: $failedPing of $($results.Count)" -ForegroundColor Red
        
        foreach ($port in $ports) {
            $openPorts = ($results | Where-Object { $_."Port_$port" -eq "Open" }).Count
            $closedPorts = $computerCount - $openPorts
            Write-Host "Port $port open: $openPorts of $($results.Count)" -ForegroundColor Green
            Write-Host "Port $port closed: $closedPorts of $($results.Count)" -ForegroundColor Red
        }
    }
    catch {
        Write-Host "Error saving file: $_" -ForegroundColor Red
        Write-Host "Attempting to save to default path..." -ForegroundColor Yellow
        try {
            $results | Export-Csv -Path $defaultPath -Encoding UTF8 -NoTypeInformation -Delimiter ';'
            Write-Host "Results saved to default path: $defaultPath" -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to save file even to default path" -ForegroundColor Red
            Write-Host "Displaying results in console:" -ForegroundColor Yellow
            $results | Format-Table -AutoSize
        }
    }
    
    Write-Host ""
    do {
        Write-Host "Select action:" -ForegroundColor Yellow
        Write-Host "1. Perform another check" -ForegroundColor White
        Write-Host "2. Exit" -ForegroundColor White
        Write-Host ""
        
        $finalChoice = Read-Host "Enter 1 or 2"
        
        switch ($finalChoice) {
            "1" {
                $continueMain = $true
                break
            }
            "2" {
                Write-Host "Script completed." -ForegroundColor Green
                exit
            }
            default {
                Write-Host "Invalid choice! Enter 1 or 2." -ForegroundColor Red
                $continueMain = $false
            }
        }
    } while (-not $continueMain)
    
} while ($continueMain)