<#
.SYNOPSIS
    Exchange Mailbox Cleanup Tool
.DESCRIPTION
    This script searches for and removes emails from all Exchange mailboxes based on either subject or sender criteria.
    It provides a two-step process: first showing a preview of found emails, then performing deletion after confirmation.
    
    FEATURES:
    - Searches all mailboxes for emails matching criteria
    - Shows preview with estimated counts before deletion
    - Requires user confirmation before actual deletion
    - Optimized deletion only from mailboxes with found items
    - Progress bar during deletion phase
    
.PARAMETER None
    All parameters are collected interactively during script execution.
    
.EXAMPLE
    PS> .\Remove-EmailsFromAllMailboxes.ps1
    
    The script will prompt for:
    1. Exchange admin credentials
    2. Exchange server name
    3. Search criteria (Subject or Sender)
    4. Search term
    5. Confirmation before deletion
    
.NOTES
    Author: Aleksandr Serkin
    Version: 2.0
    Date: 14.01.2024
    
    REQUIRED PERMISSIONS:
    - Exchange Administrator role or Mailbox Search permission
    - Ability to run Search-Mailbox cmdlet
    
    CAUTION:
    - This script PERMANENTLY deletes emails
    - Always test in non-production environment first
    - Ensure you have proper backups
    - Verify search criteria before confirming deletion
    
.LINK
    https://docs.microsoft.com/en-us/powershell/module/exchange/search-mailbox
#>

#region Setup Exchange Session
Write-Host "=== Exchange Mailbox Cleanup Tool ===" -ForegroundColor Cyan
Write-Host "Enter credentials with Exchange Search Mailbox permissions" -ForegroundColor Yellow
$UserCredential = Get-Credential

$mailserver = Read-Host "Enter your Exchange server name"
Write-Host "Connecting to Exchange server: $mailserver..." -ForegroundColor Cyan

try {
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$mailserver/PowerShell/ -Authentication Kerberos -Credential $UserCredential -ErrorAction Stop
    Import-PSSession $Session -DisableNameChecking -ErrorAction Stop
    Write-Host "Successfully connected to Exchange" -ForegroundColor Green
}
catch {
    Write-Host "Failed to connect to Exchange: $_" -ForegroundColor Red
    exit 1
}
#endregion

#region Main Menu
Write-Host "`n=== SELECT SEARCH CRITERIA ===" -ForegroundColor Yellow
Write-Host "1. Search by Email Subject"
Write-Host "2. Search by Sender Address"

do {
    $choice = Read-Host "Enter option number (1-2)"
    
    switch ($choice) {
        "1" {
            # Search by Subject
            $subject = Read-Host "Enter email subject (partial or complete)"
            if (-not [string]::IsNullOrWhiteSpace($subject)) {
                Write-Host "Searching for emails with subject containing: '$subject'..." -ForegroundColor Cyan
                
                # Phase 1: Preview Results
                Write-Host "`n=== PREVIEW RESULTS ===" -ForegroundColor Yellow
                Write-Host "Scanning all mailboxes..." -ForegroundColor Cyan
                
                try {
                    $previewResults = Get-Mailbox -ResultSize Unlimited | Search-Mailbox -SearchQuery "subject:`"$subject*`"" -EstimateResultOnly -WarningAction SilentlyContinue -ErrorAction Stop
                    
                    # Filter results to show only mailboxes with found items
                    $mailboxesWithItems = $previewResults | Where-Object { $_.ResultItemsCount -ge 1 }
                    
                    Write-Host "Search completed." -ForegroundColor Green
                    
                    if ($mailboxesWithItems.Count -gt 0) {
                        Write-Host "`nFound emails in the following mailboxes:" -ForegroundColor Green
                        $mailboxesWithItems | Select-Object @{
                            Name = "DisplayName"
                            Expression = { ($_.Identity.ToString().Split('/')[-1]) }
                        }, ResultItemsCount | Format-Table -AutoSize
                        
                        $totalFound = ($mailboxesWithItems | Measure-Object -Property ResultItemsCount -Sum).Sum
                        Write-Host "Total emails found: $totalFound in $($mailboxesWithItems.Count) mailbox(es)" -ForegroundColor Green
                        
                        # Confirmation for deletion
                        Write-Host "`n=== CONFIRM DELETION ===" -ForegroundColor Red
                        Write-Host "WARNING: This will PERMANENTLY delete emails!" -ForegroundColor Red
                        $confirm = Read-Host "Do you want to delete these emails? (Yes/No)"
                        
                        if ($confirm -eq "Yes" -or $confirm -eq "yes" -or $confirm -eq "y" -or $confirm -eq "Y") {
                            Write-Host "`nStarting deletion process..." -ForegroundColor Cyan
                            
                            # Phase 2: Perform Deletion
                            $deleteResults = @()
                            $deletedCount = 0
                            $totalToDelete = $mailboxesWithItems.Count
                            
                            # Show progress during deletion
                            foreach ($mailbox in $mailboxesWithItems) {
                                $deletedCount++
                                Write-Progress -Activity "Deleting Emails" -Status "Processing mailboxes" `
                                    -PercentComplete (($deletedCount / $totalToDelete) * 100) `
                                    -CurrentOperation "$deletedCount of $totalToDelete"
                                
                                try {
                                    $result = Search-Mailbox -Identity $mailbox.Identity -SearchQuery "subject:`"$subject*`"" -DeleteContent -Force -WarningAction SilentlyContinue -ErrorAction Stop
                                    
                                    if ($result) {
                                        $deleteResults += @{
                                            DisplayName = ($mailbox.Identity.ToString().Split('/')[-1])
                                            ResultItemsCount = $result.ResultItemsCount
                                        }
                                    }
                                }
                                catch {
                                    Write-Host "Error deleting from $($mailbox.Identity): $_" -ForegroundColor Yellow
                                }
                            }
                            
                            Write-Progress -Activity "Deleting Emails" -Completed
                            
                            if ($deleteResults.Count -gt 0) {
                                Write-Host "`n=== DELETION RESULTS ===" -ForegroundColor Green
                                Write-Host "Successfully deleted emails from mailboxes:" -ForegroundColor Green
                                $deleteResults | Select-Object DisplayName, ResultItemsCount | Format-Table -AutoSize
                                
                                $totalDeleted = ($deleteResults | Measure-Object -Property ResultItemsCount -Sum).Sum
                                Write-Host "Total emails deleted: $totalDeleted from $($deleteResults.Count) mailbox(es)" -ForegroundColor Green
                            } else {
                                Write-Host "No emails were deleted." -ForegroundColor Yellow
                            }
                        } else {
                            Write-Host "Deletion cancelled by user." -ForegroundColor Yellow
                        }
                    } else {
                        Write-Host "No emails found with the specified subject." -ForegroundColor Yellow
                    }
                }
                catch {
                    Write-Host "Error during search: $_" -ForegroundColor Red
                }
                
                $validChoice = $true
            } else {
                Write-Host "Subject cannot be empty!" -ForegroundColor Red
                $validChoice = $false
            }
        }
        
        "2" {
            # Search by Sender
            $from = Read-Host "Enter sender email address (partial or complete)"
            if (-not [string]::IsNullOrWhiteSpace($from)) {
                Write-Host "Searching for emails from sender containing: '$from'..." -ForegroundColor Cyan
                
                # Phase 1: Preview Results
                Write-Host "`n=== PREVIEW RESULTS ===" -ForegroundColor Yellow
                Write-Host "Scanning all mailboxes..." -ForegroundColor Cyan
                
                try {
                    $previewResults = Get-Mailbox -ResultSize Unlimited | Search-Mailbox -SearchQuery "from:`"$from`"" -EstimateResultOnly -WarningAction SilentlyContinue -ErrorAction Stop
                    
                    # Filter results to show only mailboxes with found items
                    $mailboxesWithItems = $previewResults | Where-Object { $_.ResultItemsCount -ge 1 }
                    
                    Write-Host "Search completed." -ForegroundColor Green
                    
                    if ($mailboxesWithItems.Count -gt 0) {
                        Write-Host "`nFound emails in the following mailboxes:" -ForegroundColor Green
                        $mailboxesWithItems | Select-Object @{
                            Name = "DisplayName"
                            Expression = { ($_.Identity.ToString().Split('/')[-1]) }
                        }, ResultItemsCount | Format-Table -AutoSize
                        
                        $totalFound = ($mailboxesWithItems | Measure-Object -Property ResultItemsCount -Sum).Sum
                        Write-Host "Total emails found: $totalFound in $($mailboxesWithItems.Count) mailbox(es)" -ForegroundColor Green
                        
                        # Confirmation for deletion
                        Write-Host "`n=== CONFIRM DELETION ===" -ForegroundColor Red
                        Write-Host "WARNING: This will PERMANENTLY delete emails!" -ForegroundColor Red
                        $confirm = Read-Host "Do you want to delete these emails? (Yes/No)"
                        
                        if ($confirm -eq "Yes" -or $confirm -eq "yes" -or $confirm -eq "y" -or $confirm -eq "Y") {
                            Write-Host "`nStarting deletion process..." -ForegroundColor Cyan
                            
                            # Phase 2: Perform Deletion
                            $deleteResults = @()
                            $deletedCount = 0
                            $totalToDelete = $mailboxesWithItems.Count
                            
                            # Show progress during deletion
                            foreach ($mailbox in $mailboxesWithItems) {
                                $deletedCount++
                                Write-Progress -Activity "Deleting Emails" -Status "Processing mailboxes" `
                                    -PercentComplete (($deletedCount / $totalToDelete) * 100) `
                                    -CurrentOperation "$deletedCount of $totalToDelete"
                                
                                try {
                                    $result = Search-Mailbox -Identity $mailbox.Identity -SearchQuery "from:`"$from`"" -DeleteContent -Force -WarningAction SilentlyContinue -ErrorAction Stop
                                    
                                    if ($result) {
                                        $deleteResults += @{
                                            DisplayName = ($mailbox.Identity.ToString().Split('/')[-1])
                                            ResultItemsCount = $result.ResultItemsCount
                                        }
                                    }
                                }
                                catch {
                                    Write-Host "Error deleting from $($mailbox.Identity): $_" -ForegroundColor Yellow
                                }
                            }
                            
                            Write-Progress -Activity "Deleting Emails" -Completed
                            
                            if ($deleteResults.Count -gt 0) {
                                Write-Host "`n=== DELETION RESULTS ===" -ForegroundColor Green
                                Write-Host "Successfully deleted emails from mailboxes:" -ForegroundColor Green
                                $deleteResults | Select-Object DisplayName, ResultItemsCount | Format-Table -AutoSize
                                
                                $totalDeleted = ($deleteResults | Measure-Object -Property ResultItemsCount -Sum).Sum
                                Write-Host "Total emails deleted: $totalDeleted from $($deleteResults.Count) mailbox(es)" -ForegroundColor Green
                            } else {
                                Write-Host "No emails were deleted." -ForegroundColor Yellow
                            }
                        } else {
                            Write-Host "Deletion cancelled by user." -ForegroundColor Yellow
                        }
                    } else {
                        Write-Host "No emails found from the specified sender." -ForegroundColor Yellow
                    }
                }
                catch {
                    Write-Host "Error during search: $_" -ForegroundColor Red
                }
                
                $validChoice = $true
            } else {
                Write-Host "Sender address cannot be empty!" -ForegroundColor Red
                $validChoice = $false
            }
        }
        
        default {
            Write-Host "Invalid option! Please enter 1 or 2." -ForegroundColor Red
            $validChoice = $false
        }
    }
} while ($validChoice -ne $true)
#endregion

#region Cleanup
# Close Exchange session
if ($Session) {
    try {
        Remove-PSSession $Session
        Write-Host "`nExchange session closed successfully." -ForegroundColor Green
    }
    catch {
        Write-Host "Warning: Could not close Exchange session properly." -ForegroundColor Yellow
    }
}

Write-Host "`n=== Script completed ===" -ForegroundColor Cyan
#endregion