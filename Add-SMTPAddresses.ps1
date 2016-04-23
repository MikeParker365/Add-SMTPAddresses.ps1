<#
.SYNOPSIS
Add-SMTPAddresses.ps1 - Add SMTP addresses to Exchange On-Premises users for a new domain name

.DESCRIPTION 
This PowerShell script will add new SMTP addresses to existing Exchange mailbox users
for a new domain. This script fills the need to make bulk email address changes
in Exchange On-premises when Email Address Policies are not applied to a large number of users.

.OUTPUTS
Results are output to a text log file.

.PARAMETER Domain
The new domain name to add SMTP addresses to each Office 365 mailbox user.

.PARAMETER MakePrimary
Specifies that the new email address should be made the primary SMTP address for the mailbox user.

.PARAMETER Commit
Specifies that the changes should be committed to the mailboxes. Without this switch no changes
will be made to mailboxes but the changes that would be made are written to a log file for evaluation.

.EXAMPLE
.\Add-SMTPAddresses.ps1 -Domain office365bootcamp.com
This will perform a test pass for adding the new alias@office365bootcamp.com as a secondary email address
to all mailboxes. Use the log file to evaluate the outcome before you re-run with the -Commit switch.

.EXAMPLE
.\Add-SMTPAddresses.ps1 -Domain office365bootcamp.com -MakePrimary
This will perform a test pass for adding the new alias@office365bootcamp.com as a primary email address
to all mailboxes. Use the log file to evaluate the outcome before you re-run with the -Commit switch.

.EXAMPLE
.\Add-SMTPAddresses.ps1 -Domain office365bootcamp.com -MakePrimary -Commit
This will add the new alias@office365bootcamp.com as a primary email address
to all mailboxes.

.NOTES
Written by: Paul Cunningham

Find me on:

* My Blog:	http://paulcunningham.me
* Twitter:	https://twitter.com/paulcunningham
* LinkedIn:	http://au.linkedin.com/in/cunninghamp/
* Github:	https://github.com/cunninghamp

For more Exchange Server tips, tricks and news
check out Exchange Server Pro.

* Website:	http://exchangeserverpro.com
* Twitter:	http://twitter.com/exchservpro

Updated by: Mike Parker

    - My blog: http://mikeparker365.co.uk/
    - Twitter: https://twitter.com/MikeParker365
    - LinkedIn: https://uk.linkedin.com/in/mikeparkero365

Change Log
V1.00, 21/05/2015 - Initial version
V2.00, 24/02/2016 - Mike Parker Initial Update - Added progress bar and count to end of script.
                  - Added necessary filter to work on only Exchange on-prem mailboxes without Email Address Policies applied.  
#>

#requires -version 2

[CmdletBinding()]
param (
	
	[Parameter( Mandatory=$true )]
	[string]$Domain,
		
	[Parameter( Mandatory=$false )]
	[switch]$Csv,
    
	[Parameter( Mandatory=$false )]
    [switch]$Commit,

    [Parameter( Mandatory=$false )]
    [switch]$MakePrimary

	)

#...................................
# Variables
#...................................

$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path

$logfile = "$myDir\Add-SMTPAddresses.log"


#...................................
# Functions
#...................................

#This function is used to write the log file
Function Write-Logfile()
{
	param( $logentry )
	$timestamp = Get-Date -DisplayHint Time
	"$timestamp $logentry" | Out-File $logfile -Append
}


#...................................
# Script
#...................................

$start = Get-Date
Write-Logfile "Script started at $start";

#Check if new domain exists in Exchange org.

$chkdom = Get-AcceptedDomain $domain

if (!($chkdom))
{
    Write-Warning "You must add the new domain name to your Office 365 tenant first."
    EXIT
}

#Get the list of mailboxes not using the email address policy in the Exchange org.

If($csv){
	Write-Logfile "Loading users to process from CSV..."

	$csvUsers = Import-Csv $Csv

	$Mailboxes = Foreach($user in $csvUsers){
		Get-Mailbox $user.EmailAddress -EA:SilentlyContinue
	}

	# Check the number of users matches the CSV
	If($Mailboxes.Count -eq $csvUsers.Count){
		Write-Logfile "Users collected successfully"

	}
	Else{
		Write-Logfile "Could not collecte all users. Please review the CSV file before continuing..."
		Break
	}
}

else{
	Write-Logfile "Processing all users without Email Address Policy in Exchange Org..."
	$Mailboxes = @(Get-Mailbox -Filter {(EmailAddressPolicyEnabled -eq $False) -and (EmailAddresses -notlike '*$tenantName')} -ResultSize Unlimited)
	}

#Set up variables for progress bar and end output
$itemCount = $mailboxes.Count
$processedCount = 1
$success = 0
$failure = 0
 
Foreach ($Mailbox in $Mailboxes)
{
    $error.clear()

   	Write-Progress -Activity "Processing.." -Status "User $processedCount of $itemCount" -PercentComplete ($processedCount / $itemCount * 100)

    try{

    Write-Host "******* Processing: $mailbox"
    Write-Logfile "******* Processing: $mailbox"

    $NewAddress = $null

    #If -MakePrimary is used the new address is made the primary SMTP address.
    #Otherwise it is only added as a secondary email address.
    if ($MakePrimary)
    {
        $NewAddress = "SMTP:" + $Mailbox.Alias + "@$Domain"
    }
    else
    {
        $NewAddress = "smtp:" + $Mailbox.Alias + "@$Domain"
    }

    #Write the current email addresses for the mailbox to the log file
    Write-Logfile ""
    Write-Logfile "Current addresses:"
    
    $addresses = @($mailbox | Select -Expand EmailAddresses)

    foreach ($address in $addresses)
    {
        Write-Logfile $address
    }

    #If -MakePrimary is used the existing primary is changed to a secondary
    if ($MakePrimary)
    {
        Write-LogFile ""
        Write-Logfile "Converting current primary address to secondary"
        $addresses = $addresses.Replace("SMTP","smtp")
    }

    #Add the new email address to the list of addresses
    Write-Logfile ""
    Write-Logfile "New email address to add is $newaddress"

    $addresses += $NewAddress

    #You must use the -Commit switch for the script to make any changes
    if ($Commit)
    {
        Write-LogFile ""
        Write-LogFile "Committing new addresses:"
        foreach ($address in $addresses)
        {
            Write-Logfile $address
        }
        Set-Mailbox -Identity $Mailbox.Alias -EmailAddresses $addresses      
    }
    else
    {
        Write-LogFile ""
        Write-LogFile "New addresses:"
        foreach ($address in $addresses)
        {
            Write-Logfile $address
        }
        Write-LogFile "Changes not committed, re-run the script with the -Commit switch when you're ready to apply the changes."
        Write-Warning "No changes made due to -Commit switch not being specified."
    }

    Write-Logfile ""
    }
    catch{
    Write-Logfile "There was an error processing $Mailbox.Alias. Please review the log."

    }
    finally{
        if(!$error){
            $success++
            }
        else{
            $failure++
            }
    }
}

Write-Logfile "$ItemCount records processed"
Write-Logfile "$success records processed successfully."
Write-Logfile "$failure records errored during processing." 

$end = Get-Date;
Write-Logfile "Script ended at $end";

$diff = New-TimeSpan -Start $start -End $end
Write-Logfile "Time taken $($diff.Hours)h : $($diff.Minutes)m : $($diff.Seconds)s ";

#...................................
# Finished
#...................................