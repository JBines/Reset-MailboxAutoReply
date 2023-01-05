<#
.SYNOPSIS
This script will reset the Cached Autoreply list allowing for a new autoreply message to be sent after 
the script is run.   
 
.DESCRIPTION
SCENARIO: By default Exchange Online only sends one auto reply per email address until the autoreply is reset. 
We created this script as some orgs have the requirement to allow one autoreply messages to be sent each day 
or week for selected mailbox users. This might be required for shared mailboxes or users who are on extended leave. 

Microsoft's recommended work around by using rules to send auto replies sends a message for each received email. 
On top of this it's difficult for users to setup and is often too 'noisy'. 

This script toggles as it were the on/off button to clear the auto reply cache and allow new auto reply messages 
to be sent to the same recipients.  

## Reset-MailboxAutoReply.ps1 [-EXOFilter <String>] [-ReportSkipped <Switch>] [-EXOManagedIdentity <Switch>] [-EXOOrganization <String>]

.PARAMETER EXOFilter
The EXOFilter parameter uses OPATH syntax to filter the results by the specified properties and values. The search 
criteria uses the syntax "Property -ComparisonOperator 'Value'". 

This parameter is not Mandatory but if not applied it will apply for to all mailboxes in the tenant. 

.PARAMETER ReportSkipped
The ReportSkipped parameter also logs output for mailboxes which were skipped because Auto Reply settings were not 
enabled. 

.PARAMETER EXOManagedIdentity
The EXOManagedIdentity switch specifies that you're using managed identity to connect. You don't need to specify a 
value with this switch.

You must use this switch with the Organization parameter.

Managed identity connections are currently supported for the following types of Azure resources:
       Azure Automation runbooks
       Azure Virtual Machines
       Azure Virtual Machine Scale Sets
       Azure Functions

.PARAMETER EXOOrganization
The Organization parameter specifies the organization when you connect using CBA or managed 
identity. You must use the primary .onmicrosoft.com domain of the organization for the value 
of this parameter.

.NOTES
Use rules to create an out of office message: 
       https://support.microsoft.com/en-us/office/use-rules-to-create-an-out-of-office-message-9f124e4a-749e-4288-a266-2d009686b403

Requires the v3 Exchange Online Module: 
       https://techcommunity.microsoft.com/t5/exchange-team-blog/exchange-online-powershell-v3-module-general-availability/ba-p/3632543?WT.mc_id=M365-MVP-9501 

Use Azure managed identities to connect to Exchange Online PowerShell: 
       https://learn.microsoft.com/en-us/powershell/exchange/connect-exo-powershell-managed-identity?view=exchange-ps

[AUTHOR]
Joshua Bines, Systems Engineer
 
[CONTRIBUTORS]
B Muller, Systems Engineer
 
Find me on:
* Web:     https://theinformationstore.com.au
* LinkedIn:  https://www.linkedin.com/in/joshua-bines-4451534
* Github:    https://github.com/jbines


[VERSION HISTORY / UPDATES]
0.0.1  - PRIVATE - Created the bare bones.
0.0.2  - PRIVATE - [Bug Fix] Fixed scope and add -filter command.
0.0.3 20190101 - JBines - Code Review. Added error handling and console logging. [Bug Fix] EndTime
0.0.4 20221223 - BMuller - [Bug Fix] Added missing 'Enabled' vs only 'Scheduled'. 
0.1.0 20230103 - JBines - [Feature] Enabled Managed Identies, support for Azure Automation, added custom filter and EXO connect checks.  

#>
[CmdletBinding(DefaultParametersetName='None')] 
Param 
(
       [Parameter(Mandatory = $False)]
       [ValidateNotNullOrEmpty()]
       [String]$EXOFilter,
       [Parameter(Mandatory = $False)]
       [ValidateNotNullOrEmpty()]
       [Switch]$ReportSkipped,
       [Parameter(Mandatory = $False)]
       [ValidateNotNullOrEmpty()]
       [Switch]$EXOManagedIdentity,
       [Parameter(Mandatory = $False)]
       [ValidateNotNullOrEmpty()]
       [String]$EXOOrganization

)

## Functions
# Logging function
# Author: Aaron Guilmette
 
function Write-Log([string[]]$Message, [string]$LogFile = $Script:LogFile, [switch]$ConsoleOutput, [ValidateSet("SUCCESS", "INFO", "WARN", "ERROR", "DEBUG")][string]$LogLevel)
{
       $Message = $Message + $Input
       If (!$LogLevel) { $LogLevel = "INFO" }
       switch ($LogLevel)
       {
              SUCCESS { $Color = "Green" }
              INFO { $Color = "White" }
              WARN { $Color = "Yellow" }
              ERROR { $Color = "Red" }
              DEBUG { $Color = "Gray" }
       }
       if ($Message -ne $null -and $Message.Length -gt 0)
       {
              $TimeStamp = [System.DateTime]::Now.ToString("yyyy-MM-dd HH:mm:ss")
              if ($LogFile -ne $null -and $LogFile -ne [System.String]::Empty)
              {
                     Out-File -Append -FilePath $LogFile -InputObject "[$TimeStamp] [$LogLevel] $Message"
              }
              if ($ConsoleOutput -eq $true)
              {
                     Write-Host "[$TimeStamp] [$LogLevel] :: $Message" -ForegroundColor $Color

                if($AutomationPSConnection -or $AutomationPSCertificate -or $EXOManagedIdentity)
                {
                     Write-Output "[$TimeStamp] [$LogLevel] :: $Message"
                }
              }
              if($LogLevel -eq "ERROR")
              {
                      Write-Error "[$TimeStamp] [$LogLevel] :: $Message"
              }  
   }
}#End function Write-Log

Function Test-CommandExists 
{

 Param ($command)

     $oldPreference = $ErrorActionPreference

     $ErrorActionPreference = 'stop'

     try {if(Get-Command $command){RETURN $true}}

     Catch {Write-Host "$command does not exist"; RETURN $false}

     Finally {$ErrorActionPreference=$oldPreference}

} #end function test-CommandExists

#If enabled with ManagedIdentity use it to connect.
if ($EXOManagedIdentity -and $EXOOrganization) {

       Connect-ExchangeOnline -ManagedIdentity -Organization $EXOOrganization

}
                        
#Check Access to Exchange Online and has all the required permissions!
If(Test-CommandExists Get-EXOMailbox,Get-MailboxAutoReplyConfiguration,Set-MailboxAutoReplyConfiguration){

       Write-Log -Message "Test-CommandExists: Access Confirmed" -LogLevel DEBUG -ConsoleOutput

} 
Else {Write-Log -Message "You are missing the PowerShell EXO CMDlets. Script requires Get-EXOMailbox,Get-MailboxAutoReplyConfiguration,Set-MailboxAutoReplyConfiguration" -LogLevel ERROR -ConsoleOutput; Break}

#Populate array with mailboxes
if ($EXOFilter) {
       $Mailboxes = Get-EXOMailbox -Filter $EXOFilter -ResultSize unlimited
       Write-Log -Message "Processing Mailboxes - Total Found: $($Mailboxes.count)" -LogLevel INFO -ConsoleOutput
}
else {
       $Mailboxes = Get-EXOMailbox -ResultSize unlimited
       Write-Log -Message "Processing Mailboxes - Total Found: $($Mailboxes.count)" -LogLevel INFO -ConsoleOutput
}

foreach ($Mailbox in $Mailboxes) {
       #Null Var for Looping errors
       $Start = $null
       $End = $null
       $MailboxUPN = $null
       $mailboxAutoReplyState = $null

       #Set Loop Var
       $mailboxUPN = $Mailbox.UserPrincipalName
       $mailboxAutoReplyState = (Get-MailboxAutoReplyConfiguration -Identity $MailboxUPN).AutoReplyState

       if (($mailboxAutoReplyState -eq "Scheduled") -or ($mailboxAutoReplyState -eq "Enabled")) {

              switch ($mailboxAutoReplyState) {
                     "Scheduled"  { 

                            $Start = (Get-MailboxAutoReplyConfiguration -Identity $MailboxUPN).StartTime
                            $End = (Get-MailboxAutoReplyConfiguration -Identity $MailboxUPN).EndTime
                            Set-MailboxAutoReplyConfiguration -Identity $MailboxUPN -AutoReplyState Disabled
                            Set-MailboxAutoReplyConfiguration -Identity $MailboxUPN -AutoReplyState scheduled -StartTime $Start -EndTime $End
                            
                            If($?){Write-Log -Message "CMDlet:Set-MailboxAutoReplyConfiguration;UPN:$MailboxUPN;State:Scheduled;StartTime:$Start;EndTime:$End" -LogLevel SUCCESS -ConsoleOutput}
                            Else{Write-Log -Message "CMDlet:Set-MailboxAutoReplyConfiguration;UPN:$MailboxUPN;State:Scheduled;StartTime:$Start;EndTime:$End" -LogLevel ERROR -ConsoleOutput}
              
                     }
                     "Enabled" { 

                            Set-MailboxAutoReplyConfiguration -Identity $MailboxUPN -AutoReplyState Disabled
                            Set-MailboxAutoReplyConfiguration -Identity $MailboxUPN -AutoReplyState Enabled
              
                            If($?){Write-Log -Message "CMDlet:Set-MailboxAutoReplyConfiguration;UPN:$MailboxUPN;State:Enabled" -LogLevel SUCCESS -ConsoleOutput}
                            Else{Write-Log -Message "CMDlet:Set-MailboxAutoReplyConfiguration;UPN:$MailboxUPN;State:Enabled" -LogLevel ERROR -ConsoleOutput}
                            
                     }
                     Default {
                            
                            If($?){Write-Log -Message "NO OPERATION;UPN:$MailboxUPN" -LogLevel ERROR -ConsoleOutput}
                     }
              }#End Switch
       }
       else {
              If($ReportSkipped){Write-Log -Message "SKIPPED - AutoReply:Disabled;UPN:$MailboxUPN" -LogLevel DEBUG -ConsoleOutput}
       }

}#End Foreach

#Script End... Yay!
