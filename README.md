# Reset-MailboxAutoReply

````powershell

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
````
