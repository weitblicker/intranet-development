# Powershell

# Install Modules

Install-Module MicrosoftTeams
Install-Module AzureAD
Install-Module -Name Microsoft.Online.SharePoint.PowerShell

# Connect to Exchange Online Powershell

$UserCredential = Get-Credential

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection

Set-ExecutionPolicy Unrestricted

Import-PSSession $Session

... do stuff ...

Remove-PSSession $Session

Set-ExecutionPolicy Restricted


# Teams

## Create new Team

Connect-AzureAD -Credential $UserCredential
Connect-MsolService -Credential $UserCredential
Connect-MicrosoftTeams  -Credential $UserCredential

... do stuff ...

Disconnect-MicrosoftTeams
Disconnect-AzureAD

## Find existing Team

$group = Get-AzureADGroup -SearchString <mailadresse>

$group = Get-MsolGroup -SearchString <mailadresse>

## Create / Edit Team

New-Team -group $group.ObjectID

Set-Team -group $group.ObjectID



# Sharepoint

Connect-SPOService -Url https://weitblicker-admin.sharepoint.com -Credential $UserCredential

## Delete Sharepoint sited from "recycle bin"

Get-SPODeletedSite
https://docs.microsoft.com/en-us/powershell/module/sharepoint-online/get-spodeletedsite?view=sharepoint-ps

Remove-SPODeletedSite -Identity <url>
https://docs.microsoft.com/en-us/powershell/module/sharepoint-online/remove-spodeletedsite?view=sharepoint-ps



# Office 365 Group = UnifiedGroup

Overview: https://support.office.com/en-us/article/manage-office-365-groups-with-powershell-aeb669aa-1770-4537-9de2-a82ac11b0540
New: https://docs.microsoft.com/en-us/powershell/module/exchange/users-and-groups/new-unifiedgroup?view=exchange-ps
Set: https://docs.microsoft.com/en-us/powershell/module/exchange/users-and-groups/Set-UnifiedGroup?view=exchange-ps
Get: https://docs.microsoft.com/en-us/powershell/module/exchange/mailboxes/get-mailbox?view=exchange-ps

## Add new Office 365 Group
New-UnifiedGroup -DisplayName "Kronos" -PrimarySmtpAddress "kronos@weitblicker.org" -Owner "bundesverband@weitblicker.org" -AccessType Private

## Set Primary Email-Address
Set-UnifiedGroup -Identity "Kronos" -PrimarySmtpAddress "intranet.kronos@weitblicker.org"

## Remove Alias Email-Address
Set-UnifiedGroup -Identity "intranet.kronos@weitblicker.org" -EmailAddress @{remove='kronos@weitblicker.onmicrosoft.com'}

## Set Alias Email-Address
Set-UnifiedGroup -Identity "intranet.kronos@weitblicker.org" -EmailAddress @{add='intranet.kronos@weitblicker.onmicrosoft.com'}

## Read Email Addresses
Get-UnifiedGroup -Identity "intranet.kronos@weitblicker.org" | FL EmailAddresses

## Delete Group from "recycle bin"

https://answers.microsoft.com/en-us/msoffice/forum/all/deleting-o365-group/2aec6a5e-35ce-4180-b44c-f85091150db1

Get-AzureADMSDeletedGroup

Remove-AzureADMSDeletedDirectoryObject –Id <objectId>



# E-Mail-aktivierte Sicherheitsgruppe = DistributionGroup

New: https://docs.microsoft.com/en-us/powershell/module/exchange/users-and-groups/new-distributiongroup?view=exchange-ps
Get: https://docs.microsoft.com/en-us/powershell/module/exchange/users-and-groups/get-distributiongroup?view=exchange-ps
Set: https://docs.microsoft.com/en-us/powershell/module/exchange/users-and-groups/Set-DistributionGroup?view=exchange-ps


## Anlegen
New-DistributionGroup -Name "Kronos Vorstand" -Type "Security" -PrimarySmtpAddress "vorstand.kronos@weitblicker.org" -ManagedBy "bundesverband@weitblicker.org"


## Add Membership in other distributiongroup
Add-DistributionGroupMember -Identity godfather@weitblicker.org -BypassSecurityGroupManagerCheck -Member mitglieder.kronos@weitblicker.org

## Remove Membership in other distributiongroup
Remove-DistributionGroupMember -Identity godfather@weitblicker.org -BypassSecurityGroupManagerCheck -Member mitglieder.kronos@weitblicker.org


# Freigegebene Postfächer = Mailbox

New: https://docs.microsoft.com/en-us/powershell/module/exchange/mailboxes/new-mailbox?view=exchange-ps
Get: https://docs.microsoft.com/en-us/powershell/module/exchange/mailboxes/get-mailbox?view=exchange-ps
Set: https://docs.microsoft.com/en-us/powershell/module/exchange/mailboxes/set-mailbox?view=exchange-ps
Add-MailboxPermission: https://docs.microsoft.com/en-us/powershell/module/exchange/mailboxes/add-mailboxpermission?view=exchange-ps

## Anlegen
New-Mailbox -Shared -Name "Kronos" -PrimarySmtpAddress "kronos@weitblicker.org";

Set-Mailbox -Identity "Kronos" -GrantSendOnBehalfTo "kronos vorstand";

Add-MailboxPermission -Identity "Kronos" -User "kronos vorstand" -AccessRights FullAccess -InheritanceType All

## Liste aller (nicht endgültig) gelöschten Postfächer

Get-Mailbox -SoftDeletedMailbox -ResultSize Unlimited | Sort-Object -Property Name


## Wiederherstellen in ein existierendes Postfach

New-MailboxRestoreRequest -SourceMailbox <mailbox id> -TargetMailbox <target> -AllowLegacyDNMismatch
- https://docs.microsoft.com/en-us/exchange/recipients/disconnected-mailboxes/restore-deleted-mailboxes?view=exchserver-2019
- https://docs.microsoft.com/en-us/powershell/module/exchange/mailboxes/new-mailboxrestorerequest?view=exchange-ps

Get-MailboxRestoreRequest


## ? = DynamicDistributionGroup

New: https://docs.microsoft.com/en-us/powershell/module/exchange/users-and-groups/new-dynamicdistributiongroup?view=exchange-ps
Set: https://docs.microsoft.com/en-us/powershell/module/exchange/users-and-groups/set-dynamicdistributiongroup?view=exchange-ps

# Helpers

## To Title Case
Get-Culture).TextInfo.ToTitleCase($city)
