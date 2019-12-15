[Threading.Thread]::CurrentThread.CurrentUICulture = 'en-US';

#set-executionpolicy remotesigned
$azureAccountName ="groupsyncserviceuser@weitblicker.org"
$azurePassword = ConvertTo-SecureString "xxx" -AsPlainText -Force

$UserCredential = New-Object System.Management.Automation.PSCredential($azureAccountName, $azurePassword)

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session


$Logfile = "C:\Logs\$((Get-Date).toString("yyyy_MM_dd_HH_mm_ss")).log"
Function TextLogger
{
   Param ([string]$logstring)
   #Write-Host $logstring
   Add-content $Logfile -value $logstring
}

# Group domain 
$domain = "weitblicker.org"

$syncingGroups=@{ mitgliedsgruppe= "mitglieder";}

# syncingCities
$cityNames = "hamburg","bayreuth","berlin", "bochum", "bonn", "duisburg-essen", "freiburg", "goettingen", "hannover", "heidelberg", "kiel", "koeln", "leipzig", "marburg", "muenchen", "muenster", "osnabrueck", "plus"

TextLogger "Start syncing all mitgliedsgruppe o365 to mitglieder security groups"
Foreach ($k in $syncingGroups.Keys)
{
    
    $source = $k
    $target = $syncingGroups[$k] 
    TextLogger ("Start syning from " + $source + " to " + $target)
    
    $UnifiedGroupFilter = 'PrimarySmtpAddress -Like "' + $source + '.*"'

    $O365BoardGroups = Get-UnifiedGroup -Filter $UnifiedGroupFilter
    Foreach ($O365BoardGroup in $O365BoardGroups)
    {
        
        $cityRegex = $O365BoardGroup.PrimarySmtpAddress -match ($source + '.(.+(?=@))@' + $domain)
        if ($cityRegex) {
            $city = $matches[1]
            if(!$cityNames.Contains($city)) {
                TextLogger ("Skipping city " + $city)
                continue;
            }
            TextLogger ("Get members of group: " + $O365BoardGroup.PrimarySmtpAddress)
            $members = Get-UnifiedGroupLinks –Identity $O365BoardGroup.Identity –LinkType Members 

            $targetSecurityMemberGroup = $target + "." + $city + "@" + $domain
            TextLogger ("Add users to: " + $targetSecurityMemberGroup)
            #$currentMembers = Get-DistributionGroupMember $targetSecurityMemberGroup | Select-Object -ExpandProperty Name | Sort-Object



            Foreach ($O365BoardGroupMember in $members)
            {
                try 
				{
                    $ErrorActionPreference = "Stop"; #Make all errors terminating
                    $memberMail = $O365BoardGroupMember.Identity + "@" + $domain
                    Add-DistributionGroupMember -Identity $targetSecurityMemberGroup -Member $memberMail -BypassSecurityGroupManagerCheck
                    TextLogger ("Added user " + $O365BoardGroupMember.Identity + " to: " +$targetSecurityMemberGroup)                  
                }
                catch 
				{
                    if($_.FullyQualifiedErrorId -match 'AlreadyExists')
                    {
                      TextLogger ("User " + $O365BoardGroupMember.Identity + " is already in: " + $targetSecurityMemberGroup)
                      }
                    elseIf($_.FullyQualifiedErrorId -match 'Cmdlet-ManagementObjectNotFoundException') 
                    {
                        TextLogger ("User " + $O365BoardGroupMember.Identity + " could not be find for group: " + $targetSecurityMemberGroup)
                    }
                      else
                      {
                          throw $_
                      }
                  }finally{
                      $ErrorActionPreference = "Continue"; #Reset the error action pref to default
                }
            }
        }
    }
}

#
# Sync Vorstands security group member to vorstandsgruppe o365 as member
#

TextLogger ("Start syncing all Vorstand security group to vorstandsgruppe o365")
$source = "vorstand"
$target = "vorstandsgruppe"
TextLogger ("Start syncing from " + $source +" to " +$target)

foreach($city in $cityNames){
    # set up mails
    $securityGroupMail = $source + "." + $city + "@" + $domain
    $o365GroupMail = $target + "." + $city + "@" + $domain

    TextLogger ("Syncing members of security group: " + $securityGroupMail + " to o365 group")

    # get members from distribution group and write to file
    $members = Get-DistributionGroupMember $securityGroupMail | Select-Object -ExpandProperty Name | Sort-Object
    $targetGroupMembers = Get-UnifiedGroupLinks –Identity $o365GroupMail –LinkType Members | Select-Object -ExpandProperty Name | Sort-Object
    Foreach ($targetGroupMember in $targetGroupMembers) {
        if($members -notcontains $targetGroupMember) {
            # user was removed since last sync
            if($targetGroupMember -like "bundesverband"){continue}  # skip bundesverband, as this is the owner and cann not be removed
            Remove-UnifiedGroupLinks -Identity $o365GroupMail -Links $targetGroupMember -LinkType Members -Confirm:$false
            TextLogger ("Removed " + $targetGroupMember + " from " + $o365GroupMail)
        }
    }
    Foreach ($member in $members) {
        if($targetGroupMembers -notcontains $member) {
            # user was added since last sync
            Add-UnifiedGroupLinks -Identity $o365GroupMail -Links $member -LinkType Members -Confirm:$false
            TextLogger ("Add " + $member + " to " + $o365GroupMail)
        }
    }
}

#
# Sync Admins and Vorstand to O356 group owners of their vorstandsgruppe and mitgliedsgruppe 
#

$targets = "mitgliedsgruppe"
foreach($target in $targets) {
    TextLogger ("Start syning from admin to " + $target + " as Owner")
    foreach($city in $cityNames){
        # set up names
        $group_name = $city + " " + $type
        $securityGroupMail = "admin." + $city + "@" + $domain
        $o365GroupMail = $target + "." + $city + "@" + $domain
        $securityGroupMailVorstand = "vorstand." + $city + "@" + $domain
         
        TextLogger ("Syncing members of security group: "+ $securityGroupMail + " and "+ $securityGroupMailVorstand + " to o365 group owner of " + $o365GroupMail)

        # get members from distribution group
        $members = Get-DistributionGroupMember $securityGroupMail | Select-Object -ExpandProperty Name | Sort-Object
        $vorstandMembers = Get-DistributionGroupMember $securityGroupMailVorstand | Select-Object -ExpandProperty Name | Sort-Object

        $currentGroupOwners = Get-UnifiedGroupLinks –Identity $o365GroupMail –LinkType Owner | Select-Object -ExpandProperty Name | Sort-Object
        
        $targetGroupOwners = @($members) + @($vorstandMembers)

        $targetGroupMembers = Get-UnifiedGroupLinks –Identity $o365GroupMail –LinkType Members | Select-Object -ExpandProperty Name | Sort-Object
        Foreach ($currentGroupOwner in $currentGroupOwners) {
            if($targetGroupOwners -notcontains $currentGroupOwner) {
                # user was removed since last sync
                if($currentGroupOwner -like "bundesverband"){continue}  # skip bundesverband, as this is the owner and cann not be removed
                Remove-UnifiedGroupLinks -Identity $o365GroupMail -Links $currentGroupOwner -LinkType Owner -Confirm:$false
                TextLogger ("Removed " + $currentGroupOwner +" as Owner from " +$o365GroupMail)
            }
        }
        Foreach ($member in $targetGroupOwners) {
            if($currentGroupOwners -notcontains $member) {
                # user was added since last sync
                Add-UnifiedGroupLinks -Identity $o365GroupMail -Links $member -LinkType Owner -Confirm:$false
                TextLogger ("Add " + $member + " as Owner to " + $o365GroupMail)
            }
        }
    }
}
 TextLogger ("Done")

Remove-PSSession $Session
