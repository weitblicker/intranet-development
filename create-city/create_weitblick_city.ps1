### Script to create new weitblick city // update existing to new group structure
###---------------------------------------------------------------------------

### Manual Init

# Execute this everytime when opening the script-editor
# Set-ExecutionPolicy -Scope Process -ExecutionPolicy Unrestricted -force

# Execute this when using the script on a computer for the first time
# Install-Module MicrosoftTeams
# Install-Module MSOnline
# Install-Module AzureAD

# Admin rights on weitblicker.org and portal.azure.com for weitblicker.org required

# WICHTIG: Zur Zeit sind keine Umlaute in Namen (O365) möglich, daher für alle Identifier IMMER Address, nicht Name verwenden

###---------------------------------------------------------------------------
### Start Session

write-host "INITIALISE" -ForegroundColor DarkYellow

try
{
    if ([string]::IsNullOrEmpty($UserCredential)) {
        $UserCredential = Get-Credential
    }

      $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Unrestricted -force
    Import-PSSession $Session -AllowClobber
}
catch
{
    write-host $_.Exception.Message -ForegroundColor DarkRed
    write-host "Login failed. Executed Powershell as admin?" -ForegroundColor Red
    $UserCredential=$NULL
    break
}

## get (new) city name
[string] $city=''; $city=(read-host "Enter the name of the Weitblick City you want to create").ToLower()

    #### @todo3: Check for blanks and other invalid signs
if ([string]::IsNullOrEmpty($city)){
    write-host "Empty or invalid City Input. Aborting."  -ForegroundColor DarkRed
    Remove-PSSession $Session
    break
}


###---------------------------------------------------------------------------
### FUNCTIONS

function Replace-Umlaute ([string]$s) {
    #source: https://www.datenteiler.de/powershell-umlaute-ersetzen/
    $UmlautObject = New-Object PSObject | Add-Member -MemberType NoteProperty -Name Name -Value $s -PassThru

    # hash tables are by default case insensitive
    # we have to create a new hash table object for case sensitivity

    $characterMap = New-Object system.collections.hashtable
    $characterMap.ä = "ae"
    $characterMap.ö = "oe"
    $characterMap.ü = "ue"
    $characterMap.ß = "ss"
    $characterMap.Ä = "Ae"
    $characterMap.Ü = "Ue"
    $characterMap.Ö = "Oe"


    foreach ($property  in 'Name') {
        foreach ($key in $characterMap.Keys) {
            $UmlautObject.$property = $UmlautObject.$property -creplace $key,$characterMap[$key]
        }
    }

    $UmlautObject.Name
}

###---------------------------------------------------------------------------
### Flags

# Namings for Mail and Groups (with/without special chars)
    #TEMP-BUGFIX - keine Umlaute in Office365
$cityUpperFull=(Get-Culture).TextInfo.ToTitleCase($city)									# Städtchen
$cityUpperSimple=Replace-Umlaute $city($cityUpperFull)										# Staedtchen
$cityLowerFull=$city																		# städtchen
$cityLowerSimple=Replace-Umlaute $city($cityLowerFull)										# staedtchen

#@todo3: Init-Fehler: "WARNUNG: Die Namen einiger importierter Befehle auf Modul "tmp_s252o0r4.fsq" enthalten nicht genehmigte Verben"

write-host "VARIABLE GENERATION" -ForegroundColor Darkyellow

[bool] $MailboxExists = $false

#Domain
$OMSCOM="@weitblicker.onmicrosoft.com"
$WBORG="@weitblicker.org"
$WBNGO="@weitblicker.ngo"
$WBONG="@weitblicker.ong"

### Mailbox
$MailboxName = $cityUpperFull																# Stadt
$MailboxAddress = $cityLowerSimple + "@weitblicker.org"										# stadt@weitblicker.org

# security groups
$SecurityGroupMemberName = $cityLowerFull + " mitglieder"									# stadt mitglieder
$SecurityGroupMemberAddress = "mitglieder." + $cityLowerSimple + $WBORG						# mitglieder.stadt@weitblicker.org
$SecurityGroupStaffName = $cityLowerFull + " vorstand"										# stadt vorstand
$SecurityGroupStaffAddress = "vorstand." + $cityLowerSimple + $WBORG						# vorstand.stadt@weitblicker.org
$SecurityGroupAdminName = $cityLowerFull + " stadtadmin"									# stadt stadtadmin
$SecurityGroupAdminAddress = "admin." + $cityLowerSimple + $WBORG							# admin.stadt@weitblicker.org
$SecurityGroupAdminAddressProxy = "stadtadmin." + $cityLowerSimple + $WBORG					# stadtadmin.stadt@weitblicker.org - Alias for admin.stadt@weitblicker.org

# office 365 groups
$UnifiedGroupMemberName = $cityUpperSimple					 								# Stadt
$UnifiedGroupMemberNameURL = "weitblick_" + $cityLowerSimple								# Weitblick_Stadt - Temp. name for Sharepoint URL Creation
$UnifiedGroupMemberAddress = "mitgliedsgruppe." + $cityLowerSimple + $WBORG					# mitgliedsgruppe.stadt@weitblicker.org
$UnifiedGroupMemberAddressOMS = "mitgliedsgruppe." + $cityLowerSimple + $OMSCOM				# mitgliedsgruppe.stadt@weitblicker.onmicrosoft.com - MS Address for O365Group
$UnifiedGroupStaffName = $cityUpperSimple + " Vorstand"										# Stadt Vorstand
$UnifiedGroupStaffAddress = "vorstandsgruppe." + $cityLowerSimple + $WBORG 					# vorstandsgruppe.stadt@weitblicker.org
$UnifiedGroupStaffAddressOMS = "vorstandsgruppe." + $cityLowerSimple + $OMSCOM					# vorstandsgruppe.stadt@weitblicker.onmicrosoft.com - MS Address for O365Group

###---------------------------------------------------------------------------
### Exit Parameters and triggers

# Trigger for Member Group and Mailbox
[bool] $MailboxExists = $false; if (Get-Mailbox $MailboxAddress -ErrorAction silentlycontinue) {$MailboxExists=$true}
[bool] $UnifiedGroupMemberExists = $false; if (Get-UnifiedGroup $UnifiedGroupMemberAddress -ErrorAction silentlycontinue) {$UnifiedGroupMemberExists=$true}

#### @todo3: user bundesverband found? else end

###---------------------------------------------------------------------------
### Begin of Function

## Creation Process for Office365 Group and Mailbox
write-host "GROUP AND MAILBOX CREATION" -ForegroundColor DarkYellow

write-host "Create non-existing Security Groups" -ForegroundColor Yellow
# create E-Mail activated security groups
if (-not (Get-DistributionGroup $SecurityGroupStaffAddress -ErrorAction silentlycontinue)) {New-DistributionGroup -Name $SecurityGroupStaffName -Type "Security" -PrimarySmtpAddress $SecurityGroupStaffAddress -ManagedBy "bundesverband@weitblicker.org" | Out-Null}
if (-not (Get-DistributionGroup $SecurityGroupMemberAddress -ErrorAction silentlycontinue)) {New-DistributionGroup -Name $SecurityGroupMemberName -Type "Security" -PrimarySmtpAddress $SecurityGroupMemberAddress -ManagedBy "bundesverband@weitblicker.org" | Out-Null}
if (-not (Get-DistributionGroup $SecurityGroupAdminAddress -ErrorAction silentlycontinue)) {New-DistributionGroup -Name $SecurityGroupAdminName -Type "Security" -PrimarySmtpAddress $SecurityGroupAdminAddress -ManagedBy "bundesverband@weitblicker.org" | Out-Null}

write-host "Create Mailbox/MemberGroup:" -ForegroundColor DarkYellow
# Create office 365 member group
if (-not $UnifiedGroupMemberExists) {
    write-host "Creating new Member group" -ForegroundColor Yellow

    # Create Group with temporary URL Name
    New-UnifiedGroup -DisplayName $UnifiedGroupMemberNameURL -PrimarySmtpAddress $UnifiedGroupMemberAddress -Owner "bundesverband@weitblicker.org" -AccessType Private | Out-Null
}

if (-not $MailboxExists) {
    write-host "Creating New Mailbox" -ForegroundColor Yellow

    New-Mailbox -Shared -Name $MailboxName -PrimarySmtpAddress $MailboxAddress | Out-Null
    #@todo3: WARNUNG: Fehler beim Replizieren von Postfach: 'xxx' zum Standort: 'EURPRD10.PROD.OUTLOOK.COM/Configuration/Sites/DB7PR10'

    # Add Exchange license to Mailbox
    Connect-MsolService -Credential $UserCredential | Out-Null
    Set-MsolUser -UserPrincipalName $MailboxAddress -UsageLocation "DE" | Out-Null
    Set-MsolUserLicense -UserPrincipalName $MailboxAddress -AddLicenses "weitblicker:STANDARDWOFFPACK" | Out-Null
}

# Create office 365 Staff group
if (-not (Get-UnifiedGroup $UnifiedGroupStaffAddress -ErrorAction silentlycontinue)) {
    write-host "Creating new Staff group" -ForegroundColor Yellow

    New-UnifiedGroup -DisplayName $UnifiedGroupStaffName -PrimarySmtpAddress $UnifiedGroupStaffAddress -Owner "bundesverband@weitblicker.org" -AccessType Private | Out-Null
}


## Set all group settings
write-host "GROUP CONFIG..." -ForegroundColor DarkYellow

# Office365 Groups
write-host "...Office365" -ForegroundColor DarkCyan

    ## Member Group Settings
    Set-UnifiedGroup -Identity $UnifiedGroupMemberAddress -PrimarySmtpAddress $UnifiedGroupMemberAddress -DisplayName $($cityUpperFull) -AccessType Private -SubscriptionEnabled:$true -AlwaysSubscribeMembersToCalendarEvents:$true -AutoSubscribeNewMembers:$true | Out-Null
    Add-UnifiedGroupLinks -Identity $UnifiedGroupMemberAddress -LinkType Owners -Links "bundesverband@weitblicker.org"
    Set-UnifiedGroup -Identity $UnifiedGroupMemberAddress -EmailAddresses @{remove=$($UnifiedGroupMemberNameURL + $WBORG)} -ErrorAction silentlycontinue | Out-Null

    # Adjust MOERA (onmicrosoft ID)
    Set-UnifiedGroup -Identity $UnifiedGroupMemberAddress -EmailAddresses @{add=$("mitgliedsgruppe." + $cityLowerSimple + $OMSCOM)}  -ErrorAction silentlycontinue | Out-Null
    Set-UnifiedGroup -Identity $UnifiedGroupMemberAddress -EmailAddresses @{remove=$UnifiedGroupMemberNameURL + $OMSCOM} -ErrorAction silentlycontinue | Out-Null

    ## Staff Group Settings
    Set-UnifiedGroup -Identity $UnifiedGroupStaffAddress -PrimarySmtpAddress $UnifiedGroupStaffAddress -DisplayName $($cityUpperFull + " Vorstand") -AccessType Private -SubscriptionEnabled:$true -AlwaysSubscribeMembersToCalendarEvents:$true -AutoSubscribeNewMembers:$true | Out-Null
    Add-UnifiedGroupLinks -Identity $UnifiedGroupStaffAddress -LinkType Owners -Links "bundesverband@weitblicker.org"
    Set-UnifiedGroup -Identity $UnifiedGroupStaffAddress -EmailAddresses @{remove=$($cityLowerSimple + "Vorstand" + $WBORG)} -ErrorAction silentlycontinue | Out-Null

    # Adjust MOERA (onmicrosoft ID)
    Set-UnifiedGroup -Identity $UnifiedGroupStaffAddress -EmailAddresses @{add= "vorstandsgruppe." + $cityLowerSimple + $OMSCOM}| Out-Null
    Set-UnifiedGroup -Identity $UnifiedGroupStaffAddress -EmailAddresses @{remove=$cityLowerSimple + "Vorstand" + $OMSCOM}| Out-Null

    #Delete ~1 Group
    #@todo3: Existiert nicht
    Set-UnifiedGroup -Identity $UnifiedGroupStaffAddress -EmailAddresses @{remove=$cityLowerSimple + "Vorstand1" + $OMSCOM}| Out-Null

    # Remove "bundesverband" as Group Member
    # Remove-DistributionGroupMember -Identity $UnifiedGroupMemberAddress -Member "bundesverband@weitblicker.org"
        #@todo2: Bundesverband als Mitglied der Gruppen entfernen
        #Geht nicht weil Owner, in der GUI jedoch möglich!

    # Set Read Permissions for all members
        #@todo2: Leseberechtigung für mitglieder@weitblicker.org (weitblick mitglieder) setzen
        # Unklar wo das geht bzw. wie die heißen
        # Add-SPOService

## E-Mail activated Security Groups
write-host "...E-Mail activated Security" -ForegroundColor DarkCyan

    # Group Settings
    Set-DistributionGroup -Identity $SecurityGroupStaffAddress  -Name $SecurityGroupStaffName -PrimarySmtpAddress $SecurityGroupStaffAddress -BypassSecurityGroupManagerCheck -ManagedBy "bundesverband@weitblicker.org" | Out-Null
    Set-DistributionGroup -Identity $SecurityGroupMemberAddress -Name $SecurityGroupMemberName -PrimarySmtpAddress $SecurityGroupMemberAddress -BypassSecurityGroupManagerCheck -ManagedBy "bundesverband@weitblicker.org" -RequireSenderAuthenticationEnabled:$true | Out-Null
    Set-DistributionGroup -Identity $SecurityGroupAdminAddress -Name $SecurityGroupAdminName -PrimarySmtpAddress $SecurityGroupAdminAddress -BypassSecurityGroupManagerCheck -ManagedBy "bundesverband@weitblicker.org" | Out-Null

    # Set parent memberships of groups
    Add-DistributionGroupMember -Identity mitglieder@weitblicker.org -BypassSecurityGroupManagerCheck -Member $SecurityGroupMemberAddress  -ErrorAction silentlycontinue | Out-Null
    Add-DistributionGroupMember -Identity vorstaende@weitblicker.org -BypassSecurityGroupManagerCheck -Member $SecurityGroupStaffAddress  -ErrorAction silentlycontinue | Out-Null
    Add-DistributionGroupMember -Identity stadtadmins@weitblicker.org -BypassSecurityGroupManagerCheck -Member $SecurityGroupAdminAddress  -ErrorAction silentlycontinue | Out-Null
    Add-DistributionGroupMember -Identity staedte@weitblicker.org -BypassSecurityGroupManagerCheck -Member $MailboxAddress  -ErrorAction silentlycontinue | Out-Null

    # set alias (proxy) for admin group
    Set-DistributionGroup -Identity $SecurityGroupAdminAddress -EmailAddress @{add=$SecurityGroupAdminAddressProxy} -ErrorAction silentlycontinue | Out-Null

# Mailbox
write-host "...Mailbox" -ForegroundColor DarkCyan
    ## City Mailbox
    Set-Mailbox -Identity $MailboxAddress -DisplayName $("Weitblick " + $MailboxName) -GrantSendOnBehalfTo $SecurityGroupStaffAddress -MessageCopyForSendOnBehalfEnabled $true -MessageCopyForSentAsEnabled $true | Out-Null
    Add-MailboxPermission -Identity $MailboxAddress -User $SecurityGroupStaffAddress -AccessRights FullAccess -InheritanceType All | Out-Null
        # @todo3: Mitgliedschaft wird nicht im Frontend angezeigt
        # Add-AzureADGroupMember??

    Set-MailBox -Identity $MailboxAddress -EmailAddresses @{add=$cityLowerSimple + $OMSCOM}

    #@todo3: Fehler beim Anlegen - Die Domäne "" kann nicht verwendet werden, da es sich nicht um eine akzeptierte Domäne für die Organisation handelt.
    #Set-MailBox -Identity $MailboxAddress -EmailAddresses @{add=$cityLowerSimple + $WBNGO}
    #Set-MailBox -Identity $MailboxAddress -EmailAddresses @{add=$cityLowerSimple + $WBONG}


# Verschoben unter die Config, damit der Teams Name korrekt ist
## Create Teams
write-host "creating teams..." -ForegroundColor DarkYellow
# connect to AzureAD and Teams

Connect-AzureAD -Credential $UserCredential  # seems to be necessary to connect to MsolService ?
Connect-MsolService -Credential $UserCredential
Connect-MicrosoftTeams -Credential $UserCredential

# Get IDs of Unified Group (O365 Group)
$UnifiedGroupMember = Get-MsolGroup -SearchString $UnifiedGroupMemberAddress
$UnifiedGroupStaff = Get-MsolGroup -SearchString $UnifiedGroupStaffAddress

# Create new Team based on Unified Group (O365 Group)
#@todo2: Prüfen ob existiert
New-Team -Group $UnifiedGroupMember.ObjectID -ErrorAction silentlycontinue | Out-Null
New-Team -Group $UnifiedGroupStaff.ObjectID -ErrorAction silentlycontinue | Out-Null

# set private
Set-Team -Group $UnifiedGroupMember.ObjectID -Visibility Private | Out-Null
Set-Team -Group $UnifiedGroupStaff.ObjectID -Visibility Private | Out-Null

# add description
#@todo3: Nur wenn leer
Set-Team -Group $UnifiedGroupMember.ObjectID -Description $("Zentrale Gruppe von Weitblick " + $cityUpperFull ) | Out-Null
Set-Team -Group $UnifiedGroupStaff.ObjectID -Description $("Gruppe des Vorstandes von Weitblick " + $cityUpperFull) | Out-Null

# Disconnect again
Disconnect-MicrosoftTeams
Disconnect-AzureAD
# no disconnect for MsolService


## Create Planner
    #@todo3: Integrate Planner Script



# call group website once from outlook view to not get 404 for weitblicker.sharepoint.com/sites/stadt
$url = "https://weitblicker.sharepoint.com/_layouts/15/groupstatus.aspx?id=" + $UnifiedGroupMember.ObjectID + "&target=site"
Start-Process -FilePath firefox $url

write-host "Aaaaaaaaaand done!" -ForegroundColor DarkYellow

## End of Function
###---------------------------------------------------------------------------

## End Session
Remove-PSSession $Session
#Set-ExecutionPolicy Restricted -force
