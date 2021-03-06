### Script to create new weitblick city // update existing to new group structure

### Manual Init
# Execute this when using the script on a computer for the first time
# Install-Module MicrosoftTeams
# Install-Module MSOnline
# Install-Module AzureAD

# Execute this when opening the script for the first
# Set-ExecutionPolicy -Scope Process -ExecutionPolicy Unrestricted -force

# Admin rights on weitblicker.org and portal.azure.com for weitblicker.org required

###---------------------------------------------------------------------------
### Start Session


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
    write-host $_.Exception.Message -ForegroundColor Yellow
    write-host "Login failed. Executed Powershell as admin?" -ForegroundColor Yellow
    $UserCredential=$NULL
    break
}

[bool]$chk = $true
while($chk=$true) { #Dauerrepeat
## get (new) city name
[string] $city=''; $city=read-host "Enter city to delete: "


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

$city=$city.ToLower()


$city_mail=Replace-Umlaute $city
$city=(Get-Culture).TextInfo.ToTitleCase($city)

#@todo3: Init-Fehler: "WARNUNG: Die Namen einiger importierter Befehle auf Modul "tmp_s252o0r4.fsq" enthalten nicht genehmigte Verben"

write-host "Generating variables..." -ForegroundColor Darkyellow


### Alias Variables
$MailboxName = $city 																# Stadt
$MailboxAddress = $city_mail + "@weitblicker.org"									# stadt@weitblicker.org
$MailboxAddressSMTP = "SMTP:" + $MailboxAddress
$MailboxAddressTemporary = $city_mail + "_temp@weitblicker.org"						# stadt_temp@weitblicker.org - to release primarySmtpAdress
$MailboxAddressTemporarySMTP = "SMTP:" + $MailboxAddressTemporary

# security groups
$SecurityGroupMemberName = $city.ToLower() + " mitglieder"							# stadt mitglieder
$SecurityGroupMemberAddress = "mitglieder." + $city_mail + "@weitblicker.org"		# mitglieder.stadt@weitblicker.org - e-mail security grp
$SecurityGroupStaffName = $city.ToLower() + " vorstand"								# stadt vorstand
$SecurityGroupStaffAddress = "vorstand." + $city_mail + "@weitblicker.org"			# vorstand.stadt@weitblicker.org
$SecurityGroupAdminName = $city.ToLower() + " stadtadmin"							# stadt stadtadmin
$SecurityGroupAdminAddress = "admin." + $city_mail + "@weitblicker.org"				# admin.stadt@weitblicker.org
$SecurityGroupAdminAddressProxy = "stadtadmin." + $city_mail + "@weitblicker.org"	# stadtadmin.stadt@weitblicker.org

# office 365 groups
$UnifiedGroupMemberName = $MailboxName										 		# Stadt
$UnifiedGroupMemberAddress = "mitgliedsgruppe." + $city_mail + "@weitblicker.org"	# mitgliedsgruppe.stadt@weitblicker.org
$UnifiedGroupStaffName = $MailboxName + " Vorstand"									# Stadt Vorstand
$UnifiedGroupStaffAddress = "vorstandsgruppe." + $city_mail + "@weitblicker.org" 	# vorstandsgruppe.stadt@weitblicker.org

Remove-DistributionGroup $GroupMemberName -Confirm:$False -BypassSecurityGroupManagerCheck -ErrorAction silentlycontinue | out-Null
Start-Sleep -Milliseconds 500
Remove-DistributionGroup $GroupStaffName -Confirm:$False -BypassSecurityGroupManagerCheck -ErrorAction silentlycontinue | out-Null
Start-Sleep -Milliseconds 500
Remove-DistributionGroup $GroupAdminName -Confirm:$False -BypassSecurityGroupManagerCheck -ErrorAction silentlycontinue | out-Null
Start-Sleep -Milliseconds 500
Remove-UnifiedGroup $UnifiedGroupMemberName -Confirm:$False -ErrorAction silentlycontinue | out-Null
Start-Sleep -Milliseconds 500
Remove-UnifiedGroup $UnifiedGroupStaffName -Confirm:$False -ErrorAction silentlycontinue | out-Null
Start-Sleep -Milliseconds 500
Remove-Mailbox $MailboxName -Confirm:$False -ErrorAction silentlycontinue | out-Null


# Nochmal mit ohne Umlaute in Namen

#temp-Bugfix - keine Umlaute
#@todo: Umlaute in Namen von Office365 Gruppen
$city=(Get-Culture).TextInfo.ToTitleCase($city_mail)

#@todo3: Init-Fehler: "WARNUNG: Die Namen einiger importierter Befehle auf Modul "tmp_s252o0r4.fsq" enthalten nicht genehmigte Verben"

write-host "Generating variables..." -ForegroundColor Darkyellow


### Alias Variables
$MailboxName = $city 																# Stadt
$MailboxAddress = $city_mail + "@weitblicker.org"									# stadt@weitblicker.org
$MailboxAddressSMTP = "SMTP:" + $MailboxAddress
$MailboxAddressTemporary = $city_mail + "_temp@weitblicker.org"						# stadt_temp@weitblicker.org - to release primarySmtpAdress
$MailboxAddressTemporarySMTP = "SMTP:" + $MailboxAddressTemporary

# security groups
$SecurityGroupMemberName = $city.ToLower() + " mitglieder"							# stadt mitglieder
$SecurityGroupMemberAddress = "mitglieder." + $city_mail + "@weitblicker.org"		# mitglieder.stadt@weitblicker.org - e-mail security grp
$SecurityGroupStaffName = $city.ToLower() + " vorstand"								# stadt vorstand
$SecurityGroupStaffAddress = "vorstand." + $city_mail + "@weitblicker.org"			# vorstand.stadt@weitblicker.org
$SecurityGroupAdminName = $city.ToLower() + " stadtadmin"							# stadt stadtadmin
$SecurityGroupAdminAddress = "admin." + $city_mail + "@weitblicker.org"				# admin.stadt@weitblicker.org
$SecurityGroupAdminAddressProxy = "stadtadmin." + $city_mail + "@weitblicker.org"	# stadtadmin.stadt@weitblicker.org

# office 365 groups
$UnifiedGroupMemberName = $MailboxName										 		# Stadt
$UnifiedGroupMemberAddress = "mitgliedsgruppe." + $city_mail + "@weitblicker.org"	# mitgliedsgruppe.stadt@weitblicker.org
$UnifiedGroupStaffName = $MailboxName + " Vorstand"									# Stadt Vorstand
$UnifiedGroupStaffAddress = "vorstandsgruppe." + $city_mail + "@weitblicker.org" 	# vorstandsgruppe.stadt@weitblicker.org

Remove-DistributionGroup $GroupMemberName -Confirm:$False -BypassSecurityGroupManagerCheck -ErrorAction silentlycontinue | out-Null
Start-Sleep -Milliseconds 200
Remove-DistributionGroup $GroupStaffName -Confirm:$False -BypassSecurityGroupManagerCheck -ErrorAction silentlycontinue | out-Null
Start-Sleep -Milliseconds 200
Remove-DistributionGroup $GroupAdminName -Confirm:$False -BypassSecurityGroupManagerCheck -ErrorAction silentlycontinue | out-Null
Start-Sleep -Milliseconds 200
Remove-UnifiedGroup $UnifiedGroupMemberName -Confirm:$False -ErrorAction silentlycontinue | out-Null
Start-Sleep -Milliseconds 200
Remove-UnifiedGroup $UnifiedGroupStaffName -Confirm:$False -ErrorAction silentlycontinue | out-Null
Start-Sleep -Milliseconds 200
Remove-Mailbox $MailboxName -Confirm:$False -ErrorAction silentlycontinue | out-Null

}

Remove-PSSession $Session