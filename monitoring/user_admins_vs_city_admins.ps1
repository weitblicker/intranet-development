# start up ... import session etc.

# Install-Module MSonline

try
{
    if ([string]::IsNullOrEmpty($UserCredential)) {
        $UserCredential = Get-Credential
    }

    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Unrestricted -force
    Import-PSSession $Session -AllowClobber
    Connect-MsolService -Credential $UserCredential
}
catch
{
    write-host $_.Exception.Message
    write-host "Login failed. Executed Powershell as admin?"
    $UserCredential=$NULL
    break
}

# get all user admins
$role = Get-MsolRole -RoleName "User Account Administrator"
$user_admins = Get-MsolRoleMember -RoleObjectId $role.objectid | Select-Object -ExpandProperty EmailAddress | Sort-Object

$cities = "bayreuth", "berlin", "bochum", "bonn", "duisburg-essen", "freiburg", "goettingen", "hamburg", "hannover", "heidelberg", "kiel", "koeln", "leipzig", "marburg", "muenchen", "muenster", "osnabrueck", "plus"
$city_admins = foreach($city in $cities){Get-DistributionGroupMember $("admin." + $city + "@weitblicker.org") | Select-Object -ExpandProperty PrimarySmtpAddress | Sort-Object}

write-host
write-host "***"
write-host "user admins not in respective city admin group"
write-host "***"
write-host
foreach($user_admin in $user_admins){
    if($city_admins -notcontains $user_admin){
        $user_admin
    }
}

write-host
write-host "***"
write-host "city admins without user admin role"
write-host "***"
write-host
foreach($city_admin in $city_admins){
    if($user_admins -notcontains $city_admin){
        $city_admin
    }
}