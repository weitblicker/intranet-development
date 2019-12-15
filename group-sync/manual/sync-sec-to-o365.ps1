# start up ... import session etc.

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
    write-host $_.Exception.Message
    write-host "Login failed. Executed Powershell as admin?"
    $UserCredential=$NULL
    break
}

# --------------------------------------------------------------
# Variables

#$cities = "bayreuth", "berlin", "bochum", "bonn", "duisburg-essen", "freiburg", "goettingen", "hamburg", "hannover", "heidelberg", "kiel", "koeln", "leipzig", "marburg", "muenchen", "muenster", "osnabrueck", "plus"

$sync_location = "./sync/"
$backup_location = "./backup/"
$date_str = Get-Date -format "yyyy-MM-dd"

# --------------------------------------------------------------
# Backup
$backup = Read-Host "should a backup be done before the sync? [y] [n]; default [y]"

if(-not ($backup -like "n")){
  # 3 Backups per city (one for each group)
  $cities = "bayreuth", "berlin", "bochum", "bonn", "duisburg-essen", "freiburg", "goettingen", "hamburg", "hannover", "heidelberg", "kiel", "koeln", "leipzig", "marburg", "muenchen", "muenster", "osnabrueck", "plus"

  foreach($city in $cities){
    foreach($type in "mitglieder", "admin", "vorstand"){
      # set up names
      $group_name = $city + " " + $type
      if($type -like "admin"){$group_name = $city + " stadt" + $type}
      $group_mail = $type + "." + $city + "@weitblicker.org"

      # set up path to backup file
      $backup_folder_path = $backup_location +"\"+ $date_str
      $backup_file_path = $backup_folder_path +"\"+ $date_str +" "+ $group_name + ".txt"
      # create folder if not already exists
      if(-not (Test-Path $backup_folder_path)){
        mkdir $backup_folder_path | Out-Null
      }

      # multiple backups on same day
      $i = 1
      while((Test-Path $backup_file_path) -and ($i -lt 100)){
        $backup_file_path = $backup_folder_path +"\"+ $date_str +"-"+$i+" "+ $group_name + ".txt"
        $i = $i + 1
      }

      # get members from distribution group and write to file
      Get-DistributionGroupMember $group_mail | Select-Object -ExpandProperty Name | Sort-Object | Set-Content $backup_file_path
    }
  }
}


# ---------------------------------------------
# sync securitygroup -> O365

$cities = "bayreuth"  # only execute single city

foreach($city in $cities){
  Write-Host "---" $city "---"
  # set up names
  $o_member_name = (Get-Culture).TextInfo.ToTitleCase($city)
  $o_member_address = "mitgliedsgruppe." + $city + "@weitblicker.org"
  $sec_member_name = $city + " mitglieder"
  $sec_member_address = "mitglieder." + $city + "@weitblicker.org"
  $o_staff_name = (Get-Culture).TextInfo.ToTitleCase($city) + " Vorstand"
  $o_staff_address = "vorstandsgruppe." + $city + "@weitblicker.org"
  $sec_staff_name =  $city + " vorstand"
  $sec_staff_address = "vorstand." + $city + "@weitblicker.org"

  # set up path to sync file from office group
  $sync_folder_path = $sync_location +"\"+ $date_str
  $sync_member_file_path = $sync_folder_path +"\"+ $date_str +" "+ $o_member_name + ".txt"
  $sync_staff_file_path = $sync_folder_path +"\"+ $date_str +" "+ $o_staff_name + ".txt"

  # create folder if not already exists
  if(-not (Test-Path $sync_folder_path)){
    Write-Host "create folder for sync files..."
    mkdir $sync_folder_path
  }

  # get members from office 365 group and write to file (overwrite existing files)
  Write-Host "write office group members to file..."
  $o_member = Get-UnifiedGroupLinks -Identity $o_member_address -LinkType Members | Select-Object -ExpandProperty Name | Sort-Object
  $o_staff = Get-UnifiedGroupLinks -Identity $o_staff_address -LinkType Members | Select-Object -ExpandProperty Name | Sort-Object



  # compare to members in security groups
  $o_member | Out-File $sync_member_file_path
  $o_staff | Out-File $sync_staff_file_path

  $sec_member = Get-DistributionGroupMember $sec_member_address | Select-Object -ExpandProperty Name | Sort-Object
  $sec_staff = Get-DistributionGroupMember $sec_staff_address | Select-Object -ExpandProperty Name | Sort-Object

  # mitglieder -> mitgliedsgruppe
  $i = 0
  foreach($user in $o_member) {
    if($sec_member -notcontains $user){
      # user was removed since last sync
      if($user -like "bundesverband"){continue}  # skip bundesverband, as this is the owner and cann not be removed
      "remove " + $user
      Remove-UnifiedGroupLinks -Identity $o_staff_address -Links $user -LinkType Members -Confirm:$false
      $i += 1
    }
  }
  # counter for added users
  $j = 0
  foreach($user in $sec_member) {
      if($o_member -notcontains $user){
        # user was added since last sync
        "add " + $user
        Add-UnifiedGroupLinks -Identity $o_member_address -Links $user -LinkType Members -Confirm:$false
        $j += 1
      }
  }
  Write-Host $("mitglieder:  removed = " + $i + "    added = " + $j)

  # vorstand -> vorstandsgruppe
  $n = 0
  foreach($user in $o_staff) {
    if($sec_staff -notcontains $user){
      # user was removed since last sync
      if($user -like "bundesverband"){continue}  # skip bundesverband, as this is the owner and cann not be removed
      "remove " + $user
      Remove-UnifiedGroupLinks -Identity $o_staff_address -Links $user -LinkType Members -Confirm:$false
      $n += 1
    }
  }
  $m = 0
  foreach($user in $sec_staff) {
      if($o_staff -notcontains $user){
        # user was added since last sync
        # todo: add user from sec group
        "add " + $user
        Add-UnifiedGroupLinks -Identity $o_staff_address -Links $user -LinkType Members -Confirm:$false
        $m += 1
      }
  }
  Write-Host $("vorstand:  removed = " + $n + "    added = " + $m)
}

Remove-PSSession -Session $Session
