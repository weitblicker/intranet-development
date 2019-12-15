# start up ... import session etc.

# get group name

$group_name = "sandburg mitglieder"
# @todo set up backup path with date and time
$date = "2019-01-19"
$folder = ".\backup\"
$file_path = $folder +"\"+ $date +"\"+ $date +" "+ $group_name + ".txt"

$member = Get-DistributionGroupMember $group_name | Select-Object -ExpandProperty Name | Sort-Object

$member_backup = Get-Content $file_path

$member
""
$member_backup
""

# check if line in backup file is member
"removed since backup, will be re-added:"
foreach($line in $member_backup) {
    if($member -notcontains $line){
      $line
    }
}

# check if member in backup file
"added since backup, will be removed:"
foreach($line in $member) {
    if($member_backup -notcontains $line){
      $line
    }
}

# ask to apply backup
$confirm = Read-Host -Prompt "Do you want to continue? [y] / [n], default: [n]"

if($confirm -like "y"){
  # check if line in backup file is member
  "Adding users..."
  foreach($line in $member_backup) {
      if($member -notcontains $line){
        Add-DistributionGroupMember $group_name -Member $line
        $line + " removed"
      }
  }

  # check if member in backup file
  "Removing users..."
  foreach($line in $member) {
      if($member_backup -notcontains $line){
        Remove-DistributionGroupMember $group_name -Member $line
        $line + " removed"
      }
  }
}
