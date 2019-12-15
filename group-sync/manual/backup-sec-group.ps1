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

# ---------------------------------------------
# actual script

$backup_location = "./backup/"
$date_str = Get-Date -format "yyyy-MM-dd"

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
      mkdir $backup_folder_path
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

Remove-PSSession $Session
