function Read-OpenFileDialog([string]$WindowTitle, [string]$InitialDirectory, [string]$Filter = "All files (*.*)|*.*", [switch]$AllowMultiSelect)
{  
    Add-Type -AssemblyName System.Windows.Forms
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Title = $WindowTitle
    if (![string]::IsNullOrWhiteSpace($InitialDirectory)) { $openFileDialog.InitialDirectory = $InitialDirectory }
    $openFileDialog.Filter = $Filter
    if ($AllowMultiSelect) { $openFileDialog.MultiSelect = $true }
    $openFileDialog.ShowHelp = $true    # Without this line the ShowDialog() function may hang depending on system configuration and running from console vs. ISE.
    $openFileDialog.ShowDialog() > $null
    if ($AllowMultiSelect) { return $openFileDialog.Filenames } else { return $openFileDialog.Filename }
}

function Read-FolderBrowserDialog([string]$Message, [string]$InitialDirectory, [switch]$NoNewFolderButton)
{
    $browseForFolderOptions = 0
    if ($NoNewFolderButton) { $browseForFolderOptions += 512 }
 
    $app = New-Object -ComObject Shell.Application
    $folder = $app.BrowseForFolder(0, $Message, $browseForFolderOptions, $InitialDirectory)
    if ($folder) { $selectedDirectory = $folder.Self.Path } else { $selectedDirectory = '' }
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) > $null
    return $selectedDirectory
}

function CheckOUExists([string]$ou)
{
    $result = ([ADSI]::Exists("LDAP://$ou"))
    return $result
}



Try

{

  Import-Module ActiveDirectory -ErrorAction Stop

}

Catch

{

  Write-Host "[ERROR]`t ActiveDirectory Module couldn't be loaded. Script will stop!"

  Exit 1

}



#----------------------------------------------------------

#VARIABLES

#----------------------------------------------------------

# $path     = Split-Path -parent $MyInvocation.MyCommand.Definition

$InputCSV = Read-OpenFileDialog -WindowTitle "Select CSV file to import" -InitialDirectory 'C:\' "CSV File (*.csv)|*.csv"
            if (![string]::IsNullOrEmpty($InputCSV)) {
            Write-Host "File Selected: $InputCSV"
            }
            Else {
            Write-Host "No File Selected"
            Exit
            }

$logdir      = Read-FolderBrowserDialog -Message "Please select a directory" -InitialDirectory 'C:\'
            if (![string]::IsNullOrEmpty($logdir)) {
            $log = "$logdir\Create-BulkADUsers-CSV.log"
            Write-Host "Log file saved at:  $logdir\Create-BulkADUsers-CSV.log"
            } 
            Else { 
            Write-Host "Log file not created."
            Exit
            } 



$date     = Get-Date

$addn     = (Get-ADDomain).DistinguishedName

$dnsroot  = (Get-ADDomain).DNSRoot



"Processing started (on " + $date + "): " | Out-File $log -append

"--------------------------------------------" | Out-File $log -append

$Title = "Are you sure you want to import these users?"
$message1 = "[A]bort or [C]ontinue"
$abort = New-Object System.Management.Automation.Host.ChoiceDescription '&Abort','Aborts the operation'
$continue = New-Object System.Management.Automation.Host.ChoiceDescription '&Continue','Continue to Import Users'
$options = [System.Management.Automation.Host.ChoiceDescription[]] ($abort,$continue)
$choice = $host.ui.PromptForChoice($title,$prompt,$options,0)


Import-Csv $InputCSV | ForEach-Object {

 $userPrincinpal = $_.samAccountName + "@$dnsroot"

 $ou = $ou = $_.ParentOU + ",$addn"

 if (!(CheckOUExists($ou))) {
 Write-Host "OU Doesnt Exists.  Creating OU"
 New-ADOrganizationalUnit -Name $_.OuName -Path $addn
 Write-Host $ou
 }
  

New-ADUser -Name $_.Name `
 -Path $ou `
 -SamAccountName  $_.samAccountName `
 -UserPrincipalName  $userPrincinpal `
 -AccountPassword (ConvertTo-SecureString "Password123" -AsPlainText -Force) `
 -ChangePasswordAtLogon $false  `
 -Enabled $true
 Write-Host "[INFO]`t Created new user : $($userPrincinpal)"

 }

#Add-ADGroupMember "Domain Admins" $_.samAccountName}