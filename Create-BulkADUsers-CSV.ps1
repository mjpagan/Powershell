function New-OrganizationalUnitFromDN
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [string]$DN
    )

    # A regex to split the DN, taking escaped commas into account
    $DNRegex = '(?<![\\]),'

    # Array to hold each component
    [String[]]$MissingOUs = @()

    # We'll need to traverse the path, level by level, let's figure out the number of possible levels 
    $Depth = ($DN -split $DNRegex).Count

    # Step through each possible parent OU
    for($i = 1;$i -le $Depth;$i++)
    {
        $NextOU = ($DN -split $DNRegex,$i)[-1]
        if($NextOU.IndexOf("OU=",[StringComparison]"CurrentCultureIgnoreCase") -ne 0 -or [ADSI]::Exists("LDAP://$NextOU"))
        {
            break
        }
        else
        {
            # OU does not exist, remember this for later
            $MissingOUs += $NextOU
        }
    }

    # Reverse the order of missing OUs, we want to create the top-most needed level first
    [array]::Reverse($MissingOUs)

    # Prepare common parameters to be passed to New-ADOrganizationalUnit
    $PSBoundParameters.Remove('DN')

    # Now create the missing part of the tree, including the desired OU
    foreach($OU in $MissingOUs)
    {
        $newOUName = (($OU -split $DNRegex,2)[0] -split "=")[1]
        $newOUPath = ($OU -split $DNRegex,2)[1]

        New-ADOrganizationalUnit -Name $newOUName -Path $newOUPath @PSBoundParameters
    }
}

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

clear
Write-Host "Select file from open file dialog.  Which may be behind this window" -foregroundcolor yellow

$InputCSV = Read-OpenFileDialog -WindowTitle "Select CSV file to import" -InitialDirectory 'C:\' "CSV File (*.csv)|*.csv"
            if (![string]::IsNullOrEmpty($InputCSV)) {
            Write-Host "File Selected: $InputCSV"
            }
            Else {
            Write-Host "No File Selected"
            Exit
            }

Write-host " "
Write-Host "Choose Directory for log file from folder browswer which may be behind this window." -ForegroundColor yellow
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

$Title = "Import Users"
$prompt = "Are you sure you want to import the csv file?"
$message1 = "[A]bort or [C]ontinue"
$abort = New-Object System.Management.Automation.Host.ChoiceDescription '&Abort','Aborts the operation'
$continue = New-Object System.Management.Automation.Host.ChoiceDescription '&Continue','Continue to Import Users'
$options = [System.Management.Automation.Host.ChoiceDescription[]] ($abort,$continue)
$choice = $host.ui.PromptForChoice($title,$prompt,$options,0)

switch ($choice)
 { 
 0 {
    write-host " "
    write-host "Import has been aborted." -ForegroundColor red
    exit
    }

 1 {

 Import-Csv $InputCSV | ForEach-Object {

 $userPrincinpal = $_.samAccountName + "@$dnsroot"

 $ou = $ou = $_.ParentOU + ",$addn"

 if (!(CheckOUExists($ou))) {
 Write-Host "OU Doesnt Exists.  Creating OU"
 
 #Add OU
 New-OrganizationalUnitFromDN $ou

 }
  
$ExistingUser = Get-ADUser -Filter {Name -eq "$_.Name"}

If ($? -eq $false) {
New-ADUser -Name $_.Name `
 -Path $ou `
 -SamAccountName  $_.samAccountName `
 -UserPrincipalName  $userPrincinpal `
 -AccountPassword (ConvertTo-SecureString "Password123" -AsPlainText -Force) `
 -ChangePasswordAtLogon $false  `
 -Enabled $true 
    If ($? -eq $true) {
        Add-Content $log -value "[INFO]`t Created new user : $($userPrincinpal)"
    } else {
        Add-Content $log -value "[ERROR]`t Error Creating new user : $($userPrincinpal)"
        Add-Content $log -value $error[0]
    }

 } else {
   Add-Content $log -value "[INFO]`t User already exists : $($userPrincinpal)"
 }
 

 }

#Add-ADGroupMember "Domain Admins" $_.samAccountName}

 }
 }


