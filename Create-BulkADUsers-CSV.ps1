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
#STATIC VARIABLES
#----------------------------------------------------------
$path     = Split-Path -parent $MyInvocation.MyCommand.Definition
$newpath  = "C:\Scripts\NewUsers.csv"
$log      = "C:\Scripts\Create-BulkADUsers-CSV.log"
$date     = Get-Date
$addn     = (Get-ADDomain).DistinguishedName
$dnsroot  = (Get-ADDomain).DNSRoot
$i        = 1

"Processing started (on " + $date + "): " | Out-File $log -append
"--------------------------------------------" | Out-File $log -append

Import-Csv $newpath | ForEach-Object {
 $userPrincinpal = $_.samAccountName + "@$dnsroot"
 $ou = $_.ParentOU + ",$addn"
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
