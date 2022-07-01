$User = ""

#Must be un under Exchange Powershell. Tested under Exchange 2016 on prem. 

Write-Host "This script disconnects a users email box, finds a disconnected box with the same display name as the users, and connects that to the users AD."
Write-Host "This is used when the users AD account is deleted and recreated by mistake and a new mailbox is created, and the user wants their old email box back."
Write-Host "The new email box is left hanging as a disconnected email box in exchange, so make sure the user does not need any email as they will loose access."



$User = Read-Host "Enter user AD login name (ID) of user"



$ErrorActionPreference = 'Stop'

$emailboxguid = (Get-MailboxStatistics $User).MailBoxGuid
$emailboxDisplayName = (Get-MailboxStatistics $User).DisplayName
$emailboxDB = (Get-MailboxStatistics $User).DataBase.Name

#Health checks
if ($emailboxguid) {
	Write-Host "GOOD: Found user $User to remove mailbox from"
} else {
	Write-Error "ERROR: Did not find $User to remove mailbox from or user does not have a mailbox"
}


Write-Host "Searching for disconnected mailbox for $emailboxDisplayName (this is slow)"

#I wish I could find a faster way. 
$disconnectMailbox = Get-MailboxDatabase | Get-MailboxStatistics | Where {$_.DisconnectDate -ne $Null -and $_.displayName -eq $emailboxDisplayName}
if($disconnectMailbox -eq $NULL)
{
    Write-Error ("ERROR: A disconnected mailbox associated with $emailboxDisplayName does not exist.")
    exit
}

#Get the info for the disconnected mail box now before we have 2 of them. 
$disconnectMailboxDisplayName = ($disconnectMailbox).DisplayName
$disconnectMailboxGuid = ($disconnectMailbox).MailboxGuid
$disconnectMailboxDatabase = ($disconnectMailbox).Database


Write-Host "GOOD: Found disconnected mailbox for $disconnectMailboxDisplayName  GUID:  $disconnectMailboxGuid  Database: $disconnectMailboxDatabase "


$UserName = (Get-ADUser -Identity $User).Name

Write-Host "----------------------------------------"

Write-Host "$UserName ($User) will have their emailbox disconnected and the disconnected email box $disconnectMailboxDisplayName reconnected to their AD account $User"
Read-Host "Press any key to go"



Write-Host "----------------------------------------"

Write-Host "Enabeling user accounts"
Enable-ADAccount -Identity $User -ErrorAction Stop
Update-StoreMailboxState -Database $emailboxDB -Identity $emailboxguid -ErrorAction Stop

Start-Sleep 5

If (-NOT (Get-ADUser -Identity $User).Enabled) {
	Write-Error "ERROR: User account $User not enabled"
	exit
}
	


Write-Host "Disabeling $User Mailbox"
Disable-Mailbox -Identity $User -ErrorAction Stop

Write-Host "Forceing Exchange to update disconnected mailboxes for $User"
Update-StoreMailboxState -Database $emailboxDB -Identity $emailboxguid -ErrorAction Stop

Write-Host "Connecting Mailbox to $User"
Enable-ADAccount -Identity $User -ErrorAction Stop
Connect-Mailbox -Identity $disconnectMailboxGuid  -Database $disconnectMailboxDatabase -User $User -ErrorAction Stop

Write-Host "Done"



