Import-Module ActiveDirectory
[Reflection.Assembly]::LoadWithPartialName("System.Web") # Allows random password generation


$TermedUser = Read-Host "Username of Termed Employee" # Username of termed employee

######## This block checks that the user exists before continuing########
Try {
$User = Get-ADUser $TermedUser -Properties * -ErrorAction Stop
    }
    Catch {
            If ($_ -like "*Cannot find an object with identity*") {
                    "User '$TermedUser' Does not exist. Aborting Script"
             }
             Else {
                "An Error Occured. Aborting Script"
                }
                Continue
            }
            "User '$($User.SamAccountName)' Exists. Continuing Script"
#########################################################################

do {
   $TempPass = [System.Web.Security.Membership]::GeneratePassword(8,2)
} until ($TempPass -match '\d') # 8 characters long with at least one number

# AD Stuff
Set-ADUser $User -enabled $false -description ("DOT:" + (Get-Date).tostring("MM.dd.yyyy")) # Disable the account, add Date of Term to description
Set-ADAccountPassword $User -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $TempPass -Force -Verbose) -PassThru # Reset the account password with random pass
Get-ADPrincipalGroupMembership $User | where {$_.name -ne 'Domain Users'} | Remove-ADGroupMember -member $User # Remove from all groups
Move-ADObject -identity $User.DistinguishedName -TargetPath "OU=Disabled Accounts,OU=NorthMarq_Capital,OU=Companies,DC=NMC,DC=Northmarq,DC=com" # Move to Disabled Accounts OU

# Exchange Stuff
Set-Mailbox $User -HiddenFromAddressListsEnabled $True # Hide from Address List
Set-CASMailbox -Identity $User -ActiveSyncEnabled $false # Disable ActiveSync
If ($User.Manager -ne $null){ # Check if the manager exists before running next 2 lines
        $Manager = Get-ADUser $User.Manager -Properties * | Select samaccountname # Get managers username
        Add-MailboxPermission -identity "$User" -user "$Manager" -AccessRights FullAccess -InheritanceType All # Give manager full permission to mailbox
    }
Get-ActiveSyncDevice -mailbox $User | Remove-ActiveSyncDevice -force # Remove all activeSync Devices
Set-MailboxAutoReplyConfiguration -Identity $User -AutoReplyState Enabled -InternalMessage "Internal auto-reply message." -ExternalMessage "External auto-reply message." # Sets Auto Replies
#}
    #Else { Write-Host "User does not exist"}