Set-ExecutionPolicy RemoteSigned
$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://starscream.nmc.northmarq.com/PowerShell/ -Authentication Kerberos -Credential $UserCredential
Import-PSSession $Session
