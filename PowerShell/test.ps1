<#
Author          : Avanade Collaboration Sl Capability
Created Date    : April 12, 2016
Description     : Exports each library item details from a given SharePoint 2010/2013 site collection .
Reference Doc   : -
#>

param(
[string]$siteCollectionURL,
[string]$spoAdmin,
[string]$spoAdminPassword,
[string]$spoAdminURL

)

Add-Type –Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.dll" 
Add-Type –Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

#Import SharePoint Online module 
Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking


#$username = 'labani@AvaIndCollabSL.onmicrosoft.com'
#$Password = '$D9920526814l'

$username = $spoAdmin
$Password = $spoAdminPassword
#$spoAdminURL = 'https://avaindcollabsl-admin.sharepoint.com/'



#$Password = Get-Content "Password.txt"

# Check to ensure Microsoft.SharePoint.PowerShell is loaded
$snapin = Get-PSSnapin | Where-Object {$_.Name -eq 'Microsoft.SharePoint.Powershell'}
if ($snapin -eq $null) 
{
    #Write-Host "Loading SharePoint Powershell Snapin"

    # Add SharePoint cmdlets reference 
    Add-PSSnapin "Microsoft.SharePoint.Powershell" -ErrorAction SilentlyContinue
}


$cred = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist $username, $(convertto-securestring $Password -asplaintext -force)
Connect-SPOService -Url $spoAdminURL -Credential $cred
Get-SPOUser -Site $siteCollectionURL | Where-Object {$_.IsSiteAdmin -eq $true} | select DisplayName





