#################################################################################################################################
# This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment. # 
# THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED,  #
# INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.               #
# We grant You a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute the object  #
# code form of the Sample Code, provided that You agree: (i) to not use Our name, logo, or trademarks to market Your software   #
# product in which the Sample Code is embedded; (ii) to include a valid copyright notice on Your software product in which the  #
# Sample Code is embedded; and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and against any claims   #
# or lawsuits, including attorneys’ fees, that arise or result from the use or distribution of the Sample Code.                 #
#################################################################################################################################

#----------------------------------------------------------------------              
#-     UPDATE VARIABLES TO REFLECT YOUR ENVIRONMENT                   -
#----------------------------------------------------------------------
    
# Provide Azure AD Application registration information for your app.
$AppID = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" 
$TenantId  = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"

#----------------------------------------------------------------------              
#-     DO NOT CHANGE ANY CODE BELOW THIS LINE                         -
#----------------------------------------------------------------------
#-                                                                    -
#-                           Author:  Dirk Buntinx                    -
#-                           Date:    12/9/2022                       -
#-                           Version: v1.0                            -
#-                                                                    -
#----------------------------------------------------------------------

Write-Host "--------------------"
Write-Host "- Script settings: -"
Write-Host "--------------------"
Write-Host " - AppID: $($AppID)"
Write-Host " - TenantId: $($TenantId)"
Write-Host "--------------------"
Write-Host "-  Start script    -"
Write-Host "--------------------"
Write-Host " - Importing EWS Managed API"
# Download the EWS Managed API from here:
# https://github.com/officedev/ews-managed-api
# In order to install, make sure you have a NuGet PackageSource for location https://www.nuget.org/api/v2 installed
# Check running Get-PackageSource
# You will find a Nuget package Source for https://api.nuget.org/v3/index.json and you will have to add another one:
# Register-PackageSource -Name MyNuGet -Location https://www.nuget.org/api/v2 -ProviderName NuGet
# Now find the WebServices package using command:
# Find-Package Microsoft.Exchange.WebServices -RequiredVersion 2.2.0 -Source MyNuGet
# Than install it by piping the previous command to the install-package command


# Import the EWS API dll
Import-Module "C:\Program Files\PackageManagement\NuGet\Packages\Microsoft.Exchange.WebServices.2.2\lib\40\Microsoft.Exchange.WebServices.dll"

#######

Write-Host " - Getting OAuth Token"

# Use the MSAL.PS library to get the OAuth token
# install the MSAL.PS library from here: https://github.com/AzureAD/MSAL.PS
$Scopes = "https://outlook.office365.com/EWS.AccessAsUser.All"
$AuthResult = Get-MsalToken -TenantId $TenantId -ClientId $AppID -Scopes $Scopes -Interactive
#Write-Host "Auth Token: $($AuthResult.AccessToken)"

Write-Host " - Creating Exchange Service object"

#Create the Exchange Service object 
$Service = [Microsoft.Exchange.WebServices.Data.ExchangeService]::new()
$Service.Url = "https://outlook.office365.com/EWS/Exchange.asmx"

#Set the recommended Instrumentation Headers
$Service.ClientRequestId = [guid]::NewGuid().ToString()
$Service.ReturnClientRequestId = $True
$Service.UserAgent = "OAuth_SampleScriptAppOnly"

#Using OAuth authentication
$Service.Credentials = [Microsoft.Exchange.WebServices.Data.OAuthCredentials]$AuthResult.AccessToken

#Using EWS Impersonation to connect to the mailbox defined in vairable $MailboxName
#$Service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)

#Set the required X-AnchorMailbox Header
#$Service.HttpHeaders.Add("X-AnchorMailbox", $MailboxName)

######

Write-Host " - Retrieving Top Level Folders"

#Create a FolderView and get only the ID + DisplayName property
$Folderview = New-Object Microsoft.Exchange.WebServices.Data.FolderView(100)
$Folderview.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.Webservices.Data.BasePropertySet]::IdOnly)
$Folderview.PropertySet.Add([Microsoft.Exchange.Webservices.Data.FolderSchema]::DisplayName)
$Folderview.Traversal = [Microsoft.Exchange.Webservices.Data.FolderTraversal]::Shallow

#Search only for Folder Directly under the MsgFolderRoot
$FoldersResult = $Service.FindFolders([Microsoft.Exchange.Webservices.Data.WellKnownFolderName]::MsgFolderRoot, $Folderview)
Write-Host " - Displaying Top Level Folders for mailbox $($MailboxName)"
#List folders result
foreach ($Folder in $FoldersResult.Folders) 
{ 
    Write-Host "$([char]9)- $($Folder.DisplayName)"
}
Write-Host "--------------------"
Write-Host "-  End script      -"
Write-Host "--------------------"