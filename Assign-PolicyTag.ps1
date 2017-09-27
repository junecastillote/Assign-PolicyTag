<#	
	.NOTES
	===========================================================================
	 Created on:   	26-Sep-2017
	 Created by:   	Tito D. Castillote Jr.
					june.castillote@gmail.com
	 Filename:     	Assign-PolicyTag.ps1
	 Version:		1.0 (26-Sep-2017)
	===========================================================================

	.LINK
		http://www.lazyexchangeadmin.com/

	.SYNOPSIS
		Stamp folder with specified Policy Tag

	.DESCRIPTION
		This script requires EWS Managed API 2.2
		
	.EXAMPLE
	
#>

#Region Variables
[array]$folders = @("Notes","Contacts","Tasks","Journal","Calendar")
$policyTag = "{4e3bf873-6524-4d41-a817-d27aa4dcc27f}"
$retentionFlags = 129
$retentionPeriod = 182
$adminUsername = "user@domain.com"
$adminPassword = "Password"
#End Region
Import-Module -Name "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"

$ews = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1)

#Set the admin credentials here
$ews.Credentials = new-object Microsoft.Exchange.WebServices.Data.WebCredentials($adminUsername,$adminPassword)

$ews.Url= new-object Uri("https://outlook.office365.com/EWS/Exchange.asmx")


gc users.txt | foreach-object {
    $EmailAddress = $_

    foreach ($FolderName in $folders)
	{
	Write-host (Get-Date) ": Searching folder $($FolderName) in" $EmailAddress

    $ews.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,$EmailAddress);

    $oFolderView = new-object Microsoft.Exchange.WebServices.Data.FolderView(1)
    $oSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,$FolderName)

    $oFindFolderResults = $ews.FindFolders([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$oSearchFilter,$oFolderView)

    if ($oFindFolderResults.TotalCount -eq 0)
    {
         Write-host (Get-Date) ":Folder $($FolderName) does not exist in" $EmailAddress "-skipped"
    }
    else
    {
        Write-host (Get-Date) ": Folder $($FolderName) found in" $EmailAddress

        #0x3019
        $PR_POLICY_TAG = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x3019,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary);

        #0x301D    
        $PR_RETENTION_FLAGS = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x301D,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);
        
        #0x301A
        $PR_RETENTION_PERIOD = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x301A,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);

        $oFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($ews,$oFindFolderResults.Folders[0].Id)
       
        $oFolder.SetExtendedProperty($PR_RETENTION_FLAGS, $retentionFlags)

        $oFolder.SetExtendedProperty($PR_RETENTION_PERIOD, $retentionPeriod)

        $PR_POLICY_TAG_GUID = new-Object Guid($policyTag);

        $oFolder.SetExtendedProperty($PR_POLICY_TAG, $PR_POLICY_TAG_GUID.ToByteArray())

        $oFolder.Update()

        Write-host (Get-Date) ": Retention policy stamped!"
    }    
    $ews.ImpersonatedUserId = $null
	}    
} 