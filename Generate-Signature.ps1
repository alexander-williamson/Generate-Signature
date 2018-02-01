# This script makes Outlook email signature from AD info
# Add this to GPO logon script. Place script and template files to sysvol share.
# This can handle html and txt templates at same time.
#
# Show all object properties: $objUser | Select-Object -Property *
$strName = $env:username
$strFilePrefix = "Company"
$arrItemsToReplace = @('FullName','title','telephoneNumber','mail') # what to replace from template
# title, description, postalCode, telephoneNumber, givenName, displayName, streetAddress, name, mail, mobile
$arrFileFormats = @('htm', 'txt') # do html and txt templates

# AD Filtering and fetching into $objUser
$objSearcher = New-Object System.DirectoryServices.DirectorySearcher
$objSearcher.Filter = "(&(objectCategory=User)(samAccountName=$strName))"
$objUser = $objSearcher.FindOne().GetDirectoryEntry()

# Set signature path and file name
$FolderLocation = [string]$Env:appdata + '\Microsoft\Signatures\'  
# if path doesn't exist then create it
if ( !$(Try { Test-Path $FolderLocation.trim() } Catch { $false }) ) { New-Item $FolderLocation -force }

# Go thru each desired formats
foreach ( $strFormat in $arrFileFormats ) {

	$fullOutputPath = $FolderLocation + $strFilePrefix + " " + [string]$objUser.name + ".$strFormat"

	# read template file
	if ( !$(Try { Test-Path ".\Template.$strFormat" } Catch { $false }) ) { continue }
	$strSignature = (Get-Content ".\Template.$strFormat")

	# replace content in template
	foreach ($element in $arrItemsToReplace) { 
		$strSignature = $strSignature -replace "%%$element%%",[string]$objUser.$element 
	}

	# write new signature
	Set-Content -path $fullOutputPath $strSignature

	# output the path
	"Successfully generated signature to: $fullOutputPath"

}
