$strName = $env:username
$strFilePrefix = "Company"

# AD Filtering and fetching into $objUser
$objSearcher = New-Object System.DirectoryServices.DirectorySearcher
$objSearcher.Filter = "(&(objectCategory=User)(samAccountName=$strName))"
$objUser = $objSearcher.FindOne().GetDirectoryEntry()

# Set signature path and file name
$FolderLocation = [string]$Env:appdata + '\Microsoft\Signatures\'  
# if path doesn't exist then create it
if ( !$(Try { Test-Path $FolderLocation.trim() } Catch { $false }) ) { New-Item $FolderLocation -force }
$fullOutputPath = $FolderLocation + $strFilePrefix + " " + [string]$objUser.name + ".htm"

# read template file
$strSignature = (Get-Content .\Template.html)

# replace content in template
foreach ($element in @('FullName','title','telephoneNumber','mail')) { 
	$strSignature = $strSignature -replace "%%$element%%",[string]$objUser.$element 
}

# write new signature
Set-Content -path $fullOutputPath $strSignature

# output the path
"Successfully generated signature to: $fullOutputPath"
