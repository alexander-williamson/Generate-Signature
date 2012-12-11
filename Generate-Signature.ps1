$strName = $env:username

# AD Filtering and fetching into $objUser
$strFilter = "(&(objectCategory=User)(samAccountName=$strName))"
$objSearcher = New-Object System.DirectoryServices.DirectorySearcher
$objSearcher.Filter = $strFilter
$objPath = $objSearcher.FindOne()
$objUser = $objPath.GetDirectoryEntry()


$strFullname = $objUser.FullName
$arrFullnameParts = $strFullname.split(" ")
$strFirstname = $arrFullnameParts[0]
$strLastname = $arrFullnameParts[1]

$strName = $objUser.FullName
$strTitle = $objUser.Title
# $strCompany = $objUser.Company
# $strCred = $objUser.info
# $strStreet = $objUser.StreetAddress
$strMobile = $objUser.telephoneNumber
# $strCity =  $objUser.l
# $strPostCode = $objUser.PostalCode
# $strCountry = $objUser.co
$strEmail = if ($objUser.mail) { $objUser.mail.ToLower() }

# $strWebsite = $objUser.wWWHomePage

$strDepartment = "IT"

$UserDataPath = $Env:appdata

# todo: auto detect keys
# if (test-path "HKCU:\\Software\\Microsoft\\Office\\11.0\\Common\\General") {
#   get-item -path HKCU:\\Software\\Microsoft\\Office\\11.0\\Common\\General | new-Itemproperty -name Signatures -value signaturesCompany -propertytype string -force
#   
# 
# if (test-path "HKCU:\\Software\\Microsoft\\Office\\12.0\\Common\\General") {
#   get-item -path HKCU:\\Software\\Microsoft\\Office\\12.0\\Common\\General | new-Itemproperty -name Signatures -value signaturesCompany -propertytype string -force
# }
# if (test-path "HKCU:\\Software\\Microsoft\\Office\\14.0\\Common\\General") {
#  get-item -path HKCU:\\Software\\Microsoft\\Office\\14.0\\Common\\General | new-Itemproperty -name Signatures -value signaturesCompany -propertytype string -force
# }

$FolderLocation = $UserDataPath + '\Microsoft\Signatures\'  
mkdir $FolderLocation -force

$fullOutputPath = $FolderLocation + $strFullname + ".htm"

# replace content in template
# todo: refactor
(Get-Content .\Template.html) | Foreach-Object {$_ -replace "%%FIRSTNAME%%", $strFirstname} | Foreach-Object {$_ -replace "%%LASTNAME%%", $strLastname} | Foreach-Object {$_ -replace "%%TITLE%%", $strTitle} | Foreach-Object {$_ -replace "%%DEPARTMENT%%", $strDepartment} | Foreach-Object {$_ -replace "%%MOBILE%%", $strMobile} | Foreach-Object {$_ -replace "%%EMAIL%%", $strEmail} | Set-Content $fullOutputPath

# output the path
""
"Successfully generated signature to: "
$fullOutputPath
""
"Press any key to continue..."
Read-Host