#######################################################################
#
# Powershell script for Silent Deploy Microsoft Visio 2007 Professional
#
# Script for Silent Deploy Microsoft Visio 2007 Professional.
# Script create xml file Acme.xml in $env:TEMP directory.
# Then script run Microsoft Visio 2007 Professional Setup.exe
# from en_office_visio_professional_2007_cd_x12-19212.iso with 
# parameter /Config $env:TEMP\Acme.xml.
# Afrer Silent Deploy file Acme.xml in $env:TEMP directory will be delete.
#  
# Version:        0.1
# Author:         Andrii Romanenko
# Website:        blogs.airra.net
# Creation Date:  12.09.2007
# Purpose/Change: Initial script development
#
# Run:
#   
# .\MsVisio2007ProSilentDeploy.ps1
#
#######################################################################

# Create Acme.xml in $env:TEMP directory.
$ExportFilePath = "$env:TEMP\Acme.xml"
$ProductKey = "12345-ABCDE-67890-FGHIJ-KLMNO"

$XmlDocument = New-Object System.Xml.XmlDocument
$xmlDeclaration = $XmlDocument.CreateXmlDeclaration("1.0",$Null,$Null)
$Null = $XmlDocument.AppendChild($xmlDeclaration)
$RootNode = $XmlDocument.CreateElement("Configuration")
$RootNodeAttribute = $RootNode.SetAttribute("Product","Professional")

$DisplayNode = $XmlDocument.CreateElement("Display")
$DisplayAttribute = $DisplayNode.SetAttribute("Level","None")
$DisplayAttribute = $DisplayNode.SetAttribute("CompletionNotice","no")
$DisplayAttribute = $DisplayNode.SetAttribute("SuppressModal","yes")
$DisplayAttribute = $DisplayNode.SetAttribute("AcceptEula","yes")

$PIDKEYNode = $XmlDocument.CreateElement("PIDKEY")
$PIDKEYAttribute = $PIDKEYNode.SetAttribute("Value",$ProductKey)

$Null = $RootNode.AppendChild($DisplayNode)
$Null = $RootNode.AppendChild($PIDKEYNode)
$Null = $XmlDocument.AppendChild($RootNode)
$XmlDocument.Save($ExportFilePath)

# Begin Silent Deploy Microsoft Visio 2007 Professional
$Message = "Begin Silent Setup Microsoft Visio 2007 Professional"
Write-Host -ForegroundColor Green $Message

$ISOPath = "E:\" 
$Params = " /config " + $env:TEMP + "\Acme.xml" 
$Command = $ISOPath + "Setup.exe" + $Params

Invoke-Expression $Command 
 
Write-Host "." -NoNewline
Start-Sleep -Seconds 60

# Checking if the application installation was successful 
Do {
         Write-Host "." -NoNewline
         Start-Sleep -Seconds 10
    } Until(Get-EventLog -LogName Application -Newest 10 | Where-Object {$_.Message -Like "*Microsoft Office Visio Professional 2007*Installation*successfully*"})

$Message = "`nSilent Setup Microsoft Visio 2007 Professional Succesfully!"
Write-Host -ForegroundColor Green $Message

# Disable the Office 2007 Welcome screen
Set-Location HKCU:
New-Item -Path "Software\Microsoft\Office\12.0\Common\" -Name "General" | Out-Null
New-ItemProperty -Path "Software\Microsoft\Office\12.0\Common\General" -Name "ShownOptIn" -Value 00000001 -PropertyType DWORD | Out-Null

$Message = "`nDisable the Office 2007 Welcome screen Succesfully!"
Write-Host -ForegroundColor Green $Message

#Removing Acme.xml
Set-Location c:
Remove-Item -Path ($env:TEMP+ "\Acme.xml") -Force

$Message = "`nRemoving Acme.xml Succesfully!"
Write-Host -ForegroundColor Green $Message
