###################################################################
#
# Powershell script for generate Notepad++ userDefineLang.xml
#
# Script for Notepad++ Powershell User Define Language support.
# Script create xml file User Define Language - userDefineLang.xml 
# in %APPDATA%\Notepad++\ directory.
#  
# Version:        0.1
# Author:         Andrii Romanenko
# Website:        blogs.airra.net
# Creation Date:  12.09.2007
# Purpose/Change: Initial script development
#
# Run:
#  
# .\NppPowershellLanguageSupport.ps1
#
####################################################################

$ExportFilePath = "$env:APPDATA\Notepad++\userDefineLang.xml"

$XmlDocument = New-Object System.Xml.XmlDocument
$xmlDeclaration = $XmlDocument.CreateXmlDeclaration("1.0",$Null,$Null)
$Null = $XmlDocument.AppendChild($xmlDeclaration)
$RootNode = $XmlDocument.CreateElement("NotepadPlus")


$UserLangNode = $XmlDocument.CreateElement("UserLang")
$UserLangAttribute = $UserLangNode.SetAttribute("name","PowerShell")
$UserLangAttribute = $UserLangNode.SetAttribute("ext","ps1")

$SettingsNode = $XmlDocument.CreateElement("Settings")

$GlobalNode = $XmlDocument.CreateElement("Global")
$GlobalAttribute = $GlobalNode.SetAttribute("caseIgnored","yes")

$TreatAsSymbolNode = $XmlDocument.CreateElement("TreatAsSymbol")
$TreatAsSymbolAttribute = $TreatAsSymbolNode.SetAttribute("comment","no")
$TreatAsSymbolAttribute = $TreatAsSymbolNode.SetAttribute("commentLine","yes")


$PrefixNode = $XmlDocument.CreateElement("PrefixNode")
$PrefixAttribute = $PrefixNode.SetAttribute("words1","yes")
$PrefixAttribute = $PrefixNode.SetAttribute("words2","no")
$PrefixAttribute = $PrefixNode.SetAttribute("words3","no")
$PrefixAttribute = $PrefixNode.SetAttribute("words4","no")


$Null = $SettingsNode.AppendChild($GlobalNode)
$Null = $SettingsNode.AppendChild($TreatAsSymbolNode)
$Null = $SettingsNode.AppendChild($PrefixNode)

$Null = $UserLangNode.AppendChild($SettingsNode)

$KeywordListsNode = $XmlDocument.CreateElement("KeywordLists")

$KeywordsNode1 = $XmlDocument.CreateElement("Keywords")
$KeywordsAttribute = $KeywordsNode1.SetAttribute("name","Delimiters")
$KeywordsTextNode1 = $XmlDocument.CreateTextNode(""00"00")
$Null = $KeywordsNode1.AppendChild($KeywordsTextNode1)

$KeywordsNode2 = $XmlDocument.CreateElement("Keywords")
$KeywordsAttribute = $KeywordsNode2.SetAttribute("name","Folder+")
$KeywordsTextNode2 = $XmlDocument.CreateTextNode("{")
$Null = $KeywordsNode2.AppendChild($KeywordsTextNode2)

$KeywordsNode3 = $XmlDocument.CreateElement("Keywords")
$KeywordsAttribute = $KeywordsNode3.SetAttribute("name","Folder-")
$KeywordsTextNode3 = $XmlDocument.CreateTextNode("}")
$Null = $KeywordsNode3.AppendChild($KeywordsTextNode3)

$KeywordsNode4 = $XmlDocument.CreateElement("Keywords")
$KeywordsAttribute = $KeywordsNode4.SetAttribute("name","Operators")
$OperatorsSymbols = '! " % & ( ) * , / : ; ? @ [ \ ] ^ ` | ~ + < = >'
$KeywordsTextNode4 = $XmlDocument.CreateTextNode($OperatorsSymbols)
$Null = $KeywordsNode4.AppendChild($KeywordsTextNode4)

$KeywordsNode5 = $XmlDocument.CreateElement("Keywords")
$KeywordsAttribute = $KeywordsNode5.SetAttribute("name","Comment")
$KeywordsTextNode5 = $XmlDocument.CreateTextNode("1 2 0#")
$Null = $KeywordsNode5.AppendChild($KeywordsTextNode5)

$KeywordsNode6 = $XmlDocument.CreateElement("Keywords")
$KeywordsAttribute = $KeywordsNode6.SetAttribute("name","Words1")
$KeywordsTextNode6 = $XmlDocument.CreateTextNode("Add-Content Add-History Add-Member Add-PSSnapin Clear-Content Clear-Item Clear-ItemProperty Clear-Variable Compare-Object ConvertFrom-SecureString Convert-Path ConvertTo-Html ConvertTo-SecureString Copy-Item Copy-ItemProperty Export-Alias Export-Clixml Export-Console Export-Csv ForEach-Object Format-Custom Format-List Format-Table Format-Wide Get-Acl Get-Alias Get-AuthenticodeSignature Get-ChildItem Get-Command Get-Content Get-Credential Get-Culture Get-Date Get-EventLog Get-ExecutionPolicy Get-Help Get-History Get-Host Get-Item Get-ItemProperty Get-Location Get-Member Get-PfxCertificate Get-Process Get-PSDrive Get-PSProvider Get-PSSnapin Get-Service Get-TraceSource Get-UICulture Get-Unique Get-Variable Get-WmiObject Group-Object Import-Alias Import-Clixml Import-Csv Invoke-Expression Invoke-History Invoke-Item Join-Path Measure-Command Measure-Object Move-Item Move-ItemProperty New-Alias New-Item New-ItemProperty New-Object New-PSDrive New-Service New-TimeSpan New-Variable Out-Default Out-File Out-Host Out-Null Out-Printer Out-String Pop-Location Push-Location Read-Host Remove-Item Remove-ItemProperty Remove-PSDrive Remove-PSSnapin Remove-Variable Rename-Item Rename-ItemProperty Resolve-Path Restart-Service Resume-Service Select-Object Select-String Set-Acl Set-Alias Set-AuthenticodeSignature Set-Content Set-Date Set-ExecutionPolicy Set-Item Set-ItemProperty Set-Location Set-PSDebug Set-Service Set-TraceSource Set-Variable Sort-Object Split-Path Start-Service Start-Sleep Start-Transcript Stop-Process Stop-Service Stop-Transcript Suspend-Service Tee-Object Test-Path Trace-Command Update-FormatData Update-TypeData Where-Object Write-Debug Write-Error Write-Host Write-Output Write-Progress Write-Verbose Write-Warning switch function if throw else while break - $ System. Management.")
$Null = $KeywordsNode6.AppendChild($KeywordsTextNode6)

$KeywordsNode7 = $XmlDocument.CreateElement("Keywords")
$KeywordsAttribute = $KeywordsNode7.SetAttribute("name","Words2")
$KeywordsTextNode7 = $XmlDocument.CreateTextNode("ac asnp clc cli clp clv cpi cpp cvpa diff epal epcsv fc fl foreach % ft fw gal gc gci gcm gdr ghy gi gl gm gp gps group gsv gsnp gu gv gwmi iex ihy ii ipal ipcsv mi mp nal ndr ni nv oh rdr ri rni rnp rp rsnp rv rvpa sal sasv sc select si sl sleep sort sp spps spsv sv tee where ? write cat cd clear cp h history kill lp ls mount mv popd ps pushd pwd r rm rmdir echo cls chdir copy del dir erase move rd ren set type")
$Null = $KeywordsNode7.AppendChild($KeywordsTextNode7)

$KeywordsNode8 = $XmlDocument.CreateElement("Keywords")
$KeywordsAttribute = $KeywordsNode8.SetAttribute("name","Words3")
$Words3 = "CIM_DataFile CIM_DirectoryContainsFile CIM_ProcessExecutable CIM_VideoControllerResolution Msft_Providers Msft_WmiProvider_Counters NetDiagnostics Win32_1394Controller Win32_1394ControllerDevice Win32_AccountSID Win32_ActionCheck Win32_ActiveRoute Win32_AllocatedResource Win32_ApplicationCommandLine Win32_ApplicationService Win32_AssociatedBattery Win32_AssociatedProcessorMemory Win32_AutochkSetting Win32_BaseBoard Win32_Battery Win32_Binary Win32_BindImageAction Win32_BIOS Win32_BootConfiguration Win32_Bus Win32_CacheMemory Win32_CDROMDrive Win32_CheckCheck Win32_CIMLogicalDeviceCIMDataFile Win32_ClassicCOMApplicationClasses Win32_ClassicCOMClass Win32_ClassicCOMClassSetting Win32_ClassicCOMClassSettings Win32_ClassInfoAction Win32_ClientApplicationSetting Win32_CodecFile Win32_COMApplicationSettings Win32_ComClassAutoEmulator Win32_ComClassEmulator Win32_CommandLineAccess Win32_ComponentCategory Win32_ComputerSystem Win32_ComputerSystemProcessor Win32_ComputerSystemProduct Win32_ComputerSystemWindowsProductActivationSetting Win32_Condition Win32_ConnectionShare Win32_ControllerHasHub Win32_CreateFolderAction Win32_CurrentProbe Win32_DCOMApplication Win32_DCOMApplicationAccessAllowedSetting Win32_DCOMApplicationLaunchAllowedSetting Win32_DCOMApplicationSetting Win32_DependentService Win32_Desktop Win32_DesktopMonitor Win32_DeviceBus Win32_DeviceMemoryAddress Win32_DfsNode Win32_DfsNodeTarget Win32_DfsTarget Win32_Directory Win32_DirectorySpecification Win32_DiskDrive Win32_DiskDrivePhysicalMedia Win32_DiskDriveToDiskPartition Win32_DiskPartition Win32_DiskQuota Win32_DisplayConfiguration Win32_DisplayControllerConfiguration Win32_DMAChannel Win32_DriverForDevice Win32_DriverVXD Win32_DuplicateFileAction Win32_Environment Win32_EnvironmentSpecification Win32_ExtensionInfoAction Win32_Fan Win32_FileSpecification Win32_FloppyController Win32_FloppyDrive Win32_FontInfoAction Win32_Group Win32_GroupInDomain Win32_GroupUser Win32_HeatPipe Win32_IDEController Win32_IDEControllerDevice Win32_ImplementedCategory Win32_InfraredDevice Win32_IniFileSpecification Win32_InstalledSoftwareElement Win32_IP4PersistedRouteTable Win32_IP4RouteTable Win32_IRQResource Win32_Keyboard Win32_LaunchCondition Win32_LoadOrderGroup Win32_LoadOrderGroupServiceDependencies Win32_LoadOrderGroupServiceMembers Win32_LocalTime Win32_LoggedOnUser Win32_LogicalDisk Win32_LogicalDiskRootDirectory Win32_LogicalDiskToPartition Win32_LogicalFileAccess Win32_LogicalFileAuditing Win32_LogicalFileGroup Win32_LogicalFileOwner Win32_LogicalFileSecuritySetting Win32_LogicalMemoryConfiguration Win32_LogicalProgramGroup Win32_LogicalProgramGroupDirectory Win32_LogicalProgramGroupItem Win32_LogicalProgramGroupItemDataFile Win32_LogicalShareAccess Win32_LogicalShareAuditing Win32_LogicalShareSecuritySetting Win32_LogonSession Win32_LogonSessionMappedDisk Win32_MappedLogicalDisk Win32_MemoryArray Win32_MemoryArrayLocation Win32_MemoryDevice Win32_MemoryDeviceArray Win32_MemoryDeviceLocation Win32_MIMEInfoAction Win32_MotherboardDevice Win32_MountPoint Win32_MoveFileAction Win32_NamedJobObject Win32_NamedJobObjectActgInfo Win32_NamedJobObjectLimit Win32_NamedJobObjectLimitSetting Win32_NamedJobObjectProcess Win32_NamedJobObjectSecLimit Win32_NamedJobObjectSecLimitSetting Win32_NamedJobObjectStatistics Win32_NetworkAdapter Win32_NetworkAdapterConfiguration Win32_NetworkAdapterSetting Win32_NetworkClient Win32_NetworkConnection Win32_NetworkLoginProfile Win32_NetworkProtocol Win32_NTDomain Win32_NTEventlogFile Win32_NTLogEvent Win32_NTLogEventComputer Win32_NTLogEventLog Win32_NTLogEventUser Win32_ODBCAttribute Win32_ODBCDataSourceAttribute Win32_ODBCDataSourceSpecification Win32_ODBCDriverAttribute Win32_ODBCDriverSoftwareElement Win32_ODBCDriverSpecification Win32_ODBCSourceAttribute Win32_ODBCTranslatorSpecification Win32_OnBoardDevice Win32_OperatingSystem Win32_OperatingSystemAutochkSetting Win32_OperatingSystemQFE Win32_OSRecoveryConfiguration Win32_PageFile Win32_PageFileElementSetting Win32_PageFileSetting Win32_PageFileUsage Win32_ParallelPort Win32_Patch Win32_PatchFile Win32_PatchPackage Win32_PCMCIAController Win32_PerfFormattedData_ContentFilter_IndexingServiceFilter Win32_PerfFormattedData_ContentIndex_IndexingService Win32_PerfFormattedData_Fax_FaxServices Win32_PerfFormattedData_IPSec_IPSecv4Driver Win32_PerfFormattedData_IPSec_IPSecv4IKE Win32_PerfFormattedData_ISAPISearch_HttpIndexingService Win32_PerfFormattedData_MSDTC_DistributedTransactionCoordinator Win32_PerfFormattedData_NETFramework_NETCLRExceptions Win32_PerfFormattedData_NETFramework_NETCLRInterop Win32_PerfFormattedData_NETFramework_NETCLRJit Win32_PerfFormattedData_NETFramework_NETCLRLoading Win32_PerfFormattedData_NETFramework_NETCLRLocksAndThreads Win32_PerfFormattedData_NETFramework_NETCLRMemory Win32_PerfFormattedData_NETFramework_NETCLRRemoting Win32_PerfFormattedData_NETFramework_NETCLRSecurity Win32_PerfFormattedData_NTDS_NTDS Win32_PerfFormattedData_PerfDisk_LogicalDisk Win32_PerfFormattedData_PerfDisk_PhysicalDisk Win32_PerfFormattedData_PerfNet_Browser Win32_PerfFormattedData_PerfNet_Redirector Win32_PerfFormattedData_PerfNet_Server Win32_PerfFormattedData_PerfNet_ServerWorkQueues Win32_PerfFormattedData_PerfOS_Cache Win32_PerfFormattedData_PerfOS_Memory Win32_PerfFormattedData_PerfOS_Objects Win32_PerfFormattedData_PerfOS_PagingFile Win32_PerfFormattedData_PerfOS_Processor Win32_PerfFormattedData_PerfOS_System Win32_PerfFormattedData_PerfProc_FullImage_Costly Win32_PerfFormattedData_PerfProc_Image_Costly Win32_PerfFormattedData_PerfProc_JobObject Win32_PerfFormattedData_PerfProc_JobObjectDetails Win32_PerfFormattedData_PerfProc_Process Win32_PerfFormattedData_PerfProc_ProcessAddressSpace_Costly Win32_PerfFormattedData_PerfProc_Thread Win32_PerfFormattedData_PerfProc_ThreadDetails_Costly Win32_PerfFormattedData_RemoteAccess_RASPort Win32_PerfFormattedData_RemoteAccess_RASTotal Win32_PerfFormattedData_RSVP_ACSRSVPInterfaces Win32_PerfFormattedData_RSVP_ACSRSVPService Win32_PerfFormattedData_Spooler_PrintQueue Win32_PerfFormattedData_TapiSrv_Telephony Win32_PerfFormattedData_Tcpip_ICMP Win32_PerfFormattedData_Tcpip_ICMPv6 Win32_PerfFormattedData_Tcpip_IP Win32_PerfFormattedData_Tcpip_IPv4 Win32_PerfFormattedData_Tcpip_IPv6 Win32_PerfFormattedData_Tcpip_NBTConnection Win32_PerfFormattedData_Tcpip_NetworkInterface Win32_PerfFormattedData_Tcpip_TCP Win32_PerfFormattedData_Tcpip_TCPv4 Win32_PerfFormattedData_Tcpip_TCPv6 Win32_PerfFormattedData_Tcpip_UDP Win32_PerfFormattedData_Tcpip_UDPv4 Win32_PerfFormattedData_Tcpip_UDPv6 Win32_PerfFormattedData_TermService_TerminalServices Win32_PerfFormattedData_TermService_TerminalServicesSession Win32_PerfRawData_ASP_ActiveServerPages Win32_PerfRawData_ContentFilter_IndexingServiceFilter Win32_PerfRawData_ContentIndex_IndexingService Win32_PerfRawData_Fax_FaxServices Win32_PerfRawData_FileReplicaConn_FileReplicaConn Win32_PerfRawData_FileReplicaSet_FileReplicaSet Win32_PerfRawData_IAS_IASAccountingClients Win32_PerfRawData_IAS_IASAccountingServer Win32_PerfRawData_IAS_IASAuthenticationClients Win32_PerfRawData_IAS_IASAuthenticationServer Win32_PerfRawData_InetInfo_InternetInformationServicesGlobal Win32_PerfRawData_IPSec_IPSecv4Driver Win32_PerfRawData_IPSec_IPSecv4IKE Win32_PerfRawData_ISAPISearch_HttpIndexingService Win32_PerfRawData_MSDTC_DistributedTransactionCoordinator Win32_PerfRawData_NETFramework_NETCLRExceptions Win32_PerfRawData_NETFramework_NETCLRInterop Win32_PerfRawData_NETFramework_NETCLRJit Win32_PerfRawData_NETFramework_NETCLRLoading Win32_PerfRawData_NETFramework_NETCLRLocksAndThreads Win32_PerfRawData_NETFramework_NETCLRMemory Win32_PerfRawData_NETFramework_NETCLRRemoting Win32_PerfRawData_NETFramework_NETCLRSecurity Win32_PerfRawData_NTDS_NTDS Win32_PerfRawData_NTFSDRV_SMTPNTFSStoreDriver Win32_PerfRawData_PerfDisk_LogicalDisk Win32_PerfRawData_PerfDisk_PhysicalDisk Win32_PerfRawData_PerfNet_Browser Win32_PerfRawData_PerfNet_Redirector Win32_PerfRawData_PerfNet_Server Win32_PerfRawData_PerfNet_ServerWorkQueues Win32_PerfRawData_PerfOS_Cache "
$Words3 = $Words3 + "Win32_PerfRawData_PerfOS_Memory Win32_PerfRawData_PerfOS_Objects Win32_PerfRawData_PerfOS_PagingFile Win32_PerfRawData_PerfOS_Processor Win32_PerfRawData_PerfOS_System Win32_PerfRawData_PerfProc_FullImage_Costly Win32_PerfRawData_PerfProc_Image_Costly Win32_PerfRawData_PerfProc_JobObject Win32_PerfRawData_PerfProc_JobObjectDetails Win32_PerfRawData_PerfProc_Process Win32_PerfRawData_PerfProc_ProcessAddressSpace_Costly Win32_PerfRawData_PerfProc_Thread Win32_PerfRawData_PerfProc_ThreadDetails_Costly Win32_PerfRawData_RemoteAccess_RASPort Win32_PerfRawData_RemoteAccess_RASTotal Win32_PerfRawData_RSVP_ACSPerRSVPService Win32_PerfRawData_RSVP_ACSRSVPInterfaces Win32_PerfRawData_RSVP_ACSRSVPService Win32_PerfRawData_SMTPSVC_SMTPServer Win32_PerfRawData_Spooler_PrintQueue Win32_PerfRawData_TapiSrv_Telephony Win32_PerfRawData_Tcpip_ICMP Win32_PerfRawData_Tcpip_ICMPv6 Win32_PerfRawData_Tcpip_IP Win32_PerfRawData_Tcpip_IPv4 Win32_PerfRawData_Tcpip_IPv6 Win32_PerfRawData_Tcpip_NBTConnection Win32_PerfRawData_Tcpip_NetworkInterface Win32_PerfRawData_Tcpip_TCP Win32_PerfRawData_Tcpip_TCPv4 Win32_PerfRawData_Tcpip_TCPv6 Win32_PerfRawData_Tcpip_UDP Win32_PerfRawData_Tcpip_UDPv4 Win32_PerfRawData_Tcpip_UDPv6 Win32_PerfRawData_TermService_TerminalServices Win32_PerfRawData_TermService_TerminalServicesSession Win32_PerfRawData_W3SVC_WebService Win32_PhysicalMedia Win32_PhysicalMemory Win32_PhysicalMemoryArray Win32_PhysicalMemoryLocation Win32_PingStatus Win32_PNPAllocatedResource Win32_PnPDevice Win32_PnPEntity Win32_PnPSignedDriver Win32_PnPSignedDriverCIMDataFile Win32_PointingDevice Win32_PortableBattery Win32_PortConnector Win32_PortResource Win32_POTSModem Win32_POTSModemToSerialPort Win32_Printer Win32_PrinterConfiguration Win32_PrinterController Win32_PrinterDriver Win32_PrinterDriverDll Win32_PrinterSetting Win32_PrinterShare Win32_PrintJob Win32_Process Win32_Processor Win32_ProductOptional Win32_ProductCheck Win32_ProductResource Win32_ProductSoftwareFeatures Win32_ProgIDSpecification Win32_ProgramGroup Win32_ProgramGroupContents Win32_Property Win32_ProtocolBinding Win32_Proxy Win32_PublishComponentAction Win32_QuickFixEngineering Win32_QuotaSetting Win32_Refrigeration Win32_Registry Win32_RegistryAction Win32_RemoveFileAction Win32_RemoveIniAction Win32_ReserveCost Win32_ScheduledJob Win32_SCSIController Win32_SCSIControllerDevice Win32_SecuritySettingOfLogicalFile Win32_SecuritySettingOfLogicalShare Win32_SelfRegModuleAction Win32_SerialPort Win32_SerialPortConfiguration Win32_SerialPortSetting Win32_ServerConnection Win32_ServerSession Win32_Service Win32_ServiceControl Win32_ServiceSpecification Win32_ServiceSpecificationService Win32_SessionConnection Win32_SessionProcess Win32_ShadowBy Win32_ShadowCopy Win32_ShadowDiffVolumeSupport Win32_ShadowFor Win32_ShadowOn Win32_ShadowProvider Win32_ShadowStorage Win32_ShadowVolumeSupport Win32_Share Win32_ShareToDirectory Win32_ShortcutAction Win32_ShortcutFile Win32_ShortcutSAP Win32_SID Win32_SoftwareElement Win32_SoftwareElementAction Win32_SoftwareElementCheck Win32_SoftwareElementCondition Win32_SoftwareElementResource Win32_SoftwareFeature Win32_SoftwareFeatureAction Win32_SoftwareFeatureCheck Win32_SoftwareFeatureParent Win32_SoftwareFeatureSoftwareElements Win32_SoundDevice Win32_StartupCommand Win32_SubDirectory Win32_SystemAccount Win32_SystemBIOS Win32_SystemBootConfiguration Win32_SystemDesktop Win32_SystemDevices Win32_SystemDriver Win32_SystemDriverPNPEntity Win32_SystemEnclosure Win32_SystemLoadOrderGroups Win32_SystemLogicalMemoryConfiguration Win32_SystemNetworkConnections Win32_SystemOperatingSystem Win32_SystemPartitions Win32_SystemProcesses Win32_SystemProgramGroups Win32_SystemResources Win32_SystemServices Win32_SystemSlot Win32_SystemSystemDriver Win32_SystemTimeZone Win32_SystemUsers Win32_TapeDrive Win32_TCPIPPrinterPort Win32_TemperatureProbe Win32_Terminal Win32_TerminalService Win32_TerminalServiceSetting Win32_TerminalServiceToSetting Win32_TerminalTerminalSetting Win32_Thread Win32_TimeZone Win32_TSAccount Win32_TSClientSetting Win32_TSEnvironmentSetting Win32_TSGeneralSetting Win32_TSLogonSetting Win32_TSNetworkAdapterListSetting Win32_TSNetworkAdapterSetting Win32_TSPermissionsSetting Win32_TSRemoteControlSetting  Win32_TSSessionDirectory Win32_TSSessionDirectorySetting Win32_TSSessionSetting Win32_TypeLibraryAction Win32_UninterruptiblePowerSupply Win32_USBController Win32_USBControllerDevice Win32_USBHub Win32_UserAccount Win32_UserDesktop Win32_UserInDomain Win32_UTCTime Win32_VideoConfiguration Win32_VideoController Win32_VideoSettings Win32_VoltageProbe Win32_Volume Win32_VolumeQuota Win32_VolumeQuotaSetting Win32_VolumeUserQuota Win32_WindowsProductActivation Win32_WMIElementSetting Win32_WMISetting"
$KeywordsTextNode8 = $XmlDocument.CreateTextNode($Words3)
$Null = $KeywordsNode8.AppendChild($KeywordsTextNode8)

$KeywordsNode9 = $XmlDocument.CreateElement("Keywords")
$KeywordsAttribute = $KeywordsNode9.SetAttribute("name","Words4")
$Words4 = @'
$ $$ $? $^ $Args $DebugPreference $Error $ErrorActionPreference $false $foreach $HistorySize $HOME $Host $Input $LastExitCode $Match $MshHome $MyInvocation $null $OFS $PSCommandPath $ShellID $StackTrace $true $_
'@
$KeywordsTextNode9 = $XmlDocument.CreateTextNode($Words4)
$Null = $KeywordsNode9.AppendChild($KeywordsTextNode9)

$Null = $KeywordListsNode.AppendChild($KeywordsNode1)
$Null = $KeywordListsNode.AppendChild($KeywordsNode2)
$Null = $KeywordListsNode.AppendChild($KeywordsNode3)
$Null = $KeywordListsNode.AppendChild($KeywordsNode4)
$Null = $KeywordListsNode.AppendChild($KeywordsNode5)
$Null = $KeywordListsNode.AppendChild($KeywordsNode6)
$Null = $KeywordListsNode.AppendChild($KeywordsNode7)
$Null = $KeywordListsNode.AppendChild($KeywordsNode8)
$Null = $KeywordListsNode.AppendChild($KeywordsNode9)
$Null = $UserLangNode.AppendChild($KeywordListsNode)

$StylesNode = $XmlDocument.CreateElement("Styles")

$WordsStyleNode1 = $XmlDocument.CreateElement("WordsStyle")
$WordsStyleAttribute1 = $WordsStyleNode1.SetAttribute("name","DEFAULT")
$WordsStyleAttribute1 = $WordsStyleNode1.SetAttribute("fontStyle","0")
$WordsStyleAttribute1 = $WordsStyleNode1.SetAttribute("fontName","")
$WordsStyleAttribute1 = $WordsStyleNode1.SetAttribute("bgColor","FFFFFF")
$WordsStyleAttribute1 = $WordsStyleNode1.SetAttribute("fgColor","000000")
$WordsStyleAttribute1 = $WordsStyleNode1.SetAttribute("styleID","11")

$WordsStyleNode2 = $XmlDocument.CreateElement("WordsStyle")
$WordsStyleAttribute2 = $WordsStyleNode2.SetAttribute("name","FOLDEROPEN")
$WordsStyleAttribute2 = $WordsStyleNode2.SetAttribute("fontStyle","0")
$WordsStyleAttribute2 = $WordsStyleNode2.SetAttribute("fontName","")
$WordsStyleAttribute2 = $WordsStyleNode2.SetAttribute("bgColor","FFFFFF")
$WordsStyleAttribute2 = $WordsStyleNode2.SetAttribute("fgColor","000000")
$WordsStyleAttribute2 = $WordsStyleNode2.SetAttribute("styleID","12")

$WordsStyleNode3 = $XmlDocument.CreateElement("WordsStyle")
$WordsStyleAttribute3 = $WordsStyleNode3.SetAttribute("name","FOLDERCLOSE")
$WordsStyleAttribute3 = $WordsStyleNode3.SetAttribute("fontStyle","0")
$WordsStyleAttribute3 = $WordsStyleNode3.SetAttribute("fontName","")
$WordsStyleAttribute3 = $WordsStyleNode3.SetAttribute("bgColor","FFFFFF")
$WordsStyleAttribute3 = $WordsStyleNode3.SetAttribute("fgColor","000000")
$WordsStyleAttribute3 = $WordsStyleNode3.SetAttribute("styleID","13")

$WordsStyleNode4 = $XmlDocument.CreateElement("WordsStyle")
$WordsStyleAttribute4 = $WordsStyleNode4.SetAttribute("name","KEYWORD1")
$WordsStyleAttribute4 = $WordsStyleNode4.SetAttribute("fontStyle","0")
$WordsStyleAttribute4 = $WordsStyleNode4.SetAttribute("fontName","")
$WordsStyleAttribute4 = $WordsStyleNode4.SetAttribute("bgColor","FFFFFF")
$WordsStyleAttribute4 = $WordsStyleNode4.SetAttribute("fgColor","0000FF")
$WordsStyleAttribute4 = $WordsStyleNode4.SetAttribute("styleID","5")

$WordsStyleNode5 = $XmlDocument.CreateElement("WordsStyle")
$WordsStyleAttribute5 = $WordsStyleNode5.SetAttribute("name","KEYWORD2")
$WordsStyleAttribute5 = $WordsStyleNode5.SetAttribute("fontStyle","0")
$WordsStyleAttribute5 = $WordsStyleNode5.SetAttribute("fontName","")
$WordsStyleAttribute5 = $WordsStyleNode5.SetAttribute("bgColor","FFFFFF")
$WordsStyleAttribute5 = $WordsStyleNode5.SetAttribute("fgColor","8000FF")
$WordsStyleAttribute5 = $WordsStyleNode5.SetAttribute("styleID","6")

$WordsStyleNode6 = $XmlDocument.CreateElement("WordsStyle")
$WordsStyleAttribute6 = $WordsStyleNode6.SetAttribute("name","KEYWORD3")
$WordsStyleAttribute6 = $WordsStyleNode6.SetAttribute("fontStyle","0")
$WordsStyleAttribute6 = $WordsStyleNode6.SetAttribute("fontName","")
$WordsStyleAttribute6 = $WordsStyleNode6.SetAttribute("bgColor","FFFFFF")
$WordsStyleAttribute6 = $WordsStyleNode6.SetAttribute("fgColor","FF00FF")
$WordsStyleAttribute6 = $WordsStyleNode6.SetAttribute("styleID","7")

$WordsStyleNode7 = $XmlDocument.CreateElement("WordsStyle")
$WordsStyleAttribute7 = $WordsStyleNode7.SetAttribute("name","KEYWORD4")
$WordsStyleAttribute7 = $WordsStyleNode7.SetAttribute("fontStyle","0")
$WordsStyleAttribute7 = $WordsStyleNode7.SetAttribute("fontName","")
$WordsStyleAttribute7 = $WordsStyleNode7.SetAttribute("bgColor","FFFFFF")
$WordsStyleAttribute7 = $WordsStyleNode7.SetAttribute("fgColor","FF4500")
$WordsStyleAttribute7 = $WordsStyleNode7.SetAttribute("styleID","8")

$WordsStyleNode8 = $XmlDocument.CreateElement("WordsStyle")
$WordsStyleAttribute8 = $WordsStyleNode8.SetAttribute("name","COMMENT")
$WordsStyleAttribute8 = $WordsStyleNode8.SetAttribute("fontStyle","0")
$WordsStyleAttribute8 = $WordsStyleNode8.SetAttribute("fontName","")
$WordsStyleAttribute8 = $WordsStyleNode8.SetAttribute("bgColor","FFFFFF")
$WordsStyleAttribute8 = $WordsStyleNode8.SetAttribute("fgColor","006400")
$WordsStyleAttribute8 = $WordsStyleNode8.SetAttribute("styleID","1")

$WordsStyleNode9 = $XmlDocument.CreateElement("WordsStyle")
$WordsStyleAttribute9 = $WordsStyleNode9.SetAttribute("name","COMMENT LINE")
$WordsStyleAttribute9 = $WordsStyleNode9.SetAttribute("fontStyle","0")
$WordsStyleAttribute9 = $WordsStyleNode9.SetAttribute("fontName","")
$WordsStyleAttribute9 = $WordsStyleNode9.SetAttribute("bgColor","FFFFFF")
$WordsStyleAttribute9 = $WordsStyleNode9.SetAttribute("fgColor","006400")
$WordsStyleAttribute9 = $WordsStyleNode9.SetAttribute("styleID","2")

$WordsStyleNode10 = $XmlDocument.CreateElement("WordsStyle")
$WordsStyleAttribute10 = $WordsStyleNode10.SetAttribute("name","NUMBER")
$WordsStyleAttribute10 = $WordsStyleNode10.SetAttribute("fontStyle","0")
$WordsStyleAttribute10 = $WordsStyleNode10.SetAttribute("fontName","")
$WordsStyleAttribute10 = $WordsStyleNode10.SetAttribute("bgColor","FFFFFF")
$WordsStyleAttribute10 = $WordsStyleNode10.SetAttribute("fgColor","800080")
$WordsStyleAttribute10 = $WordsStyleNode10.SetAttribute("styleID","4")

$WordsStyleNode11 = $XmlDocument.CreateElement("WordsStyle")
$WordsStyleAttribute11 = $WordsStyleNode11.SetAttribute("name","OPERATOR")
$WordsStyleAttribute11 = $WordsStyleNode11.SetAttribute("fontStyle","0")
$WordsStyleAttribute11 = $WordsStyleNode11.SetAttribute("fontName","")
$WordsStyleAttribute11 = $WordsStyleNode11.SetAttribute("bgColor","FFFFFF")
$WordsStyleAttribute11 = $WordsStyleNode11.SetAttribute("fgColor","a9a9a9")
$WordsStyleAttribute11 = $WordsStyleNode11.SetAttribute("styleID","10")

$WordsStyleNode12 = $XmlDocument.CreateElement("WordsStyle")
$WordsStyleAttribute12 = $WordsStyleNode12.SetAttribute("name","DELIMINER1")
$WordsStyleAttribute12 = $WordsStyleNode12.SetAttribute("fontStyle","0")
$WordsStyleAttribute12 = $WordsStyleNode12.SetAttribute("fontName","")
$WordsStyleAttribute12 = $WordsStyleNode12.SetAttribute("bgColor","FFFFFF")
$WordsStyleAttribute12 = $WordsStyleNode12.SetAttribute("fgColor","8b0000")
$WordsStyleAttribute12 = $WordsStyleNode12.SetAttribute("styleID","14")

$WordsStyleNode13 = $XmlDocument.CreateElement("WordsStyle")
$WordsStyleAttribute13 = $WordsStyleNode13.SetAttribute("name","DELIMINER2")
$WordsStyleAttribute13 = $WordsStyleNode13.SetAttribute("fontStyle","0")
$WordsStyleAttribute13 = $WordsStyleNode13.SetAttribute("fontName","")
$WordsStyleAttribute13 = $WordsStyleNode13.SetAttribute("bgColor","FFFFFF")
$WordsStyleAttribute13 = $WordsStyleNode13.SetAttribute("fgColor","000000")
$WordsStyleAttribute13 = $WordsStyleNode13.SetAttribute("styleID","15")

$WordsStyleNode14 = $XmlDocument.CreateElement("WordsStyle")
$WordsStyleAttribute14 = $WordsStyleNode14.SetAttribute("name","DELIMINER3")
$WordsStyleAttribute14 = $WordsStyleNode14.SetAttribute("fontStyle","0")
$WordsStyleAttribute14 = $WordsStyleNode14.SetAttribute("fontName","")
$WordsStyleAttribute14 = $WordsStyleNode14.SetAttribute("bgColor","FFFFFF")
$WordsStyleAttribute14 = $WordsStyleNode14.SetAttribute("fgColor","000000")
$WordsStyleAttribute14 = $WordsStyleNode14.SetAttribute("styleID","16")

$Null = $StylesNode.AppendChild($WordsStyleNode1)
$Null = $StylesNode.AppendChild($WordsStyleNode2)
$Null = $StylesNode.AppendChild($WordsStyleNode3)
$Null = $StylesNode.AppendChild($WordsStyleNode4)
$Null = $StylesNode.AppendChild($WordsStyleNode5)
$Null = $StylesNode.AppendChild($WordsStyleNode6)
$Null = $StylesNode.AppendChild($WordsStyleNode7)
$Null = $StylesNode.AppendChild($WordsStyleNode8)
$Null = $StylesNode.AppendChild($WordsStyleNode9)
$Null = $StylesNode.AppendChild($WordsStyleNode10)
$Null = $StylesNode.AppendChild($WordsStyleNode11)
$Null = $StylesNode.AppendChild($WordsStyleNode12)
$Null = $StylesNode.AppendChild($WordsStyleNode13)
$Null = $StylesNode.AppendChild($WordsStyleNode14)

$Null = $UserLangNode.AppendChild($StylesNode)

$Null = $RootNode.AppendChild($UserLangNode)
$Null = $XmlDocument.AppendChild($RootNode)
$XmlDocument.Save($ExportFilePath)

$NewContent = Get-Content $ExportFilePath | Foreach-Object {$_ -replace "&", "&" -replace "<", "<" -replace ">", ">"} 
$NewContent | Set-Content $ExportFilePath
