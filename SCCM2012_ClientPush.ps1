$PSDefaultParameterValues =@{"get-cimclass:namespace"="Root\SMS\site_DEX";"get-cimclass:computername"="DexSCCM";"get-cimInstance:computername"="DexSCCM";"get-ciminstance:namespace"="Root\SMS\site_DEX";"get-wmiobject:namespace"="Root\SMS\site_DEX";"get-WMIObject:computername"="DexSCCM"}
#$PSDefaultParameterValues =@{"get-cimclass:namespace"="Root\SMS\site_DEX";"get-cimclass:computername"="DexSCCM";"get-cimInstance:computername"="DexSCCM";"get-ciminstance:namespace"="Root\SMS\site_DEX"}
#select * From SMS_SCI_SysResUse where FileType=2 and SiteCode='DEX' and RoleName = 'SMS Management Point'
# SMS_SCI_Component.FileType=2,ItemName="SMS_CLIENT_CONFIG_MANAGER|DexSCCM.dexter.com",ItemType="Component",SiteCode="DEX"


#import the ConfigMgr Module
Import-Module -Name "$(split-path $Env:SMS_ADMIN_UI_PATH)\ConfigurationManager.psd1"

#set the location to the CMSite
Set-Location -Path DEX:

#loading up the default param values for Get-CIMInstance and Get-CIMClass
$PSDefaultParameterValues =@{"get-cimclass:namespace"="Root\SMS\site_DEX";"get-cimclass:computername"="DexSCCM";"get-cimInstance:computername"="DexSCCM";"get-ciminstance:namespace"="Root\SMS\site_DEX"}

#region Confgure the Client Installation Settings

Set-CMClientPushInstallation -SiteCode DEX -EnableAutomaticClientPushInstallation $true -EnableSystemTypeServer $true -EnableSystemTypeWorkstation $true -EnableSystemTypeConfigurationManager $true -InstallClientToDomainController $true -ChosenAccount "Dexter\Administrator" -InstallationProperty "SMSSITECODE=DEX"

#endregion Confgure the Client Installation Settings

#region Configure Boundary Groups and under references tab specify the Server for Content Location - Otherwise client won't know where to get the Content from

#create the new Boundary group
New-CMBoundaryGroup -Name "DexLabBG" -Description "Boundary Group for my LAB" -DefaultSiteCode DEX -Verbose

#add the boundary to the Boundary group
Add-CMBoundaryToGroup -Boundary (Get-CMBoundary ) -BoundaryGroup (Get-CMBoundaryGroup ) -Verbose

#Add the Distribution Point as the Content Location for the Boundary Group
get-cimclass -ClassName SMS_BoundaryGroup | select -ExpandProperty CimclassMethods

$BoundaryGroup = Get-CimInstance -ClassName SMS_BoundaryGroup 

#Add the Site System for the Content Location -- WMI Way
Invoke-CimMethod -InputObject $BoundaryGroup -MethodName AddSiteSystem -Arguments @{ServerNALPath = [string[]]'["Display=\\DexSCCM.dexter.com\"]MSWNET:["SMS_SITE=DEX"]\\DexSCCM.dexter.com\'; Flags=([System.UInt32[]]0) } -Verbose

#endregion Configure Boundary Groups and under references tab specify the Server for Content Location - Otherwise client won't know where to get the Content from

#region Distribute the Config Mgr Client Package to ensure that the Package resides there
Start-CMContentDistribution -PackageName "Configuration Manager Client Package" -DistributionPointGroupName "Dex LAB DP group"

#To verify the Integrity of the Package -- WMI Way
#get the reference of the package stored in the DP
$package = Get-CimInstance -ClassName SMS_DistributionPoint -Filter 'PackageID="DEX00002"'

#use the VerifyMethod() on the class SMS_DistributionPoint
Invoke-CimMethod -ClassName SMS_DistributionPoint -Name VerifyPackage -Arguments @{PackageID=$package.PackageID;NALPath=$package.ServerNALPath} -ComputerName dexsccm -Namespace root/sms/site_DEX -Verbose 

#endregion Distribute the Config Mgr Client Package to ensure that the Package resides there

<#


PS DEX:\> Get-CimClass -ClassName SMS_ObjectContainerItem | select -ExpandProperty cimclassmethods | fl *


Name       : MoveMembers
ReturnType : SInt32
Parameters : {InstanceKeys, ContainerNodeID, TargetContainerNodeID, ObjectType}
Qualifiers : {deprecated, Description, implemented, static}

Name       : MoveMembersEx
ReturnType : SInt32
Parameters : {InstanceKeys, ContainerNodeID, TargetContainerNodeID, ObjectTypeName}
Qualifiers : {Description, implemented, static}



#>
