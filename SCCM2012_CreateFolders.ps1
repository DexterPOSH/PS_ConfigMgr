$PSDefaultParameterValues =@{"get-cimclass:namespace"="Root\SMS\site_DEX";"get-cimclass:computername"="DexSCCM";"get-cimInstance:computername"="DexSCCM";"get-ciminstance:namespace"="Root\SMS\site_DEX";"get-wmiobject:namespace"="Root\SMS\site_DEX";"get-WMIObject:computername"="DexSCCM"}

#import the ConfigMgr Module
Import-Module -Name "$(split-path $Env:SMS_ADMIN_UI_PATH)\ConfigurationManager.psd1"

#set the location to the CMSite
Set-Location -Path DEX:


Get-CimClass -ClassName *container*

Get-CimInstance -ClassName SMS_ObjectContainerNode


#Create a new Folder using WMI
 $POSHFolder = New-CimInstance -ClassName SMS_ObjectContainerNode -Property @{Name="PowerShell";ObjectType=6000;ParentContainerNodeid=0;SourceSite="DEX"} -Namespace root/sms/site_DEX -ComputerName DexSCCM -Verbose

#Move the Application to the Folder

$Application  = Get-CMApplication -Name "PowerShell Community Extensions"

Invoke-CimMethod -ClassName SMS_ObjectContainerItem -MethodName MoveMembersEx -Arguments @{InstanceKeys=[string[]]$Application.ModelName;ContainerNodeID=[System.UInt32]0;TargetContainerNodeID=[System.UInt32]($POSHFolder.ContainerNodeID);ObjectTypeName="SMS_ApplicationLatest"} -Namespace root/sms/site_DEX -ComputerName DexSCCM -Verbose
#How to know what goes in the InstanceKeys ??


