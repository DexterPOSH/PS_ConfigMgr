$PSDefaultParameterValues =@{"get-cimclass:namespace"="Root\SMS\site_DEX";"get-cimclass:computername"="DexSCCM";"get-cimInstance:computername"="DexSCCM";"get-ciminstance:namespace"="Root\SMS\site_DEX";"get-wmiobject:namespace"="Root\SMS\site_DEX";"get-WMIObject:computername"="DexSCCM"}
<#
Const disc_user    = "SMS_AD_USER_DISCOVERY_AGENT" 
Const disc_group   = "SMS_AD_SECURITY_GROUP_DISCOVERY_AGENT" 
Const disc_system  = "SMS_AD_SYSTEM_DISCOVERY_AGENT" 
Const disc_sysgrp  = "SMS_AD_SYSTEM_GROUP_DISCOVERY_AGENT" 
Const disc_network = "SMS_NETWORK_DISCOVERY" 
Const disc_heart   = "SMS_SITE_CONTROL_MANAGER" 


#>


Import-Module -Name "$(split-path $Env:SMS_ADMIN_UI_PATH)\ConfigurationManager.psd1"


#region Active Directory Forest Discovery
 
#create a Schedule Token 
$Schedule = New-CMSchedule -RecurInterval Days -RecurCount 7

#Enable the Active Directory Forest Discovery 
Set-CMDiscoveryMethod -ActiveDirectoryForestDiscovery -SiteCode DEX -Enabled:$true -PollingSchedule $Schedule -EnableActiveDirectorySiteBoundaryCreation:$true -EnableSubnetBoundaryCreation:$true

#To run AD Forest Disovery now
Invoke-CMForestDiscovery -SiteCode DEX -Verbose


#endregion Active Directory Forest Discovery





#region AD Group Discovery

#set the default parameter values for CIM cmdlet..don't want to write them again and again
$PSDefaultParameterValues = @{"get-cimInstance:computername"="DexSCCM";"get-ciminstance:namespace"="Root\SMS\site_DEX"}

#set the Discovery Scopes

$GroupDiscovery = get-ciminstance -classname SMS_SCI_Component -filter 'componentname ="SMS_AD_SECURITY_GROUP_DISCOVERY_AGENT"'
$ADContainerProp = $GroupDiscovery.PropLists | where {$_.PropertyListName -eq "AD Containers" }
#$ADContainerProp.Values = "All My AD groups",0,0,1
$ADContainerProp.Values = "dex test",0,0,1  #Name, Type Setting (Location [0] or Group [1]),Recursive,don't know what this does

#need to add new Embedded Property to the Props specifying the Search Base...we can overwrite the already existing one too.
$NewProp = New-CimInstance -ClientOnly -Namespace "root/sms/site_dex" -ClassName SMS_EmbeddedPropertyList -Property @{PropertyListName="Search Bases:dex test";Values=[string[]]"LDAP://DC=dexter,DC=com"}
$GroupDiscovery.PropLists += $NewProp

#set the Changes back to the CIM Instance
get-ciminstance -classname SMS_SCI_Component -filter 'componentname ="SMS_AD_SECURITY_GROUP_DISCOVERY_AGENT"' | Set-CimInstance -Property @{PropLists=$GroupDiscovery.PropLists}


#Use the Cmdlet to configure rest of the options
Set-CMDiscoveryMethod -SiteCode DEX -ActiveDirectoryGroupDiscovery -Enabled $true -EnableDeltaDiscovery $true -DeltaDiscoveryIntervalMinutes 5  -EnableFilteringExpiredLogon $true -TimeSinceLastLogonDays 90 -EnableFilteringExpiredPassword $true -TimeSinceLastPasswordUpdateDays 90 -DiscoverDistributionGroupsMembership $true



#need to restart the Service
(Get-Service SMS_SITE_COMPONENT_MANAGER -ComputerName dexsccm).stop()

Start-Sleep -Seconds 10

(Get-Service SMS_SITE_COMPONENT_MANAGER -ComputerName dexsccm).start()

#endregion AD Group Discovery





#region AD System Discovery

#set the default parameter values for CIM cmdlet
$PSDefaultParameterValues = @{"get-cimInstance:computername"="DexSCCM";"get-ciminstance:namespace"="Root\SMS\site_DEX"}

#Set the System Discovery
$Schedule = New-CMSchedule -Start "2014/02/13 20:20:00" -RecurInterval minutes -RecurCount 10
Set-CMDiscoveryMethod -SiteCode DEX -ActiveDirectorySystemDiscovery -Enabled $true  -EnableFilteringExpiredLogon $true -TimeSinceLastLogonDays 90 -EnableFilteringExpiredPassword $true -TimeSinceLastPasswordUpdateDays 90 -PollingSchedule $Schedule

#To set the AD Containers 
$Sysdiscovery = get-ciminstance -classname SMS_SCI_Component -filter 'componentname ="sms_ad_system_discovery_agent"'

$ADContainerProp =$Sysdiscovery.PropLists | where {$_.PropertyListName -eq "AD Containers" }

$ADContainerProp.Values = "LDAP://CN=System,DC=Dexter,DC=Com",1,1 # Values ---- Ldap path of the Container, Recursive search, Discover objects within groups

#set the changes back to the CIM Instance
Get-CimInstance -classname SMS_SCI_Component -filter 'componentname ="sms_ad_system_discovery_agent"' | Set-CimInstance -Property @{PropLists=$Sysdiscovery.PropLists}

#need to restart the Service
(Get-Service SMS_SITE_COMPONENT_MANAGER -ComputerName dexsccm).stop()

Start-Sleep -Seconds 10

(Get-Service SMS_SITE_COMPONENT_MANAGER -ComputerName dexsccm).start()

#endregion AD System Discovery





#region AD User Discovery

$UserDiscovery = get-ciminstance -classname SMS_SCI_Component -filter 'componentname ="SMS_AD_USER_DISCOVERY_AGENT"'
$ADContainerProp =$UserDiscovery.PropLists | where {$_.PropertyListName -eq "AD Containers" }
$ADContainerProp.Values = "LDAP://CN=Users,DC=Dexter,DC=Com",0,0

Get-CimInstance -classname SMS_SCI_Component -filter 'componentname ="SMS_AD_USER_DISCOVERY_AGENT"' | Set-CimInstance -Property @{PropLists=$UserDiscovery.PropLists}

$Schedule = New-CMSchedule -Start "2014/02/15 12:00:10" -RecurInterval Minutes -RecurCount 10 
Set-CMDiscoveryMethod -ActiveDirectoryUserDiscovery -SiteCode DEX -Enabled $true -PollingSchedule $Schedule -EnableDeltaDiscovery $true -DeltaDiscoveryIntervalMinutes 10 


#need to restart the Service
(Get-Service SMS_SITE_COMPONENT_MANAGER -ComputerName dexsccm).stop()

Start-Sleep -Seconds 10

(Get-Service SMS_SITE_COMPONENT_MANAGER -ComputerName dexsccm).start()

#endregion AD User Discovery



#region HeartBeat Discovery

$Schedule = New-CMSchedule -Start "2014/02/16 10:30:00" -DurationInterval Minutes -DurationCount 10
Set-CMDiscoveryMethod -Heartbeat -SiteCode DEX -Enabled $True -PollingSchedule $Schedule 

#endregion HeartBeat Discovery



#region Network Discovery

Set-CMDiscoveryMethod -NetworkDiscovery -SiteCode DEX -NetworkDiscoveryType Topology -Enabled $true 

#endregion Network Discovery