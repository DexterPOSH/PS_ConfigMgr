#region Add Catalog Website Service Point and Wesbite Point
$PSDefaultParameterValues =@{"get-cimclass:namespace"="Root\SMS\site_DEX";"get-cimclass:computername"="DexSCCM";"get-cimInstance:computername"="DexSCCM";"get-ciminstance:namespace"="Root\SMS\site_DEX";"get-wmiobject:namespace"="Root\SMS\site_DEX";"get-WMIObject:computername"="DexSCCM"}
Import-Module -Name "$(split-path $Env:SMS_ADMIN_UI_PATH)\ConfigurationManager.psd1"

Add-CMApplicationCatalogWebServicePoint -PortNumber 80 -SiteCode "CM1" -SiteSystemServerName "DexSCCM.dexter.com" -CommunicationType HTTP #-IISWebsite "Default Web Site" -WebApplicationName "CMApplicationCatalogSvc" -Verbose

Add-CMApplicationCatalogWebsitePoint -SiteSystemServerName 'DexSCCM' -SiteCode DEX -SiteSystemServerNameConfiguredForApplicationCatalogWebServicePoint 'dexsccm' -ConfiguredAsHttpConnection -PortForHttpConnection 80 -IISWebsite "Default Web Site" -WebApplicationName "CMApplicationCatalogSvc" -OrganizationName "Dexter's LAB" -NetbiosName "DexSCCM" -ColorBlue 52 -ColorRed 168 -Verbose


#endregion



#region Create Custom Client Device Settings

#First create the New Client settings 
New-CMClientSetting -Name "Custom Client Device Settings" -Type Device -Description "Custom Client settings: following Windows-Noob --> DexterPOSH" 

# Decrease the Priority
Set-CMClientSetting -Priority Decrease -Name "Custom Client Device Settings"

#Now start configuring the Client Policy, Computer Agent and Software Updates settings
# 1. Configure Client Policy Settings
Set-CMClientSetting -Name "Custom Client Device Settings" -PolicyPollingInterval 5  -EnableUserPolicyPolling $true -EnableUserPolicyOnInternet $false

# 2. Configure the Computer Agent
Set-CMClientSetting -Name "Custom Client Device Settings" -PowerShellExecutionPolicy Bypass -InitialReminderHoursInterval 48 -InterimReminderHoursInterval 4 -FinalReminderMinutesInterval 15  -PortalUrl "http://dexsccm.dexter.com/CMApplicationCatalog/" -AddPortalToTrustedSiteList $true -AllowPortalToHaveElevatedTrust $true -BrandingTitle "Dexter's LAB" -InstallRestriction AllUsers -DisplayNewProgramNotification $true 

#deploy the above settings to the Collection
Start-CMClientSettingDeployment -ClientSettingName "Custom Client Device Settings" -CollectionName "All Systems"


#endregion Create Custom Device Client Settings 



#region Create Custom User Device Settings --WMI Way

#create the new instance of the Class pass to it relevant properties like Name, Type , Priority, Description
$ClientUserSetting = New-CimInstance -ClassName SMS_ClientSettings -Property @{Name="Custom Client User Settings";Type=2;Priority=2;Description="Custom Client User Settings -- WMI Way"} -Namespace root/sms/site_DEX -ComputerName DexSCCM -verbose

#Now we need to add the User Device Affinity to the class we created as one of the AgentConfgurations
$ClientUserSetting.AgentConfigurations += New-CimInstance -ClassName SMS_TargetingAgentConfig -Property @{AgentID=10;AllowUserAffinity=1} -Namespace root/sms/Site_DEX -ComputerName DexSCCM -Verbose

#Now once done set back the property to the ConfigMgr Server
Set-CimInstance -InputObject $ClientUserSetting 

#deploy the above settings to the Collection
Start-CMClientSettingDeployment -ClientSettingName "Custom Client Device Settings" -CollectionName "All Systems"

#endregion Create Custom User Device Settings















#region BITS Settings

#WMI Class SMS_BITS2Config 


#endregion BITS Settings


#region Computer Restart Settings

#WMI Class SMS_ClientRestartAgentConfig

#endregion


#region Computer Settings

#WMI SMS_ClientSettings 
# WMI Class -- SMS_ConfigMgrClientAgentConfig

#endregion

<#
#first create SMS_ClientSetting
Execute SQL =select  all SMS_ClientSettingsDefault.AssignmentCount,SMS_ClientSettingsDefault.CreatedBy,SMS_ClientSettingsDefault.DateCreated,SMS_ClientSettingsDefault.DateModified,SMS_ClientSettingsDefault.Description,SMS_ClientSettingsDefault.Enabled,SMS_ClientSettingsDefault.Flags,SMS_ClientSettingsDefault.LastModifiedBy,SMS_ClientSettingsDefault.Name,SMS_ClientSettingsDefault.Priority,SMS_ClientSettingsDefault.ID,SMS_ClientSettingsDefault.SourceSite,SMS_ClientSettingsDefault.Type,SMS_ClientSettingsDefault.UniqueID from vSMS_ClientSettingsDefault AS SMS_ClientSettingsDefault

Execute WQL  =select Priority from SMS_ClientSettings where Priority > 1
Execute SQL =select  all SMS_ClientSettings.Priority from vSMS_ClientSettings AS SMS_ClientSettings  where SMS_ClientSettings.Priority > 1

#create a new CLient settings
New-CimInstance -ClassName SMS_ClientSettings -Namespace root/sms/site_DEX -Property @{Name="Dex Settings";Priority=12} -ComputerName dexsccm 


#>