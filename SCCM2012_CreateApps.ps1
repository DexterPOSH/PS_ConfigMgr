$PSDefaultParameterValues =@{"get-cimclass:namespace"="Root\SMS\site_DEX";"get-cimclass:computername"="DexSCCM";"get-cimInstance:computername"="DexSCCM";"get-ciminstance:namespace"="Root\SMS\site_DEX";"get-wmiobject:namespace"="Root\SMS\site_DEX";"get-WMIObject:computername"="DexSCCM"}
Import-Module -Name "$(split-path $Env:SMS_ADMIN_UI_PATH)\ConfigurationManager.psd1"


function Import-SCCMAssemblies($SCCMAdminConsolePath){
$path = "$SCCMAdminConsolePath\Microsoft.ConfigurationManagement.ApplicationManagement.dll"
if (Test-Path $path) { $t = [System.Reflection.Assembly]::LoadFrom($path)}
 
$path = "$SCCMAdminConsolePath\Microsoft.ConfigurationManagement.ApplicationManagement.Extender.dll"
if (Test-Path $path) { $t = [System.Reflection.Assembly]::LoadFrom($path)}
 
$path = "$SCCMAdminConsolePath\Microsoft.ConfigurationManagement.ApplicationManagement.MsiInstallerdll"
if (Test-Path $path) { $t = [System.Reflection.Assembly]::LoadFrom($path)}
}

Import-SCCMAssemblies -SCCMAdminConsolePath $(Split-Path -Path $Env:SMS_ADMIN_UI_PATH)

$SiteID = (Invoke-CimMethod -Name GetSiteId -ClassName SMS_Identification -ComputerName dexsccm -Namespace root/sms/site_DEX).Siteid

$Scopeid = "ScopedID_" + $SiteID



#Create an unique id for the application and the deployment type
$newApplicationID = "Application_" + [guid]::NewGuid().ToString() 
$newDeploymentTypeID = "DeploymentType_" + [guid]::NewGuid().ToString()     
  
#Create SCCM 2012 object id for application and deploymenttyo 
$newApplicationID = New-Object Microsoft.ConfigurationManagement.ApplicationManagement.ObjectID($scopeid,$newApplicationID)   
$newDeploymentTypeID = New-Object Microsoft.ConfigurationManagement.ApplicationManagement.ObjectID($scopeid , $newDeploymentTypeID) 
     
#Create all the objects neccessary for the creation of the application
$newApplication =  New-Object Microsoft.ConfigurationManagement.ApplicationManagement.Application($newApplicationID)    
$newDeploymentType = New-Object  Microsoft.ConfigurationManagement.ApplicationManagement.DeploymentType($newDeploymentTypeID,"MSI")
$newDisplayInfo  =  New-Object Microsoft.ConfigurationManagement.ApplicationManagement.AppDisplayInfo 
$newApplicationContent = New-Object  Microsoft.ConfigurationManagement.ApplicationManagement.Content
$newContentFile = New-Object  Microsoft.ConfigurationManagement.ApplicationManagement.ContentFile
  
#Setting Display Info
$newDisplayInfo.Title = $Application.DisplayName
$newDisplayInfo.Language = $newApplication.DisplayInfo.DefaultLanguage
$newDisplayInfo.Description = $Application.DisplayName
$newDisplayInfo.Version = $Application.PRVersion 
$newApplication.DisplayInfo.Add($newDisplayInfo)
 
#Setting default Language must be set and displayinfo must exists
$newApplication.DisplayInfo.DefaultLanguage = "en-US"
 
 
$newApplication.Title = $APP.Title
$newApplication.Version = 1
 
#Deployment Type msi installer will be used
$newDeploymentType.Title = "Deploy $($APP.DisplayName)"
$newDeploymentType.Version = 1 
$newDeploymentType.Installer.ProductCode = $APP.ProductCode 
$newDeploymentType.Installer.InstallCommandLine = "Msiexec /i $($APP.MSIName) /qb-! Reboot=ReallySuppress"
$newDeploymentType.Installer.InstallFolder  ="\"  
$newDeploymentType.Installer.UninstallCommandLine = "Msiexec /x $($APP.ProductCode) /qb-! Reboot=ReallySuppress"
 
#Add the msi as content to the application
 
#UPDATE: Add all content to the application
$newApplicationContent = [Microsoft.ConfigurationManagement.ApplicationManagement.ContentImporter]::CreateContentFromFolder($APPUNCPath)
 
    #DELETE: $newContentFile.Name = $APP.MSIName
    #DELETE: $newApplicationContent.Files.Add($newContentFile)
    #DELETE: $newApplicationContent.Location = $APP.UNCPath
    $newApplicationContent.OnSlowNetwork = "Download"
 
 
 
$newDeploymentType.Installer.Contents.Add($newApplicationContent )
$newApplication.DeploymentTypes.Add($newDeploymentType)
 
#Serialize the object to an xml file
$newApplicationXML = [Microsoft.ConfigurationManagement.ApplicationManagement.Serialization.SccmSerializer]::SerializeToSTring($newApplication,$true)

$applicationClass = [WMICLASS]"\\$($SccmServer.Machine)\$($SccmServer.Namespace):SMS_Application"
$newApplication = $applicationClass.createInstance()
 
$newApplication.SDMPackageXML = $newApplicationXML
$tmp = $newApplication.Put()
 
#Reload the application to get the daat
$newApplication.Get()









#my tries

#region Create Basic App & Deploy it

#create a new Application
New-CMApplication -Name "Quest Active Roles Managment Shell for AD" -Description "Quest Snapin for AD Admins (64-bit only)" -SoftwareVersion "1.51" -AutoInstall $true 

#create a New Application Category
New-CMCategory -CategoryType AppCategories -Name "ADAdministration" -Verbose

#Set new properties on the Application Object Created.
Set-CMApplication -Name "Quest Active Roles Managment Shell for AD"  -LocalizedApplicationName "Quest AD Snapin"  -LocalizedApplicationDescription "PowerShell Snapin to be used by AD Admins" -AppCategories "ADAdministration" -SendToProtectedDistributionPoint $true 

#Add the Deployment type automatically from the MSI 
Add-CMDeploymentType -ApplicationName "Quest Active Roles Managment Shell for AD" -InstallationFileLocation "\\dexsccm\Packages\QuestADSnapin\Quest_ActiveRolesManagementShellforActiveDirectoryx64_151.msi" -MsiInstaller -AutoIdentifyFromInstallationFile -ForceForUnknownPublisher $true -InstallationBehaviorType InstallForSystem

#Distribute the Content to the DP Group
Start-CMContentDistribution -ApplicationName "Quest Active Roles Managment Shell for AD" -DistributionPointGroupName "Dex LAB DP group" -Verbose

#create the Device Collection
New-CMDeviceCollection -Name "Quest Active Roles Managment Shell for AD" -Comment "All the Machines where Quest AD Snapin is sent to" -LimitingCollectionName "All Systems"  -RefreshType Periodic -RefreshSchedule (New-CMSchedule -Start (get-date) -RecurInterval Days -RecurCount 7) 

#Add the Direct Membership Rule to add a Resource as a member to the Collection
Add-CMDeviceCollectionDirectMembershipRule -CollectionName "Quest Active Roles Managment Shell for AD"  -Resource (Get-CMDevice -Name "DexterDC") -Verbose

#start the Deployment
Start-CMApplicationDeployment -CollectionName "Quest Active Roles Managment Shell for AD" -Name "Quest Active Roles Managment Shell for AD" -DeployAction Install -DeployPurpose Available -UserNotification DisplayAll -AvaliableDate (get-date) -AvaliableTime (get-date) -TimeBaseOn LocalTime  -Verbose

#Run the Deployment Summarization
Invoke-CMDeploymentSummarization -CollectionName "Quest Active Roles Managment Shell for AD" -Verbose

Invoke-CMClientNotification -DeviceCollectionName "Quest Active Roles Managment Shell for AD" -NotificationType RequestMachinePolicyNow -Verbose

#endregion Create Basic App & Deploy it