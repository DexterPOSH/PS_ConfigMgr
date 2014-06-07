$PSDefaultParameterValues =@{"get-cimclass:namespace"="Root\SMS\site_DEX";"get-cimclass:computername"="DexSCCM";"get-cimInstance:computername"="DexSCCM";"get-ciminstance:namespace"="Root\SMS\site_DEX";"get-wmiobject:namespace"="Root\SMS\site_DEX";"get-WMIObject:computername"="DexSCCM"}
#set the Supersedence
<# No Need to use the WMI Class
#WMI Classes found
# SMS_Apprelation_flat  --> represents the flattened application relation. This includes direct and indirect relations. 
# SMS_DeploymentInfo  --> represents information for all types of deployment

$SuperSed = New-CimInstance -ClassName "SMS_AppRelation_Flat" -Namespace "Root/SMS/Site_DEX" -ComputerName DexSCCM -Property @{
                                                                                                                        FromApplicationCIID=[uint32]$ApplicationName.CI_ID;
                                                                                                                        FromDeploymentTypeCIID=[uint32]16778516;
                                                                                                                        Level=[uint32] 1;
       #857103                                                                                                                 RelationType=[uint32]15;
                                                                                                                        ToApplicationCIID=[uint32]16778416;
                                                                                                                        ToDeploymentTypeCIID=[uint32]16778417} 

#>

#Load the ConfigurationManager Module
Import-Module -Name "$(split-path $Env:SMS_ADMIN_UI_PATH)\ConfigurationManager.psd1"

#Load the Default Parameter Values for Get-WMIObject cmdlet
$PSDefaultParameterValues =@{"Get-wmiobject:namespace"="Root\SMS\site_DEX";"Get-WMIObject:computername"="DexSCCM"}

#load the Application Management DLL
Add-Type -Path "$(Split-Path  $Env:SMS_ADMIN_UI_PATH)\Microsoft.ConfigurationManagement.ApplicationManagement.dll"


#region Configure SuperSedence

$ApplicationName = "Notepad++ 6.5.1"
$application = [wmi](Get-WmiObject -Query "select * from sms_application where LocalizedDisplayName='$ApplicationName' AND ISLatest='true'").__PATH


#Deserialize the SDMPackageXML
$Deserializedstuff = [Microsoft.ConfigurationManagement.ApplicationManagement.Serialization.SccmSerializer]::DeserializeFromString($application.SDMPackageXML)


#try creating the Supersedence

#Name of the Application which will be superseded
$Name = "Notepad++ 6.2.3"

#Reference to the above application
$supersededapplication = [wmi](Get-WmiObject -Query "select * from sms_application where LocalizedDisplayName='$Name' AND ISLatest='true'").__PATH

#deserialize the XML
$supersededDeserializedstuff = [Microsoft.ConfigurationManagement.ApplicationManagement.Serialization.SccmSerializer]::DeserializeFromString($supersededapplication.SDMPackageXML)

# set the Desired State for the Superseded Application's Deployment type to "prohibit" from running
$DTDesiredState = [Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DeploymentTypeDesiredState]::Prohibited  


#Store the arguments before hand
$ApplicationAuthoringScopeId = ($supersededapplication.CI_UniqueID -split "/")[0]
$ApplicationLogicalName = ($supersededapplication.CI_UniqueID -split "/")[1]
$ApplicationVersion =  $supersededapplication.SourceCIVersion
$DeploymentTypeAuthoringScopeId = $supersededDeserializedstuff.DeploymentTypes.scope
$DeploymentTypeLogicalName = $supersededDeserializedstuff.DeploymentTypes.name
$DeploymentTypeVersion = $supersededDeserializedstuff.DeploymentTypes.Version
$uninstall = $false  #this determines if the superseded Application needs to be uninstalled when the new Application is pushed

#create the intent expression
$intentExpression = new-object Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DeploymentTypeIntentExpression -ArgumentList  $ApplicationAuthoringScopeId, $ApplicationLogicalName, $ApplicationVersion, $DeploymentTypeAuthoringScopeId, $DeploymentTypeLogicalName, $DeploymentTypeVersion, $DTDesiredState, $uninstall

# Create the Severity None
$severity = [Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Rules.NoncomplianceSeverity]::None

# Create the Empty Rule Context
$RuleContext = New-Object -TypeName Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Rules.RuleScope
$RuleContext.ConfigurationItemLogicalName = "Dexter's Supersedence Rule"
$RuleContext.ConfigurationItemVersion = "1.0"

#Create the new DeploymentType Rule
$DTRUle = New-Object -TypeName Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Rules.DeploymentTypeRule -ArgumentList $severity, $null, $intentExpression

#add the supersedence to the deployment type
$Deserializedstuff.DeploymentTypes[0].Supersedes.Add($DTRUle) 

# Serialize the XML 
$newappxml = [Microsoft.ConfigurationManagement.ApplicationManagement.Serialization.SccmSerializer]::Serialize($Deserializedstuff, $false)

#Set-CimInstance -InputObject $application -Property @{SDMPackageXML=$newappxml} -Verbose

#set the property back on the local copy of the Object
$application.SDMPackageXML = $newappxml

#Now time to set the changes back to the ConfigMgr
$application.Put()




















#set the Error Codes 
$test = New-Object -TypeName Microsoft.ConfigurationManagement.ApplicationManagement.ExitCode
$test.Code = 20
#$test.Class = [Microsoft.ConfigurationManagement.ApplicationManagement.ExitCodeClass]"Failure" #Default is failure
$test.Name = "Setting VM Size failed"
$test.Code = 21

$test1 = New-Object -TypeName Microsoft.ConfigurationManagement.ApplicationManagement.ExitCode
$test1.Code = 22
#$test2.Class = [Microsoft.ConfigurationManagement.ApplicationManagement.ExitCodeClass]"Failure"
$test1.Name = "Resetting the VM Size to System managed Failed"

$Deserializedstuff.DeploymentTypes.Installer.ExitCodes.add($test)

$Deserializedstuff.DeploymentTypes.Installer.ExitCodes.add($test1)

$newappxml = [Microsoft.ConfigurationManagement.ApplicationManagement.Serialization.SccmSerializer]::Serialize($Deserializedstuff, $false)

#Set-CimInstance -InputObject $application -Property @{SDMPackageXML=$newappxml} -Verbose

$application.SDMPackageXML = $newappxml
$application.Put()

#endregion add the extra exit codes 21