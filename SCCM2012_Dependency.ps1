#Load the ConfigurationManager Module
Import-Module -Name "$(split-path $Env:SMS_ADMIN_UI_PATH)\ConfigurationManager.psd1"


#Load the Default Parameter Values for Get-WMIObject cmdlet
$PSDefaultParameterValues =@{"Get-wmiobject:namespace"="Root\SMS\site_DEX";"Get-WMIObject:computername"="DexSCCM"}

#load the Application Management DLL
Add-Type -Path "$(Split-Path  $Env:SMS_ADMIN_UI_PATH)\Microsoft.ConfigurationManagement.ApplicationManagement.dll"
Add-Type -Path "$(Split-Path  $Env:SMS_ADMIN_UI_PATH)\Microsoft.ConfigurationManagement.ApplicationManagement.MsiInstaller.dll"


#Creating Type Accelerators - for making assemlby refrences easier later
$accelerators = [PSObject].Assembly.GetType('System.Management.Automation.TypeAccelerators')
$accelerators::Add('SccmSerializer',[type]'Microsoft.ConfigurationManagement.ApplicationManagement.Serialization.SccmSerializer')


#region Application 1
$ApplicationName1 = "Notepad++ 6.2.3"

#get direct reference to the Application's WMI Instance
$application1 = [wmi](Get-WmiObject -Query "select * from sms_application where LocalizedDisplayName='$ApplicationName1' AND ISLatest='true'").__PATH

#Deserialize the SDMPackageXML
$App1Deserializedstuff = [SccmSerializer]::DeserializeFromString($application1.SDMPackageXML)

#endregion Application 1


#region Application 2 

#Name of the Application which will be added as a dependency
$ApplicationName2 = ".NET4"

#Reference to the above application
$application2 = [wmi](Get-WmiObject -Query "select * from sms_application where LocalizedDisplayName='$ApplicationName2' AND ISLatest='true'").__PATH

#deserialize the XML
$App2Deserializedstuff = [SccmSerializer]::DeserializeFromString($application2.SDMPackageXML)

#endregion



# Create the Severity None
$severity = [Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Rules.NoncomplianceSeverity]::Critical

#Create the Annotation - Name & description of the Dependency 
$annotation = New-Object -TypeName Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Rules.Annotation
$annotation.DisplayName.Text = "DependencyName"

# Create the Empty Rule Context
$RuleContext = New-Object -TypeName Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Rules.RuleScope




# set the Desired State as "Required" or Mandatory
$DTDesiredState = [Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DeploymentTypeDesiredState]::Required  


#Store the arguments before hand
$ApplicationAuthoringScopeId = ($application2.CI_UniqueID -split "/")[0]
$ApplicationLogicalName = ($application2.CI_UniqueID -split "/")[1]
$ApplicationVersion =  $application2.SourceCIVersion
$DeploymentTypeAuthoringScopeId = $App2Deserializedstuff.DeploymentTypes.scope
$DeploymentTypeLogicalName = $App2Deserializedstuff.DeploymentTypes.name
$DeploymentTypeVersion = $App2Deserializedstuff.DeploymentTypes.Version
$AutoInstall = $True  #this determines if the dependency Application needs to be uninstalled when the new Application is pushed

#create the intent expression which will be addded to the Operand
$intentExpression = new-object Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DeploymentTypeIntentExpression -ArgumentList  $ApplicationAuthoringScopeId, $ApplicationLogicalName, $ApplicationVersion, $DeploymentTypeAuthoringScopeId, $DeploymentTypeLogicalName, $DeploymentTypeVersion, $DTDesiredState, $AutoInstall

#create the new OR operator 
$OrOperator = [Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator]::Or 

#Create the Operand - Note the typename of this one
$operand = New-Object  Microsoft.ConfigurationManagement.DesiredConfigurationManagement.CustomCollection[Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DeploymentTypeIntentExpression]

#add the Intent Expression to the Operand
$operand.Add($intentExpression)

#Now the Operator and Operand are added by the Expression
$BaseExpression = New-Object -TypeName Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DeploymentTypeExpression -ArgumentList $OrOperator,$operand



#Create the new DeploymentType Rule
$DTRUle = New-Object -TypeName Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Rules.DeploymentTypeRule -ArgumentList $("DTRule_"+[guid]::NewGuid().Guid),$severity, $annotation, $BaseExpression


#add the supersedence to the deployment type
$App1Deserializedstuff.DeploymentTypes[0].Dependencies.Add($DTRUle)


# Serialize the XML 
$newappxml = [SccmSerializer]::Serialize($App1Deserializedstuff, $false)

#Set-CimInstance -InputObject $application -Property @{SDMPackageXML=$newappxml} -Verbose

#set the property back on the local copy of the Object
$application1.SDMPackageXML = $newappxml

#Now time to set the changes back to the ConfigMgr
$application1.Put()

