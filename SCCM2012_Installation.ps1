
<#
.LINKS
    http://www.powershellmagazine.com/2012/05/14/managing-group-policy-with-powershell/

    Add AD User/Group to the Local Admin Group
    http://gallery.technet.microsoft.com/scriptcenter/Add-AD-UserGroup-to-Local-fe5e9239

#>

#region Create AD Users

#Import the Module
Import-Module -Name ActiveDirectory

$newUsers = "SMSadmin","Testuser","Testuser2","Testuser3","DomJoin","ReportsUser","ClientInstall","SCCMNAA"

#Create a Common Password..this is a Demo Environment
$Password =  ConvertTo-SecureString -String "P@ssw0rd2" -AsPlainText  -Force
foreach ($newuser in $newUsers)
{
    New-ADUser -SamAccountName $newUser -Name $newuser -AccountPassword $Password -PassThru | Enable-ADAccount -Verbose
}

#endregion Create AD Users


#region Give AD Users Local Admin Access

#need to add the AD Users [ClientInstall,SMSadmin] to the Local Admin Group
  
  ([ADSI]"WinNT://DexSCCM/Administrators,group").add("WinNT://Dexter/ClientInstall")

  ([ADSI]"WinNT://DexSCCM/Administrators,group").add("WinNT://Dexter/SMSAdmin")

#enregion Give AD Users Local Admin Access


#region create Conatiner 'System Management'

#create the Container
New-ADObject -path 'CN=System,DC=dexter,DC=com' -Type container -name 'System Management' -PassThru

#get the Default Naming context
$root = (Get-ADRootDSE).defaultNamingContext

#store the ACL for the Container System Management
$acl = get-acl "AD:CN=System Management,CN=System,$root"

#get the Computer AD Object
$SCCMComputerAccount = Get-ADComputer -Identity DexSCCM

#Create an ACE to give the Computer Account Full access to the Container "System Management" and the child Objects
$All = [System.DirectoryServices.ActiveDirectorySecurityInheritance]::SelfAndChildren
$ace = new-object System.DirectoryServices.ActiveDirectoryAccessRule $SCCMComputerAccount.SID, "GenericAll", "Allow", $All

#add the ACE to the ACL
$acl.AddAccessRule($ace)

#Set the modified ACL back to the Container "System Management" 
Set-acl -aclobject $acl "ad:CN=System Management,CN=System,$root"

#endregion 

#region Extend AD Schema

& E:\VLab_Software\SystemCenter21012R2\Extracted\ConfigMgr_Extract\SMSSETUP\BIN\X64\extadsch.exe


#endregion

#region Open TCP Ports 1433 and 4022 for SQL Replication

#Create  a new GPO
New-GPO -Name SCCM_FireWall_Rule -Comment "This is to allow port 1433 and 4022 for sql replication" -Domain dexter.com -Verbose

#Open the NETGPO session to add the firewall rules to it
$GPOSession = Open-NetGPO -PolicyStore dexter.com\SCCM_FireWall_Rule -DomainController DexterDC -Verbose

#Add a new Firewall rules to the GPO session
New-NetFirewallRule -name AllowPort1433 -DisplayName "ALlow port 1433 for SQL Replication" -Profile Domain -Direction Inbound -Protocol TCP -LocalPort 1433 -LocalAddress 10.1.1.1/24  -GPOSession $GPOSession -Verbose	
New-NetFirewallRule -name AllowPort4022 -DisplayName "ALlow port 4022 for SQL Replication" -Profile Domain -Direction Inbound -Protocol TCP -LocalPort 4022 -LocalAddress 10.1.1.1/24  -GPOSession $GPOSession -Verbose	



#Save the New GPO session
Save-NetGPO -GPOSession  $GPOSession -Verbose	

#Now link the GPO to the Domain 
New-GPLink -Name SCCM_FireWall_Rule -Target "DC=Dexter,DC=Com" -LinkEnabled Yes

#Update the Group Policy on all the Computers...Not that many in my LAB
Get-ADComputer -Filter * -SearchBase "CN=Computers,DC=Dexter,DC=Com" | ForEach-Object -Process {Invoke-GPUpdate -RandomDelayInMinutes 0 -Force }

#endregion

#region Install required features

#create a PSSession to the remote Server..soon going to get SCCM installed
$Session = New-PSSession -ComputerName DexSCCM

$featuresneeded = "Web-Common-Http","Web-Static-Content","Web-Default-Doc","Web-Dir-Browsing","Web-Http-Errors","Web-Http-Redirect","Web-App-Dev","Web-Asp-Net","Web-Net-Ext","Web-ASP","Web-ISAPI-Ext","Web-ISAPI-Filter","Web-Health","Web-Http-Logging","Web-Log-Libraries","Web-Request-Monitor","Web-Http-Tracing","Web-Security","Web-Basic-Auth","Web-Windows-Auth","Web-Url-Auth","Web-Filtering","Web-IP-Security","Web-Performance","Web-Stat-Compression","Web-Mgmt-Tools","Web-Mgmt-Console","Web-Scripting-Tools","Web-Mgmt-Service","Web-Mgmt-Compat","Web-Metabase","Web-WMI","Web-Lgcy-Scripting","Web-Lgcy-Mgmt-Console"

Invoke-Command -Session $Session -ScriptBlock {Get-WindowsFeature -Name $using:FeaturesNeeded | where {$_.installed -eq $false } | Add-WindowsFeature }


#endregion


#regiond download .NET 4

$url = "http://download.microsoft.com/download/9/5/A/95A9616B-7A37-4AF6-BC36-D6EA96C8DAAE/dotNetFx40_Full_x86_x64.exe"

Start-BitsTransfer -Source $url -Destination E:\VLab_Software -Asynchronous

#endregion

#region Add BITS & RDC
Invoke-Command -Session $session -ScriptBlock {Get-WindowsFeature -Name "BITS","RDC" | Add-WindowsFeature -Verbose }

#endregion


