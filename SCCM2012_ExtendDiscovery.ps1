<#

PS DEX:\> $name= 'C:\Program Files (x86)\Microsoft System Center 2012 R2 Configuration Manager SDK\Redistributables\Micr
osoft.ConfigurationManagement.Messaging.dll'
PS DEX:\> [System.Reflection.Assembly]::LoadFrom($name)

GAC    Version        Location
---    -------        --------
False  v4.0.30319     C:\Program Files (x86)\Microsoft System Center 2012 R2 Configuration Manager SDK\Redistributab...

PS DEX:\> $object = New-Object -TypeName Microsoft.ConfigurationManagement.Messaging.Messages.Server.DiscoveryDataRecord
File("testing")
PS DEX:\> $object


AgentName          : testing
FileSuffix         : DDR
Architecture       : System
SiteCode           :
BuildNumber        : 5.00.7711.0000
IsSigned           : False
SigningCertificate :
Settings           : Microsoft.ConfigurationManagement.Messaging.Framework.MessageFileSettings
Trusted            : False

$object.AddStringProperty("Netbios Name",'key',64,'DexFakeMachine1')                                              
  49 $object.AddStringProperty("AD Site Name",'None',64,'Dexters LAB')                                                 
  50 $object.AddStringPropertyArray                                                                                    
  51 $object.AddStringPropertyArray("IP Addresses",'Array',64,'10.1.1.99','10.1.1.100')                                
  52 $object.AddStringPropertyArray("MAC Addresses",'Array',64,"00:02:A5:B1:11:68","00:02:A5:B1:11:69")                
  53 $object.SerializeToFile                                                                                           
  54 $object.SerializeToFile("C:\test\DexTest.ddr")                                                                    
  55 $object.SerializeToFile("C:\DexTest.ddr")                                                                         
  56 cd C:\                                                                                                            
  57 $object.SerializeToFile(".\DexTest.ddr")                                                                          
  58 $object.SerializeToFile(".")                                                                                      
  59 ii .                                                                                                              
  60 notepad .\KAIGOSKF.DDR                                                                                            
  61 $object.SerializeToInbox                                                                                          
  62 $object.SerializeToInbox("DexSCCM.dexter.com") 

#>


#region Create DDRs using the COM Interface

#load the DLL 
& regsvr32.exe 'C:\Program Files (x86)\Microsoft System Center 2012 R2 Configuration Manager SDK\Redistributables\amd64\smsrsgenctl.dll'

$Computer = "DextestMachine1"
#explicitly cast this as a String Array if single element present
[string[]]$IPAddress = "10.1.1.100"
[String[]]$MACAddress = "00:02:A7:B2:23:88"

#create the COM Object
$SMSDisc = New-Object -ComObject SMSResGen.SMSResGen.1

#Specify the Architecture, Name for the DDR and the SiteCode
$SMSDisc.DDRNew("System","myCustomAgent","DEX")

#Add the String Property Netbios Name - (Propertyname,Value,Width,DiscoveryFlag specifying that it is Key[Hex Value])
$SMSDisc.DDRAddString("Netbios Name", $Computer, 64, 0x8)

#Add the String Array Property IP addresses - (Propertyname,Value[],Width,DiscoveryFlag specifying it is an Array)
$SMSDisc.DDRAddStringArray("IP Addresses", $IPAddress, 64, 0x10)
$SMSDisc.DDRAddStringArray("MAC Addresses", $MACAddress, 64, 0x10)

#save the DDR to the Desktop
#NOte that the 
$SMSDisc.DDRWrite([System.Environment]::GetFolderPath("Desktop") + "\$Computer.DDR")

#now copy this DDR file to your SCCM Server DDM.Box inbox 
Copy-Item -path  "C:\Users\Administrator\Desktop\DextestMachine1.DDR" -Destination  "\\dexsccm\C$\Program Files\Microsoft Configuration Manager\inboxes\ddm.box" -Verbose


#endregion Create DDRs using the COM Interface




#region Create DDRs Using DiscoveryDataRecordFile Class

#this DLL loads up using Add-Type
Add-Type -Path 'C:\Program Files (x86)\Microsoft System Center 2012 R2 Configuration Manager SDK\Redistributables\Microsoft.ConfigurationManagement.Messaging.dll'

#initialize a new instance of the Object
$DDRFile = New-Object -TypeName  Microsoft.ConfigurationManagement.Messaging.Messages.Server.DiscoveryDataRecordFile("TestingDDR")

#By default the Architecture is System which refers to the SMS_R_system Class

#add the Key Property (Netbios Name is the Key Property for the System Architecture)
#AddStringProperty Method Defintion says - Property name, DiscoveryFlags, Width, Value
#You can see the enumeration for the DiscoveryFlags by using : [System.Enum]::GetValues("Microsoft.ConfigurationManagement.Messaging.Messages.Server.DdmDiscoveryFlags")
$DDRFile.AddStringProperty("Netbios Name",'key',64,'DexFakeMachine1')  

#add the AD Site Name String Property....have a look at the Method Definitions (use this in your Console For Ex: $DDRFile | Get-Member )
$DDRFile.AddStringProperty("AD Site Name",'None',64,'Dexters LAB')

#Add Array Properties
$DDRFile.AddStringPropertyArray("IP Addresses",'Array',64,[string[]]'10.1.1.99')   #Typecast as a string array if one value                          
$DDRFile.AddStringPropertyArray("MAC Addresses",'Array',64,"00:02:A5:B1:11:68","00:02:A5:B1:11:69")

#After populating the Properties just use the Method SerializeToInbox to send it to the SCCM server for Processing
$DDRFile.SerializeToInbox('DexSCCM.dexter.com') #SCCM Server name is passed if you are not working on it

#endregion Create DDRs Using DataSicoveryRecordFile Class
