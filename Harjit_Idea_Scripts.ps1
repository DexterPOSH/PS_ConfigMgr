$PSDefaultParameterValues =@{"get-cimclass:namespace"="Root\SMS\site_DEX";"get-cimclass:computername"="DexSCCM";"get-cimInstance:computername"="DexSCCM";"get-ciminstance:namespace"="Root\SMS\site_DEX";"get-wmiobject:namespace"="Root\SMS\site_DEX";"get-WMIObject:computername"="DexSCCM"}


#region Script that can tell me which collection a particular system or several Systems belong to ?

$name = 'DexChef'

#$CIMinstance = Get-CimInstance -ClassName SMS_R_System -Filter "Name='$name'"
$CIMinstance = Get-CimInstance -Query "Select ResourceID FROM SMS_R_System WHERE Name='$name'"

$Allcollections = Get-CimInstance -Query "Select CollectionID,MemberClassName from SMS_Collection"
foreach ($collection in $Allcollections)
{

    if (Get-CimInstance -Query "Select * FROM $($collection.MemberClassname) WHERE ResourceID='$($CIMinstance.ResourceID)'")
    {
        "It is in collection ($collection.name)"
    }

}

<#

$ResID = (Get-CMDevice -Name "CLTwin7").ResourceID

   2: $Collections = (Get-WmiObject -Class sms_fullcollectionmembership -Namespace root\sms\site_HQ1 -Filter "ResourceID = '$($ResID)'").CollectionID

   3: foreach ($Collection in $Collections)

   4:     {

   5:         Get-CMDeviceCollection -CollectionId $Collection | select Name, CollectionID

   6:     }

#>


 function Get-ResourceCollectionMemberhip
 {
    <#
        .Synopsis
        Function to Add Packages to the DP
        .DESCRIPTION
        THis Function will connect to the SCCM Server SMS namespace and then Add the Package IDs
        passed to the Function for the specified DP name.
        .EXAMPLE
        PS> Add-SCCMDPContent -PackageID DEX123AB,DEX145CD -DPname DexDP -Computername DexSCCMServer  

        This will remove the Packages with Package IDs [ DEX123AB,DEX145CD] from the Distribution Point "DexDP".
        .INPUTS
        System.String[]
        .OUTPUTS
        System.Management.Automation.PSCustomObject
        .NOTES
        Author - DexterPOSH (Deepak Singh Dhami)

        Credits - MVP David O'Brien 
                  [http://www.david-obrien.net/2014/01/find-configmgr-collection-membership-client-via-powershell/]

    #>

     [CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact="Low")]
     [OutputType([int])]
     Param
     (
         #Specify the Device names
         [Parameter(Mandatory=$true,
                    ValueFromPipelineByPropertyName=$true,
                    Position=0)]
         [string[]]$Name,

         #Supply the SCCM Site Server hosting SMS Namespace Provider. Default - LocalMachine
        [Parameter()]
        [Alias('SCCMServer')]
        [String]$ComputerName = $env:COMPUTERNAME,

        #Specify the Resource type you are querying for e.g User/Device. [Default - Device]
        [Parameter()]
        [ValidateSet("User","Device")]
        $ResourceType="Device"
         
     )
 
     Begin
     {
        Write-Verbose -Message '[BEGIN] : Starting the Function'
        try
        {
            $sccmProvider = Get-WmiObject -Query 'select * from SMS_ProviderLocation where ProviderForLocalSite = true' -Namespace 'root\sms' -ComputerName $ComputerName -ErrorAction Stop
            # Split up the namespace path
            $Splits = $sccmProvider.NamespacePath -split '\\', 4
            Write-Verbose  -Message "Provider is located on $($sccmProvider.Machine) in namespace $($splits[3])"
 
            # Create a new hash to be passed on later
            $hash = @{'ComputerName' = $ComputerName;'ErrorAction' = 'Stop'}                      
                        
        }
        catch
        {
            Write-Warning  -Message 'Something went wrong while getting the SMS ProviderLocation or SMS_DistributionPoint Class Object'
            throw $Error[0].Exception
        }
     }
     Process
     {
        #foreach Device name in the input
        foreach($Resourcename in $name) 
        {
            switch -Exact ($ResourceType)
            {
                "Device" {
                    try
                    {
                        $ResourceID = Get-CimInstance -Query "Select ResourceID FROM SMS_CombinedDeviceResources WHERE NAME='$Resourcename'"  @hash | Select-Object -ExpandProperty  ResourcedID 
                        $MemberofCollections = Get-CimInstance -Query "Select CollectionID FROM SMS_FullCollectionMembership WHERE ResourceID='$ResourceID'" -ErrorVariable ResourceQuery @hash
                        
                        foreach ($Collection in $MemberofCollections)
                        {
                            $collectionInfo = Get-CimInstance -Query "Select Name FROM SMS_Collection WHERE CollectionID='$($Collection.CollectionID)'"
                            [PSCustomobject]@{
                                Name = $Resourcename;
                                ResourceType = 'Device';
                                CollectionName = $collectionInfo.Name;
                                CollectionID = $Collection.CollectionID
                            }
                        }
                        
                     }
                     catch
                     {
                        Write-Warning -Message $_.exception 

                     }
                }

                "User" {
                    

                }
            }
     }
     End
     {
     }
 }

#endregion

#region script to list the Servers ad when they were last rebooted which would indicate when they recieved last patches 


#endregion


#region Script to select one or Several applications and publish them to DPs


#endregion


#region Script to deploy one or several Apps and deploy them to one collection or several



#endregion