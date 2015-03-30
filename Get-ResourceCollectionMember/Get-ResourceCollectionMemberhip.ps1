function Get-ResourceCollectionMemberhip
 {
<#
    .Synopsis
    Function to retrieve the Collection Names a User is part of.
    .DESCRIPTION
    This Function will connect to the SCCM Server SMS namespace and then get all the Collection the Resource is part of.
    It supports querying both the User and Device Resourcetype.
    .EXAMPLE
    PS>Get-ResourceCollectionMemberhip -Name dexterposh -ComputerName dexsccm -ResourceType User 

    Name                                          ResourceType                                  CollectionName                               CollectionID                                
    ----                                          ------------                                  --------------                               ------------                                
    dexterposh                                    User                                          TestUserCollection                           DEX00032                                    
    dexterposh                                    User                                          All Users                                    SMS00002                                    
    dexterposh                                    User                                          All Users and User Groups                    SMS00004    
    
    One has to specify the ResourceType User , if you are looking for the Collection Membership of User resources.                                
    .EXAMPLE
    PS>Get-ResourceCollectionMemberhip -Name dexchef -ComputerName dexsccm  

    Name                                          ResourceType                                  CollectionName                               CollectionID                                
    ----                                          ------------                                  --------------                               ------------                                
    dexchef                                       Device                                        Server2012                                   DEX00038                                    
    dexchef                                       Device                                        Server2008                                   DEX00039                                    
    dexchef                                       Device                                        All Systems                                  SMS00001                                    
    dexchef                                       Device                                        All Desktop and Server Clients               SMSDM003 

    Note - By default the Function looks for the Device ResourceType.
    .INPUTS
    System.String[]
    .OUTPUTS
    System.Management.Automation.PSCustomObject[]
    .NOTES
    Author - DexterPOSH (Deepak Singh Dhami)
    blog - www.dexterposh.com

    Credits - MVP David O'Brien 
                Adapted and extended the logic in the below blog post.
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
            $sccmProvider = Get-CimInstance -Query 'select * from SMS_ProviderLocation where ProviderForLocalSite = true' -Namespace 'root\sms' -ComputerName $ComputerName -ErrorAction Stop
            # Split up the namespace path
            $Splits = $sccmProvider.NamespacePath -split '\\', 4
            Write-Verbose  -Message "Provider is located on $($sccmProvider.Machine) in namespace $($splits[3])"
 
            # Create a new hash to be passed on later
            $hash = @{'ComputerName' = $ComputerName;'NameSpace' = $Splits[3];'ErrorAction' = 'Stop'}                      
                        
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
            Write-Verbose -Message "[PROCESS] : Processing the $Resourcename"
            switch -Exact ($ResourceType)
            {
                "Device" {
                    
                    Write-Verbose -Message "[PROCESS] : ResourceType matched to Device"
                    try {
                        
                        #Get the Resource ID for the Resource using the input Name
                        Write-Verbose -Message "[PROCESS] : Querying the SMS_CombinedDeviceResources Class to fetch the ResourceID"
                        $ResourceID = Get-CimInstance -Query "Select ResourceID FROM SMS_CombinedDeviceResources WHERE NAME='$Resourcename'"  @hash | Select-Object -ExpandProperty  ResourceID 
                        
                        #Once we have the ResourceID thenwe can use it to get the CollectionIDs of the Collection the Resource is member of
                        Write-Verbose -Message "[PROCESS] : Fetching the CollectionIDs the Resource is part of"
                        $MemberofCollections = Get-CimInstance -Query "Select CollectionID FROM SMS_FullCollectionMembership WHERE ResourceID='$ResourceID'" -ErrorVariable ResourceQuery @hash
                        
                        #Iterate the list of CollectionIDs the Resource is membre of and Spit out a Custom Object
                        Write-Verbose -Message "[PROCESS] : Iterating over fetched CollectionIds now"
                        foreach ($Collection in $MemberofCollections) {                        
                            
                            Write-Verbose -Message "[PROCESS] : Getting the Name for CollectionID $($Collection.CollectionID)"
                            $collectionInfo = Get-CimInstance -Query "Select Name FROM SMS_Collection WHERE CollectionID='$($Collection.CollectionID)'" @hash
                            
                            [PSCustomobject]@{
                                Name = $Resourcename;
                                ResourceType = 'Device';
                                CollectionName = $collectionInfo.Name;
                                CollectionID = $Collection.CollectionID
                            }
                        }
                        
                     }
                     catch {                     
                        Write-Warning -Message $_.exception 
                     }
                }

                "User" {
                    
                    Write-Verbose -Message "[PROCESS] : ResourceType matched to User"
                    try {
                        
                        #Get the Resource ID for the Resource using the input Name. Note that with the User Resources , forced to use the WQL LIKE Operator
                        #Because the LIKE operator can return more than one instance of the Objects.
                        Write-Verbose -Message "[PROCESS] : Querying the SMS_CombinedUserResources Class to fetch the ResourceID"
                        $ResourceIDs = @(Get-CimInstance -Query "Select ResourceID FROM SMS_CombinedUserResources WHERE Name LIKE'%$Resourcename%'"  @hash)
                        
                                                                      
                        foreach ($ResourceID in $ResourceIDs) {
                            
                            #Once we have the ResourceID then we can use it to get the CollectionIDs of the Collection the Resource is member of
                            Write-Verbose -Message "[PROCESS] : Fetching the CollectionIDs the Resource is part of"
                            $MemberofCollections = Get-CimInstance -Query "Select CollectionID FROM SMS_FullCollectionMembership WHERE ResourceID='$($ResourceID.ResourceID)'" -ErrorVariable ResourceQuery @hash
                        
                            #Iterate the list of CollectionIDs the Resource is membre of and Spit out a Custom Object
                            Write-Verbose -Message "[PROCESS] : Iterating over fetched CollectionIds now"
                            foreach ($Collection in $MemberofCollections) { 
                                                   
                                Write-Verbose -Message "[PROCESS] : Getting the Name for CollectionID $($Collection.CollectionID)"
                                $collectionInfo = Get-CimInstance -Query "Select Name FROM SMS_Collection WHERE CollectionID='$($Collection.CollectionID)'" @hash
                            
                                [PSCustomobject]@{
                                    Name = $Resourcename;
                                    ResourceType = 'User';
                                    CollectionName = $collectionInfo.Name;
                                    CollectionID = $Collection.CollectionID
                                    }
                            }#end foreach($collection in $memberofCollections)
                        }
                        

                    }
                    catch {
                         Write-Warning -Message $_.exception 
                    }
                }
            }
     }
     }
     End
     {
        Write-Verbose -Message '[END] : Ending the Function'
     }
 }