
function Add-SCCMDPContent
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

    #>

    [CmdletBinding()]
    [OutputType([PSObject])]
    Param
    (
        # Specify the Package IDs which need to be removed from the DP
        [Parameter(Mandatory,
                   ValueFromPipelineByPropertyName,
                   Position = 0)]
        [string[]]$PackageID,

        # Pass the DP name where cleanup is to be done
        [Parameter(Mandatory = $true)]
        [String]$DPName,

        #Supply the SCCM Site Server hosting SMS Namespace Provider
        [Parameter()]
        [Alias('SCCMServer')]
        [String]$ComputerName
    )

    Begin
    {
        Write-Verbose -Message '[BEGIN] Starting the Function'
        try
        {
            $sccmProvider = Get-WmiObject -Query 'select * from SMS_ProviderLocation where ProviderForLocalSite = true' -Namespace 'root\sms' -ComputerName $ComputerName -ErrorAction Stop
            # Split up the namespace path
            $Splits = $sccmProvider.NamespacePath -split '\\', 4
            $NALPath = '["Display=\\{0}\"]MSWNET:["SMS_SITE={1}"]\\{0}\' -f $sccmProvider.Machine,$sccmProvider.SiteCode
            Write-Verbose  -Message "Provider is located on $($sccmProvider.Machine) in namespace $($splits[3])"
 
            # Create a new hash to be passed on later
            $hash = @{'ComputerName' = $ComputerName;'NameSpace' = $Splits[3];'ErrorAction' = 'Stop'}
            
            $WMIClass = Get-WmiObject -Class = 'SMS_DistributionPoint' -List @hash
        }
        catch
        {
            Write-Warning  -Message 'Something went wrong while getting the SMS ProviderLocation or SMS_DistributionPoint Class Object'
            throw $Error[0].Exception
        }
    }
    Process
    {
        
            
            Write-Verbose -Message "[PROCESS] Working to add packages to DP --> $DPName  "
            
            #get all the packages in the Distribution Point
            foreach ($ID in $PackageID)
            {
                TRY {
                    if (Get-WmiObject -Query "Select PackageID from SMS
                    $newInstance = $WMIClass.createInstance()
                    $newinstance.PackageID = $ID
                    $newinstance.ServerNALPath = $NALPath
                    $newinstance.SiteCode = $sccmProvider.sitecode
                    $newinstance.put()     
                
                }                                                 
                CATCH
                {
                    Write-Warning  -Message "[PROCESS] Something went wrong while adding the Package with PackageID $ID  from $DPname"
                    Throw $_.exception 
                }
            }#End Foreach-Object
            
        }#End Process
    End
    {
        Write-Verbose  -Message '[END] Ending the Function'
    }
}
