
#requires -version 3.0
Set-StrictMode -Version latest
$VerbosePreference = 'continue' #setting this as all the verbose messages are displayed on the background console window (hidden by default)

<#
Credits - Below links have been very helpful
http://stackoverflow.com/questions/21935398/powershell-observablecollection-predicate-filter
#>

#region XAML definition
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="POSH-Deploy v1 - by DexterPOSH" Height="748" Width="1272" ResizeMode="NoResize" Background="#FF5397B4">
    <Grid Height="713" Width="1253" Background="#FF5397B4">
        <Label Content="Choose action[ Add/ Remove machines to collection]" Height="28" HorizontalAlignment="Left" Margin="10,10,0,0" Name="labelAction" VerticalAlignment="Top" Width="320" />
        <Label Content="Enter Machine Names (one per line)" Height="28" HorizontalAlignment="Left" Margin="421,10,0,0" Name="labelMachine" VerticalAlignment="Top" Width="222" />
        <Button Content="Action" Height="67" HorizontalAlignment="Left" Margin="699,31,0,0" Name="buttonaction" VerticalAlignment="Top" Width="75" ToolTip="Add/ Remove Machine name from the Query Rules" />
        <TextBox Height="113" HorizontalAlignment="Left" Margin="421,31,0,0" Name="textBoxComputer" VerticalAlignment="Top" Width="203" AcceptsReturn="True" VerticalScrollBarVisibility="Auto" ToolTip="Add Machine Name (one per line)" />
        <CheckBox Content="Add Name to Query" Height="30" HorizontalAlignment="Left" Margin="34,42,0,0" Name="checkBoxAdd" VerticalAlignment="Top" Width="144" ToolTip="Check this if want to add machine names to the Collection" />
        <CheckBox Content="Remove Name from Query" Height="30" HorizontalAlignment="Left" Margin="228,42,0,0" Name="checkBoxRemove" VerticalAlignment="Top" Width="162" />
        <StackPanel Height="515" HorizontalAlignment="Left" Margin="12,152,0,0" Name="stackPanel1" VerticalAlignment="Top" Width="762">
            <DataGrid  Height="517" Name="dataGridApps" Width="753" RowHeight="30" ColumnWidth="100" GridLinesVisibility="Horizontal" HeadersVisibility="All" AutoGenerateColumns="True" ClipToBounds="True" HorizontalContentAlignment="Center" FontSize="13" ToolTip="Double Click on a Row to select" IsReadOnly="True" />
        </StackPanel>
        <DataGrid AutoGenerateColumns="True" Height="305" HorizontalAlignment="Left" Margin="812,362,0,0" Name="dataGridSelectedApps" VerticalAlignment="Top" Width="415" SelectionUnit="CellOrRowHeader" ToolTip="double click on a row to un-select it" IsReadOnly="True" />
        <Label Content="default -current user" Height="38" HorizontalAlignment="Left" Margin="1007,0,0,0" Name="labelCreds" VerticalAlignment="Top" Width="134" />
        <TextBox Height="30" HorizontalAlignment="Left" Margin="807,31,0,0" Name="textBoxServer" VerticalAlignment="Top" Width="182" ToolTip="Input the SCCM server having SMS namespace provider installed" />
        <Label Content="SCCM Server (SMS Namespace)" Height="25" HorizontalAlignment="Left" Margin="810,0,0,0" Name="labelServer" VerticalAlignment="Top" Width="179" />
        <Label Content="Select Collections" Height="30" HorizontalAlignment="Left" Margin="16,116,0,0" Name="labelSelectApps" VerticalAlignment="Top" Width="119" />
        <Label Content="Log Information" Height="32" HorizontalAlignment="Left" Margin="809,79,0,0" Name="labelLog" VerticalAlignment="Top" Width="230" />
        <Label Content="Selected Collections" Height="37" HorizontalAlignment="Left" Margin="807,319,0,0" Name="labelSelectedApps" VerticalAlignment="Top" Width="205" />
        <CheckBox Content="Show PS Window" Height="26" HorizontalAlignment="Left" Margin="810,673,0,0" Name="checkBoxPSWindow" VerticalAlignment="Top" Width="136" ToolTip="Shows the PS Console running in background" />
        <Button Content="Test SMS Connection" Height="28" HorizontalAlignment="Left" Margin="1113,33,0,0" Name="buttonTestSMSConnection" VerticalAlignment="Top" Width="127" ToolTip="Tests the connectivity to the SMS Namespace (using the creds specified)" />
        <TextBox Height="28" HorizontalAlignment="Left" Margin="253,116,0,0" Name="textBoxSearch" VerticalAlignment="Top" Width="150" />
        <Label Content="Search Collections" Height="38" HorizontalAlignment="Left" Margin="136,116,0,0" Name="labelSearch" VerticalAlignment="Top" Width="111" />
        <Button Content="Copy" Height="30" HorizontalAlignment="Left" Margin="941,319,0,0" Name="buttonCopy" VerticalAlignment="Top" Width="71" ToolTip="Copies the App names in the clipboard" />
        <Button Content="Sync Collections List" Height="25" HorizontalAlignment="Left" Margin="34,676,0,0" Name="buttonSyncApps" IsEnabled = "false" VerticalAlignment="Top" Width="118" ToolTip="Will fetch all the collections on the SMS Server " />
        <Button Content="Clear" Height="30" HorizontalAlignment="Left" Margin="1029,319,0,0" Name="buttonClear" VerticalAlignment="Top" Width="75" ToolTip="Clears the selected app list" />
        <CheckBox Content=" Alternate Creds" Height="27" HorizontalAlignment="Left" Margin="1003,34,0,0" Name="checkBoxCred" VerticalAlignment="Top" Width="101" />
        <Button Content="Browse" Height="23" HorizontalAlignment="Left" Margin="630,124,0,0" Name="buttonBrowse" VerticalAlignment="Top" Width="72" ToolTip="Browse for file with Machine Names" />
        <Label Content="Created by DexterPOSH" Height="31" HorizontalAlignment="Left" Margin="1089,676,0,0" Name="labelName" VerticalAlignment="Top" Width="164" FontSize="14" Foreground="Red"></Label>
        <Image Height="36" HorizontalAlignment="Left" Margin="1044,673,0,0" Name="image1" Stretch="Fill" VerticalAlignment="Top" Width="37" />
        <Button Content="Collection Integrity Check" Height="25" HorizontalAlignment="Left" Margin="186,676,0,0" Name="buttonIntegrity" IsEnabled = "false" VerticalAlignment="Top" Width="166" ToolTip="Does a check on the last 3 collections in PS_deploy.csv"/>
        <ListBox Height="208" HorizontalAlignment="Left" Margin="800,105,0,0" Name="listBoxLog" VerticalAlignment="Top" Width="427" />
    </Grid>
</Window>

"@
Add-Type -AssemblyName PresentationFramework 
Add-Type -AssemblyName System.Windows.Forms
$reader=(New-Object System.Xml.XmlNodeReader $xaml)
$Window=[Windows.Markup.XamlReader]::Load( $reader )

#endregion XAML definition

#region Setup log for the changes made to the Collections

Write-Host -ForegroundColor Red "########################----POSH Deploy----######################## "	
Write-Host -ForegroundColor 'Cyan'  "Welcome to the ConfigMgr (SCCM) deployment tool.!!!"
Write-Host -ForegroundColor 'Cyan' "Designed and Created by Deepak Singh Dhami (@DexterPOSH)"
$host.UI.RawUI.WindowTitle = "POSH Deploy by DexterPOSH"

 
$PSdeployfile = "$([System.Environment]::GetFolderPath('MyDocuments'))\PS_Deploy.csv"


#As part of the recovery plan the Tool will log all the deployments to a file PS_Deploy.csv (in User MyDocuments) until the size of it grows older than 10MB
if (Test-Path -Path $PSDeployFile -Type leaf )
{
		
	if ( $((Get-Item -Path 	$PSDeployFile).length/1MB) -ge 10MB)
	{
		Write-Verbose -message "[POSH Deploy] Size exceeded 10MB on $(get-date)..so taking backup"
        Rename-Item -Path $PSDeployFile  -NewName PS_Deploy_bak.csv -Force
	}
	else
	{
		Write-Verbose -message "[POSH Deploy] Size of PS_Deploy.csv is below 5 MB...continuing"
	}
    
}
else
{
	Write-Verbose -Message "[POSH Deploy] PS_Deploy.csv not found....creating one in $([System.Environment]::GetFolderPath('MyDocuments'))"
    New-Item -Path $PSDeployFile -ItemType file -Verbose
    
}



#endregion


#region Connect to Control
$buttonaction = $Window.FindName("buttonaction")
$textBoxComputer = $Window.FindName('textBoxComputer')
$checkBoxAdd = $Window.FindName('checkBoxAdd')
$checkBoxRemove = $Window.FindName('checkBoxRemove')
$buttonBrowse = $Window.FindName("buttonBrowse")
$dataGridApps = $Window.FindName('dataGridApps')
$dataGridSelectedApps = $Window.FindName('dataGridSelectedApps')
$textBoxServer = $Window.FindName('textBoxServer')
$checkBoxPSWindow = $Window.FindName('checkBoxPSWindow')
$buttonTestSMSConnection = $Window.FindName('buttonTestSMSConnection')
$textBoxSearch = $Window.FindName('textBoxSearch')
$buttonCopy = $Window.FindName('buttonCopy')
$buttonSyncApps = $Window.FindName('buttonSyncApps')
$buttonclear = $Window.FindName('buttonClear')
$checkBoxCred = $Window.FindName('checkBoxCred')
$DexLabel = $Window.FindName('labelName')
$ListBoxLog = $Window.findName('listBoxLog')
$buttonIntegrity=$Window.FindName('buttonIntegrity')
#endregion Connect to Control


#region Customize the GUI
$LogCollection = New-Object System.Collections.ObjectModel.ObservableCollection[String] 
$ListBoxLog.ItemsSource =  [System.Windows.Data.CollectionViewSource]::GetDefaultView($LogCollection)

$LogCollection.Add("Welcome to POSH-Deploy, Enter your SCCM Server name & Hit Test Connection to begin")

$dataGridApps.background="LightGray" 
$dataGridApps.RowBackground="LightYellow" 
$dataGridApps.AlternatingRowBackground="LightBlue"
$buttonaction.Background = 'Yellow'

$dataGridSelectedApps.background="LightGray" 
$dataGridSelectedApps.RowBackground="LightGreen" 
$dataGridSelectedApps.AlternatingRowBackground="LightYellow"

$textBoxServer.text = $env:COMPUTERNAME

$collectionView = New-Object System.Collections.ObjectModel.ObservableCollection[object] 
if (Test-Path -Path "$([System.Environment]::GetFolderPath('MyDocuments'))\Collection.csv" )
{
     Import-Csv -Path "$([System.Environment]::GetFolderPath('MyDocuments'))\Collection.csv"| ForEach-Object -Process {$collectionView.Add($_)}
}
else
{
    $LogCollection.Add("Collection.csv not found. After Test Connection, Hit Sync Collection List")
}

$view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($collectionView)
$filter = ''
$view.Filter = {param($item) $item -match $filter}
$view.Refresh()


$dataGridApps.ItemsSource = $view
$buttonaction.IsEnabled = $false #This will be disabled at start unless the Test SMS Connection is successful


#endregion customize the GUI



#region Function definitions
#region helper Functions
			
Function Invoke-SCCMServiceCheck
{
	[CmdletBinding()]
	#[OutputType([System.Int32])]
	param(
		[Parameter(Position=0, Mandatory=$true,
					helpmessage="Enter the ComputerNames to refresh machine policy on",
				    ValueFromPipeline=$true,
				    ValueFromPipelineByPropertyName=$true
					)]
		[ValidateNotNullOrEmpty()]
		[System.String[]]
		$ComputerName
		)
    BEGIN
    {
       Function Set-RemoteService
       {
        [CmdletBinding()]
        [OutputType([PSObject])]
        param(
	        [Parameter(Position=0, Mandatory=$false,
				        helpmessage="Enter the ComputerNames to Chec & Fix Services on",
				        ValueFromPipeline=$true,
				        ValueFromPipelineByPropertyName=$true
				        )]
	        [String[]]$ComputerName=$env:COMPUTERNAME,

            [Parameter(Mandatory=$true,
                        helpmessage="Enter the Service Name (Accepts WQL wildacard)")]
            [string]$Name,

    
            [Parameter(Mandatory=$true,helpmessage="Enter the state of the Service to be set")]
            [ValidateSet("Running","Stopped")]
            [string]$state,

            [Parameter(Mandatory=$true,helpmessage="Enter the Startup Type of the Service to be set")]
            [ValidateSet("Automatic","Manual","Disabled")]
            [string]$startupType
	        )
                BEGIN
                {
                    Write-Verbose -Message "[Invoke-SCCMServiceCheck][Set-RemoteService] - Starting the Function."      
                              
                }
		        PROCESS 
		        {
			        foreach ($computer in $computername )
			        {
				        Write-Verbose -Message "[Invoke-SCCMServiceCheck][Set-RemoteService] - Checking if $Computer is online"
				        if (Test-Connection -ComputerName $Computer -Count 2 -Quiet)
                        {
                            Write-Verbose -message "[Invoke-SCCMServiceCheck][Set-RemoteService] - $Computer is online"
                            #region try to set the required state and StartupType of the Service
                            try
                            {
                                $service = Get-WmiObject -Class Win32_Service -ComputerName $Computer -Filter  "Name LIKE '$Name'"  -ErrorAction Stop
                                #Check the State and set it
                                if ( $service.State -ne "$state")
                                {
                                    #Set the State of the Remote Service
                                    switch -exact ($state)
                                    {
                                        'Running' 
                                        {
                                            $changestateaction = $service.startService()
                                            Start-Sleep -Seconds 2 #it will require some time to process action
                                            if ($changestateaction.ReturnValue -ne 0 )
                                            {
                                                $err = Invoke-Expression  "net helpmsg $($changestateaction.ReturnValue)" 
                                                 Write-Warning -message  "[Invoke-SCCMServiceCheck][Set-RemoteService] - $Computer couldn't change state to $state `nWMI Call Returned :$err" 
                                         
                                            }
                                            break
                                     
                                        }
                                    
                                        'Stopped' 
                                        {
                                            $changestateaction = $service.stopService()
                                            Start-Sleep -Seconds 2 
                                            if ($changestateaction.ReturnValue -ne 0 )
                                            {
                                                $err = Invoke-Expression  "net helpmsg $($changestateaction.ReturnValue)" 
                                                Write-Warning -message  "[Invoke-SCCMServiceCheck][Set-RemoteService] -  $Computer couldn't change state to $state `nWMI Call Returned :$err" 
                                            }
                                            break
                                        }
                                    
                                    } #end switch
                                } #end if

                                #Check the StartMode and set it
                                if ($service.startMode -ne $startupType)
                                {
                            
                                    #set the Start Mode of the Remote Service
                                    $changemodeaction = $service.ChangeStartMode("$startupType")
                                    Start-Sleep -Seconds 2
                                    if ($changemodeaction.ReturnValue -ne 0 )
                                    {
                                        $err = Invoke-Expression  "net helpmsg $($changemodeaction.ReturnValue)" 
                                        Write-Warning -message  "[Invoke-SCCMServiceCheck][Set-RemoteService] -  $Computer couldn't change startmode to $startupType `nWMI Call Returned :$err" 
                                    }
                                
                                } #end if
                                                     
                                #Write the Object to the Pipeline
                                Get-WmiObject -Class Win32_Service -ComputerName $Computer -Filter  "Name LIKE '$Name'" -ErrorAction Stop | Select-Object -Property @{Label="ComputerName";Expression={"$($_.__SERVER)"}},@{Label="ServiceName";Expression={$_.Name}},StartMode,State       
            
                            }#end try
                            catch
                            {
                                Write-Warning -Message "[Invoke-SCCMServiceCheck][Set-RemoteService] - $Computer :: $_.exception"
                            } #end catch

                            #endregion try to set the required state and StartupType of the Service											
			        }
                    else
                    {
                        Write-Verbose -Message "[Invoke-SCCMServiceCheck][Set-RemoteService] - $Computer is Offline"
                    }
						
		        } #end foreach ($computer in $Computername)
	        }#end PROCESS
            END
            {
                Write-Verbose -Message "[Invoke-SCCMServiceCheck][Set-RemoteService] - Ending the Function"
                 
            }

        } 
    }
	PROCESS 
	{
		foreach ($computer in $computername )
		{
			try
			{
				#Automatic Updates Service set to Running(Auto)
				Set-RemoteService -ComputerName $computer -Name Wuauserv -StartupType Automatic -state Running -ErrorAction Stop
				Write-Verbose "Invoke-SCCMServiceCheck : $computer --> Automatic Updates Service Checked"
									
				#WMI Service set to Running(Auto)
				Set-RemoteService -ComputerName $computer -Name Winmgmt -StartupType Automatic -state Running -ErrorAction Stop
				Write-Verbose "Invoke-SCCMServiceCheck : $computer --> Windows Management Service Checked"
									
				#Remote Registry Service set to Running(Auto)
				Set-RemoteService -ComputerName $computer -Name RemoteRegistry -StartupType Automatic -state Running -ErrorAction stop
				Write-Verbose "Invoke-SCCMServiceCheck : $computer --> Remote Registry Service Checked"
									
				#SMS Agent Host (CcmExec) Service set to Running(Auto)
				Set-RemoteService -ComputerName $computer -Name CcmExec -StartupType Automatic -state Running -ErrorAction stop
				Write-Verbose "Invoke-SCCMServiceCheck : $computer --> SMS Agent Host Service Checked"
									
				#BITS  Service set to Running(Auto)
				Set-RemoteService -ComputerName $computer -Name Bits -StartupType Automatic -state Running -ErrorAction Stop
				Write-Verbose "Invoke-SCCMServiceCheck : $computer --> BITS service Checked"
									
									
			}
			catch
			{
				Write-Warning "Invoke-SCCMServiceCheck : $computer --> One of the SMS service is not running & could not be fixed. "
			}
							
		} #end Foreach
	} #end PROCESS
}
		
		
Function Invoke-MachinePolicyRefresh 
{
[CmdletBinding()]
#[OutputType([System.Int32])]
param(
	[Parameter(Position=0, Mandatory=$true,
				helpmessage="Enter the ComputerNames to refresh machine policy on",
				ValueFromPipeline=$true,
				ValueFromPipelineByPropertyName=$true
				)]
	[ValidateNotNullOrEmpty()]
	[System.String[]]
	$ComputerName
	)
	PROCESS 
	{
		foreach ($computer in $computername )
	    {
			    try
			    {
				   
                    $SMSCli = Get-WmiObject -Class SMS_Client -Namespace root\ccm -List #used Get-WMIObject as this will be handled with alternate creds
                    #$SMSCli = [wmiclass]"\\$computer\root\ccm:sms_client"
				    $SMSCli.TriggerSchedule('{00000000-0000-0000-0000-000000000021}') | Out-Null
				    Write-Verbose " Invoke-MachinePolicyRefresh: $computer --> Machine Policy refreshed"
                    
			    }
			    catch
			    {
				    Write-Warning " Invoke-MachinePolicyRefresh: $computer --> Could not Refresh Machine Policy"
			    }
	    } #end foreach
	}#end PROCESS
}


Function Connect-SCCMServer {
    # Connect to one SCCM server
    [CmdletBinding()]
    PARAM (
        [Parameter(Mandatory=$false,HelpMessage="SCCM Server Name or FQDN",ValueFromPipeline=$true)][Alias("ServerName","FQDN")][String] $SCCMServer = (Get-Content env:computername),
        [Parameter(Mandatory=$false,HelpMessage="Credentials to use" )][System.Management.Automation.PSCredential]$credential = $null
    )
 
    PROCESS {

           
            Write-Verbose -Message "[Connect-SCCMServer] Trying to connect to  SCCM server WMI"
            #region open a CIM session
            $Params = @{ComputerName = $SCCMServer;ErrorAction = 'Stop'}          
            if ($credential)
            {
                 $Params.Add('Credential',$credential)
            }
                
            try
            {
                 
             
            #region create the Hash to be used later for WMI queries   
                $sccmProvider = Get-WmiObject -query "select * from SMS_ProviderLocation where ProviderForLocalSite = true" -Namespace "root\sms" @Params
                # Split up the namespace path
                $Splits = $sccmProvider.NamespacePath -split "\\", 4
                Write-Verbose "[Connect-SCCMServer] Provider is located on $($sccmProvider.Machine) in namespace $($splits[3])"
 
                # Create a new hash to be passed on later
                $WMIHash= @{"ComputerName"=$SCCMServer;"NameSpace"=$Splits[3];"ErrorAction"="Stop"}
                
                    if ( $checkBoxCred.IsChecked)
                    {
                        Write-Verbose -Message "[POSH Deploy] Setting the PSDefaultParameterValues for Get-WMIObject to use alternate Creds"
                        $LogCollection.Add("Setting the PSDefaultParameterValues for Get-WMIObject to use alternate Creds")
                        $Script:PSDefaultParameterValues = @{"Get-WMIObject:Credential"=$Script:Cred} #setting the Credentials for all the WMI Calls as CIM Session can't be used
                    }
                Write-Output -InputObject $WMIHash
                
                #endregion create the Hash to be used later for CIM queries
            }
            catch
            {
               # Write-Warning "[Connect-SCCMServer] Something went wrong"
                $LogCollection.Add("Something went wrong while connecting to SMS Provider. Try again")
                Write-Warning -Message "[Connect-SCCMServer] $_.exception"
            }

            Write-Verbose -Message "[Connect-SCCMServer] Ending the Function"
    }
}


#endregion 
		
		
#region Hide and Show Console definitions
# Credits to - http://powershell.cz/2013/04/04/hide-and-show-console-window-from-gui/
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();
 
[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
		 
function Show-Console {
	$consolePtr = [Console.Window]::GetConsoleWindow()
	#5 show
	[Console.Window]::ShowWindow($consolePtr, 5)
}
		 
function Hide-Console {
	$consolePtr = [Console.Window]::GetConsoleWindow()
	#0 hide
	[Console.Window]::ShowWindow($consolePtr, 0)
}
	
#endregion Hide and Show Console definitions


#region Add-MachineNametoSCCMCollection
function Add-MachineToSCCMCollection
	{
	[CmdletBinding()]
	[OutputType([PSObject])]
	Param
	(
        # Enter the Computer Name
		[Parameter(Mandatory,
		            helpmessage="Enter the ComputerNames array to remove from the Collection",
		            ValueFromPipeline=$true,
		            ValueFromPipelineByPropertyName=$true
		            )]
		[Alias("CN","computer")]
		[String[]]
		$computername,

        #Specify the Collection ID
		[Parameter()]
		[validatenotnullorempty()]
		#[ValidatePattern('^[A-Za-z]{3}\w{5}$')]
		[string]
		$CollectionId,

		# Specify the CIM Hash generated by Connect-SCCMServer
		[Parameter(Mandatory)]
        [validatenotnullorempty()]
		[hashtable]$WMIHash        
		
		
	)
	#Set-StrictMode -Version 2
	Begin
	{
		
	    Write-Verbose -Message "Add-MachineToSCCMCollection: Starting the function"         
        
        #ScriptBlock to create a new Query Rule  
		$createNewAutomatedQueryRule = { 
            Write-Verbose "Add-MachineToSCCMCollection: Auotmated_QuerRule not found for this collection..creating one"
					
			$QueryRuleClass = Get-WmiObject -Class SMS_CollectionRuleQuery -List @WMIHash #Returns back the class Object
            $TempQueryRule = $QueryRuleClass.createinstance()
            $TempQueryRule.QueryExpression = 'select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System where SMS_R_System.NetbiosName in ("null")'
            $TempQueryRule.RuleName = 'Automated_QueryRule'
            $null = $Collection.AddMembershipRule($TempQueryRule) 
            $Collection.get()
            
			$QueryRules = $Collection.CollectionRules			
			Write-Verbose "Add-MachineToSCCMCollection: Auotmated_QuerRule Created and saved for the Collection"
			$Automated_QueryRule = $QueryRules | where {$_.RuleName -eq "Automated_QueryRule"} | Sort-Object -Property {$_.QueryExpression.Length} | Select-Object -First 1
			Write-Output -InputObject $Automated_QueryRule		
        }
	}
			
	Process
	{
				
		try 
		    {
                $Collection = Get-WmiObject -Class SMS_Collection -Filter "CollectionID='$collectionid'" @Script:WMIHash
		    
		        Write-Verbose "Add-MachineToSCCMCollection: Queried the Collection $($collection.name) with CollectionId : $collectionid  successfully"
				$Collection.Get() #Invoke Get Method to get the Lazy Properties back

			    #get the Collection Rules in an array 
			    $QueryRules = @($Collection.CollectionRules)
		        
		        Write-Verbose "Add-MachineToSCCMCollection: Queried the QueryRules on the Collection successfully"
		    }
		          
		catch 
		    {
		        Write-Warning -Message "Add-MachineToSCCMCollection: Something went wrong while querying the Collection and the Collection Rules on it"
                $LogCollection.Add("Add-MachineToSCCMCollection: $_.exception")
		    }
				
		
		#The Script won't make any changes in the existing query but will only look for a QueryRule by name "Automated_QueryRule		
		If ($QueryRules) 
        {	
            # If already Query Rules exist..then look among them for "automated_queryrule"	
		    If (!($Automated_QueryRule = $QueryRules | where { ($_.__Class -eq "SMS_CollectionRuleQuery")-and ($_.RuleName -eq "Automated_QueryRule")} | Sort-Object -Property {$_.QueryExpression.Length} | Select-Object -First 1))
		    {
							
			   $Automated_QueryRule = & $createNewAutomatedQueryRule
		    }
        }
        else
        {
            #If there are no Query Rules for the Collection go ahead and create one
            $Automated_QueryRule = & $createNewAutomatedQueryRule
        }
				
				
		#make the QueryExpression ready for adding machinenames to it
		$tempQueryExpression = $Automated_QueryRule.QueryExpression
				
		#remove the last ')' in the QueryRule
		$tempQueryExpression = $tempQueryExpression.remove(($tempQueryExpression.length -1 ))
		
		Foreach ($computer in $computername)
		{	
			if ($tempQueryExpression -match $computer)
			{
				#this does a basic check to see if the machine name is already addded in the Automated_QueryRule
				#Won't check the other QueryRules as ...one can always run the Remove-MachineFromCollection to remove the machine name from all QUeryRules and then add
				Write-Verbose "Add-MachineToSCCMCollection: The machine name $computer is already there in the Automated_QueryRule"
			}
			else 
			{
				Write-Verbose -Message "Add-MachineToSCCMCollection: Adding machine name $computer to the Automated_QueryRule."
				#Adds machine name in the body at the end of the query
				$Tempqueryexpression = $tempqueryexpression + ",`"$computer`""
											
			}
								
		} #end Foreach ($computer in $computername)
			
	    
		#add the  last ')' in the QueryRule
		$tempQueryExpression = $tempQueryExpression + ')'
		
        #One more check done here to validate the QueryExpression
		if ( (Invoke-WmiMethod -Class SMS_CollectionRuleQuery -Name ValidateQuery -ArgumentList $tempQueryExpression @WMIHash).returnvalue )
		{	
            Write-Verbose -Message "Add-MachineToSCCMCollection: The QueryExpression is Validated"
        }
        else
        {
            $LogCollection.add("Add-MachineToSCCMCollection: The QueryExpression created is not valid" )
        }
		
		if ($($Automated_QueryRule.queryExpression).length -lt $tempQueryExpression.length)
		{
			#after the text manipulation, save the new query...only if the QueryExpression has been modified
			$Automated_QueryRule.queryexpression = $Tempqueryexpression

            #region log the changes

	        Write-Verbose -Message "Add-MachineToSCCMCollection: Taking the backup in POSH_Deploy.csv"
	        #Now the QueryExpression has been modified...So before proceeeding take a backup
	        try
            {
                [pscustomobject]@{"Collection"=$collection.Name ;
	                            "CollectionId"=$CollectionId;
								"Action" = 'Add';
	                            "MachineNames"=[string]$computername;
	                            "QueryName"='Automated_QueryRule'; #Adding machine names always done to this QueryRule
	                            "QueryId"=$Automated_QueryRule.QueryId;
                                "QueryExpression"=$Automated_QueryRule.QueryExpression
	                        }| Export-Csv -NoTypeInformation -Path $PSdeployfile -Append
	        
            #Put the QueryRule in PS_Audit.csv
            Add-Content -Value "$($Collection.Name); $($Automated_QueryRule.queryexpression)" -Path C:\Temp\PS_Audit.csv 

	        }
            catch
            {
                #If the file is POSH_Deploy.csv is already opened then don't proceed
                $LogCollection.Add("Add-MachineToSCCMCollection: $_.Exception")
            }
	        #endregion log the changes

				    
            #region delete the previous Query Rule

			#before deleting need to make sure the collection is ready ....Precaution
			while ((Get-WmiObject -Query "Select CurrentStatus from SMS_Collection WHERE CollectionID='$collectionid'" @WMIHash).CurrentStatus -ne 1 )
			{
			    start-sleep -Seconds 1
                Write-Verbose -Message "Add-MachineToSCCMCollection: Collection is not in the ready state, So sleeping for 1 second"
		
			} 
			        
			#do a check before modifying the collection ---Is the Collection ready ? Someone could be editing the Collection Rules
			if(( Get-WmiObject -Query "Select CurrentStatus from SMS_Collection WHERE CollectionID='$collectionid'" @WMIHash).CurrentStatus -eq 1 )
			{
			    #means the collection is ready for the update
                
	            Write-Verbose -Message "Add-MachineToSCCMCollection: The Automated_QueryRule's QueryExpression is validated....Saving it now"
	            try 
			    {
                    # deleting the QueryRule.
			        $null = $collection.DeleteMembershipRule($Automated_QueryRule) 
			        Write-Verbose "Add-MachineToSCCMCollection: Invoked Method DeleteMembershipRule on the Collection"
			        $null = $collection.RequestRefresh() 
			        Write-Verbose "Add-MachineToSCCMCollection: Invoked Method RequestRefresh on the Collection"
			    }
			    catch
			    {
			        Write-Error "Couldn't invoke method DeleteMembershipRule on the Collection with ID $collectionid"
			        $LogCollection.Add("Add-MachineToSCCMCollection: $_.exception")
			    }
	            
		
			}
			#endregion

			#region Add a new Query Rule
			#before deleting need to make sure the collection is ready 
			do 
			{
			    start-sleep -Seconds 1
		
			} until ((Get-WmiObject -Query "Select CurrentStatus from SMS_Collection WHERE CollectionID='$collectionid'" @WMIHash).CurrentStatus -eq 1 )
			        
			#do a check before modifying the collection ---Is the Collection ready ?
			        
			if((Get-WmiObject -Query "Select CurrentStatus from SMS_Collection WHERE CollectionID='$collectionid'" @WMIHash).CurrentStatus -eq 1)
			{
			    #means the collection is ready for the update
			    try 
			    {
			        $null = $collection.AddmembershipRule($Automated_QueryRule) 
			        Write-Verbose "Add-MachineToSCCMCollection: Invoked Method AddMembershipRule on the Collection"
			        $null = $collection.RequestRefresh() 
			        Write-Verbose "Add-MachineToSCCMCollection: Invoked Method RequestRefresh on the Collection"
														
			    }
			    catch
			    {
			        Write-Warning "Add-MachineToSCCMCollection: Couldn't invoke method AddMembershipRule on the Collection with ID $collectionid"
			        $LogCollection.Add("Add-MachineToSCCMCollection: $_.exception ")
			    }
            
			}
			#endregion Add a new Query Rule	        
           
	           
		} #End  if ($($Automated_QueryRule.queryExpression).length -lt $tempQueryExpression.length)
			        
		       		
	} #end Process block
				
	End				
	{
		    Write-Verbose "Add-MachineToSCCMCollection: Ending the function"
	}	
	}

#endregion


#region Remove-MachineNamefromSCCMCollection

function Remove-MachineFromSCCMCollection
	{
	[CmdletBinding()]
	[OutputType([PSObject])]
	Param
	(
        # Enter the Computer Name
		[Parameter(Mandatory,
		            helpmessage="Enter the ComputerNames array to remove from the Collection",
		            ValueFromPipeline=$true,
		            ValueFromPipelineByPropertyName=$true
		            )]
		[Alias("CN","computer")]
		[String[]]
		$computername,

        #Specify the Collection ID
		[Parameter()]
		[validatenotnullorempty()]
		#[ValidatePattern('^[A-Za-z]{3}\w{5}$')]
		[string]
		$CollectionId,

		# Specify the CIM Hash generated by Connect-SCCMServer
		[Parameter(Mandatory)]
        [validatenotnullorempty()]
		[hashtable]$WMIHash  
	)
	
	Begin
	{
		Write-Verbose -Message "Remove-MachineFromSCCMCollection: Starting the function"	   
		        
	}
			
	Process
	{
				
		try 
		    {
		       $Collection = Get-WmiObject -Class SMS_Collection -Filter "CollectionID='$collectionid'" @WMIHash
		    
		        Write-Verbose "Remove-MachineFromSCCMCollection: Queried the Collection $($collection.name) with CollectionId : $collectionid  successfully"
				$Collection.Get() #Invoke Get Method to get the Lazy Properties back

			    #create an empty array to hold QueryRules
			    $QueryRules = @()
                $QueryRules = $Collection.CollectionRules | Where-Object -FilterScript {$_.__Class -eq "SMS_CollectionRuleQuery"} #Take only QueryMembershipRule
		        
		        Write-Verbose "Remove-MachineFromSCCMCollection: Queried the QueryRules on the Collection successfully"
            }
		          
		catch 
		    {
		        Write-Warning -Message "Remove-MachineFromSCCMCollection: Something went wrong while querying the Collection and the Collection Rules on it"
                $LogCollection.Add("Remove-MachineFromSCCMCollection: $_.exception ")
		    }
					
		Foreach ($query in $QueryRules)
		{		       					
			$queryexpression = $query.QueryExpression
			#save the Original Length of the QuerExpression to later check if there were any changes made
			$OriginalLength = $queryexpression.length
            
			#region make changes to the each query and save them
		
			Foreach ($computer in $computername)
			{	
				$templength = $queryexpression.length
			            
				#replaces machine name in the body and end of the query
				$queryexpression = $queryexpression -replace ",`"$computer`"",''
						
				<#this takes care if the machine name is at starting of the queryexpression ...see the below query example
					
				select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,
				SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System 
				where SMS_R_System.NetbiosName in ("Machinename","OFPMXL93219T0",".....)
						
				#>
				$queryexpression = $queryexpression -replace "`"$computer`",",''
			}	
						
							
			
		
			#after the text manipulation, save the new query
			$query.queryexpression = $queryexpression
			#endregion 
		
           
            #region log the changes
			if ($query.queryexpression.length -ne $OriginalLength)
			{
			    
	            Write-Verbose -Message "Remove-MachineFromSCCMCollection: Taking the backup in C:\Temp\POSH_Deploy.csv"
	            #Now the QueryExpression has been modified...So before proceeeding take a backup
	            try
                {
                    [pscustomobject]@{
	                                "Collection"=$collection.Name 
									"CollectionID"=$CollectionId
	                                "Action" = 'Remove'
	                                "MachineNames"=[string]$computername
	                                "QueryName"=$query.RuleName 
	                                "QueryId"=$query.QueryId
                                    "QueryExpression"=$query.QueryExpression
	                            } | Export-Csv -NoTypeInformation -Path $PSdeployfile -Append
	
	            }
                catch
                {
                    $LogCollection.Add("Remove-MachineFromSCCMCollection: $_.exception")
                }
	            #endregion log the changes
	            
	
	            Write-Verbose "Remove-MachineFromSCCMCollection: The QueryExpression seems to be modified...saving the modified one now"
				
			    #region delete the previous Query Rule
			    #before deleting need to make sure the collection is ready 
			    do 
			    {
			        start-sleep -Seconds 1
		
			    } until ((Get-WmiObject -Query "Select CurrentStatus from SMS_Collection WHERE CollectionID='$collectionid'" @WMIHash).CurrentStatus -eq 1 )
			            
			    #do a check before modifying the collection ---Is the Collection ready ?
			    if((Get-WmiObject -Query "Select CurrentStatus from SMS_Collection WHERE CollectionID='$collectionid'" @WMIHash).CurrentStatus -eq 1)
			    {
			        if ( (Invoke-WmiMethod -Class SMS_CollectionRuleQuery -Name ValidateQuery -ArgumentList $query.QueryExpression @WMIHash).returnvalue -eq $true)
					{
	                    Write-Verbose -Message "Remove-MachineFromSCCMCollection: Query Validation Succeeded"
	                    try 
			            {
			                $Null = $collection.DeleteMembershipRule($query)
			                Write-Verbose "Remove-MachineFromSCCMCollection: Invoked Method DeleteMembershipRule on the Collection"
			                $Null = $collection.RequestRefresh()
			                Write-Verbose "Remove-MachineFromSCCMCollection: Invoked Method RequestRefresh on the Collection"
			            }
			            catch
			            {
			                Write-Error "Remove-MachineFromSCCMCollection: Couldn't invoke method DeleteMembershipRule on the Collection with ID $collectionid"
			                $LogCollection.Add("Remove-MachineFromSCCMCollection: Couldn't delete the QueryRule in the Collection")
			            }
	                }
	                else
	                {
	                    Write-Verbose -Message "Remove-MachineFromSCCMCollection: Query you created is incorrect."
	                    $LogCollection.Add("QueryExpression couldn't be validated. Invalid WQL Syntax")
	
		            }
			    }
			    #endregion
		
			    #region Add a new Query Rule
			    #before deleting need to make sure the collection is ready 
			    do 
			    {
			        start-sleep -Seconds 1
		
			    } until ((Get-WmiObject -Query "Select CurrentStatus from SMS_Collection WHERE CollectionID='$collectionid'" @WMIHash).CurrentStatus -eq 1 )
			            
			    #do a check before modifying the collection ---Is the Collection ready ?
			            
			    if((Get-WmiObject -Query "Select CurrentStatus from SMS_Collection WHERE CollectionID='$collectionid'" @WMIHash).CurrentStatus -eq 1)
			    {
			        #means the collection is ready for the update
			        try 
			        {
			            $Null = $collection.AddmembershipRule($query)
			            Write-Verbose "Invoked Method AddMembershipRule on the Collection"
			            $Null = $collection.RequestRefresh()
			            Write-Verbose "Invoked Method RequestRefresh on the Collection"
								
						
			        }
			        catch
			        {
			            Write-Warning "Remove-MachineFromSCCMCollection: Couldn't invoke method AddMembershipRule on the Collection with ID $collectionid"
			            $LogCollection.Add("Remove-MachineFromSCCMCollection: $_.exception")
			        }
			    }
			             
			    #endregion
		
			} #end if ($query.queryexpression.length -ne $OriginalLength)
			        
		}#end foreach ($query in $queryrules)
		
		} #end Process block
				
	End
	{
		    Write-Verbose "Remove-MachineFromSCCMCollection: Ending the function"
	}
	}


#endregion Remove-MachineNamefromSCCMCollection


#region Invoke-CollectionQueryRuleIntegrityCheck
function Invoke-CollectionQueryRuleIntegrityCheck
{
	[CmdletBinding()]
	[OutputType([PSObject])]
	Param
	(
        # Specify the Path for PSDeploy.csv [Default - Looks in User's MyDocuments ]
		[Parameter(ValueFromPipeline=$true,
		            ValueFromPipelineByPropertyName=$true
		            )]
		[String]
		$PSDeployCSVPath="$([System.Environment]::GetFolderPath('MyDocuments'))\PS_Deploy.csv",

        # Specify the no of Collections to perform Integrity Check on [last modified - Default is 3]
		[Parameter()]
        [validatenotnullorempty()]
		[int]$Count=3        
		
		
	)
	#Set-StrictMode -Version 2
	BEGIN
	{
		Write-Verbose -Message "[Invoke-CollectionQueryRuleIntegrityCheck] - Starting the Function "
        if (Test-Path -Path $PSDeployCSVPath -PathType Leaf)
        {
            Write-Verbose -Message "[Invoke-CollectionQueryRuleIntegrityCheck] - $PSDeployCSVPath File Found"
        }
        else
        {
            $LogCollection.add("[Invoke-CollectionQueryRuleIntegrityCheck] - $PSDeployCSVPath File not found.")
            
        }

    }
    PROCESS
    {
        $Groups = Import-Csv -Path $PSDeployCSVPath | Select-Object -Last $Count | Group-Object -Property CollectionId
        $Collections = @()
        $Collections += $Groups | where Count -eq 1 | Select-Object -ExpandProperty Group
        $Collections += $Groups | where Count -gt 1 | ForEach-Object -Process { $_.Group | Select-Object -Last 1}
         Write-Verbose -Message "[Invoke-CollectionQueryRuleIntegrityCheck] - Integrity Check will be performed on last $Count Collections in the $PSDeployCSVPath"
         
         foreach ($collection in $Collections)
         {
            $CMCollection = Get-WmiObject -Class SMS_Collection -Filter "CollectionID='$($collection.collectionid)'" @Script:WMIHash
            $CMCollection.Get() #Invoke Get Method to get the Lazy Properties back

		    #get the Collection Rules in an array 
		    $QueryRules = @($CMCollection.CollectionRules)
            $AutomatedQueryRule = @($QueryRules | Where-Object -FilterScript  {($_.RuleName -eq $($collection.QueryName))})

            If ($AutomatedQueryRule.Count -gt 1)
            {
                #There might be multiple Automated QueryRules in the Collection. So at that point delete all of them and start over by creatiing just one from the PS_Deploy.csv
                Write-Warning -Message "[Invoke-CollectionQueryRuleIntegrityCheck] - Found $($AutomatedQueryRule.Count) Automated QueryRules. Deleting Everything and creating new one from PS_deploy.csv "
                $AutomatedQueryRule | ForEach-Object -Process { $null = $CMCollection.deleteMembershipRule($_)  }
                Write-Verbose -Message "[Invoke-CollectionQueryRuleIntegrityCheck] - Deleted multiple instances of the $($collection.QueryName)"
                $AutomatedQueryRule = $null #set it to Null so that below code takes care of creating one
            }
           
            

            if ($AutomatedQueryRule)
            {
                
                #Found a matching rule on the collection...compare the Query Expression
                Write-Verbose -Message "[Invoke-CollectionQueryRuleIntegrityCheck] - Found the matching Query Rule in the Collection"
                if ($AutomatedQueryRule.QueryExpression -eq $collection.QueryExpression)
                {
                    #QueryExpression matches..Integrity Checked
                    Write-Verbose -Message "[Invoke-CollectionQueryRuleIntegrityCheck] - Integrity Check succeeded"
                }
                else
                {
                    #QueryExpression didn't match..Delete the old one and create new one
                    Write-Warning -Message "[Invoke-CollectionQueryRuleIntegrityCheck] - Integrity Check Failed . QueryExpressions not matching. Creating new QueryRule"
                    TRY
                    {
                        if ((Invoke-WmiMethod -Class SMS_CollectionRuleQuery -Name ValidateQuery -ArgumentList $Collection.QueryExpression @Script:WMIHash).returnvalue )
                        {
                            $AutomatedQueryRule | ForEach-Object -Process {$null = $CMCollection.DeleteMembershipRule($_) } #We wrapped it as an array
                            Write-Verbose -Message "[Invoke-CollectionQueryRuleIntegrityCheck] - Deleted instance of the $($collection.QueryName) with mismatching QueryExpression"
                            $QueryRuleClass = Get-WmiObject -Class SMS_CollectionRuleQuery -List @Script:WMIHash #Returns back the class Object
                            $TempQueryRule = $QueryRuleClass.createinstance()
                            $TempQueryRule.QueryExpression = $Collection.QueryExpression
                            $TempQueryRule.RuleName = 'Automated_QueryRule'
                
                            $null = $CMCollection.AddMembershipRule($TempQueryRule) 
                            Write-Verbose -Message "[Invoke-CollectionQueryRuleIntegrityCheck] - Added new Automated_QueryRule derived from PS_Deploy.csv"
                        }
                        else
                        {
                            Write-Error -Message "[Invoke-CollectionQueryRuleIntegrityCheck] - the AutomatedQueryRule was found but the QueryExpression is not valid. Resolve Manually."
                        }
                    }
                    CATCH
                    {
                        Write-Warning -Message "[Invoke-CollectionQueryRuleIntegrityCheck] - Something went wrong while deleting and then adding the QueryRule for $($Collection.Collection)"
                        $LogCollection.Add("[Invoke-CollectionQueryRuleIntegrityCheck]  $_.exception")
                    }

                }
                
            }
            else
            {
                
                #Matching Query Rule not found create one
                Write-Warning -Message "[Invoke-CollectionQueryRuleIntegrityCheck] - Didn't Found the matching Query Rule in the Collection"
                
                TRY
                {
                    #Validate the QueryExpression before saving it
                    if ((Invoke-WmiMethod -Class SMS_CollectionRuleQuery -Name ValidateQuery -ArgumentList $Collection.QueryExpression @Script:WMIHash).returnvalue )
                    {
                        $QueryRuleClass = Get-WmiObject -Class SMS_CollectionRuleQuery -List @Script:WMIHash #Returns back the class Object
                        $TempQueryRule = $QueryRuleClass.createinstance()
                        $TempQueryRule.QueryExpression = $Collection.QueryExpression
                        $TempQueryRule.RuleName = 'Automated_QueryRule'
                
                        $Null = $CMCollection.AddMembershipRule($TempQueryRule) 
                        Write-Verbose -Message "[Invoke-CollectionQueryRuleIntegrityCheck] - Added new Automated_QueryRule derived from PS_Deploy.csv"
                    }
                    else
                    {
                       $LogCollection.Add("[Invoke-CollectionQueryRuleIntegrityCheck] - QueryExpression in POSHDeploy.csv is not valid for $($Collection.collection)")
                    }

                    Write-Verbose -Message "[Invoke-CollectionQueryRuleIntegrityCheck] - Added the Rule. Logging this to POSH_Deploy.csv now"
                }
                CATCH
                {
                    Write-Warning -Message "[Invoke-CollectionQueryRuleIntegrityCheck] - Something went wrong while adding the QueryRule for $($Collection.Collection)"
                    $LogCollection.Add("[Invoke-CollectionQueryRuleIntegrityCheck]: $_.exception")
                }
            }
        
    }
    
}
END
    {

    }
}
#endregion 
#endregion Function definitions



#region Events

#region Button Action

#On click, change window background color
$buttonaction.Add_Click({
        #$Window.IsEnabled = $false 
		$buttonAction.IsEnabled = $false 
        $Window.cursor = [System.Windows.Input.Cursors]::Wait
		#Check to ensure that the applications are selected
		
        
		if ( ! $script:apps  )
		{
			#$Null = [System.Windows.Forms.MessageBox]::Show("Choose an application to perform action on" , "Warning") 
            $LogCollection.Add("Please select a collection to work with")
		}
		else
        {
		  #if (! $buttonBrowse.fil
            $script:ComputerName = New-Object System.Collections.ArrayList
		    $temp = $($textBoxComputer.Text) -split "`n"
		    $temp = @($temp | ForEach-Object -Process { $_.trim()}) #remove spaces from the machine names
		    $script:ComputerName = [System.Collections.ArrayList]$temp
            
            #below code will remove the not resloving machines names from the $Script:ComputerName
             For ($i=0;  $i -lt $script:ComputerName.Count; $i++)
             {
                try
                    {
                                       
                        $Null = [system.net.dns]::Resolve("$($script:ComputerName[$i])") 
						if (Test-Connection -ComputerName $($script:ComputerName[$i]) -Count 2 -Quiet)
						{
							#$OS = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $($script:ComputerName[$i])
							Write-Verbose -Message "[POSH Deploy] $($script:ComputerName[$i]) is online " #; OSInfo : $($OS.Caption) $($OS.OSArchitecture)"
							$LogCollection.Add("$($script:ComputerName[$i])--> Online ")
                            
						}
						else
						{	
                            Write-Verbose  "[POSH Deploy] $($script:ComputerName[$i]) is offline "
							$LogCollection.Add("$($script:ComputerName[$i]) --> Offline")
						}
                    } 
                catch 
                    { 
                        # $null = something will suppress the Red exception that is thrown..remove it and see what happens
                        
                        Write-Verbose "Ignore the error as the incorrect machine name is being removed from the Input"
                        if ($script:ComputerName[$i] )
						{
							$LogCollection.Add("$($script:ComputerName[$i]) --> Wrong Machine (Will be removed from Input)")
						}
						#$textboxInfo.ForeColor = 'Red'
						$null = $script:ComputerName.Remove("$($script:ComputerName[$i])")  
                        $i-- #need to decrease index by vone each time a machine name is removed
                    }
            
            }
            #show the final list of resolvable machine names in the domain. The query shouldn't get a machine name which can't be resolved.
             $LogCollection.Add("Final list of resolvable machine names --> $script:ComputerName")
            
		    #based on the checkbox selected perform the action
		    if ($buttonAction.Content -eq "ADD")
		    {
			    #Call the Function to deploy the apps
			        $LogCollection.Add("Starting the ADD action on the selected collections")          
            	    $script:apps | 
                                ForEach-Object -Process {
                                            $LogCollection.Add("Starting Function Add-MachineToSCCMCollection for Collection $($_.name)")
                                            Add-MachineToSCCMCollection -CollectionId $_.CollectionId -computername $script:ComputerName -WMIHash $Script:WMIHash
                                            $LogCollection.Add("Ended Function Add-MachineToSCCMCollection for Collection $($_.name)")
                                }
                    $LogCollection.Add("ADD action completed on selected collections")
			     #region to check for the SCCM related Services are started and Machine Policy refresh
			
			    Write-Verbose -Message "Add-MachineToSCCMCollection: All the machines added..now doing post advertisement tasks"
				if ($script:ComputerName)
				{
					$script:ComputerName | 
					    ForEach -Process {
					
						    if (Test-Connection -ComputerName $_ -Count 2 -Quiet)
						    {
							    Write-Verbose "$_ --> is online doing Policy Refresh and Service Check"
                                $LogCollection.Add("Trying to do Service Check & Policy refresh on the machine $_")
							    Invoke-SCCMServiceCheck -ComputerName $_ -verbose
							    Invoke-MachinePolicyRefresh -ComputerName $_ -verbose
						    }
						    else
						    {
							    Write-Warning "$_ --> is not online skipping Policy Refresh and Service Check"
						    }
	                    
	                        #now have to do some check on the SCCM Client on the Machinename supplied
				            $MachineStatus = Get-WmiObject -Query "Select Client,Active FROM SMS_R_System WHERE NetbiosName LIKE '$_'" @Script:WMIhash
                            
						
				            #Check the SCCMClient property
				            if (($MachineStatus.client -eq 1) -or ($MachineStatus.active -eq 1))
				            {
					            Write-Verbose "$_ seems to have Client Installed"
					           
				            }
				            else
				            {
					            Write-Warning "$_ doesn't seem to have Client Installed on it...Raise a ticket to install it"
					            $LogCollection.Add("Warning : The SCCM Client on the machine $_ is not installed or active..Raise a ticket to install it")
					            
					            
				            }							            
	                    	
	                }#end foreach -process
				}#end if (script:Computername)
			#endregion
			
            }
			
				
		    elseif ($buttonAction.Content -eq "REMOVE")
		    {
			    #Call the Function to remove the apps
                $LogCollection.Add("Starting REMOVE action on selected applications")
			    $script:apps | ForEach-Object -Process { 
                                                    $LogCollection.Add("Starting Function Remove-MachineFromSCCMCollection for Collection $($_.name)")
                                                    Remove-MachineFromSCCMCollection -computername $script:ComputerName -CollectionId $_.CollectionId -WMIHash $Script:WMIHash
                                                    $LogCollection.Add("Ending Function Remove-MachineFromSCCMCollection for Collection $($_.name)") 
                                                     }
			
			    
			    $buttonAction.IsEnabled = $true
                
            }
		    
            else
		    {
			    $LogCollection.Add("Warning : Choose an Action First")
                #$Null = [System.Windows.Forms.MessageBox]::Show("Choose an action first" , "Warning") 
			
		    }	
	
	} #end else
		     
			$buttonAction.IsEnabled = $true
            $Window.cursor = [System.Windows.Input.Cursors]::Arrow
            Write-Host -ForegroundColor 'Red' "########################----POSH Deploy----######################## "	
            Write-Host -ForegroundColor 'Cyan'  "Welcome to the SCCM deployment tool.!!!"
            Write-Host -ForegroundColor 'Cyan' "Designed and Created by @DexterPOSH"		
	

})

#Make the mouse act like something is happening
$buttonaction.Add_MouseEnter({
    $Window.Cursor = [System.Windows.Input.Cursors]::Hand
})
#Switch back to regular mouse
$buttonaction.Add_MouseLeave({
    $Window.Cursor = [System.Windows.Input.Cursors]::Arrow
})

#endregion Button Action

#region CheckBoxes


$checkBoxAdd.Add_Checked({
    
    $checkBoxRemove.IsEnabled = $false
    $buttonaction.content = "ADD"
    $buttonaction.Background = 'Green'
    $Window.Background = '#FF3DBE5A'
})
$checkBoxAdd.Add_UnChecked({
   
    $checkBoxRemove.IsEnabled = $true
    $buttonaction.content = "Action"
    $buttonaction.Background = 'Yellow'
    $Window.Background = '#FF5397B4'
})

$checkBoxRemove.Add_Checked({
    
    $checkBoxAdd.IsEnabled = $false
    $buttonaction.content = "REMOVE"
    $buttonaction.Background = 'Red'
    $Window.Background = '#FFE59943'

})
$checkBoxRemove.Add_UnChecked({
    $checkBoxAdd.IsEnabled = $True
    $buttonaction.content = "Action"
    $buttonaction.Background = 'Yellow'
    $Window.Background = '#FF5397B4'

})
#endregion Checkboxes

#region Data Grid events

$dataGridApps.Add_MouseDoubleClick({
    if ( [System.String]::IsNullorEmpty($dataGridApps.CurrentItem)) { return}
    $script:apps += ($dataGridApps.CurrentItem); 
    $dataGridSelectedApps.ItemsSource = $script:apps
    
})


$dataGridSelectedApps.Add_MouseDoubleClick({
    $script:apps = @($script:apps | Where-Object {$_ -ne $($dataGridSelectedApps.CurrentItem) })
    $dataGridSelectedApps.ItemsSource = $script:apps
})

$buttonClear.add_click({
    $script:apps = @()
    $dataGridSelectedApps.ItemsSource = $script:apps
})

$buttonCopy.Add_click({
    $script:apps | Select-Object -ExpandProperty Name | clip.exe
})

#endregion Data Grid Events


#checkBoxPSWindow
$checkBoxPSWindow.Add_Checked({Show-Console})
$checkBoxPSWindow.Add_UnChecked({Hide-Console})

#checkboxCred
$checkBoxCred.Add_Checked({$script:cred= Get-Credential -Message "Enter the Credential for the User with Access to the SMS Namespace "; $LogCollection.Add("Credentials added for all the subsequent operations")})
$checkBoxCred.Add_UnChecked({Remove-Variable -Name cred -Scope Script})

#region test the SMS Connection, create a CIM hash
$buttonTestSMSConnection.add_click({
    $LogCollection.Add("Trying to Connect to the SMS Namespace on the SCCM server")
    $Window.cursor = [System.Windows.Input.Cursors]::Wait
    $hash = @{"SCCMServer"=$($textBoxServer.text)}

    if (($checkBoxCred.IsChecked) -and ($Script:Cred))
    {
        $hash.add('Credential',$Script:Cred) #add the Credentials Object if the checkbox is checked
    }
   
    if ($Script:WMIHash = Connect-SCCMServer @hash  ) #add the support to supply creds 
    {
      
        $LogCollection.Add("Successfully Connected to the SMS Namespace on the server $($textBoxServer.text). WMI Hash created" )
        $buttonaction.IsEnabled = $true
        $textBoxServer.Background = "#FFD1F8A9"
        $buttonSyncApps.IsEnabled = $true
        $buttonIntegrity.IsEnabled = $true
    }
    else
    {
       
        $LogCollection.Add("Can't connect to the SMS Namespace on the server $($textBoxServer.text).`nVerify the Server has SMS NameSpace Provider installed or supply alternate credentials")
                
    }
    $Window.cursor = [System.Windows.Input.Cursors]::Arrow
})

#endregion 

#Sync the App list
$buttonSyncApps.add_click({
   
        $LogCollection.Add("Wait for few seconds..performing a sync & downloading the Collection List ")
        $query = 'SELECT Name , Comment, CollectionID FROM SMS_Collection WHERE Name NOT LIKE "All%" order by Name' #Filter out the Collections like "All Systems", "All*"

        $AllCollections = Get-WmiObject -Query $query @WMIhash

        $AllCollections | Select-Object -Property Name,Comment,CollectionID | Export-Csv -Notype "$([System.Environment]::GetFolderPath('MyDocuments'))\Collection.csv" –Force
            #MASTER.CSV has been updated ...time to reload the GridView 
        $collectionView.Clear()
        Import-Csv -Path "$([System.Environment]::GetFolderPath('MyDocuments'))\Collection.csv"| ForEach-Object -Process {$collectionView.Add($_)}
        $LogCollection.Add("New Collections have been added..Note the collections matching pattern ALL* are dropped")
        $view.Refresh()
    
})

#Do a Check for the Integrity
$buttonIntegrity.add_click({ 
    $LogCollection.Add("last 3 Collections in the PS_Deploy.csv will be checked/ Fixed for integrity")
    $Window.cursor = [System.Windows.Input.Cursors]::Wait
    Invoke-CollectionQueryRuleIntegrityCheck 
    $LogCollection.Add("Integrity Check/ Fix done")
     $Window.cursor = [System.Windows.Input.Cursors]::Arrow

})

#Button Browse event
$buttonBrowse.add_click({
        $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
        InitialDirectory = [Environment]::GetFolderPath('Desktop')
        Filter = 'Text files (*.txt)|*.txt'
    }
    [void]$FileBrowser.ShowDialog()
    $computers = Get-Content -path ($FileBrowser.FileName) 

    foreach ($computer in $computers)
    {
        $textBoxComputer.AppendText("$computer`n")
    }
})

#Make the grid searchable
$textBoxSearch.add_TextChanged({$filter = $textBoxSearch.text; $view.Refresh()})


#Select Apps on the double click of the item in Data Grid
$script:apps = @() #New-Object System.Collections.Generic.List[System.Management.Automation.PSCustomObject]





$DexLabel.add_MouseDoubleClick({
    #TODO: Place custom script here
        $DexLabel.ToolTip = 'Double Click to visit my Blog'
        $url =  "http://dexterposh.blogspot.com"
        $ie = New-Object -ComObject InternetExplorer.Application
        $ie.navigate($url) 
        $ie.visible = $true
})

$DexLabel.Add_MouseEnter({
    $Window.Cursor = [System.Windows.Input.Cursors]::Hand
    $DexLabel.Foreground ='Cyan'
    $DexLabel.ToolTip = 'Double Click to visit my blog'

})
#Switch back to regular mouse
$DexLabel.Add_MouseLeave({
    $Window.Cursor = [System.Windows.Input.Cursors]::Arrow
     $DexLabel.Foreground ='Red'
})

#endregion Events

#Start
$Null = $Window.ShowDialog() 
