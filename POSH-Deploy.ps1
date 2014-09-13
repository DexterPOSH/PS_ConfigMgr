#http://stackoverflow.com/questions/21935398/powershell-observablecollection-predicate-filter
#requires -version 3.0
#requires -runasadministrator
Set-StrictMode -Version latest
$VerbosePreference = 'continue' #setting this as all the verbose messages are displayed on the background console window (hidden by default)

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
        <Button Content="Sync Collections List" Height="25" HorizontalAlignment="Left" Margin="34,676,0,0" Name="buttonSyncApps" VerticalAlignment="Top" Width="118" ToolTip="Will fetch all the collections on the SMS Server " />
        <Button Content="Clear" Height="30" HorizontalAlignment="Left" Margin="1029,319,0,0" Name="buttonClear" VerticalAlignment="Top" Width="75" ToolTip="Clears the selected app list" />
        <CheckBox Content=" Alternate Creds" Height="27" HorizontalAlignment="Left" Margin="1003,34,0,0" Name="checkBoxCred" VerticalAlignment="Top" Width="101" />
        <Button Content="Browse" Height="23" HorizontalAlignment="Left" Margin="630,124,0,0" Name="buttonBrowse" VerticalAlignment="Top" Width="72" ToolTip="Browse for file with Machine Names" />
        <Label Content="Created by DexterPOSH" Height="31" HorizontalAlignment="Left" Margin="1089,676,0,0" Name="labelName" VerticalAlignment="Top" Width="164" FontSize="14" Foreground="Red"></Label>
        <Image Height="36" HorizontalAlignment="Left" Margin="1044,673,0,0" Name="image1" Stretch="Fill" VerticalAlignment="Top" Width="37" />
        <TextBox Height="199" HorizontalAlignment="Left" Margin="803,112,0,0" Name="textBoxlog" VerticalAlignment="Top" Width="433" Background="#FFD1F8A9" />
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



$PSdeployfile = "$([System.Environment]::GetFolderPath('Desktop'))\PS_Deploy.csv"
$PSAuditFile = "$([System.Environment]::GetFolderPath('Desktop'))\PS_Audit.csv"

#As part of the recovery plan the Tool will log all the deployments to a file PS_Deploy.csv (in User Desktop) until the size of it grows older than 10MB
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
	Write-Verbose -Message "[POSH Deploy] PS_Deploy.csv not found....creating one in $([System.Environment]::GetFolderPath('Desktop'))"
    New-Item -Path $PSDeployFile -ItemType file -Verbose
}


#All the Queries before deleting will be saved in User Desktop in a file named PS_Audit.csv 
if (Test-Path -Path $PSAuditFile -Type leaf )
{
	#put the date in it
		
	if ( $((Get-Item -Path 	$PSAuditFile).length/1MB) -ge 10MB)
	{
		"Size exceeded 10 MB on $(get-date)..so renaming it to QueryBackup.bak and creating new PS_Audit.csv"	
        Rename-Item -Path $PSAuditFile -NewName "QueryBackup.bak" -Force -Verbose
        New-Item -Path $PSAuditFile -ItemType file -Verbose
	}
	else
	{
		Write-Verbose -Message "[POSH Deploy] Size of PS_Audit.csv is below 10MB...continuing"
	}
}
else
{
	Write-Verbose -Message "[POSH Deploy] PS_Audit.csv not found....creating one in $([System.Environment]::GetFolderPath('Desktop'))"
    New-Item -Path $PSAuditFile -ItemType file -Verbose
	
}

#endregion


#region Connect to Control
$buttonaction = $Window.FindName("buttonaction")
$textBoxComputer = $Window.FindName('textBoxComputer')
$checkBoxAdd = $Window.FindName('checkBoxAdd')
$checkBoxRemove = $Window.FindName('checkBoxRemove')

$dataGridApps = $Window.FindName('dataGridApps')
$dataGridSelectedApps = $Window.FindName('dataGridSelectedApps')
$textBoxServer = $Window.FindName('textBoxServer')
$checkBoxPSWindow = $Window.FindName('checkBoxPSWindow')
$buttonTestSMSConnection = $Window.FindName('buttonTestSMSConnection')
$textBoxSearch = $Window.FindName('textBoxSearch')
$buttonCopy = $Window.FindName('buttonCopy')
$buttonSyncApps = $Window.FindName('buttonSyncApps')
$checkBoxCred = $Window.FindName('checkBoxCred')
$DexLabel = $Window.FindName('labelName')
$textBoxlog = $Window.findName('textBoxlog')
#endregion Connect to Control


#region Customize the GUI
$dataGridApps.background="LightGray" 
$dataGridApps.RowBackground="LightYellow" 
$dataGridApps.AlternatingRowBackground="LightBlue"
$buttonaction.Background = 'Yellow'

$dataGridSelectedApps.background="LightGray" 
$dataGridSelectedApps.RowBackground="LightGreen" 
$dataGridSelectedApps.AlternatingRowBackground="LightYellow"

$textBoxServer.text = $env:COMPUTERNAME

$collectionView = New-Object System.Collections.ObjectModel.ObservableCollection[object] 
Import-Csv -Path C:\temp\MASTER.CSV| ForEach-Object -Process {$collectionView.Add($_)}

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
                    $ErrorActionPreference = 'stop'               
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
                 $ErrorActionPreference = 'Continue' #Setting it back, Just in case someone is running this from ISE
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

            Get-CimSession | Remove-CimSession #cleaning up in case anyone clicks more than once on this
            Write-Verbose -Message "[Connect-SCCMServer] Trying to open a CIM Session with SCCM server"
            #region open a CIM session
            $CIMSessionParams = @{ComputerName = $SCCMServer;ErrorAction = 'Stop'}          
            if ($credential)
            {
                 $CIMSessionParams.Add('Credential',$credential)
            }
                
            try
            {
                If ((Test-WSMan -ComputerName $SCCMServer -ErrorAction SilentlyContinue).ProductVersion -match 'Stack: 3.0')
                {
                    Write-Verbose -Message "[Connect-SCCMServer] WSMAN is responsive"
                    $CimSession = New-CimSession @CIMSessionParams
                    $CimProtocol = $CimSession.protocol
                    Write-Verbose -Message "[Connect-SCCMServer] [$CimProtocol] CIM SESSION - Opened"
                } 
 
                else 
                {
                    Write-Verbose -Message "[Connect-SCCMServer] Attempting to connect with protocol: DCOM"
                    $CIMSessionParams.SessionOption = New-CimSessionOption -Protocol Dcom
                    $CimSession = New-CimSession @CIMSessionParams
                    $CimProtocol = $CimSession.protocol
 
                    Write-Verbose -Message "[Connect-SCCMServer] [$CimProtocol] CIM SESSION - Opened"
                }
       
 
            #endregion open a CIM session
 
            #region create the Hash to be used later for CIM queries   
                $sccmProvider = Get-CimInstance -query "select * from SMS_ProviderLocation where ProviderForLocalSite = true" -Namespace "root\sms" -CimSession $CimSession -ErrorAction Stop
                # Split up the namespace path
                $Splits = $sccmProvider.NamespacePath -split "\\", 4
                Write-Verbose "[Connect-SCCMServer] Provider is located on $($sccmProvider.Machine) in namespace $($splits[3])"
 
                # Create a new hash to be passed on later
                $CIMHash= @{"CimSession"=$CimSession;"NameSpace"=$Splits[3];"ErrorAction"="Stop";"Verbose"=$false}
               
                Write-Output -InputObject $CIMHash
                
                #endregion create the Hash to be used later for CIM queries
            }
            catch
            {
                Write-Warning "[Connect-SCCMServer] Something went wrong"
                throw $_.Exception
            }

            Write-Verbose -Message "[Connect-SCCMServer] Ending the Function"
    }
}


Function Out-TextBoxLog
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory,ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [string[]]$text,

        [Parameter()]
        [string]$Colour="#FFD1F8A9"
        )
    $textBoxlog.Text = "$text" + [System.Environment]::NewLine
    $textBoxlog.Background = $Colour
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
		[ValidatePattern('^[A-Za-z]{3}\w{5}$')]
		[string]
		$CollectionId,

		# Specify the CIM Hash generated by Connect-SCCMServer
		[Parameter(Mandatory)]
        [validatenotnullorempty()]
		[hashtable]$CIMHash        
		
		
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
            $Collection.AddMembershipRule($TempQueryRule) | Out-Null
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
                $Collection = Get-WmiObject -Class SMS_Collection -Filter "CollectionID='$collectionid'" @WMIHash
		    
		        Write-Verbose "Add-MachineToSCCMCollection: Queried the Collection $($collection.name) with CollectionId : $collectionid  successfully"
				$Collection.Get() #Invoke Get Method to get the Lazy Properties back

			    #get the Collection Rules in an array 
			    $QueryRules = @($Collection.CollectionRules)
		        
		        Write-Verbose "Add-MachineToSCCMCollection: Queried the QueryRules on the Collection successfully"
		    }
		          
		catch 
		    {
		        Write-Warning -Message "Add-MachineToSCCMCollection: Something went wrong while querying the Collection and the Collection Rules on it"
                Throw $_.exception 
		    }
				
		
		#The Script won't make any changes in the existing query but will only look for a QueryRule by name "Automated_QueryRule		
		If ($QueryRules) 
        {	
            # If already Query Rules exist..then look among them for "automated_queryrule"	
		    If (!($Automated_QueryRule = $QueryRules | where {$_.RuleName -eq "Automated_QueryRule"} | Sort-Object -Property {$_.QueryExpression.Length} | Select-Object -First 1))
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
            Write-Error -Message "Add-MachineToSCCMCollection: The QueryExpression created is not valid" -ErrorAction Stop
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
                throw $_.Exception
            }
	        #endregion log the changes

				    
            #region delete the previous Query Rule

			#before deleting need to make sure the collection is ready ....Precaution
			while ((Get-CimInstance -Query "Select CurrentStatus from SMS_Collection WHERE CollectionID='$collectionid'" @CIMHash).CurrentStatus -ne 1 )
			{
			    start-sleep -Seconds 1
                Write-Verbose -Message "Add-MachineToSCCMCollection: Collection is not in the ready state, So sleeping for 1 second"
		
			} 
			        
			#do a check before modifying the collection ---Is the Collection ready ? Someone could be editing the Collection Rules
			if(( Get-CimInstance -Query "Select CurrentStatus from SMS_Collection WHERE CollectionID='$collectionid'" @CIMHash).CurrentStatus -eq 1 )
			{
			    #means the collection is ready for the update
                
	            Write-Host -ForegroundColor Green "Add-MachineToSCCMCollection: The Automated_QueryRule's QueryExpression is validated....Saving it now"
	            try 
			    {
                    #Before deleting the QueryExpression..take the backup i
			        $collection.DeleteMembershipRule($Automated_QueryRule) | Out-Null
			        Write-Verbose "Add-MachineToSCCMCollection: Invoked Method DeleteMembershipRule on the Collection"
			        $collection.RequestRefresh() | Out-Null
			        Write-Verbose "Add-MachineToSCCMCollection: Invoked Method RequestRefresh on the Collection"
			    }
			    catch
			    {
			        Write-Error "Couldn't invoke method DeleteMembershipRule on the Collection with ID $collectionid"
			        throw $_.exception
			    }
	            
		
			}
			#endregion

			#region Add a new Query Rule
			#before deleting need to make sure the collection is ready 
			do 
			{
			    start-sleep -Seconds 1
		
			} until ((Get-CimInstance -Query "Select CurrentStatus from SMS_Collection WHERE CollectionID='$collectionid'" @CIMHash).CurrentStatus -eq 1 )
			        
			#do a check before modifying the collection ---Is the Collection ready ?
			        
			if((Get-CimInstance -Query "Select CurrentStatus from SMS_Collection WHERE CollectionID='$collectionid'" @CIMHash).CurrentStatus -eq 1)
			{
			    #means the collection is ready for the update
			    try 
			    {
			        $collection.AddmembershipRule($Automated_QueryRule) | Out-Null
			        Write-Verbose "Add-MachineToSCCMCollection: Invoked Method AddMembershipRule on the Collection"
			        $collection.RequestRefresh() | Out-Null
			        Write-Verbose "Add-MachineToSCCMCollection: Invoked Method RequestRefresh on the Collection"
														
			    }
			    catch
			    {
			        Write-Error "Add-MachineToSCCMCollection: Couldn't invoke method AddMembershipRule on the Collection with ID $collectionid"
			        throw $_.exception 
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
		[ValidatePattern('^[A-Za-z]{3}\w{5}$')]
		[string]
		$CollectionId,

		# Specify the CIM Hash generated by Connect-SCCMServer
		[Parameter(Mandatory)]
        [validatenotnullorempty()]
		[hashtable]$CIMHash  
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
                $QueryRules = $Collection.CollectionRules
		        
		        Write-Verbose "Remove-MachineFromSCCMCollection: Queried the QueryRules on the Collection successfully"
            }
		          
		catch 
		    {
		        Write-Warning -Message "Remove-MachineFromSCCMCollection: Something went wrong while querying the Collection and the Collection Rules on it"
                Throw $_.exception		    
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
                    throw $_.exception 
                }
	            #endregion log the changes
	            
	
	            Write-Verbose "Remove-MachineFromSCCMCollection: The QueryExpression seems to be modified...saving the modified one now"
				
			    #region delete the previous Query Rule
			    #before deleting need to make sure the collection is ready 
			    do 
			    {
			        start-sleep -Seconds 1
		
			    } until ((Get-CimInstance -Query "Select CurrentStatus from SMS_Collection WHERE CollectionID='$collectionid'" @CIMHash).CurrentStatus -eq 1 )
			            
			    #do a check before modifying the collection ---Is the Collection ready ?
			    if((Get-CimInstance -Query "Select CurrentStatus from SMS_Collection WHERE CollectionID='$collectionid'" @CIMHash).CurrentStatus -eq 1)
			    {
			        if ( (Invoke-WmiMethod -Class SMS_CollectionRuleQuery -Name ValidateQuery -ArgumentList $query.QueryExpression @WMIHash).returnvalue -eq $true)
					{
	                    Write-Verbose -Message "Remove-MachineFromSCCMCollection: Query Validation Succeeded"
	                    try 
			            {
			                $collection.DeleteMembershipRule($query) | Out-Null
			                Write-Verbose "Remove-MachineFromSCCMCollection: Invoked Method DeleteMembershipRule on the Collection"
			                $collection.RequestRefresh() | Out-Null
			                Write-Verbose "Remove-MachineFromSCCMCollection: Invoked Method RequestRefresh on the Collection"
			            }
			            catch
			            {
			                Write-Error "Remove-MachineFromSCCMCollection: Couldn't invoke method DeleteMembershipRule on the Collection with ID $collectionid"
			                throw "Remove-MachineFromSCCMCollection: Couldn't delete the QueryRule in the Collection"
			            }
	                }
	                else
	                {
	                    Write-Host -ForegroundColor Red "Remove-MachineFromSCCMCollection: Query you created is incorrect."
	                    throw "Remove-MachineFromSCCMCollection: QueryExpression couldn't be validated. Invalid WQL Syntax"
	
		            }
			    }
			    #endregion
		
			    #region Add a new Query Rule
			    #before deleting need to make sure the collection is ready 
			    do 
			    {
			        start-sleep -Seconds 1
		
			    } until ((Get-CimInstance -Query "Select CurrentStatus from SMS_Collection WHERE CollectionID='$collectionid'" @CIMHash).CurrentStatus -eq 1 )
			            
			    #do a check before modifying the collection ---Is the Collection ready ?
			            
			    if((Get-CimInstance -Query "Select CurrentStatus from SMS_Collection WHERE CollectionID='$collectionid'" @CIMHash).CurrentStatus -eq 1)
			    {
			        #means the collection is ready for the update
			        try 
			        {
			            $collection.AddmembershipRule($query) | Out-Null
			            Write-Verbose "Invoked Method AddMembershipRule on the Collection"
			            $collection.RequestRefresh() | Out-Null
			            Write-Verbose "Invoked Method RequestRefresh on the Collection"
								
						
			        }
			        catch
			        {
			            Write-Error "Remove-MachineFromSCCMCollection: Couldn't invoke method AddMembershipRule on the Collection with ID $collectionid"
			            throw "Remove-MachineFromSCCMCollection: Couldn't add the QueryRule in the Collection"
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

#endregion Function definitions



#region Events

#region Button Action

#On click, change window background color
$buttonaction.Add_Click({
        $Window.IsEnabled = $false 
		$buttonAction.IsEnabled = $false 
		
		#Check to ensure that the applications are selected
		
        $textBoxlog.clear() #clear the Log box
		if ( ! $script:apps  )
		{
			[System.Windows.Forms.MessageBox]::Show("Choose an application to perform action on" , "Warning") | Out-Null
		}
		else
        {
		  
            $script:ComputerName = New-Object System.Collections.ArrayList
		    $temp = $($textBoxComputer.Text) -split "`r`n"
		    $temp = @($temp | ForEach-Object -Process { $_.trim(" ")}) #remove spaces from the machine names
		    $script:ComputerName = [System.Collections.ArrayList]$temp

            #below code will remove the not resloving machines names from the $Script:ComputerName
             For ($i=0;  $i -lt $script:ComputerName.Count; $i++)
             {
                try
                    {
                                       
                        [system.net.dns]::Resolve("$($script:ComputerName[$i])") | Out-Null
						if (Test-Connection -ComputerName $($script:ComputerName[$i]) -Count 2 -Quiet)
						{
							#$OS = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $($script:ComputerName[$i])
							Write-Host -ForegroundColor 'Green' "[POSH Deploy] $($script:ComputerName[$i]) is online " #; OSInfo : $($OS.Caption) $($OS.OSArchitecture)"
							Out-TextBoxLog -text "$($script:ComputerName[$i])--> Online `n" 
                            
						}
						else
						{	
                            write-host -ForegroundColor Red "$($script:ComputerName[$i]) is offline "
							Out-TextBoxLog -text "$($script:ComputerName[$i]) --> Offline `n"
						}
                    } 
                catch 
                    { 
                        # $null = something will suppress the Red exception that is thrown..remove it and see what happens
                        #[System.Windows.Forms.MessageBox]::Show("Machine name $($script:ComputerName[$i]) is wrong (not resolving)..skipping it" , "Warning") | Out-Null
                        Write-Verbose "Ignore the error as the incorrect machine name is being removed from the Input"
                        if ($script:ComputerName[$i] )
						{
							$TextBoxLog.Text += "$($script:ComputerName[$i]) --> Wrong Machine " + [System.Environment]::NewLine
						}
						#$textboxInfo.ForeColor = 'Red'
						$null = $script:ComputerName.Remove("$($script:ComputerName[$i])")  
                        $i-- #need to decrease index by vone each time a machine name is removed
                    }
            
            }
            #show the final list of resolvable machine names in the domain. The query shouldn't get a machine name which can't be resolved.
             Out-TextBoxLog -Text "Final list of resolvable machine names -- $script:ComputerName `n" 
            
		    #based on the checkbox selected perform the action
		    if ($buttonAction.Content -eq "ADD")
		    {
			    #Call the Function to deploy the apps
			                    
            	    $script:apps | 
                                ForEach-Object -Process {
                                   
                                            Add-MachineToSCCMCollection -CollectionId $_.CollectionId -computername $script:ComputerName -CIMHash $Script:CIMHash
                                    
                                }
			     #region to check for the SCCM related Services are started and Machine Policy refresh
			
			    Write-Verbose -Message "Add-MachineToSCCMCollection: All the machines added..now doing post advertisement tasks"
				if ($script:ComputerName)
				{
					$script:ComputerName | 
					    ForEach -Process {
					
						    if (Test-Connection -ComputerName $_ -Count 2 -Quiet)
						    {
							    Write-Verbose "$_ --> is online doing Policy Refresh and Service Check"
							    Invoke-SCCMServiceCheck -ComputerName $_ -verbose
							    Invoke-MachinePolicyRefresh -ComputerName $_ -verbose
						    }
						    else
						    {
							    Write-Warning "$_ --> is not online skipping Policy Refresh and Service Check"
						    }
	                    
	                        #now have to do some check on the SCCM Client on the Machinename supplied
				            $MachineStatus = Get-CimInstance -Query "Select Client,Active FROM SMS_R_System WHERE NetbiosName LIKE '$_'" @CIMHash
                            
						
				            #Check the SCCMClient property
				            if (($MachineStatus.client -eq 1) -or ($MachineStatus.active -eq 1))
				            {
					            Write-Verbose "$_ seems to have Client Installed"
					           
				            }
				            else
				            {
					            Write-Warning "$_ doesn't seem to have Client Installed on it...Raise a ticket to install it"
					            
					            [System.Windows.Forms.MessageBox]::Show("The SCCM Client on the machine $_ is not installed or active..Raise a ticket to install it" , "Warning") | Out-Null
					            
				            }							            
	                    	
	                }#end foreach -process
				}#end if (script:Computername)
			#endregion
			
            }
			
				
		    elseif ($buttonAction.Content -eq "REMOVE")
		    {
			    #Call the Function to remove the apps
			    $script:apps | ForEach-Object -Process { Remove-MachineFromSCCMCollection -computername $script:ComputerName -CollectionId $_.CollectionId -CIMHash $Script:CIMhash }
			
			    $Window.IsEnabled = $true
			    $buttonAction.IsEnabled = $true
                
            }
		    
            else
		    {
			    [System.Windows.Forms.MessageBox]::Show("Choose an action first" , "Warning") | Out-Null	
			
		    }	
	
	} #end else
		    $Window.IsEnabled = $true  
			$buttonAction.IsEnabled = $true
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
    $script:apps += ($dataGridApps.CurrentItem); 
    $dataGridSelectedApps.ItemsSource = $script:apps
})


$dataGridSelectedApps.Add_MouseDoubleClick({
    $script:apps = @($script:apps | Where-Object {$_ -ne $($dataGridSelectedApps.CurrentItem) })
    $dataGridSelectedApps.ItemsSource = $script:apps
})


#endregion Data Grid Events


#checkBoxPSWindow
$checkBoxPSWindow.Add_Checked({Show-Console})
$checkBoxPSWindow.Add_UnChecked({Hide-Console})

#checkbox
$checkBoxCred.Add_Checked({$script:cred= Get-Credential -Message "Enter the Credential for the User with Access to the SMS Namespace "})
$checkBoxCred.Add_UnChecked({Remove-Variable -Name cred -Scope Script})

#test the SMS Connection, create a CIM hash
$buttonTestSMSConnection.add_click({
    
    $hash = @{"SCCMServer"=$($textBoxServer.text)}
    if (($checkBoxCred.IsChecked) -and ($Script:Cred))
    {
        $hash.add('Credential',$Script:Cred) #add the Credentials Object if the checkbox is checked
    }
   
    if ($Script:CIMHash = Connect-SCCMServer @hash  ) #add the support to supply creds 
    {
        #create the WMI Hash from the CIMHash 
        $Script:WMIHash = $Script:CIMHash.Clone()
        $Script:WMIHash.Remove('CimSession')
        $Script:WMIHash.Add('ComputerName',$($CIMHash.CimSession.ComputerName))
        if ( $checkBoxCred.IsChecked)
        {
            Write-Verbose -Message "[POSH Deploy] Setting the PSDefaultParameterValues for Get-WMIObject to use alternate Creds"
            Out-TextBoxLog -text "Setting the PSDefaultParameterValues for Get-WMIObject to use alternate Creds"
            $Script:PSDefaultParameterValues = @{"Get-WMIObject:Credential"=$Script:Cred} #setting the Credentials for all the WMI Calls as CIM Session can't be used
        }
        
        Out-TextBoxLog -text "Successfully Connected to the SMS Namespace on the server $($textBoxServer.text). CIM Hash created" 
        $buttonaction.IsEnabled = $true
        $textBoxServer.Background = "#FFD1F8A9"
    }
    else
    {
       
        Out-TextBoxLog -text "Can't connect to the SMS Namespace on the server $($textBoxServer.text). `n Verify the Server has SMS NameSpace Provider installed or supply alternate credentials"  -colour "#FFFF8686"
        $textBoxServer.Background = "#FFFF8686"
        throw "Can't connect to the SCCM server"
    }
})

#Sync the App list
$buttonSyncApps.add_click({
    if ($Script:CIMHash)
    {
        $query = 'SELECT Name , Comment, CollectionID FROM SMS_Collection WHERE Name NOT LIKE "All%" order by Name' #Filter out the Collections like "All Systems", "All*"

        $AllCollections = Get-CimInstance -Query $query @CIMhash

        $AllCollections | Select-Object -Property Name,Comment,CollectionID | Export-Csv -Notype C:\temp\master.csv –Force
            #MASTER.CSV has been updated ...time to reload the GridView 
        $collectionView.Clear()
        Import-Csv -Path C:\temp\MASTER.CSV| ForEach-Object -Process {$collectionView.Add($_)}
        $view.Refresh()
    }
    else
    {
        #pop up a dialog box to connect to SCCM Server
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
$Window.ShowDialog() | Out-Null
