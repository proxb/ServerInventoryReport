Param (
    [parameter()]
    $SQLServer = 'vSQL'
)
#region Helper Functions
function Out-DataTable {
    [CmdletBinding()]
    param([Parameter(Position=0, Mandatory=$true, ValueFromPipeline = $true)] [PSObject[]]$InputObject)

    Begin
    {
    function Get-Type {
        param($type)

        $types = @(
        'System.Boolean',
        'System.Byte[]',
        'System.Byte',
        'System.Char',
        'System.Datetime',
        'System.Decimal',
        'System.Double',
        'System.Guid',
        'System.Int16',
        'System.Int32',
        'System.Int64',
        'System.Single',
        'System.UInt16',
        'System.UInt32',
        'System.UInt64')

        if ( $types -contains $type ) {
            Write-Output "$type"
        }
        else {
            Write-Output 'System.String'
        
        }
    } #Get-Type
        $dt = new-object Data.datatable  
        $First = $true 
    }
    Process
    {
        foreach ($object in $InputObject)
        {
            $DR = $DT.NewRow()  
            foreach($property in $object.PsObject.get_properties())
            {  
                if ($first)
                {  
                    $Col =  new-object Data.DataColumn  
                    $Col.ColumnName = $property.Name.ToString()  
                    if ($property.value)
                    {
                        if ($property.value -isnot [System.DBNull]) {
                            $Col.DataType = [System.Type]::GetType("$(Get-Type $property.TypeNameOfValue)")
                         }
                    }
                    $DT.Columns.Add($Col)
                }  
                if ($property.Gettype().IsArray) {
                    $DR.Item($property.Name) =$property.value | ConvertTo-XML -AS String -NoTypeInformation -Depth 1
                }  
               else {
                    If ($Property.Value) {
                        $DR.Item($Property.Name) = $Property.Value
                    } Else {
                        $DR.Item($Property.Name)=[DBNull]::Value
                    }
                }
            }  
            $DT.Rows.Add($DR)  
            $First = $false
        }
    } 
     
    End
    {
        Write-Output @(,($dt))
    }

}

Function Write-DataTable {
    [CmdletBinding()]
    param(
    [Parameter(Position=0, Mandatory=$true)] 
    [string]$Computername,
    [Parameter(Position=1, Mandatory=$true)] 
    [string]$Database,
    [Parameter(Position=2, Mandatory=$true)] 
    [string]$TableName,
    [Parameter(Position=3, Mandatory=$true)] 
    $Data,
    [Parameter(Position=4)] 
    [string]$Username,
    [Parameter(Position=5)] 
    [string]$Password,
    [Parameter(Position=6)] 
    [Int32]$BatchSize=50000,
    [Parameter(Position=7)] 
    [Int32]$QueryTimeout=0,
    [Parameter(Position=8)] 
    [Int32]$ConnectionTimeout=15
    )
    
    $SQLConnection = new-object System.Data.SqlClient.SQLConnection

    If ($Username) { 
        $ConnectionString = "Server={0};Database={1};User ID={2};Password={3};Trusted_Connection=False;Connect Timeout={4}" -f $Computername,$Database,$Username,$Password,$ConnectionTimeout 
    }
    Else { 
        $ConnectionString = "Server={0};Database={1};Integrated Security=True;Connect Timeout={2}" -f $Computername,$Database,$ConnectionTimeout 
    }

    $SQLConnection.ConnectionString = $ConnectionString

    Try {
        $SQLConnection.Open()
        $bulkCopy = New-Object Data.SqlClient.SqlBulkCopy -ArgumentList $SQLConnection, ([System.Data.SqlClient.SqlBulkCopyOptions]::TableLock),$Null
        $bulkCopy.DestinationTableName = $tableName
        $bulkCopy.BatchSize = $BatchSize
        $bulkCopy.BulkCopyTimeout = $QueryTimeOut
        $bulkCopy.WriteToServer($Data)        
    }
    Catch {
        Write-Error "$($TableName): $($_)"
    }
    Finally {
        $SQLConnection.Close()
    }
}

Function Get-Server {
    [cmdletbinding(DefaultParameterSetName='All')]
    Param (
        [parameter(ParameterSetName='DomainController')]
        [switch]$DomainController,
        [parameter(ParameterSetName='MemberServer')]
        [switch]$MemberServer
    )
    Write-Verbose "Parameter Set: $($PSCmdlet.ParameterSetName)"
    Switch ($PSCmdlet.ParameterSetName) {
        'All' {
            $ldapFilter = "(&(objectCategory=computer)(OperatingSystem=Windows*Server*))"
        }
        'DomainController' {
            $ldapFilter = "(&(objectCategory=computer)(OperatingSystem=Windows*Server*)(userAccountControl:1.2.840.113556.1.4.803:=8192))"
        }
        'MemberServer' {
            $ldapFilter = "(&(objectCategory=computer)(OperatingSystem=Windows*Server*)(!(userAccountControl:1.2.840.113556.1.4.803:=8192)))"
        }
    }
    $searcher = [adsisearcher]""
    $Searcher.Filter = $ldapFilter
    $Searcher.pagesize = 10
    $searcher.sizelimit = 5000
    $searcher.PropertiesToLoad.Add("name") | Out-Null
    $Searcher.sort.propertyname='name'
    $searcher.Sort.Direction = 'Ascending'
    $Searcher.FindAll() | ForEach {
        $_.Properties.name
    }
}

Function Invoke-SQLCmd {    
    [cmdletbinding(
        DefaultParameterSetName = 'NoCred',
        SupportsShouldProcess = $True,
        ConfirmImpact = 'Low'
    )]
    Param (
        [parameter()]
        [string]$Computername = 'vSQL',
        
        [parameter()]
        [string]$Database = 'Master',    
        
        [parameter()]
        [string]$TSQL,

        [parameter()]
        [int]$ConnectionTimeout = 30,

        [parameter()]
        [int]$QueryTimeout = 120,

        [parameter()]
        [System.Collections.ICollection]$SQLParameter,

        [parameter(ParameterSetName='Cred')]
        [Alias('RunAs')]        
        [System.Management.Automation.Credential()]$Credential = [System.Management.Automation.PSCredential]::Empty,

        [parameter()]
        [ValidateSet('Query','NonQuery')]
        [string]$CommandType = 'Query'
    )
    If ($PSBoundParameters.ContainsKey('Debug')) {
        $DebugPreference = 'Continue'
    }
    $PSBoundParameters.GetEnumerator() | ForEach {
        Write-Debug $_
    }
    #region Make Connection
    Write-Verbose "Building connection string"
    $Connection=new-object System.Data.SqlClient.SQLConnection 
    If ($PSBoundParameters.ContainsKey('Verbose')) {
        $Handler = [System.Data.SqlClient.SqlInfoMessageEventHandler] {
            Param($sender, $event) 
            Write-Verbose $event.Message -Verbose
        }
        $Connection.add_InfoMessage($Handler)
        $Connection.FireInfoMessageEventOnUserErrors=$True  
    }
    Switch ($PSCmdlet.ParameterSetName) {
        'Cred' {
            $ConnectionString = "Server={0};Database={1};User ID={2};Password={3};Trusted_Connection=False;Connect Timeout={4}" -f $Computername,
                                                                                        $Database,$Credential.Username,
                                                                                        $Credential.GetNetworkCredential().password,$ConnectionTimeout   
            Remove-Variable Credential
        }
        'NoCred' {
            $ConnectionString = "Server={0};Database={1};Integrated Security=True;Connect Timeout={2}" -f $Computername,$Database,$ConnectionTimeout                 
        }
    }   
    $Connection.ConnectionString=$ConnectionString
    Write-Verbose "Opening connection to $($Computername)"
    $Connection.Open()
    #endregion Make Connection

    #region Initiate Query
    Write-Verbose "Initiating query -> $Tsql"
    $Command=new-object system.Data.SqlClient.SqlCommand($Tsql,$Connection)
    If ($PSBoundParameters.ContainsKey('SQLParameter')) {
        $SqlParameter.GetEnumerator() | ForEach {
            Write-Verbose "Adding SQL Parameter: $($_.Key) with Value: $($_.Value)"
            If ($_.Value -ne $null) { 
                [void]$Command.Parameters.AddWithValue($_.Key, $_.Value) 
            }
            Else { 
                [void]$Command.Parameters.AddWithValue($_.Key, [DBNull]::Value) 
            }
        }
    }
    $Command.CommandTimeout=$QueryTimeout
    If ($PSCmdlet.ShouldProcess("Computername: $($Computername) - Database: $($Database)",'Run TSQL operation')) {
        Switch ($CommandType) {
            'Query' {
                Write-Verbose "Performing Query operation"
                $DataSet=New-Object system.Data.DataSet
                $DataAdapter=New-Object system.Data.SqlClient.SqlDataAdapter($Command)
                [void]$DataAdapter.fill($DataSet)
                $DataSet.Tables
            }
            'NonQuery' {
                Write-Verbose "Performing Non-Query operation"
                [void]$Command.ExecuteNonQuery()
            }
        }
    }
    #endregion Initiate Query    

    #region Close connection
    Write-Verbose "Closing connection"
    $Connection.Close()        
    #endregion Close connection
}

Function Get-LocalUser {
    [Cmdletbinding()] 
    Param( 
        [Parameter()] 
        [String[]]$Computername = $Computername
    )
    Function Convert-UserFlag {
        Param ($UserFlag)
        $List = New-Object System.Collections.ArrayList
        Switch ($UserFlag) {
            ($UserFlag -BOR 0x0001)  {[void]$List.Add('SCRIPT')}
            ($UserFlag -BOR 0x0002)  {[void]$List.Add('ACCOUNTDISABLE')}
            ($UserFlag -BOR 0x0008)  {[void]$List.Add('HOMEDIR_REQUIRED')}
            ($UserFlag -BOR 0x0010)  {[void]$List.Add('LOCKOUT')}
            ($UserFlag -BOR 0x0020)  {[void]$List.Add('PASSWD_NOTREQD')}
            ($UserFlag -BOR 0x0040)  {[void]$List.Add('PASSWD_CANT_CHANGE')}
            ($UserFlag -BOR 0x0080)  {[void]$List.Add('ENCRYPTED_TEXT_PWD_ALLOWED')}
            ($UserFlag -BOR 0x0100)  {[void]$List.Add('TEMP_DUPLICATE_ACCOUNT')}
            ($UserFlag -BOR 0x0200)  {[void]$List.Add('NORMAL_ACCOUNT')}
            ($UserFlag -BOR 0x0800)  {[void]$List.Add('INTERDOMAIN_TRUST_ACCOUNT')}
            ($UserFlag -BOR 0x1000)  {[void]$List.Add('WORKSTATION_TRUST_ACCOUNT')}
            ($UserFlag -BOR 0x2000)  {[void]$List.Add('SERVER_TRUST_ACCOUNT')}
            ($UserFlag -BOR 0x10000)  {[void]$List.Add('DONT_EXPIRE_PASSWORD')}
            ($UserFlag -BOR 0x20000)  {[void]$List.Add('MNS_LOGON_ACCOUNT')}
            ($UserFlag -BOR 0x40000)  {[void]$List.Add('SMARTCARD_REQUIRED')}
            ($UserFlag -BOR 0x80000)  {[void]$List.Add('TRUSTED_FOR_DELEGATION')}
            ($UserFlag -BOR 0x100000)  {[void]$List.Add('NOT_DELEGATED')}
            ($UserFlag -BOR 0x200000)  {[void]$List.Add('USE_DES_KEY_ONLY')}
            ($UserFlag -BOR 0x400000)  {[void]$List.Add('DONT_REQ_PREAUTH')}
            ($UserFlag -BOR 0x800000)  {[void]$List.Add('PASSWORD_EXPIRED')}
            ($UserFlag -BOR 0x1000000)  {[void]$List.Add('TRUSTED_TO_AUTH_FOR_DELEGATION')}
            ($UserFlag -BOR 0x04000000)  {[void]$List.Add('PARTIAL_SECRETS_ACCOUNT')}
        }
        $List -join '; '
    }
    Function ConvertTo-SID {
        Param([byte[]]$BinarySID)
        (New-Object System.Security.Principal.SecurityIdentifier($BinarySID,0)).Value
    }
    $adsi = [ADSI]"WinNT://$Computername"
    $adsi.Children | where {$_.SchemaClassName -eq 'user'} |
    Select @{L='Computername';E={$Computername}}, @{L='Name';E={$_.Name[0]}}, 
    @{L='PasswordAge';E={("{0:N0}" -f ($_.PasswordAge[0]/86400))}}, 
    @{L='LastLogin';E={If ($_.LastLogin[0] -is [datetime]){$_.LastLogin[0]}Else{$Null}}}, 
    @{L='SID';E={(ConvertTo-SID -BinarySID $_.ObjectSID[0])}}, 
    @{L='UserFlags';E={(Convert-UserFlag -UserFlag $_.UserFlags[0])}}
}

Function Get-LocalGroup {
    [Cmdletbinding()] 
    Param( 
        [Parameter()] 
        [String[]]$Computername = $Computername
    )
    Function ConvertTo-SID {
        Param([byte[]]$BinarySID)
        (New-Object System.Security.Principal.SecurityIdentifier($BinarySID,0)).Value
    }
    Function Get-LocalGroupMember {
        Param ($Group)
        $group.Invoke('members') | ForEach {
            $_.GetType().InvokeMember("Name", 'GetProperty', $null, $_, $null)
        }
    }
    $adsi = [ADSI]"WinNT://$Computername"
    $adsi.Children | where {$_.SchemaClassName -eq 'group'} | 
    Select @{L='Computername';E={$Computername}},@{L='Name';E={$_.Name[0]}},
    @{L='Members';E={((Get-LocalGroupMember -Group $_)) -join '; '}},
    @{L='SID';E={(ConvertTo-SID -BinarySID $_.ObjectSID[0])}},
    @{L='GroupType';E={$GroupType[[int]$_.GroupType[0]]}}
}

Function Get-SecurityUpdate {
    [Cmdletbinding()] 
    Param( 
        [Parameter()] 
        [String[]]$Computername = $Computername
    )              
    ForEach ($Computer in $Computername){ 
        $Paths = @("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall","SOFTWARE\\Wow6432node\\Microsoft\\Windows\\CurrentVersion\\Uninstall")         
        ForEach($Path in $Paths) { 
            #Create an instance of the Registry Object and open the HKLM base key 
            Try { 
                $reg=[microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$Computer) 
            } Catch { 
                $_ 
                Continue 
            } 
            Try {
                #Drill down into the Uninstall key using the OpenSubKey Method 
                $regkey=$reg.OpenSubKey($Path)  
                #Retrieve an array of string that contain all the subkey names 
                $subkeys=$regkey.GetSubKeyNames()      
                #Open each Subkey and use GetValue Method to return the required values for each 
                ForEach ($key in $subkeys){   
                    $thisKey=$Path+"\\"+$key   
                    $thisSubKey=$reg.OpenSubKey($thisKey)   
                    # prevent Objects with empty DisplayName 
                    $DisplayName = $thisSubKey.getValue("DisplayName")
                    If ($DisplayName -AND $DisplayName -match '^Update for|rollup|^Security Update|^Service Pack|^HotFix') {
                        $Date = $thisSubKey.GetValue('InstallDate')
                        If ($Date) {
                            Write-Verbose $Date 
                            $Date = $Date -replace '(\d{4})(\d{2})(\d{2})','$1-$2-$3'
                            Write-Verbose $Date 
                            $Date = Get-Date $Date
                        } 
                        If ($DisplayName -match '(?<DisplayName>.*)\((?<KB>KB.*?)\).*') {
                            $DisplayName = $Matches.DisplayName
                            $HotFixID = $Matches.KB
                        }
                        Switch -Wildcard ($DisplayName) {
                            "Service Pack*" {$Description = 'Service Pack'}
                            "Hotfix*" {$Description = 'Hotfix'}
                            "Update*" {$Description = 'Update'}
                            "Security Update*" {$Description = 'Security Update'}
                            Default {$Description = 'Unknown'}
                        }
                        # create New Object with empty Properties 
                        $Object = [pscustomobject] @{
                            Type = $Description
                            HotFixID = $HotFixID
                            InstalledOn = $Date
                            Description = $DisplayName
                        }
                        $Object
                    } 
                }   
                $reg.Close() 
            } Catch {}                  
        }  
    }  
}

Function Get-Software {
    [OutputType('System.Software.Inventory')]
    [Cmdletbinding()] 
    Param( 
        [Parameter(ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)] 
        [String[]]$Computername=$env:COMPUTERNAME
    )         
    Begin {
    }
    Process {     
        ForEach ($Computer in $Computername){ 
            If (Test-Connection -ComputerName $Computer -Count 1 -Quiet) {
                $Paths = @("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall","SOFTWARE\\Wow6432node\\Microsoft\\Windows\\CurrentVersion\\Uninstall")         
                ForEach($Path in $Paths) { 
                    Write-Verbose "Checking Path: $Path"
                    # Create an instance of the Registry Object and open the HKLM base key 
                    Try { 
                        $reg=[microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$Computer,'Registry64') 
                    } Catch { 
                        Write-Error $_ 
                        Continue 
                    } 
                    # Drill down into the Uninstall key using the OpenSubKey Method 
                    Try {
                        $regkey=$reg.OpenSubKey($Path)  
                        # Retrieve an array of string that contain all the subkey names 
                        $subkeys=$regkey.GetSubKeyNames()      
                        # Open each Subkey and use GetValue Method to return the required values for each 
                        ForEach ($key in $subkeys){   
                            Write-Verbose "Key: $Key"
                            $thisKey=$Path+"\\"+$key 
                            Try {  
                                $thisSubKey=$reg.OpenSubKey($thisKey)   
                                # Prevent Objects with empty DisplayName 
                                $DisplayName = $thisSubKey.getValue("DisplayName")
                                If ($DisplayName -AND $DisplayName -notmatch '^Update for|rollup|^Security Update|^Service Pack|^HotFix') {
                                    $Date = $thisSubKey.GetValue('InstallDate')
                                    If ($Date) {
                                        Try {
                                            $Date = [datetime]::ParseExact($Date, 'yyyyMMdd', $Null)
                                        } Catch{
				                            Write-Warning "$($Computer): $_ <$($Date)>"
                                            $Date = $Null
                                        }
                                    } 
                                    # Create New Object with empty Properties 
                                    $Publisher = Try {
                                        $thisSubKey.GetValue('Publisher').Trim()
                                    } 
                                    Catch {
                                        $thisSubKey.GetValue('Publisher')
                                    }
                                    $Version = Try {
                                        #Some weirdness with trailing [char]0 on some strings
                                        $thisSubKey.GetValue('DisplayVersion').TrimEnd(([char[]](32,0)))
                                    } 
                                    Catch {
                                        $thisSubKey.GetValue('DisplayVersion')
                                    }
                                    $UninstallString = Try {
                                        $thisSubKey.GetValue('UninstallString').Trim()
                                    } 
                                    Catch {
                                        $thisSubKey.GetValue('UninstallString')
                                    }
                                    $InstallLocation = Try {
                                        $thisSubKey.GetValue('InstallLocation').Trim()
                                    } 
                                    Catch {
                                        $thisSubKey.GetValue('InstallLocation')
                                    }
                                    $InstallSource = Try {
                                        $thisSubKey.GetValue('InstallSource').Trim()
                                    } 
                                    Catch {
                                        $thisSubKey.GetValue('InstallSource')
                                    }
                                    $HelpLink = Try {
                                        $thisSubKey.GetValue('HelpLink').Trim()
                                    } 
                                    Catch {
                                        $thisSubKey.GetValue('HelpLink')
                                    }
                                    $Object = [pscustomobject]@{
                                        Computername = $Computer
                                        DisplayName = $DisplayName
                                        Version = $Version
                                        InstallDate = $Date
                                        Publisher = $Publisher
                                        UninstallString = $UninstallString
                                        InstallLocation = $InstallLocation
                                        InstallSource = $InstallSource
                                        HelpLink = $thisSubKey.GetValue('HelpLink')
                                        EstimatedSizeMB = [decimal]([math]::Round(($thisSubKey.GetValue('EstimatedSize')*1024)/1MB,2))
                                    }
                                    $Object.pstypenames.insert(0,'System.Software.Inventory')
                                    Write-Output $Object
                                }
                            } Catch {
                                Write-Warning "$Key : $_"
                            }   
                        }
                    } Catch {}   
                    $reg.Close() 
                }                  
            } Else {
                Write-Error "$($Computer): unable to reach remote system!"
            }
        } 
    } 
} 

Function Get-UserShareDACL {
    [cmdletbinding()]
    Param(
        [Parameter()]
        $Computername = $Computername                     
    )                   
    Try {    
        Write-Verbose "Computer: $($Computername)"
        #Retrieve share information from comptuer
        $Shares = Get-WmiObject -Class Win32_LogicalShareSecuritySetting -ComputerName $Computername -ea stop
        ForEach ($Share in $Shares) {
            $MoreShare = $Share.GetRelated('Win32_Share')
            Write-Verbose "Share: $($Share.name)"
            #Try to get the security descriptor
            $SecurityDescriptor = $Share.GetSecurityDescriptor()
            #Iterate through each descriptor
            ForEach ($DACL in $SecurityDescriptor.Descriptor.DACL) {
                [pscustomobject] @{
                    Computername = $Computername
                    Name = $Share.Name
                    Path = $MoreShare.Path
                    Type = $ShareType[[int]$MoreShare.Type]
                    Description = $MoreShare.Description
                    DACLName = $DACL.Trustee.Name
                    AccessRight = $AccessMask[[int]$DACL.AccessMask]
                    AccessType = $AceType[[int]$DACL.AceType]                    
                }
            }
        }
    }
    #Catch any errors                
    Catch {}                                                    
}

Function Get-AdminShare {
    [cmdletbinding()]
    Param (
        $Computername = $Computername
    )
    $WMIParams = @{
        Computername = $Computername
        Class = 'Win32_Share'
        Property = 'Name', 'Path', 'Description', 'Type'
        ErrorAction = 'Stop'
        Filter = "Type='2147483651' OR Type='2147483646' OR Type='2147483647' OR Type='2147483648'"
    }
    Get-WmiObject @WMIParams | Select-Object Name, Path, Description, 
    @{L='Type';E={$ShareType[[int64]$_.Type]}}
}

Function Convert-ChassisType {
    Param ([int[]]$ChassisType)
    $List = New-Object System.Collections.ArrayList
    Switch ($ChassisType) {
        0x0001  {[void]$List.Add('Other')}
        0x0002  {[void]$List.Add('Unknown')}
        0x0003  {[void]$List.Add('Desktop')}
        0x0004  {[void]$List.Add('Low Profile Desktop')}
        0x0005  {[void]$List.Add('Pizza Box')}
        0x0006  {[void]$List.Add('Mini Tower')}
        0x0007  {[void]$List.Add('Tower')}
        0x0008  {[void]$List.Add('Portable')}
        0x0009  {[void]$List.Add('Laptop')}
        0x000A  {[void]$List.Add('Notebook')}
        0x000B  {[void]$List.Add('Hand Held')}
        0x000C  {[void]$List.Add('Docking Station')}
        0x000D  {[void]$List.Add('All in One')}
        0x000E  {[void]$List.Add('Sub Notebook')}
        0x000F  {[void]$List.Add('Space-Saving')}
        0x0010  {[void]$List.Add('Lunch Box')}
        0x0011  {[void]$List.Add('Main System Chassis')}
        0x0012  {[void]$List.Add('Expansion Chassis')}
        0x0013  {[void]$List.Add('Subchassis')}
        0x0014  {[void]$List.Add('Bus Expansion Chassis')}
        0x0015  {[void]$List.Add('Peripheral Chassis')}
        0x0016  {[void]$List.Add('Storage Chassis')}
        0x0017  {[void]$List.Add('Rack Mount Chassis')}
        0x0018  {[void]$List.Add('Sealed-Case PC')}
    }
    $List -join ', '
}

Function Get-ScheduledTask {   
    [cmdletbinding()]
    Param (    
        [parameter(ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [string[]]$Computername = $env:COMPUTERNAME
    )
    Begin {
        $ST = new-object -com("Schedule.Service")
    }
    Process {
        ForEach ($Computer in $Computername) {
            Try {
                $st.Connect($Computer)
                $root=  $st.GetFolder("\")
                @($root.GetTasks(0)) | ForEach {
                    $xml = ([xml]$_.xml).task
                    [pscustomobject] @{
                        Computername = $Computer
                        Task = $_.Name
                        Author = $xml.RegistrationInfo.Author
                        RunAs = $xml.Principals.Principal.UserId                        
                        Enabled = $_.Enabled
                        State = Switch ($_.State) {
                            0 {'Unknown'}
                            1 {'Disabled'}
                            2 {'Queued'}
                            3 {'Ready'}
                            4 {'Running'}
                        }
                        LastTaskResult = Switch ($_.LastTaskResult) {
                            0x0 {"Successfully completed"}
                            0x1 {"Incorrect function called"}
                            0x2 {"File not found"}
                            0xa {"Environment is not correct"}
                            0x41300 {"Task is ready to run at its next scheduled time"}
                            0x41301 {"Task is currently running"}
                            0x41302 {"Task is disabled"}
                            0x41303 {"Task has not yet run"}
                            0x41304 {"There are no more runs scheduled for this task"}
                            0x41306 {"Task is terminated"}
                            0x00041307 {"Either the task has no triggers or the existing triggers are disabled or not set"}
                            0x00041308 {"Event triggers do not have set run times"}
                            0x80041309 {"A task's trigger is not found"}
                            0x8004130A {"One or more of the properties required to run this task have not been set"}
                            0x8004130B {"There is no running instance of the task"}
                            0x8004130C {"The Task * SCHEDuler service is not installed on this computer"}
                            0x8004130D {"The task object could not be opened"}
                            0x8004130E {"The object is either an invalid task object or is not a task object"}
                            0x8004130F {"No account information could be found in the Task * SCHEDuler security database for the task indicated"}
                            0x80041310 {"Unable to establish existence of the account specified"}
                            0x80041311 {"Corruption was detected in the Task * SCHEDuler security database"}
                            0x80041312 {"Task * SCHEDuler security services are available only on Windows NT"}
                            0x80041313 {"The task object version is either unsupported or invalid"}
                            0x80041314 {"The task has been configured with an unsupported combination of account settings and run time options"}
                            0x80041315 {"The Task * SCHEDuler Service is not running"}
                            0x80041316 {"The task XML contains an unexpected node"}
                            0x80041317 {"The task XML contains an element or attribute from an unexpected namespace"}
                            0x80041318 {"The task XML contains a value which is incorrectly formatted or out of range"}
                            0x80041319 {"The task XML is missing a required element or attribute"}
                            0x8004131A {"The task XML is malformed"}
                            0x0004131B {"The task is registered, but not all specified triggers will start the task"}
                            0x0004131C {"The task is registered, but may fail to start"}
                            0x8004131D {"The task XML contains too many nodes of the same type"}
                            0x8004131E {"The task cannot be started after the trigger end boundary"}
                            0x8004131F {"An instance of this task is already running"}
                            0x80041320 {"The task will not run because the user is not logged on"}
                            0x80041321 {"The task image is corrupt or has been tampered with"}
                            0x80041322 {"The Task * SCHEDuler service is not available"}
                            0x80041323 {"The Task * SCHEDuler service is too busy to handle your request"}
                            0x80041324 {"The Task * SCHEDuler service attempted to run the task, but the task did not run due to one of the constraints in the task definition"}
                            0x00041325 {"The Task * SCHEDuler service has asked the task to run"}
                            0x80041326 {"The task is disabled"}
                            0x80041327 {"The task has properties that are not compatible with earlier versions of Windows"}
                            0x80041328 {"The task settings do not allow the task to start on demand"}
                            Default {[string]$_}
                        }
                        Command = $xml.Actions.Exec.Command
                        Arguments = $xml.Actions.Exec.Arguments
                        StartDirectory =$xml.Actions.Exec.WorkingDirectory
                        Hidden = $xml.Settings.Hidden
                    }
                }
            } Catch {
                Write-Warning ("{0}: {1}" -f $Computer, $_.Exception.Message)
            }
        }
    }
}        

#endregion Helper Functions

#region Data Gathering
$InventoryDate = Get-Date
$ServerGroup = 'MemberServer'

Get-Server -MemberServer | Start-RSJob -Name {$_} -FunctionsToLoad Get-ScheduledTask, Invoke-SQLCmd,Get-LocalGroup,Get-LocalUser,Get-SecurityUpdate,
    Get-Software,Get-AdminShare,Get-UserShareDACL,Convert-ChassisType,Write-DataTable, Out-DataTable -ScriptBlock {
    Write-Verbose "[$($_)] - Initializing" -Verbose
    #region Variables
    $Computername = $_
    $SQLParams = @{
        Computername = $Using:SQLServer
        Database = 'ServerInventory'
        CommandType = 'NonQuery'
        ErrorAction = 'Stop'
        SQLParameter = @{
            '@Computername' = $Computername
        }
        Verbose = $True
    }
    $Date = $Using:InventoryDate
    #endregion Variables

    #region Lookups
    $DomainRole = @{
        0x0 = 'Standalone Workstation'
        0x1 = 'Member Workstation'
        0x2 = 'Standalone Server'
        0x3 = 'Member Server'
        0x4 = 'Backup Domain Controller'
        0x5 = 'Primary Domain Controller'
    }
    $DriveType = @{
        0x0 = 'Unknown'
        0x1 = 'No Root Directory'
        0x2 = 'Removable Disk'
        0x3 = 'Local Disk'
        0x4 = 'Network Drive'
        0x5 = 'Compact Disk'
        0x6 = 'RAM Disk'
    }
    $GroupType = @{
        0x2 = 'Global Group'
        0x4 = 'Local Group'
        0x8 = 'Universal Group'
        2147483648 = 'Security Enabled'
    }
    $ShareType = @{
        0x0 = 'Disk Drive'
        0x1 = 'Print Queue'
        0x2 = 'Device'
        0x3 = 'IPC'
        2147483648 = 'Disk Drive Admin'
        2147483647 = 'Print Queue Admin'
        2147483646 = 'Device Admin'
        2147483651 = 'IPC Admin'
    }
    $AceType = @{
        0x0 = 'Access Allowed'
        0x1 = 'Access Denied'
        0x2 = 'Audit'
    }
    $AceFlags = @{
        0x1 = 'OBJECT_INHERIT_ACE'
        0X2 = 'CONTAINER_INHERIT_ACE'
        0X4 = 'NO_PROPAGATE_ACE'
        0X8 = 'INHERIT_ONLY_ACE'
        0X10 = 'INHERITED_ACE'
    }
    $AccessMask = @{
        0x1F01FF = "FullControl"
        0x120089 = "Read"
        0x12019F = "Read, Write"
        0x1200A9 = "ReadAndExecute"
        1610612736 = "ReadAndExecuteExtended"
        0x1301BF = "ReadAndExecute, Modify, Write"
        0x1201BF = "ReadAndExecute, Write"
        0x10000000 = "FullControl (Sub Only)"
    }
    $ProcessorType = @{
        0x1 = 'Other'
        0x2 = 'Unknown'
        0x3 = 'Central Processor'
        0x4 = 'Math Processor'
        0x5 = 'DSP Processor'
        0x6 = 'Video Processor'
    }
    $TypeDetail = @{
        0x1 = 'Reserved'
        0x2 = 'Other'
        0x4 = 'Unknown'
        0x8 = 'Fast-paged'
        0x10 = 'Static column'
        0x20 = 'Pseudo-static'
        0x40 = 'RAMBUS'
        0x80 = 'Synchronous'
        0x100 = 'CMOS'
        0x200 = 'EDO'
        0x400 = 'Window DRAM'
        0x800 = 'Cache DRAM'
        0x1000 = 'Nonvolatile'
    }
    $MemoryType = @{
        0x0 = 'Unknown'
        0x1 = 'Other'
        0x2 = 'DRAM'
        0x3 = 'Synchronous DRAM'
        0x4 = 'Cache DRAM'
        0x5 = 'EDO'
        0x6 = 'EDRAM'
        0x7 = 'VRAM'
        0x8 = 'SRAM'
        0x9 = 'RAM'
        0xA = 'ROM'
        0xB = 'Flash'
        0xC = 'EEPROM'
        0xD = 'FEPROM'
        0xE = 'EPROM'
        0xF = 'CDRAM'
        0x10 = '3DRAM'
        0x11 = 'SDRAM'
        0x12 = 'SGRAM'
        0x13 = 'RDRAM'
        0x14 = 'DDR'
        0x15 = 'DDR-2'
    }
    $DebugType = @{
        0x0 = 'None'
        0x1 = 'Complete Memory Dump'
        0x2 = 'Kernel Memory Dump'
        0x3 = 'Small Memory Dump'
    }
    $AUOptions = @{
        0x2 = 'Notify before download'
        0x3 = 'Automatically download and notify of installation.'
        0x4 = 'Automatic download and scheduled installation.'
        0x5 = 'Automatic Updates is required, but end users can configure it.'
    }
    $DefaultNetworkRole = @{
        0x0 = 'ClusterNetworkRoleNone'
        0x1 = 'ClusterNetworkRoleInternalUse'
        0x2 = 'ClusterNetworkRoleClientAccess'
        0x3 = 'ClusterNetworkRoleInternalAndClient'
    }
    $ClusState = @{
        -1 = 'StateUnknown'
        0x0 = 'Inherited'
        0x1 = 'Initializing'
        0x2 = 'Online'
        0x3 = 'Offline'
        0x4 = 'Failed'
        0x80 = 'Pending'
        0x81 = 'Online Pending'
        0x82 = 'Offline Pending'
    }
    $NetState = @{
        -1 = 'StateUnknown'
        0x0 = 'Unavailable'
        0x1 = 'Failed'
        0x2 = 'Unreachable'
        0x3 = 'Up'
    }
    $FailbackType = @{
        0x0 = 'ClusterGroupPreventFailback'
        0x1 = 'ClusterGroupAllowFailback'
    }
    #endregion Lookups
    
    #region General
    Try {
        $CS = Get-WmiObject Win32_ComputerSystem -ComputerName $Computername -ErrorAction Stop
        $Enclosure = Get-WmiObject Win32_SystemEnclosure -ComputerName $Computername
        $BIOS = Get-WmiObject Win32_Bios -ComputerName $Computername 
        $General = [pscustomobject]@{
            Computername = $Computername
            Manufacturer = $CS.Manufacturer
            Model = $CS.Model
            SystemType = $CS.SystemType
            SerialNumber = $Enclosure.SerialNumber
            ChassisType = (Convert-ChassisType $Enclosure.ChassisTypes)
            Description = $Enclosure.Description
            BIOSManufacturer = $BIOS.Manufacturer
            BIOSName = $BIOS.Name
            BIOSSerialNumber = $BIOS.SerialNumber
            BIOSVersion = $BIOS.SMBIOSBIOSVersion
            InventoryDate = $Date
        }
        If ($General) {
            Write-Verbose "[$Computername - $(Get-Date)] Removing old data" -Verbose
            $SQLParams.CommandType = 'NonQuery'
            $SQLParams.TSQL = "DELETE FROM tbGeneral WHERE Computername = @Computername"            
            Invoke-SQLCmd @SQLParams
            If ($Return.DefaultView) {
                Write-Verbose "[$Computername - $(Get-Date)] Throwing error" -Verbose
                Throw 'FAIL'
            }
            Else {
                $SQLParams.CommandType = 'NonQuery'
                $DataTable = $General | Out-DataTable
                Write-Verbose "[$Computername - $(Get-Date)] Updating data" -Verbose
                Write-DataTable -Computername $Using:SQLServer -Database ServerInventory -TableName tbGeneral -Data $DataTable -ErrorAction Stop
            }
        }
    }
    Catch {
        Write-Verbose "WARNING - [$Computername - $(Get-Date)]" -Verbose
        Write-Warning $_
        BREAK
    }
    #endregion General

    #region OperatingSystem
    ## TODO: Include Product Key?
    $TimeZone = Get-WmiObject Win32_TimeZone -ComputerName $Computername -ErrorAction Stop
    $OS = Get-WmiObject Win32_OperatingSystem -ComputerName $Computername 
    $PageFile = Get-WMIObject win32_PageFile -ComputerName $Computername 
    $LastReboot = Try {
        $OS.ConvertToDateTime($OS.LastBootUpTime)
    } 
    Catch {
    }
    $OperatingSystem = [pscustomobject]@{
        Computername = $Computername
        Caption = $OS.caption
        Version = $OS.version
        ServicePack = ("{0}.{1}" -f $OS.ServicePackMajorVersion, $OS.ServicePackMinorVersion)
        LastReboot = $LastReboot
        OSArchitecture = $OS.OSArchitecture
        TimeZone = $TimeZone.Caption
        PageFile = $PageFile.Name
        PageFileSizeGB = ("{0:N2}" -f ($PageFile.FileSize /1GB))
        InventoryDate = $Date
    }
    If ($OperatingSystem) {
        $SQLParams.TSQL = "DELETE FROM tbOperatingSystem WHERE Computername = @Computername"
        Invoke-SQLCmd @SQLParams
        $DataTable = $OperatingSystem | Out-DataTable
        Write-DataTable -Computername $Using:SQLServer -Database ServerInventory -TableName tbOperatingSystem -Data $DataTable
    }
    #endregion OperatingSystem

    #region Memory
    $Memory = @(Get-WmiObject Win32_PhysicalMemory -ComputerName $Computername -ErrorAction Stop| ForEach {
        [pscustomobject]@{
            Computername = $Computername
            DeviceID = $_.tag
            MemoryType = $MemoryType[[int]$_.MemoryType]
            "Capacity(GB)" = "{0}" -f ($_.capacity/1GB)
            TypeDetail = $TypeDetail[[int]$_.TypeDetail]
            Locator = $_.DeviceLocator
            InventoryDate = $Date
        }
    })

    If ($Memory) {
        $SQLParams.TSQL = "DELETE FROM tbMemory WHERE Computername = @Computername"
        Invoke-SQLCmd @SQLParams
        $DataTable = $Memory | Out-DataTable
        Write-DataTable -Computername $Using:SQLServer -Database ServerInventory -TableName tbMemory -Data $DataTable
    }
    #endregion Memory

    #region Network
    $Network = @(Get-WmiObject Win32_NetworkAdapterConfiguration -Filter "IPEnabled='True'" -ComputerName $Computername -ErrorAction Stop) | ForEach {
        [pscustomobject]@{
            Computername = $Computername
            DeviceName = $_.Caption
            DHCPEnabled = $_.DHCPEnabled
            MACAddress = $_.MACAddress    
            IPAddress = ($_.IpAddress  -join '; ')
            SubnetMask = ($_.IPSubnet  -join '; ')
            DefaultGateway = ($_.DefaultIPGateway -join '; ')
            DNSServers = ($_.DNSServerSearchOrder  -join '; ')
            InventoryDate = $Date
        }
    }

    If ($Network) {
        $SQLParams.TSQL = "DELETE FROM tbNetwork WHERE Computername = @Computername"
        Invoke-SQLCmd @SQLParams
        $DataTable = $Network | Out-DataTable
        Write-DataTable -Computername $Using:SQLServer -Database ServerInventory -TableName tbNetwork -Data $DataTable
    }
    #endregion Network

    #region CPU
    $Processor = @(Get-WmiObject Win32_Processor -ComputerName $Computername -ErrorAction Stop | ForEach {
        [pscustomobject]@{
            Computername = $Computername
            DeviceID = $_.DeviceID
            Description = $_.Description
            ProcessorType = $ProcessorType[$_.processortype]
            CoreCount = $_.NumberofCores
            NumLogicalProcessors = $_.NumberOfLogicalProcessors
            MaxSpeed = ("{0:N2} GHz" -f ($_.MaxClockSpeed/1000))
            InventoryDate = $Date
        }
    })

    If ($Processor) {
        $SQLParams.TSQL = "DELETE FROM tbProcessor WHERE Computername = @Computername"
        Invoke-SQLCmd @SQLParams
        $DataTable = $Processor | Out-DataTable
        Write-DataTable -Computername $Using:SQLServer -Database ServerInventory -TableName tbProcessor -Data $DataTable
    }
    #endregion CPU

    #region Drives
    $Disk = @(Get-WmiObject Win32_Volume -Filter "(Not Name LIKE '\\\\?\\%')" -ComputerName $Computername -ErrorAction Stop | ForEach {
        [pscustomobject]@{
            Computername = $Computername
            Drive = $_.Name
            DriveType = $DriveType[[int]$_.DriveType]
            Label = $_.label
            FileSystem = $_.FileSystem
            FreeSpaceGB = "{0:N2}" -f ($_.FreeSpace /1GB)
            CapacityGB = "{0:N2}" -f ($_.Capacity/1GB)
            PercentFree = ($_.FreeSpace/$_.Capacity)
            InventoryDate = $Date
        }
    })

    If ($Disk) {
        $SQLParams.TSQL = "DELETE FROM tbDrives WHERE Computername = @Computername"
        Invoke-SQLCmd @SQLParams
        $DataTable = $Disk | Out-DataTable
        Write-DataTable -Computername $Using:SQLServer -Database ServerInventory -TableName tbDrives -Data $DataTable
    }
    #endregion Drives

    #region AdminShares
    $AdminShare = @(Get-AdminShare -Computername $Computername -ErrorAction Stop | ForEach {
        [pscustomobject]@{
            Computername = $Computername
            Name = $_.Name
            Path = $_.Path
            Type = $_.Type  
            InventoryDate = $Date  
        }
    })

    If ($AdminShare) {
        $SQLParams.TSQL = "DELETE FROM tbAdminShare WHERE Computername = @Computername"
        Invoke-SQLCmd @SQLParams
        $DataTable = $AdminShare | Out-DataTable
        Write-DataTable -Computername $Using:SQLServer -Database ServerInventory -TableName tbAdminShare -Data $DataTable
    }
    #endregion AdminShares

    #region UserShares
    $UserShare = Try {
    Get-UserShareDACL -Computername $Computername -ErrorAction Stop | Select *,@{L='InventoryDate';E={$Date}}
    } 
    Catch {}
    If ($UserShare) {
        $SQLParams.TSQL = "DELETE FROM tbUserShare WHERE Computername = @Computername"
        Invoke-SQLCmd @SQLParams
        $DataTable = $UserShare | Out-DataTable
        Write-DataTable -Computername $Using:SQLServer -Database ServerInventory -TableName tbUserShare -Data $DataTable
    }
    #endregion UserShares

    #region Local Users
    If ($Using:ServerGroup -eq 'MemberServer') {
        $Users = @(Get-LocalUser -ComputerName $Computername -ErrorAction Stop) | Select *,@{L='InventoryDate';E={$Date}}

        If ($Users) {
            $SQLParams.TSQL = "DELETE FROM tbUsers WHERE Computername = @Computername"
            Invoke-SQLCmd @SQLParams
            $DataTable = $Users | Out-DataTable
            Write-DataTable -Computername $Using:SQLServer -Database ServerInventory -TableName tbUsers -Data $DataTable
        }
    }
    #endregion Local Users

    #region Local Groups
    If ($Using:ServerGroup -eq 'MemberServer') {
        $Groups = @(Get-LocalGroup -ComputerName $Computername -ErrorAction Stop) | Select *,@{L='InventoryDate';E={$Date}}

        If ($Groups) {
            $SQLParams.TSQL = "DELETE FROM tbGroups WHERE Computername = @Computername"
            Invoke-SQLCmd @SQLParams
            $DataTable = $Groups | Out-DataTable
            Write-DataTable -Computername $Using:SQLServer -Database ServerInventory -TableName tbGroups -Data $DataTable
        }
    }
    #endregion Local Groups

    #region Server Roles
    $ServerRoles = Try {
        Get-WmiObject Win32_ServerFeature -ComputerName $Computername -ErrorAction Stop | ForEach {
            [pscustomobject]@{
                Computername = $Computername
                ID = $_.Id
                Name = $_.Name
                InventoryDate = $Date
            }
        }
    } 
    Catch {}

    If ($ServerRoles) {
        $SQLParams.TSQL = "DELETE FROM tbServerRoles WHERE Computername = @Computername"
        Invoke-SQLCmd @SQLParams
        $DataTable = $ServerRoles | Out-DataTable
        Write-DataTable -Computername $Using:SQLServer -Database ServerInventory -TableName tbServerRoles -Data $DataTable
    }
    #endregion Server Roles

    #region Scheduled Tasks
    $ScheduledTasks = Get-ScheduledTask -Computername $Computername -ErrorAction Stop | Select *,@{L='InventoryDate';E={$Date}}

    If ($ScheduledTasks) {
        $SQLParams.TSQL = "DELETE FROM tbScheduledTasks WHERE Computername = @Computername"
        Invoke-SQLCmd @SQLParams
        $DataTable = $ScheduledTasks | Out-DataTable
        Write-DataTable -Computername $Using:SQLServer -Database ServerInventory -TableName tbScheduledTasks -Data $DataTable
    }
    #endregion Scheduled Tasks

    #region Software
    $Software = Get-Software -Computername $Computername -ErrorAction Stop | Sort DisplayName | Select *,@{L='InventoryDate';E={$Date}}

    If ($Software) {
        $SQLParams.TSQL = "DELETE FROM tbSoftware WHERE Computername = @Computername"
        Invoke-SQLCmd @SQLParams
        $DataTable = $Software | Out-DataTable
        Write-DataTable -Computername $Using:SQLServer -Database ServerInventory -TableName tbSoftware -Data $DataTable
    }
    #endregion Software

    #region Updates
    $Updates = Get-SecurityUpdate -Computername $Computername -ErrorAction Stop | 
        Select @{L='Computername';E={$Computername}},Description, HotFixID, InstalledOn, Type,@{L='InventoryDate';E={$Date}} | 
        Group-Object HotFixID | ForEach {$_.Group | Sort-Object -Unique DisplayName}
    $Hotfixes = Get-HotFix -ComputerName $Computername | ForEach {
        Switch -Wildcard ($_.Description) {
            "Service Pack*" {$Type = 'Service Pack'}
            "Hotfix*" {$Type = 'Hotfix'}
            "Update*" {$Type = 'Update'}
            "Security Update*" {$Type = 'Security Update'}
            Default {$Type = 'Unknown'}
        }
        [pscustomobject]@{
            Computername = $Computername
            Description = $_.Description
            HotFixID = $_.HotFixID
            InstalledOn = $_.InstalledOn
            Type = $Type
            InventoryDate = $Date
        }
    } 
    $TotalUpdates = $Hotfixes + $Updates

    If ($TotalUpdates) {
        $SQLParams.TSQL = "DELETE FROM tbUpdates WHERE Computername = @Computername"
        Invoke-SQLCmd @SQLParams
        $DataTable = $TotalUpdates | Out-DataTable
        Write-DataTable -Computername $Using:SQLServer -Database ServerInventory -TableName tbUpdates -Data $DataTable
    }
    #endregion Updates
} | Wait-RSJob -ShowProgress | Remove-RSJob
#endregion Data Gathering