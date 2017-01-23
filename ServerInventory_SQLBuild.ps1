[cmdletbinding(
)]
Param (
        [string]$Computername = 'vsql',
        
        [parameter()]
        [string]$Database = 'Master'       
)
#region Helper Functions
Function Invoke-SQLCmd {    
    [cmdletbinding(
        DefaultParameterSetName = 'NoCred',
        SupportsShouldProcess = $True,
        ConfirmImpact = 'Low'
    )]
    Param (
        [parameter()]
        [string]$Computername = 'vsql',
        
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
    Write-Verbose "Initiating query"
    $Command=new-object system.Data.SqlClient.SqlCommand($Tsql,$Connection)
    If ($PSBoundParameters.ContainsKey('SQLParameter')) {
        $SqlParameter.GetEnumerator() | ForEach {
            Write-Debug "Adding SQL Parameter: $($_.Key) with Value: $($_.Value)"
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
#endregion Helper Functions

#region Check/Create for database
$Database = 'ServerInventory'
$Computername = 'vsql'
$SQLParams = @{
    Computername = $Computername
    Database = 'Master'
    CommandType = 'Query'
    ErrorAction = 'Stop'
    SQLParameter = @{
        '@DatabaseName' = $Database
    }
    Verbose = $True
    Confirm=$False
}
$SQLParams.TSQL = "SELECT Name FROM sys.databases WHERE Name = @DatabaseName"
$Results = Invoke-SQLCmd @SQLParams
If ($Results.Name -eq $Null) {
    #Proceed with building the database and Table
    $SQLParams.CommandType='NonQuery'
    $SQLParams.Remove('SQLParameter')
    $SQLParams.TSQL = "CREATE Database $Database"
    Invoke-SQLCmd @SQLParams
} 
#endregion Check/Create for database

#region Create Tables

#region General Table
$Table = 'tbGeneral'
$SQLParams.CommandType = 'Query'
$SQLParams.SQLParameter = @{
    '@TableName' = $Table
}
$SQLParams.Database = 'ServerInventory'    
$SQLParams.TSQL = "SELECT TABLE_NAME AS Name FROM information_schema.tables WHERE TABLE_NAME = @TableName"
$Results = Invoke-SQLCmd @SQLParams
If ($Results.Name -eq $Null) {
    #Create the table
    $SQLParams.Remove('SQLParameter')
    $SQLParams.CommandType='NonQuery'
    $SQLParams.TSQL = "CREATE TABLE $Table  (        
        ComputerName nvarchar (256), 
        Manufacturer nvarchar (256), 
        Model  nvarchar (256),
        SystemType nvarchar (256) ,
        SerialNumber nvarchar (256), 
        ChassisType nvarchar (256),
        Description nvarchar (256),
        BIOSManufacturer nvarchar (256),
        BIOSName nvarchar (256),
        BIOSSerialNumber nvarchar (256),
        BIOSVersion nvarchar (256),
        InventoryDate datetime
    )"
    Invoke-SQLCmd @SQLParams
}
#endregion General Table

#region OperatingSystem Table
$Table = 'tbOperatingSystem'
$SQLParams.CommandType = 'Query'
$SQLParams.SQLParameter = @{
    '@TableName' = $Table
}
$SQLParams.Database = 'ServerInventory'    
$SQLParams.TSQL = "SELECT TABLE_NAME AS Name FROM information_schema.tables WHERE TABLE_NAME = @TableName"
$Results = Invoke-SQLCmd @SQLParams
If ($Results.Name -eq $Null) {
    #Create the table
    $SQLParams.Remove('SQLParameter')
    $SQLParams.CommandType='NonQuery'
    $SQLParams.TSQL = "CREATE TABLE $Table  (        
        ComputerName nvarchar (256), 
        Caption nvarchar (256), 
        Version nvarchar (256),
        ServicePack nvarchar (256) ,
        LastReboot datetime, 
        OSArchitecture nvarchar (25),
        TimeZone nvarchar (256),
        PageFile nvarchar (256),
        PageFileSizeGB decimal (20),
        InventoryDate datetime
    )"
    Invoke-SQLCmd @SQLParams
}
#endregion OperatingSystem Table

#region Memory Table
$Table = 'tbMemory'
$SQLParams.CommandType = 'Query'
$SQLParams.SQLParameter = @{
    '@TableName' = $Table
}
$SQLParams.Database = 'ServerInventory'    
$SQLParams.TSQL = "SELECT TABLE_NAME AS Name FROM information_schema.tables WHERE TABLE_NAME = @TableName"
$Results = Invoke-SQLCmd @SQLParams
If ($Results.Name -eq $Null) {
    #Create the table
    $SQLParams.Remove('SQLParameter')
    $SQLParams.CommandType='NonQuery'
    $SQLParams.TSQL = "CREATE TABLE $Table  (        
        ComputerName nvarchar (256), 
        DeviceID nvarchar (256), 
        MemoryType  nvarchar (256),
        CapacityGB decimal (20),
        TypeDetail nvarchar (256), 
        Locator nvarchar (256),
        InventoryDate datetime
    )"
    Invoke-SQLCmd @SQLParams
}
#endregion Memory Table

#region Processor Table
$Table = 'tbProcessor'
$SQLParams.CommandType = 'Query'
$SQLParams.SQLParameter = @{
    '@TableName' = $Table
}
$SQLParams.Database = 'ServerInventory'    
$SQLParams.TSQL = "SELECT TABLE_NAME AS Name FROM information_schema.tables WHERE TABLE_NAME = @TableName"
$Results = Invoke-SQLCmd @SQLParams
If ($Results.Name -eq $Null) {
    #Create the table
    $SQLParams.Remove('SQLParameter')
    $SQLParams.CommandType='NonQuery'
    $SQLParams.TSQL = "CREATE TABLE $Table  (        
        ComputerName nvarchar (256), 
        DeviceID nvarchar (256), 
        Description nvarchar (256),
        ProcessorType nvarchar (256),
        CoreCount decimal (5), 
        NumLogicalProcessors decimal (5),
        MaxSpeed nvarchar (50),
        InventoryDate datetime,
    )"
    Invoke-SQLCmd @SQLParams
}
#endregion Processor Table

#region Network Table
$Table = 'tbNetwork'
$SQLParams.CommandType = 'Query'
$SQLParams.SQLParameter = @{
    '@TableName' = $Table
}
$SQLParams.Database = 'ServerInventory'    
$SQLParams.TSQL = "SELECT TABLE_NAME AS Name FROM information_schema.tables WHERE TABLE_NAME = @TableName"
$Results = Invoke-SQLCmd @SQLParams
If ($Results.Name -eq $Null) {
    #Create the table
    $SQLParams.Remove('SQLParameter')
    $SQLParams.CommandType='NonQuery'
    $SQLParams.TSQL = "CREATE TABLE $Table  (        
        ComputerName nvarchar (256), 
        DeviceName nvarchar (256), 
        DHCPEnabled bit,
        MACAddress nvarchar (50),
        IPAddress nvarchar (MAX), 
        SubnetMask nvarchar (MAX),
        DefaultGateway nvarchar (200),
        DNSServers nvarchar (MAX),
        InventoryDate datetime
    )"
    Invoke-SQLCmd @SQLParams
}
#endregion Network Table

#region Drives Table
$Table = 'tbDrives'
$SQLParams.CommandType = 'Query'
$SQLParams.SQLParameter = @{
    '@TableName' = $Table
}
$SQLParams.Database = 'ServerInventory'    
$SQLParams.TSQL = "SELECT TABLE_NAME AS Name FROM information_schema.tables WHERE TABLE_NAME = @TableName"
$Results = Invoke-SQLCmd @SQLParams
If ($Results.Name -eq $Null) {
    #Create the table
    $SQLParams.Remove('SQLParameter')
    $SQLParams.CommandType='NonQuery'
    $SQLParams.TSQL = "CREATE TABLE $Table  (        
        ComputerName nvarchar (256), 
        Drive nvarchar (256), 
        DriveType nvarchar (256),
        Label nvarchar (256),
        FileSystem nvarchar (256), 
        FreeSpaceGB decimal (10,3),
        CapacityGB decimal (10,3),
        PercentFree decimal (4,4),
        InventoryDate datetime
    )"
    Invoke-SQLCmd @SQLParams
}
#endregion Drives Table

#region AdminShare Table
$Table = 'tbAdminShare'
$SQLParams.CommandType = 'Query'
$SQLParams.SQLParameter = @{
    '@TableName' = $Table
}
$SQLParams.Database = 'ServerInventory'    
$SQLParams.TSQL = "SELECT TABLE_NAME AS Name FROM information_schema.tables WHERE TABLE_NAME = @TableName"
$Results = Invoke-SQLCmd @SQLParams
If ($Results.Name -eq $Null) {
    #Create the table
    $SQLParams.Remove('SQLParameter')
    $SQLParams.CommandType='NonQuery'
    $SQLParams.TSQL = "CREATE TABLE $Table  (        
        ComputerName nvarchar (256), 
        Name nvarchar (256), 
        Path nvarchar (256),
        Type nvarchar (256),
        InventoryDate datetime,
    )"
    Invoke-SQLCmd @SQLParams
}
#endregion AdminShare Table

#region UserShare Table
$Table = 'tbUserShare'
$SQLParams.CommandType = 'Query'
$SQLParams.SQLParameter = @{
    '@TableName' = $Table
}
$SQLParams.Database = 'ServerInventory'    
$SQLParams.TSQL = "SELECT TABLE_NAME AS Name FROM information_schema.tables WHERE TABLE_NAME = @TableName"
$Results = Invoke-SQLCmd @SQLParams
If ($Results.Name -eq $Null) {
    #Create the table
    $SQLParams.Remove('SQLParameter')
    $SQLParams.CommandType='NonQuery'
    $SQLParams.TSQL = "CREATE TABLE $Table  (        
        ComputerName nvarchar (256), 
        Name nvarchar (256), 
        Path nvarchar (256),
        Type nvarchar (256),
        Description nvarchar (256),
        DACLName nvarchar (256),
        AccessRight nvarchar (MAX),
        AccessType nvarchar (MAX),
        InventoryDate datetime
    )"
    Invoke-SQLCmd @SQLParams
}
#endregion UserShare Table

#region Users Table
$Table = 'tbUsers'
$SQLParams.CommandType = 'Query'
$SQLParams.SQLParameter = @{
    '@TableName' = $Table
}
$SQLParams.Database = 'ServerInventory'    
$SQLParams.TSQL = "SELECT TABLE_NAME AS Name FROM information_schema.tables WHERE TABLE_NAME = @TableName"
$Results = Invoke-SQLCmd @SQLParams
If ($Results.Name -eq $Null) {
    #Create the table
    $SQLParams.Remove('SQLParameter')
    $SQLParams.CommandType='NonQuery'
    $SQLParams.TSQL = "CREATE TABLE $Table  (        
        ComputerName nvarchar (256), 
        Name nvarchar (256), 
        PasswordAge decimal (5),
        LastLogin datetime,
        SID nvarchar (100),
        UserFlags nvarchar (MAX),
        InventoryDate datetime
    )"
    Invoke-SQLCmd @SQLParams
}
#endregion Users Table

#region Groups Table
$Table = 'tbGroups'
$SQLParams.CommandType = 'Query'
$SQLParams.SQLParameter = @{
    '@TableName' = $Table
}
$SQLParams.Database = 'ServerInventory'    
$SQLParams.TSQL = "SELECT TABLE_NAME AS Name FROM information_schema.tables WHERE TABLE_NAME = @TableName"
$Results = Invoke-SQLCmd @SQLParams
If ($Results.Name -eq $Null) {
    #Create the table
    $SQLParams.Remove('SQLParameter')
    $SQLParams.CommandType='NonQuery'
    $SQLParams.TSQL = "CREATE TABLE $Table  (        
        ComputerName nvarchar (256), 
        Name nvarchar (256), 
        Members nvarchar (MAX),
        SID nvarchar (100),
        GroupType nvarchar (256),
        InventoryDate datetime
    )"
    Invoke-SQLCmd @SQLParams
}
#endregion Groups Table

#region ServerRoles Table
$Table = 'tbServerRoles'
$SQLParams.CommandType = 'Query'
$SQLParams.SQLParameter = @{
    '@TableName' = $Table
}
$SQLParams.Database = 'ServerInventory'    
$SQLParams.TSQL = "SELECT TABLE_NAME AS Name FROM information_schema.tables WHERE TABLE_NAME = @TableName"
$Results = Invoke-SQLCmd @SQLParams
If ($Results.Name -eq $Null) {
    #Create the table
    $SQLParams.Remove('SQLParameter')
    $SQLParams.CommandType='NonQuery'
    $SQLParams.TSQL = "CREATE TABLE $Table  (        
        ComputerName nvarchar (256), 
        Id decimal (5), 
        Name nvarchar (256),
        InventoryDate datetime
    )"
    Invoke-SQLCmd @SQLParams
}
#endregion ServerRoles Table

#region Software Table
$Table = 'tbSoftware'
$SQLParams.CommandType = 'Query'
$SQLParams.SQLParameter = @{
    '@TableName' = $Table
}
$SQLParams.Database = 'ServerInventory'    
$SQLParams.TSQL = "SELECT TABLE_NAME AS Name FROM information_schema.tables WHERE TABLE_NAME = @TableName"
$Results = Invoke-SQLCmd @SQLParams
If ($Results.Name -eq $Null) {
    #Create the table
    $SQLParams.Remove('SQLParameter')
    $SQLParams.CommandType='NonQuery'
    $SQLParams.TSQL = "CREATE TABLE $Table  (        
        ComputerName nvarchar (256), 
        DisplayName nvarchar (500),
        Version nvarchar (256),
        InstallDate datetime,
        Publisher nvarchar (256),
        UninstallString nvarchar (500),
        InstallLocation nvarchar (500),
        InstallSource nvarchar (500),
        HelpLink nvarchar (256),
        EstimatedSize decimal (10),
        InventoryDate datetime
    )"
    Invoke-SQLCmd @SQLParams
}
#endregion Software Table

#region Updates Table
$Table = 'tbUpdates'
$SQLParams.CommandType = 'Query'
$SQLParams.SQLParameter = @{
    '@TableName' = $Table
}
$SQLParams.Database = 'ServerInventory'    
$SQLParams.TSQL = "SELECT TABLE_NAME AS Name FROM information_schema.tables WHERE TABLE_NAME = @TableName"
$Results = Invoke-SQLCmd @SQLParams
If ($Results.Name -eq $Null) {
    #Create the table
    $SQLParams.Remove('SQLParameter')
    $SQLParams.CommandType='NonQuery'
    $SQLParams.TSQL = "CREATE TABLE $Table  (        
        ComputerName nvarchar (256), 
        Description nvarchar (256),
        HotFixID nvarchar (50),
        InstalledOn datetime,
        Type nvarchar (256),
        InventoryDate datetime
    )"
    Invoke-SQLCmd @SQLParams
}
#endregion Updates Table

#region ScheduledTasks Table
$Table = 'tbScheduledTasks'
$SQLParams.CommandType = 'Query'
$SQLParams.SQLParameter = @{
    '@TableName' = $Table
}
$SQLParams.Database = 'ServerInventory'    
$SQLParams.TSQL = "SELECT TABLE_NAME AS Name FROM information_schema.tables WHERE TABLE_NAME = @TableName"
$Results = Invoke-SQLCmd @SQLParams
If ($Results.Name -eq $Null) {
    #Create the table
    $SQLParams.Remove('SQLParameter')
    $SQLParams.CommandType='NonQuery'
    $SQLParams.TSQL = "CREATE TABLE $Table  (        
        ComputerName nvarchar (256), 
        Task nvarchar (256),
        Author nvarchar (256),
        RunAs nvarchar (256),
        Enabled bit,
        State nvarchar (256),
        LastTaskResult nvarchar (500),
        Command nvarchar (MAX),
        Arguments nvarchar (MAX),
        StartDirectory nvarchar (256),
        Hidden bit,
        InventoryDate datetime
    )"
    Invoke-SQLCmd @SQLParams
}
#endregion ScheduledTasks Table

#region Services Table
$Table = 'tbServices'
$SQLParams.CommandType = 'Query'
$SQLParams.SQLParameter = @{
    '@TableName' = $Table
}
$SQLParams.Database = 'ServerInventory'    
$SQLParams.TSQL = "SELECT TABLE_NAME AS Name FROM information_schema.tables WHERE TABLE_NAME = @TableName"
$Results = Invoke-SQLCmd @SQLParams
If ($Results.Name -eq $Null) {
    #Create the table
    $SQLParams.Remove('SQLParameter')
    $SQLParams.CommandType='NonQuery'
    $SQLParams.TSQL = "CREATE TABLE $Table  (        
        ComputerName nvarchar (256), 
        Name nvarchar (256),
        DisplayName nvarchar (256),
        Description nvarchar (256),
        IsDelayedAutoStart bit,
        SIDType nvarchar (256),
        Privileges nvarchar (500),
        ShutDownTimeout int,
        Type nvarchar (MAX),
        State nvarchar (256),
        Controls nvarchar (256),
        Win32ExitCode int,
        ServiceExitCode int,
        ProcessID int,
        ServiceFlags nvarchar (256),
        StartMode nvarchar (256),
        ErrorControl nvarchar (256),
        FilePath nvarchar (256),
        LoadOrderGroup nvarchar (256),
        Dependancies nvarchar (256),
        StartName nvarchar (256),
        InventoryDate datetime
    )"
    Invoke-SQLCmd @SQLParams
}
#endregion #region Services Table
#endregion Create Tables