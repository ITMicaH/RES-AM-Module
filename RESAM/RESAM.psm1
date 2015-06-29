
#region HelperFunctions

function Search-SQLDatabase
{
    Param (
        $SearchValue
    )
    $Query = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_CATALOG='$($RESAM_DB_Connection.DataBase)'"
    $AllTables = Invoke-SQLQuery -Query $Query
    If (!$AllTables)
    {
        throw "Unable to retreive tables from '$($RESAM_DB_Connection.DataBase)'."
    }
    foreach ($Table in $AllTables.TABLE_NAME)
    {
        $Columns = Invoke-SQLQuery -Query "select * from INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = N'$Table'"
        foreach ($Column in $Columns.COLUMN_NAME)
        {
            $result = Invoke-SQLQuery -Query "select * from dbo.$Table WHERE $Column LIKE '$($SearchValue.replace('*','%'))'"
            If ($result)
            {
                New-Object -TypeName psobject -Property @{
                    'TableName' = $Table
                    'Column' = $Column
                    'Row' = $result
                }
            } # end IF
        } # end foreach Column
    } # end foreach Table
}

function Invoke-SQLQuery
{
    [CmdletBinding()]
    [OutputType([int])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $Query,

        [Parameter(Mandatory=$False,
                   ValueFromPipeline = $false,
                   Position=1)]
        [string]
        $Type
    )

    Begin
    {
        If (!$RESAM_DB_Connection)
        {
            Throw "No open connection to a RES Automation Manager database detected. Run command Connect-RESAMDatabase first."
        }
    }
    Process
    {
        $command = $RESAM_DB_Connection.CreateCommand()
        $command.CommandText = $Query

        Write-Verbose "Running SQL query '$query'"
        #try
        #{
            $result = $command.ExecuteReader()
            $CustomTable = new-object "System.Data.DataTable"
            $CustomTable.Load($result)
            If ($Type)
            {
                $CustomTable | ConvertTo-PSObject -Type $Type
            }
            else
            {
                $CustomTable | ConvertTo-PSObject
            }
        #}
        #catch {}
        try
        {
            $result.close()
        }
        catch{}
    }
    End
    {
    }
}

function ConvertTo-PSObject
{
    [CmdletBinding()]
    [OutputType([int])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipeline = $true,
                   Position=0)]
        $InputObject,
        
        [Parameter(Mandatory=$False,
                   ValueFromPipeline = $false,
                   Position=1)]
        [string]
        $Type
    )

    Process
    {
        Write-Verbose "Creating custom object for output."
        $Properties = $InputObject | Get-Member -MemberType Property |
         select -ExpandProperty Name
        $ht = @{}
        foreach ($Property in $Properties)
        {
            $NewProp = $Property -replace '^(str|lng|ysn|dtm|img)',''
            $Value = $InputObject.$Property
            If ($NewProp -eq 'Status')
            {
                switch ($Value)
                {
                    '0' {$Value = 'Offline'}
                    '1' {$Value = 'Online'}
                }
            }
            if ($InputObject.$Property.GetType().Name -eq 'Byte[]')
            {
                $Value = ConvertFrom-ByteArray $Value
            }
            Write-Verbose "Creating output object."
            $ht.Add($NewProp,$Value)
        }
        $Object = New-Object -TypeName psobject -Property $ht
        If ($Type)
        {
            $Object.PSObject.TypeNames.Insert(0,"RES.AutomationManager.$Type")
        }
        $Object
    }
}
function ConvertFrom-ByteArray
{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true,
        ValueFromPipeline=$true,
        Position=0)]
        [System.Byte[]]
        $ByteArray
    )
    
    Write-Verbose "Processing Byte Array..."
    $NewArray = $ByteArray | ?{$_ -ne 0}
    [xml]$XML = [System.Text.Encoding]::ASCII.GetString($NewArray)
                
    $Object = New-Object -TypeName psobject
    $Properties = $XML | Get-Member -MemberType Property | ?{$_.Name -ne 'xml'}
    foreach ($Property in $Properties)
    {
        $Name = $Property.Name
        Write-Verbose "Adding property $Name to object."
        $Object | Add-Member -MemberType NoteProperty -Name $Name -Value $XML.$Name
    }
    $Object
    Write-Verbose "Finished processing array."
}

function Get-RESAMAgentTeams
{
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipelineByPropertyName=$true,
                   Position = 0)]
        [Alias('WUIDAgent')]
        [guid]
        $GUID
    )
    process
    {
        $Query = "select * from dbo.tblTeamAgents WHERE AgentGUID = '$($GUID.tostring())'"

        Invoke-SQLQuery $Query | Get-RESAMTeam
    }
}

function Add-RESAMFolderName
{
    [CmdletBinding()]
    Param (
    [Parameter(ValueFromPipeline=$true)]
    $InputObject)


    process
    {
        If ($InputObject.FolderGuid)
        {
            $Folder = $InputObject.FolderGuid | Get-RESAMFolder
            $InputObject | Add-Member -MemberType NoteProperty -Name FolderName -Value $Folder.Name
        }
        $InputObject
    }
}

function Optimize-RESAMAgent
{
    [CmdletBinding()]
    Param (
    [Parameter(ValueFromPipeline=$true)]
    $Agent)

    Process
    {
        Write-Verbose "Optimizing agent $($Agent.Name)."
        
        $Info = $Agent.Info | ?{$_ -ne 0}
        [xml]$XML = [System.Text.Encoding]::ASCII.GetString($Info)
        $Agent.Info = $XML.LAN
        
        $Properties = $Agent.Properties | ?{$_ -ne 0}
        [xml]$XML = [System.Text.Encoding]::ASCII.GetString($Properties)
        $Agent.Properties = $XML.LAN

    }
}

function Optimize-RESAMFolder
{
    [CmdletBinding()]
    Param (
    [Parameter(ValueFromPipeline=$true)]
    $Folder)


    process
    {
        $Folder.Name = $Folder.Name.Trim()
        switch ($Folder.FolderType)
        {
            1 {$Folder.FolderType = 'Module'}
            2 {$Folder.FolderType = 'Resource'}
            3 {$Folder.FolderType = 'Project'}
            5 {$Folder.FolderType = 'RunBook'}
            6 {$Folder.FolderType = 'Team'}
        }
        If ($Folder.ParentFolderGUID.tostring())
        {
            $Query = "select * from dbo.tblFolders WHERE FolderGUID = '$($Folder.ParentFolderGUID.tostring())'"
            $ParentFolder = Invoke-SQLQuery $Query
        }
        $Folder | Add-Member -MemberType NoteProperty -Name ParentFolderName -Value $ParentFolder.Name.trim()
        $Folder
    }
}

function Optimize-RESAMConnector
{
    [CmdletBinding()]
    Param (
    [Parameter(ValueFromPipeline=$true)]
    $Connector)


    process
    {
        switch ($Connector.Type)
        {
            1 {
                $Connector | Add-Member -MemberType NoteProperty -Name ConnectorFor -Value 'Database Servers'
                switch ($Connector.Flags)
                {
                    1  {$Connector.Type = 'Microsoft SQL Server'}
                    2  {$Connector.Type = 'Oracle'}
                    3  {$Connector.Type = 'Microsoft SQL Server;Oracle'}
                    4  {$Connector.Type = 'IBM DB2'}
                    5  {$Connector.Type = 'Microsoft SQL Server;IBM DB2'}
                    6  {$Connector.Type = 'Oracle;IBM DB2'}
                    7  {$Connector.Type = 'Microsoft SQL Server;Oracle;IBM DB2'}
                    8  {$Connector.Type = 'MySQL'}
                    9  {$Connector.Type = 'Microsoft SQL Server;MySQL'}
                    10 {$Connector.Type = 'Oracle;MySQL'}
                    11 {$Connector.Type = 'Microsoft SQL Server;Oracle;MySQL'}
                    12 {$Connector.Type = 'IBM DB2;MySQL'}
                    13 {$Connector.Type = 'Microsoft SQL Server;IBM DB2;MySQL'}
                    14 {$Connector.Type = 'Oracle;IBM DB2;MySQL'}
                    15 {$Connector.Type = 'Microsoft SQL Server;Oracle;IBM DB2;MySQL'}
                }
              }
            2 { 
                $Connector | Add-Member -MemberType NoteProperty -Name ConnectorFor -Value 'Virtualization Hosts'
                switch ($Connector.Flags)
                {
                    1 {$Connector.Type = 'VMWare ESX/vSphere'}
                }
              }
            3 { 
                $Connector | Add-Member -MemberType NoteProperty -Name ConnectorFor -Value 'Mail Servers'
                switch ($Connector.Flags)
                {
                    1 {$Connector.Type = 'Microsoft Exchange'}
                }
              }
            4 { 
                $Connector | Add-Member -MemberType NoteProperty -Name ConnectorFor -Value 'Directory Servers'
                switch ($Connector.Flags)
                {
                    1 {$Connector.Type = 'Microsoft Active Directory'}
                }
              }
            5 { 
                $Connector | Add-Member -MemberType NoteProperty -Name ConnectorFor -Value 'Remote Hosts'
                switch ($Connector.Flags)
                {
                    1 {$Connector.Type = 'Secure Shell'}
                }
              }
            6 { 
                $Connector | Add-Member -MemberType NoteProperty -Name ConnectorFor -Value 'Small Business Servers'
                switch ($Connector.Flags)
                {
                    0 {$Connector.Type = ''}
                }
              }
        }

        $Connector
    }
}

Function ConvertTo-LocalTime
{
    Param(
        [DateTime]
        $UTCTime
    )

    $strCurrentTimeZone = (Get-WmiObject win32_timezone).StandardName
    $TZ = [System.TimeZoneInfo]::FindSystemTimeZoneById($strCurrentTimeZone)
    [System.TimeZoneInfo]::ConvertTimeFromUtc($UTCTime, $TZ)
}

function Optimize-RESAMJob
{
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline=$true,
                   Position = 0)]
        $InputObject
    )
    process
    {
        switch ($InputObject.JobInvoker)
        {
            1  {
                If (!$InputObject.JobInvokerInfo.ToString())
                    {
                        $InputObject.JobInvokerInfo = 'User'
                    }
                }
            2  {$InputObject.JobInvokerInfo = 'Recurring schedule'}
            5  {$InputObject.JobInvokerInfo = 'RES Workspace Manager'}
            7  {$InputObject.JobInvokerInfo = 'New Agent'}
            8  {$InputObject.JobInvokerInfo = 'Boot'}
            9  {$InputObject.JobInvokerInfo = 'Project/Runbook'}
        }
        switch ($InputObject.Status)
        {
            -1        {$InputObject.Status = 'On Hold'}
            'Offline' {$InputObject.Status = 'Scheduled'}
            'Online'  {$InputObject.Status = 'Active'}
            2         {$InputObject.Status = 'Aborting'}
            3         {$InputObject.Status = 'Aborted'}
            4         {$InputObject.Status = 'Completed'}
            5         {$InputObject.Status = 'Failed'}
            6         {$InputObject.Status = 'Failed Halted'}
            7         {$InputObject.Status = 'Cancelled'}
            8         {$InputObject.Status = 'Completed with Errors'}
            9         {$InputObject.Status = 'Skipped'}
        }
        Write-Verbose "Converting dates to local time."
        $InputObject.StartDateTime = ConvertTo-LocalTime $InputObject.StartDateTime
        $InputObject.StopDateTime = ConvertTo-LocalTime $InputObject.StopDateTime
        $InputObject
    }
}


#endregion HelperFunctions

<#
.Synopsis
    Connect to RES Automation Manager SQL Database.
.DESCRIPTION
    Sets up a connection to a RES Automation Manager SQL Database. The connection is saved in a
    variable called RESAM_DB_Connection. You can only connect to one database at a time. 
.PARAMETER Datasource
    Name of the SQL datasource to connect to
.PARAMETER DatabaseName
    Name of the RES Automation Manager Database.
.PARAMETER Credential
    Credentials for the connection. Accepts PSCredentials or a username. The user must have 
    read privileges on the database.
.PARAMETER PassThru
    Returns the connection object.
.EXAMPLE
    Connect-RESAMDatabase -DataSource SRV-SQL-01 -DatabaseName RES-AM -Credential RES-AM
    Sets up a connection to database 'RES-AM' on the default SQL Instance on 'SRV-SQL-01'.
    A credential prompt will appear to ask for the password of user 'RES-AM'.
.EXAMPLE
    $Cred = Get-Credential
    C:\PS>Connect-RESAMDatabase -DataSource SRV-SQL-01\RES -DatabaseName RES-AM -Credential $Cred -Passthru
    
      
    Sets up a connection to database 'RES-AM' on the 'RES' Instance on SQL server 'SRV-SQL-01'.
    The connection object will be displayed.
.NOTES
    Author        : Michaja van der Zouwen
    Version       : 1.0
    Creation Date : 25-6-2015
.LINK
   http://itmicah.wordpress.com
#>
function Connect-RESAMDatabase
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true,
                   Position=0)]
        [string]
        $DataSource,
        [Parameter(Mandatory=$true,
                   Position=1)]
        [Alias('DBName')]
        [string]
        $DatabaseName,
        [Parameter(Mandatory=$true,
                   Position=2)]
        $Credential,

        [switch]
        $PassThru
    )

    If ($Credential) {
        Write-Verbose "Processing credentials."
        $Credential = Get-Credential $Credential -Message "Please enter credentials to connect to database '$DatabaseName'"
    }

    Write-Verbose "Connecting to database $DatabaseName on $DataSource..."
    $connectionString = "Server=$dataSource;uid=$($Credential.username);pwd=$($Credential.GetNetworkCredential().password);Database=$DatabaseName;Integrated Security=False;"
    $global:RESAM_DB_Connection = New-Object System.Data.SqlClient.SqlConnection
    $RESAM_DB_Connection.ConnectionString = $connectionString
    $RESAM_DB_Connection.Open()
    Write-Verbose 'Connection established'

    If ($PassThru)
    {
        $RESAM_DB_Connection
    }
}

<#
.Synopsis
    Disconnect from RES Automation Manager Database.
.DESCRIPTION
    Closes the connection to a RES Automation Manager Database.
.PARAMETER Connection
    Name of the SQL datasource to connect to
.EXAMPLE
    Disconnect-RESAMDatabase
    Closes connection to the currently connected database.
.NOTES
    Author        : Michaja van der Zouwen
    Version       : 1.0
    Creation Date : 25-6-2015
.LINK
   http://itmicah.wordpress.com
#>
function Disconnect-RESAMDatabase
{
    Param (
        [System.Data.SqlClient.SqlConnection]
        $Connection
    )
    If ($Connection)
    {
        Write-Verbose ""
        $connection.Close()
    }
    ElseIf ($RESAM_DB_Connection)
    {
        $RESAM_DB_Connection.Close()
    }
    Remove-Variable -Scope Global -Name RESAM_DB_Connection
}

<#
.Synopsis
    Get RES Automation Manager Agent objects.
.DESCRIPTION
    Get RES Automation Manager Agent objects from the RES Automation 
    Manager Database.
.PARAMETER Name
    Name of the Agent.
.PARAMETER GUID
    GUID of the Agent.
.PARAMETER Team
    Team object or guid of the team the agent should be member of
.EXAMPLE
    Get-RESAMAgent -Name PC1234
    Displays information on RES Automation Manager agent PC1234
.EXAMPLE
    Get-RESAMTeam -Name Team1 | Get-RESAMAgent
    Displays information on RES Automation Manager agent that are member
    of team 'Team1'
.NOTES
    Author        : Michaja van der Zouwen
    Version       : 1.0
    Creation Date : 25-6-2015
.LINK
   http://itmicah.wordpress.com
#>
function Get-RESAMAgent
{
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipelineByPropertyName=$true,
                   Position = 0)]
        [string]
        $Name,
        [Parameter(ValueFromPipelineByPropertyName=$true,
                   Position = 1)]
        [Alias('WUIDAgent')]
        [Alias('AgentGUID')]
        [guid]
        $GUID,

        [Parameter(ValueFromPipelineByPropertyName=$true,
                   Position = 2)]
        [Alias('TeamGUID')]
        [ValidateScript({
            If ($_.PSObject.TypeNames -contains 'RES.AutomationManager.Team' -or
             $_ -is [guid])
             {
                $true
             }
             else
             {
                throw "Object type should be 'RES.AutomationManager.Team'."
             }
        })]
        $Team
    )

    process
    {
        if ($Team)
        {
            $Query = "select * from dbo.tblTeamAgents WHERE TeamGUID = '$($Team.GUID)'"
            Invoke-SQLQuery $Query -Type TeamAgent | %{
                $Query = "select * from dbo.tblAgents WHERE WUIDAgent = '$($_.AgentGUID)'"
                Invoke-SQLQuery $Query -Type Agent
            }
            return
        }    
        if ($GUID)
        {
            $Query = "select * from dbo.tblAgents WHERE WUIDAgent = '$GUID'"
        }        
        elseif ($Name)
        {
            $Query = "select * from dbo.tblAgents WHERE strName LIKE '$($Name.replace('*','%'))'"
        }
        else
        {
            $Query = "select * from dbo.tblAgents"
        }
        Invoke-SQLQuery $Query -Type Agent #| Optimize-RESAMAgent
    }
}

<#
.Synopsis
    Get RES Automation Manager Team objects.
.DESCRIPTION
    Get RES Automation Manager Team objects from the RES Automation 
    Manager Database.
.PARAMETER Name
    Name of the Team.
.PARAMETER GUID
    GUID of the Team.
.EXAMPLE
    Get-RESAMTeam -Name Team1
    Displays information on RES Automation Manager team 'Team1'
.NOTES
    Author        : Michaja van der Zouwen
    Version       : 1.0
    Creation Date : 25-6-2015
.LINK
   http://itmicah.wordpress.com
#>
function Get-RESAMTeam
{
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipelineByPropertyName=$true,
                   Position = 0)]
        [string]
        $Name,
        [Parameter(ValueFromPipelineByPropertyName=$true,
                   Position = 1)]
        [Alias('TeamGUID')]
        [guid]
        $GUID,

        [Parameter(ValueFromPipelineByPropertyName=$true,
                   Position = 2)]
        [Alias('WUIDAgent')]
        [ValidateScript({
            If ($_.PSObject.TypeNames -contains 'RES.AutomationManager.Agent' -or
             $_ -is [guid])
             {
                $true
             }
             else
             {
                throw "Object type should be 'RES.AutomationManager.Agent'."
             }
        })]
        $Agent
    )
    process
    {
        if ($Agent)
        {
            If ($Agent -isnot [guid]) 
            {
                $Agent = $Agent.WUIDAgent
            }
            $Query = "select * from dbo.tblTeamAgents WHERE AgentGUID = '$Agent.GUID'"
            Invoke-SQLQuery $Query -Type AgentTeam | %{
                $Query = "select * from dbo.tblTeams WHERE GUID = '$($_.TeamGUID)'"
                Invoke-SQLQuery $Query -Type Team
            }
            return
        }
        If ($GUID)
        {
            $Query = "select * from dbo.tblTeams WHERE GUID = '$($GUID.tostring())'"
        }
        elseif ($Name)
        {
            $Query = "select * from dbo.tblTeams WHERE strName LIKE '$($Name.replace('*','%'))'"
        }
        else
        {
            $Query = "select * from dbo.tblTeams"
        }

        Invoke-SQLQuery $Query -Type Team
    }
}

function Get-RESAMAudit
{
    [CmdletBinding(DefaultParameterSetName='Default')]
    param (
        [Parameter(ValueFromPipelineByPropertyName=$true,
                   ParameterSetName='Default',
                   Position = 0)]
        [Parameter(ValueFromPipelineByPropertyName=$true,
                   ParameterSetName='TimeSpan',
                   Position = 0)]
        [string]
        $Action,

        [Parameter(ValueFromPipelineByPropertyName=$true,
                   ParameterSetName='TimeSpan',
                   Position = 1)]
        [Alias('from')]
        [datetime]
        $StartDate,

        [Parameter(ValueFromPipelineByPropertyName=$true,
                   ParameterSetName='TimeSpan',
                   Position = 2)]
        [Alias('Until')]
        [datetime]
        $EndDate,

        [Parameter(ValueFromPipelineByPropertyName=$true,
                   ParameterSetName='Default',
                   Position = 3)]
        [Parameter(ValueFromPipelineByPropertyName=$true,
                   ParameterSetName='TimeSpan',
                   Position = 3)]
        [string]
        $WindowsAccount,

        [Parameter(ValueFromPipelineByPropertyName=$true,
                   ParameterSetName='Default')]
        [int]
        $Last
    )
    begin
    {
        If ($Last)
        {
            $LastNr = "TOP $Last"
        }
        elseif (!$StartDate -and !$EndDate)
        {
            $LastNr = "TOP 1000"
            Write-Warning "Only the last 1000 audits will be displayed. If more are required use the '-Last' parameter."
        }
    }
    process
    {
        $Query = "select $LastNr strObjectDescription,
strAction,
strActionDescription,
dtmDateTime,
strWindowsAccount,
strWISDOMAccount,
strComputerName,
strComputerDomain,
strComputerIP,
strComputerMAC from dbo.tblAudits"

        $Filter = @()
        If ($Action)
        {
            $Filter += "strAction = '$Action'"
        }
        
        If ($WindowsAccount)
        {
            $Filter += "strWindowsAccount LIKE '$($WindowsAccount.Replace('*','%'))'"
        }

        If ($StartDate -and !$EndDate)
        {
            $EndDate = Get-Date
        }
        If ($EndDate -and !$StartDate)
        {
            $FirstAudit = "select TOP 1 dtmDateTime from dbo.tblAudits order by dtmDateTime ASC"
            $StartDate = Invoke-SQLQuery $FirstAudit | select -ExpandProperty DateTime
        }

        If ($Filter)
        {
            $Filter = $Filter -join ' AND '
            $Query = "$Query WHERE $Filter"
        }
        $Query = "$Query order by dtmDateTime DESC"

        If ($StartDate)
        {
            Invoke-SQLQuery $Query -Type Audit | ?{$_.DateTime -ge $StartDate -and $_.DateTime -le $EndDate}
        }
        else 
        {
            Invoke-SQLQuery $Query -Type Audit
        }
    }
}

function Get-RESAMDispatcher
{
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipelineByPropertyName=$true,
                   Position = 0)]
        [string]
        $Name,
        [Parameter(ValueFromPipelineByPropertyName=$true,
                   Position = 1)]
        [Alias('WUIDDispatcher')]
        [guid]
        $GUID
    )
    process
    {
        If ($GUID)
        {
            $Query = "select * from dbo.tblDispatchers WHERE WUIDDispatcher = '$($GUID.tostring())'"
        }
        elseif ($Name)
        {
            $Query = "select * from dbo.tblDispatchers WHERE strName LIKE '$($Name.replace('*','%'))'"
        }
        else
        {
            $Query = "select * from dbo.tblDispatchers"
        }

        Invoke-SQLQuery $Query -Type Dispatcher
    }
}

function Get-RESAMFolder
{
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipelineByPropertyName=$true,
                   Position = 0)]
        [string]
        $Name,
        [Parameter(ValueFromPipelineByPropertyName=$true,
                   Position = 1)]
        [Alias('FolderGUID')]
        [guid]
        $GUID
    )
    process
    {
        If ($GUID)
        {
            $Query = "select * from dbo.tblFolders WHERE FolderGUID = '$($GUID.tostring())'"
        }
        elseif ($Name)
        {
            $Query = "select * from dbo.tblFolders WHERE strName LIKE '$($Name.replace('*','%'))'"
        }
        else
        {
            $Query = "select * from dbo.tblFolders"
        }

        Invoke-SQLQuery $Query -Type Folder | Optimize-RESAMFolder
    }
}

function Get-RESAMModule
{
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipelineByPropertyName=$true,
                   Position = 0)]
        [Alias('strName')]
        [string]
        $Name,

        [Parameter(ValueFromPipelineByPropertyName=$true,
                   Position = 1)]
        [guid]
        $GUID
    )
    process
    {
        If ($GUID)
        {
            Write-Verbose "Running query based on GUID $GUID."
            $Query = "select * from dbo.tblModules WHERE GUID = '$($GUID.tostring())'"
        }
        Elseif ($Name)
        {
            Write-Verbose "Running query based on name $Name."
            $Query = "select * from dbo.tblModules WHERE strName LIKE '$($Name.replace('*','%'))'"
        }
        else
        {
            $Query = "select * from dbo.tblModules"
        }

        Invoke-SQLQuery $Query -Type Module | Add-RESAMFolderName
    }
}

function Get-RESAMProject
{
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipelineByPropertyName=$true,
                   Position = 0)]
        [Alias('strName')]
        [string]
        $Name,

        [Parameter(ValueFromPipelineByPropertyName=$true,
                   Position = 1)]
        [guid]
        $GUID
    )
    process
    {
        If ($GUID)
        {
            Write-Verbose "Running query based on GUID $GUID."
            $Query = "select * from dbo.tblProjects WHERE GUID = '$($GUID.tostring())'"
        }
        Elseif ($Name)
        {
            Write-Verbose "Running query based on name $Name."
            $Query = "select * from dbo.tblProjects WHERE strName LIKE '$($Name.replace('*','%'))'"
        }
        else
        {
            $Query = "select * from dbo.tblProjects"
        }

        Invoke-SQLQuery $Query -Type Project | Add-RESAMFolderName
    }
}

function Get-RESAMRunBook
{
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipelineByPropertyName=$true,
                   Position = 0)]
        [Alias('strName')]
        [string]
        $Name,

        [Parameter(ValueFromPipelineByPropertyName=$true,
                   Position = 1)]
        [Alias('WUIDAgent')]
        [guid]
        $GUID
    )
    process
    {
        If ($GUID)
        {
            Write-Verbose "Running query based on GUID $GUID."
            $Query = "select * from dbo.tblRunBooks WHERE GUID = '$($GUID.tostring())'"
        }
        Elseif ($Name)
        {
            Write-Verbose "Running query based on name $Name."
            $Query = "select * from dbo.tblRunBooks WHERE strName LIKE '$($Name.replace('*','%'))'"
        }
        else
        {
            $Query = "select * from dbo.tblRunBooks"
        }

        Invoke-SQLQuery $Query -Type RunBook | Add-RESAMFolderName
    }
}

function Get-RESAMResource
{
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipelineByPropertyName=$true,
                   Position = 0)]
        [Alias('strProductName')]
        [string]
        $Name,

        [Parameter(ValueFromPipelineByPropertyName=$true,
                   Position = 1)]
        [guid]
        $GUID
    )
    process
    {
        If ($GUID)
        {
            Write-Verbose "Running query based on GUID $GUID."
            $Query = "select * from dbo.tblResources WHERE GUID = '$($GUID.tostring())'"
        }
        Elseif ($Name)
        {
            Write-Verbose "Running query based on name $Name."
            $Query = "select * from dbo.tblResources WHERE strProductName LIKE '$($Name.replace('*','%'))'"
        }
        else
        {
            $Query = "select * from dbo.tblResources"
        }

        Invoke-SQLQuery $Query -Type Resource | Add-RESAMFolderName
    }
}

function Get-RESAMConnector
{
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipelineByPropertyName=$true,
                   Position = 0)]
        [Alias('strTarget')]
        [string]
        $Target,

        [Parameter(ValueFromPipelineByPropertyName=$true,
                   Position = 1)]
        [guid]
        $GUID,

        [Parameter(ValueFromPipelineByPropertyName=$true,
                   Position = 2)]
        [ValidateSet('Exchange','ActiveDirectory','SecureShell')]
        [string]
        $Type
    )

    begin
    {
        Switch ($Type) {
            'DataBase'       {$TypeNr = 1}
            'Virtualization' {$TypeNr = 2}
            'Mail'           {$TypeNr = 3}
            'Directory'      {$TypeNr = 4}
            'RemoteHosts'    {$TypeNr = 5}
            'SmallBusiness'  {$TypeNr = 6}
        }
    }
    process
    {
        $Filter = @()
        If ($GUID)
        {
            Write-Verbose "Running query based on GUID $GUID."
            $Filter += "GUID = '$($GUID.tostring())'"
        }
        Elseif ($Target)
        {
            Write-Verbose "Running query based on target $Target."
            $Filter += "strTarget LIKE '$($Target.replace('*','%'))'"
        }
        If ($Type)
        {
            $Filter += "lngType = $TypeNr"
        }
        $Query = "select * from dbo.tblConnectors"
        If ($Filter)
        {
            $Query = "$Query WHERE $($Filter -join ' AND ')"
        }

        Invoke-SQLQuery $Query -Type Connector | Optimize-RESAMConnector
    }
}

function Get-RESAMConsole
{
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipelineByPropertyName=$true,
                   Position = 0)]
        [Alias('strName')]
        [string]
        $Name,

        [Parameter(ValueFromPipelineByPropertyName=$true,
                   Position = 1)]
        [guid]
        $GUID
    )
    process
    {
        If ($GUID)
        {
            Write-Verbose "Running query based on GUID $GUID."
            $Query = "select * from dbo.tblConsoles WHERE GUID = '$($GUID.tostring())'"
        }
        Elseif ($Name)
        {
            Write-Verbose "Running query based on name $Name."
            $Query = "select * from dbo.tblConsoles WHERE strName LIKE '$($Name.replace('*','%'))'"
        }
        else
        {
            $Query = "select * from dbo.tblConsoles"
        }

        Invoke-SQLQuery $Query -Type Console | %{
            $Console = $_
            switch ($Console.SystemType)
            {
                1 {$Console.SystemType = 'Client'}
                2 {$Console.SystemType = 'Server'}
            }
            $Console
        }
    }
}

function Get-RESAMDatabaseLevel
{
    [CmdletBinding()]
    param ()

    process
    {
        $Query = "select * from dbo.tblDBLevel"
        
        Invoke-SQLQuery $Query -Type DBlevel | Select -ExpandProperty DBLevel
    }
}

function Get-RESAMJob
{
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipelineByPropertyName=$true,
                   Position = 0)]
        [Alias('strDescription')]
        [string]
        $Description,

        [Parameter(ValueFromPipelineByPropertyName=$true,
                   Position = 1)]
        [Alias('MasterJobGUID')]
        [guid]
        $GUID,

        [Parameter(ValueFromPipelineByPropertyName=$true,
                   Position = 2)]
        [Alias('Agent')]
        [Alias('Team')]
        [string]
        $Who,

        [Parameter(ValueFromPipelineByPropertyName=$true,
                   Position = 3)]
        [guid]
        $ModuleGUID,
        
        [switch]
        $Scheduled,

        [switch]
        $Active,

        [int]
        $Last
    )
    begin
    {
        If ($Last)
        {
            $LastNr = "TOP $Last"
        }
        else
        {
            $LastNr = "TOP 1000"
            Write-Warning "Only the last 1000 jobs will be displayed. If more are required use the '-Last' parameter."
        }
    }
    process
    {
        $Filter = @()
        If ($Scheduled)
        {
            $Filter += "(lngStatus = 0 OR lngStatus = -1)"
            $Filter += "RecurringJobGUID IS NULL"
        }
        If ($ModuleGUID)
        {
            $Filter += "ModuleGUID = '$ModuleGUID'"
        }
        else
        {
            $Filter += "(lngJobInvoker = 1 OR lngJobInvoker = 5)"
        }
        If ($GUID -and !$ModuleGUID)
        {
            Write-Verbose "Running query based on GUID $GUID."
            $Filter += "MasterJobGUID = '$($GUID.tostring())'"
        }
        Elseif ($Description -and !$ModuleGUID)
        {
            Write-Verbose "Running query based on description '$Description'."
            $Filter += "strDescription LIKE '$($Description.replace('*','%'))'"
        }
        If ($Who)
        {
            If ($Who -notmatch '\*')
            {
                $Who = "*$Who*" #Jobs can have multiple agents
            }
            $Filter += "strWho LIKE '$($Who.Replace('*','%'))'"
        }
        If ($Active)
        {
            $Filter += "lngStatus = 1"
        }

        $Query = "select $LastNr * from dbo.tblMasterJob"
        If ($Filter)
        {
            $Filter = $Filter -join ' AND '
            $Query = "$Query WHERE $Filter"
        }

        $Query = "$Query order by dtmStartDateTime DESC"
        Invoke-SQLQuery $Query -Type Job | Optimize-RESAMJob
    }
}

function Get-RESAMQueryResult
{
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipelineByPropertyName=$true,
                   Position = 0)]
        [Alias('strAgent')]
        [string]
        $Agent,

        [Parameter(ValueFromPipelineByPropertyName=$true,
                   Position = 1)]
        [Alias('QueryGUID')]
        [guid]
        $GUID,

        [Parameter(ValueFromPipelineByPropertyName=$true,
                   Position = 2)]
        [guid]
        $MasterJobGUID,

        [int]
        $Last
    )
    begin
    {
        If ($Last)
        {
            $LastNr = "TOP $Last"
        }
        else
        {
            $LastNr = "TOP 1000"
            Write-Warning "Only the last 1000 jobs will be displayed. If more are required use the '-Last' parameter."
        }
    }
    process
    {
        
        $Filter = @()
        If ($MasterJobGUID)
        {
            Write-Verbose "Running query based on MasterJobGUID $MasterJobGUID."
            $Filter += "MasterJobGUID = '$MasterJobGUID'"
        }
        ElseIf ($GUID)
        {
            Write-Verbose "Running query based on GUID $GUID."
            $Filter += "GUID = '$($GUID.tostring())'"
        }
        Elseif ($Agent)
        {
            Write-Verbose "Running query based on Agent $Agent."
            $Filter += "strAgent LIKE '$($Agent.replace('*','%'))'"
        }
        
        $Query = "select * from dbo.tblQueryResults"
        If ($Filter)
        {
            $Filter = $Filter -join ' AND '
            $Query = "$Query WHERE $Filter"
        }
        $Query = "$Query order by dtmDateTime DESC"
        Invoke-SQLQuery $Query -Type QueryResult
    }
}
