#region HelperFunctions

# Invokes a query on the RES AM Database.
function Invoke-SQLQuery
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $Query,

        [Parameter(Mandatory=$False,
                   ValueFromPipeline = $false,
                   Position=1)]
        [string]
        $Type,

        [bool]
        $Full = $true
    )

    Begin
    {
        If (!$RESAM_DB_Connection)
        {
            Throw "No connection to a RES Automation Manager database detected. Run command Connect-RESAMDatabase first."
        }
        elseif ($RESAM_DB_Connection.State -eq 'Closed')
        {
            Write-Verbose 'Connection to the database is closed. Re-opening connection...'
            try
            {
                $RESAM_DB_Connection.Open()
            }
            catch
            {
                Write-Verbose "Error re-opening connection. Removing connection variable."
                Remove-Variable -Scope Global -Name RESAM_DB_Connection
                throw "Unable to re-open connection to the database. Please reconnect using the Connect-RESAMDatabase commandlet. Error is $($_.exception)."
            }
        }
    }
    Process
    {
        $command = $RESAM_DB_Connection.CreateCommand()
        $command.CommandText = $Query

        Write-Verbose "Running SQL query '$query'"
        try
        {
            $result = $command.ExecuteReader()
        }
        catch
        {
            $RESAM_DB_Connection.Close()
            $RESAM_DB_Connection.Open()
            $result = $command.ExecuteReader()
        }
        $CustomTable = new-object "System.Data.DataTable"
        try{
            $CustomTable.Load($result)
        }
        catch{
            $_
        }
        If ($Type)
        {
            $CustomTable | ConvertTo-RESAMObject -Type $Type -Full:$Full
        }
        else
        {
            $CustomTable | ConvertTo-RESAMObject -Full:$Full
        }

        $result.close()
    }
    End
    {
        Write-Verbose "Finished running SQL query."
    }
}

# Converts a SQL query result object to a RES AM object.
function ConvertTo-RESAMObject
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline = $true,
                   Position=0)]
        $InputObject,
        
        [Parameter(Mandatory=$False,
                   ValueFromPipeline = $false,
                   Position=1)]
        [string]
        $Type,

        [bool]
        $Full = $true

    )
    Begin
    {
        Write-Verbose "Creating custom object for output."
    }
    Process
    {
        switch ($InputObject.GetType().Name)
        {
            PSCustomObject {$MemberType = 'NoteProperty'}
            Default        {$MemberType = 'Property'}
        }
        $Properties = $InputObject | Get-Member -MemberType $MemberType |
         select -ExpandProperty Name
        $ht = @{}
        foreach ($Property in $Properties)
        {
            $NewProp = $Property -replace '^(str|lng|ysn|dtm|img)','' -replace '^Obj',''
            $Value = $InputObject.$Property
            If ($Type -ne 'Parameter')
            {
                switch ($Value)
                {
                    yes {$Value = $true}
                    no  {$Value = $False}
                }
            }
            If ($NewProp -eq 'Status')
            {
                switch ($Value)
                {
                    '0' {$Value = 'Offline'}
                    '1' {$Value = 'Online'}
                }
            }
            if ($InputObject.$Property)
            {
                if ($InputObject.$Property.GetType().Name -eq 'Byte[]')
                {
                    If ($Full)
                    {
                        $Value = ConvertFrom-ByteArray $Value
                    }
                    else
                    {
                        $Value = "Use '-Full' parameter for details"
                    }
                }
            }
            If ($Property -eq 'imgWho')
            {
                $NewProp = 'WhoGUID'
            }
            If ($InputObject.$Property -is [datetime])
            {
                If ($Value | Get-Member -Name ToLocalTime)
                {
                    $Value = $Value.ToLocalTime()
                }
                else
                {
                    $Value = ConvertTo-LocalTime $Value
                }
            }
            If ($InputObject.$Property -is [string])
            {
                Try
                {
                    $Value = $Value.substring(0,1).toupper() + $Value.substring(1)
                }
                catch{}
            }
            $NewProp = $NewProp.substring(0,1).toupper() + $NewProp.substring(1)
            try
            {
                $ht.Add($NewProp,$Value)
            }
            catch
            {
                Write-Debug "Error adding property '$NewProp'"
            }
        }
        $Object = New-Object -TypeName psobject -Property $ht
        If ($Type)
        {
            $Object.PSObject.TypeNames.Insert(0,"RES.AutomationManager.$Type")
        }
        $Object
    }
}

# Converts a ByteArray to text characters.
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
    $Text = [System.Text.Encoding]::Unicode.GetString($ByteArray)
    Try {
        [xml]$XML = $Text
        $Object = New-Object -TypeName psobject
        $Properties = $XML | Get-Member -MemberType Property | ?{$_.Name -ne 'xml'}
        foreach ($Property in $Properties)
        {
            $Name = $Property.Name
            Write-Verbose "Adding property $Name to object."
            $Object | Add-Member -MemberType NoteProperty -Name $Name -Value $XML.$Name
        }
    }
    Catch {
        Write-Verbose "Not able to convert array to XML object."
        $Object = Try{
            Write-Verbose "Attempting to cast object as GUID."
            If ($Text -as [guid])
            {
                Write-Verbose "Object is indeed a GUID."
            }
            else
            {
                Write-Verbose "Object is not a GUID."
                Write-Verbose "Casting object as a string value."
                $Text
            }
        }
        catch {
            throw 'Unknown error occurred.'
        }
    }
    $Object
    Write-Verbose "Finished processing array."
}

# Get RES Automation Manager folder objects.
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

# Translates a folder guid to a name and adds the name to an object.
function Add-RESAMFolderName
{
    [CmdletBinding()]
    Param (
        [Parameter(ValueFromPipeline=$true)]
        $InputObject
    )

    process
    {
        If ($InputObject.FolderGuid)
        {
            $Folder = $InputObject.FolderGuid | Get-RESAMFolder
            $InputObject | Add-Member -MemberType NoteProperty -Name FolderName -Value $Folder.Name
        }
        If ($InputObject.ProjectGUID)
        {
            $Query = "select ModuleGUID,lngOrder,ysnEnabled from dbo.tblProjectModules where ProjectGUID = '$($InputObject.ProjectGUID)'"
            $Modules = Invoke-SQLQuery -Query $Query | select Order,ModuleGUID,Enabled | sort order
            $InputObject | Add-Member -MemberType NoteProperty -Name Modules -Value $Modules
        }
        $InputObject
    }
}

# Optimizes an agent object.
function Optimize-RESAMAgent
{
    [CmdletBinding()]
    Param (
        [Parameter(ValueFromPipeline=$true)]
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

    Process
    {
        Write-Verbose "Optimizing agent $($Agent.Name)."
        
        If ($Agent.PrimaryTeamGUID)
        {
            Write-Verbose "Adding PrimaryTeam member."
            $Query = "select strName from dbo.tblTeams WHERE GUID = '$($Agent.PrimaryTeamGUID)'"
            $PrimaryTeam = Invoke-SQLQuery $Query
            $Agent | Add-Member -MemberType NoteProperty -Name PrimaryTeam -Value $PrimaryTeam.Name
        }

        Write-Verbose "Adding Teams member."
        $Query = "select TeamGUID from dbo.tblTeamAgents WHERE AgentGUID = '$($Agent.WUIDAgent)'"
        $Teams = Invoke-SQLQuery $Query | %{
            $Query = "select strName from dbo.tblTeams WHERE GUID = '$($_.TeamGUID)'"
            Invoke-SQLQuery $Query
        }
        $Agent | Add-Member -MemberType NoteProperty -Name Teams -Value $Teams.Name

        Write-Verbose "Checking agent for duplicates."
        $Query = "SELECT strName, COUNT(strName) AS #Duplicates
                  FROM dbo.tblAgents
                  group by strName
                  having COUNT(strName) > 1"
        $Duplicates = Invoke-SQLQuery $Query
        If ($Duplicates.Name -contains $Agent.Name)
        {
            $Agent | Add-Member -MemberType NoteProperty -Name HasDuplicates -Value $True
        }
        else
        {
            $Agent | Add-Member -MemberType NoteProperty -Name HasDuplicates -Value $False
        }
        $Agent
    }
}

# Optimizes a folder object, gives meaning to number values.
function Optimize-RESAMFolder
{
    [CmdletBinding()]
    Param (
        [Parameter(ValueFromPipeline=$true)]
        $Folder
    )
    
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
            $Folder | Add-Member -MemberType NoteProperty -Name ParentFolderName -Value $ParentFolder.Name.trim()
        }
        $Folder
    }
}

# Optimizes a connector object, gives meaning to number values.
function Optimize-RESAMConnector
{
    [CmdletBinding()]
    Param (
        [Parameter(ValueFromPipeline=$true)]
        $Connector
    )
    
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
            7 { 
                $Connector | Add-Member -MemberType NoteProperty -Name ConnectorFor -Value 'Web Service Hosts'
                switch ($Connector.Flags)
                {
                    0 {$Connector.Type = 'Web Service'}
                }
              }
        }

        $Connector
    }
}

# Converts UTC to local time.
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

# Optimizes the job object, gives meaning to number values.
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
        Write-Verbose "Processing Job Invoker..."
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
            9  {$InputObject.JobInvokerInfo = 'Runbook'}
        }
        Write-Verbose "Job Invoker is '$($InputObject.JobInvokerInfo)'."

        Write-Verbose "Processing status..."
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
        Write-Verbose "Status is '$($InputObject.Status)'"
        #Write-Verbose "Converting dates to local time."
        #$InputObject.StartDateTime = ConvertTo-LocalTime $InputObject.StartDateTime
        #$InputObject.StopDateTime = ConvertTo-LocalTime $InputObject.StopDateTime
        $InputObject
    }
}

# Invokes a method using the REST Api
function Invoke-RESAMRestMethod 
{
    [CmdletBinding()]
	param(
        [Parameter(Mandatory=$True)]
	    [string]
        $Uri,

        [Parameter(Mandatory=$True)]
        [ValidateSet("GET","PUT","POST")] 
	    [string]
        $Method,

        [Parameter(Mandatory=$True)]
        $Credential,
	    
        [System.Object]
        $Body
	)

	begin
    {
        If ($Credential) {
            Write-Verbose "Processing credentials."
            $Message = "Please enter RES Automation Manager credentials to connect to the Dispatcher."
            switch ($Credential.GetType().Name)
            {
                'PSCredential' {}
                'String' {$Credential = Get-Credential $Credential -Message $Message}
            }
        }
    }
	process {
		$Splat = @{
			Uri = $Uri
			Credential = $Credential
			Method = $Method
			ContentType = "application/json"
			SessionVariable = "Script:ResAMSession"
		}
		if($Body){
			$Splat.Add("Body",$Body)
		}
		
		Invoke-RestMethod @Splat
	}
}

# Retreives only used parameters using the webapi
function Get-RESAMInputParameter
{
    [CmdletBinding()]
	param(
        [Parameter(Mandatory=$True)]
		[String]
        $Dispatcher,

        [Parameter(Mandatory=$True)]
	    $Credential,

        [Parameter(Mandatory=$True,
                   ValueFromPipeline=$True)]
		[PSObject]
        $What,

        [Switch]
        $Raw = $false
	)
	begin
    {
        If ($Credential) {
            Write-Verbose "Processing credentials."
            $Message = "Please enter RES Automation Manager credentials to connect to the Dispatcher."
            switch ($Credential.GetType().Name)
            {
                'PSCredential' {}
                'String' {$Credential = Get-Credential $Credential -Message $Message}
            }
        }
    }
	process {
		$endPoint = "Dispatcher/SchedulingService/what"
        $Type = $What.PSObject.TypeNames | ?{$_ -like 'RES*'}
		$uri = "http://$Dispatcher/$($endPoint)/$($Type.Split('.')[-1])s/$($What.GUID)/inputparameters"
		$pREST = @{
			Uri = $Uri
			Method = "GET"
			Credential = $Credential
		}
#
# Only parameters that are actually used in any of the module tasks will be returned !
#
		$result = Invoke-RESAMRestMethod @pREST
        if($Raw){$result}
        else{$result.JobParameters}
	}
}

#endregion HelperFunctions

#.ExternalHelp RESAM.Help.xml
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

        [Parameter(Mandatory=$false,
                   Position=2)]
        $Credential,

        [switch]
        $PassThru
    )

    If ($Credential) {
        Write-Verbose "Processing credentials."
        $Message = "Please enter credentials to connect to database '$DatabaseName'."
        switch ($Credential.GetType().Name)
        {
            'PSCredential' {}
            'String' {$Credential = Get-Credential $Credential -Message $Message}
        }
    }

    Write-Verbose "Connecting to database $DatabaseName on $DataSource..."
    $connectionString = "Server=$dataSource;Database=$DatabaseName"
    If ($Credential)
    {
        $connectionString = "$connectionString;uid=$($Credential.username);pwd=$($Credential.GetNetworkCredential().password);Integrated Security=False;"
    }
    else
    {
        $connectionString = "$connectionString;Integrated Security=sspi;"
    }
    $global:RESAM_DB_Connection = New-Object System.Data.SqlClient.SqlConnection
    $RESAM_DB_Connection.ConnectionString = $connectionString
    $RESAM_DB_Connection.Open()
    Write-Verbose 'Connection established'

    If ($PassThru)
    {
        $RESAM_DB_Connection
    }
}

#.ExternalHelp RESAM.Help.xml
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

#.ExternalHelp RESAM.Help.xml
function Get-RESAMAgent
{
    [CmdletBinding(DefaultParameterSetName='Default')]

    param (
        [Parameter(ValueFromPipelineByPropertyName=$true,
                   ParameterSetName='Default',
                   Position = 0)]
        [Alias('Agent')]
        [string]
        $Name,

        [Parameter(ValueFromPipelineByPropertyName=$true,
                   ParameterSetName='Default',
                   Position = 1)]
        [Alias('Who')]
        [Alias('WUIDAgent')]
        [Alias('AgentGUID')]
        [guid]
        $GUID,

        [Parameter(ValueFromPipelineByPropertyName=$true,
                   ParameterSetName='Default',
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
        $Team,

        [Parameter(ParameterSetName='Default')]
        [switch]
        $Full,

        [Parameter(ParameterSetName='Default')]
        [ValidateSet('Online','Offline')]
        [string]
        $Status,

        [Parameter(ParameterSetName='Duplicates')]
        [switch]
        $HasDuplicates
    )

    process
    {
        if ($HasDuplicates)
        {
            Write-Verbose "Checking agent for duplicates."
            $Query = "SELECT strName, COUNT(strName) AS #Duplicates
                      FROM dbo.tblAgents
                      group by strName
                      having COUNT(strName) > 1"
            Invoke-SQLQuery $Query -Type Duplicate
            return
        }
        if ($Team)
        {
            $Query = "select * from dbo.tblTeamAgents WHERE TeamGUID = '$($Team.GUID)'"
            Invoke-SQLQuery $Query -Type TeamAgent | %{
                $Query = "select * from dbo.tblAgents WHERE WUIDAgent = '$($_.AgentGUID)'"
                Invoke-SQLQuery $Query -Type Agent -Full:$Full | Optimize-RESAMAgent
            }
            return
        }
        $Filter = @()
        if ($GUID)
        {
            $filter += "WUIDAgent = '$GUID'"
        }        
        elseif ($Name)
        {
            $filter += "strName LIKE '$($Name.replace('*','%'))'"
        }
        Switch ($Status)
        {
            Online {$Filter += "lngStatus = 1"}
            Offline {$Filter += "lngStatus = 0"}
        }
        $Query = "select * from dbo.tblAgents"
        If ($Filter)
        {
            $Filter = $Filter -join ' AND '
            $Query = "$Query WHERE $Filter"
        }
        Invoke-SQLQuery $Query -Type Agent -Full:$Full | Optimize-RESAMAgent
    }
}

#.ExternalHelp RESAM.Help.xml
function Remove-RESAMAgent
{
    [CmdletBinding(SupportsShouldProcess=$true,
                  ConfirmImpact='High')]
    Param
    (
        # Name of agent to remove
        [Parameter(Mandatory=$true, 
                   ValueFromPipelineByPropertyName=$true, 
                   Position=0)]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Alias('Agent')]
        [string]
        $Name,

        # GUID of agent to remove
        [Parameter(Mandatory=$false, 
                   ValueFromPipelineByPropertyName=$true, 
                   Position=1)]
        [Alias('WUIDAgent')]
        [Alias('AgentGUID')]
        [guid]
        $GUID,

        # How to handle duplicate agents
        [ValidateSet('Abort','Skip','RemoveAll','KeepLatest')]
        [string]
        $DuplicatesPreference,

        # Remove job history as well
        [switch]
        $IncludeJobHistory
    )
    Begin
    {
        $Agents = @()
    }
    Process
    {
        Write-Verbose "Checking agent existence..."
        If ($GUID)
        {
            $Agents += (Get-RESAMAgent -GUID $GUID)
        }
        elseif ($Name)
        {
            $Agents += (Get-RESAMAgent -Name $Name)
        }
    }
    End
    {
        If (!$Agents)
        {
            throw 'Agent(s) not found in database.'
        }
        else
        {
            Write-Verbose "Found $($Agents.Count) Agent(s) in the database."
        }
        Write-Verbose 'Checking for duplicate agent names...'
        $GroupAgents = $Agents | group Name
        foreach ($AgentName in $GroupAgents)
        {
            If ($AgentName.Count -gt 1)
            {
                Write-Verbose "Multiple agents detected named '$($AgentName.Name)'!"

                If ($DuplicatesPreference)
                {
                    Write-Verbose "Using '$DuplicatesPreference' method to handle duplicates."
                    switch ($DuplicatesPreference)
                    {
                        'Abort'      {$Proceed = 0}
                        'Skip'       {$Proceed = 1}
                        'RemoveAll'  {$Proceed = 2}
                        'KeepLatest' {$Proceed = 3}
                    }
                }
                elseif ($StoredSetting)
                {
                    Write-Verbose "Using previous method to handle duplicates."
                    $Proceed = $StoredSetting
                }
                elseif($ConfirmPreference -eq 'None')
                {
                    Write-Verbose ""
                    $Proceed = 2
                }
                else
                {
                    $Title = "Duplicate agent names detected."
                    $Message = "There are $($AgentName.Count) agents named '$($AgentName.Name)'!`nHow would you like to proceed?"
                
                    $Abort = New-Object System.Management.Automation.Host.ChoiceDescription "&Abort",
                        "Abort all operations."
                    $Skip = New-Object System.Management.Automation.Host.ChoiceDescription "&Skip agents",
                        "Skip these agents and continue."
                    $RemoveAll = New-Object System.Management.Automation.Host.ChoiceDescription "&Remove all",
                        "Remove all agents."
                    $KeepLatest = New-Object System.Management.Automation.Host.ChoiceDescription "&Keep latest",
                        "Remove all agents except the latest one."
                    $Options = [System.Management.Automation.Host.ChoiceDescription[]]($Abort,$Skip,$RemoveAll,$KeepLatest)
                
                    $Proceed = $Host.ui.PromptForChoice($Title, $Message, $Options, 0)
                }

                switch ($Proceed)
                {
                    0   {throw "Operation cancelled!"}
                    1   {
                            Write-Verbose "Removing agents named '$($AgentName.Name)' from array..."
                            $Agents = $Agents | ?{$_.Name -ne $AgentName.Name}
                        }
                    3   {
                            Write-Verbose 'Removing latest deployed agent from array...'
                            $SkipAgent = $AgentName.Group | sort DeployedOn | select -Last 1
                            $Agents = $Agents | ?{$_ -ne $SkipAgent}
                        }
                }

                If (!$DuplicatesPreference -and 
                    !$StoredSetting -and 
                    $AgentName -ne $GroupAgents[-1])
                {
                    $Title = "Duplicate agent names detected."
                    $Message = "Would you like to apply this setting to all duplicate agent names?"
                
                    $Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes",
                        "Remember this setting."
                
                    $No = New-Object System.Management.Automation.Host.ChoiceDescription "&No",
                        "Ask again later."
                    $Options = [System.Management.Automation.Host.ChoiceDescription[]]($Yes,$No)

                    $Remember = $Host.ui.PromptForChoice($Title, $Message, $Options, 0)
                    switch ($Remember)
                    {
                        0 {$StoredSetting = $Proceed}
                    }
                }
            }
        }
        If ($Agents)
        {
            $WUIDs = @()
            foreach ($Agent in $Agents)
            {
                If ($Agent.Status -eq 'Online')
                {
                    Write-Error "Agent '$($Agent.Name)' is currently online. Only offline agents can be removed from the database."

                }
                else
                {
                    $WUIDs += "WUIDAgent = '$($Agent.WUIDAgent)'"
                }
            }
            $Filter = $WUIDs -join ' OR '
            $Query = "DELETE FROM dbo.tblAgents WHERE $Filter"

            if ($pscmdlet.ShouldProcess("$($WUIDs.Count) RES AM agent(s)", "Remove from database"))
            {
                foreach ($Agent in $Agents)
                {
                    Write-Verbose "Removing agent '$($Agent.name)' from the database..."
                }
                Invoke-SQLQuery $Query -ErrorAction Stop
            }
            If ($IncludeJobHistory)
            {
                If ((Get-RESAMDatabaseLevel) -ge 61)
                {
                    $SQLTable = 'dbo.tblJobsHistory'
                }
                else
                {
                    $SQLTable = 'dbo.tblJobs'
                }
                $Filter = $Filter -replace 'WUIDAgent','AgentGUID'
                $Query = "SELECT strAgent FROM $SQLTable WHERE $Filter"
                $Jobs = Invoke-SQLQuery $Query
                if ($pscmdlet.ShouldProcess("$($Jobs.Count) RES AM jobs", "Remove history from database"))
                {
                    $Query = "DELETE FROM $SQLTable WHERE $Filter"
                    Invoke-SQLQuery $Query
                }
            }
        }
        else
        {
            Write-Verbose 'No agents left to remove.'
        }
        Write-Verbose 'Finished'
    }
}

#.ExternalHelp RESAM.Help.xml
function Get-RESAMTeam
{
    [CmdletBinding(DefaultParameterSetName='Default')]

    param (
        [Parameter(Position = 0,
                   ParameterSetName = 'Default')]
        [string]
        $Name,

        [Parameter(ValueFromPipelineByPropertyName=$true,
                   Position = 1,
                   ParameterSetName = 'Default')]
        [Alias('TeamGUID')]
        [guid]
        $GUID,

        [Parameter(ValueFromPipelineByPropertyName=$true,
                   DontShow=$true,
                   ParameterSetName = 'ByAgent')]
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
        $Agent,

        [switch]
        $Full
    )
    process
    {
        if ($Agent)
        {
            If ($Agent -isnot [guid]) 
            {
                $Agent = $Agent.WUIDAgent
            }
            $Query = "select TeamGUID from dbo.tblTeamAgents WHERE AgentGUID = '$($Agent.GUID)'"
            $AgentTeams = Invoke-SQLQuery $Query -Type AgentTeam
            $Filter = @()
            foreach ($GUID in $AgentTeams.TeamGUID)
            {
                $Filter += "GUID = '$GUID'"
            }
            $Filter = $Filter -join ' OR '
            $Query = "select * from dbo.tblTeams WHERE $Filter"
        }
        elseif ($GUID)
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

        Invoke-SQLQuery $Query -Type Team -Full:$Full
    }
}

#.ExternalHelp RESAM.Help.xml
function Get-RESAMAudit
{
    [CmdletBinding(DefaultParameterSetName='Default')]

    param (
        #Query specific action only
        [Parameter(ValueFromPipelineByPropertyName=$true,
                   ParameterSetName='Default',
                   Position = 0)]
        [Parameter(ValueFromPipelineByPropertyName=$true,
                   ParameterSetName='TimeSpan',
                   Position = 0)]
        [ValidateSet('Add','Delete','Edit','Edit (details)','Other','Primary Team changed','Register','Sign in','Sign out')]
        [string]
        $Action,

        #From what time/date
        [Parameter(ValueFromPipelineByPropertyName=$true,
                   ParameterSetName='TimeSpan',
                   Position = 1)]
        [Alias('From')]
        [Alias('Start')]
        [datetime]
        $StartDate,

        #To time/date
        [Parameter(ValueFromPipelineByPropertyName=$true,
                   ParameterSetName='TimeSpan',
                   Position = 2)]
        [Alias('Until')]
        [Alias('End')]
        [datetime]
        $EndDate,

        #Audits from a single account
        [Parameter(ValueFromPipelineByPropertyName=$true,
                   ParameterSetName='Default',
                   Position = 3)]
        [Parameter(ValueFromPipelineByPropertyName=$true,
                   ParameterSetName='TimeSpan',
                   Position = 3)]
        [string]
        $WindowsAccount,

        #Limit result
        [Parameter(ValueFromPipelineByPropertyName=$true,
                   ParameterSetName='Default')]
        [int]
        $Last = 1000
    )
    begin
    {
        If (!$PSBoundParameters.ContainsKey('Last') -and !$StartDate -and !$EndDate)
        {
            Write-Warning "Only the last 1000 jobs will be displayed. If more are required use the '-Last' parameter."
        }
        $LastNr = "TOP $Last"
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

#.ExternalHelp RESAM.Help.xml
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
        $GUID,

        [ValidateSet('Online','Offline')]
        [string]
        $Status,

        [switch]
        $Full
    )
    process
    {
        Switch ($Status)
        {
            Online {$Filter = "lngStatus = 1"}
            Offline {$Filter = "lngStatus = 0"}
        }
        If ($GUID)
        {
            $Query = "select * from dbo.tblDispatchers WHERE WUIDDispatcher = '$($GUID.tostring())'"
            If ($Filter)
            {
                $Query = "$Query AND $Filter"
            }
        }
        elseif ($Name)
        {
            $Query = "select * from dbo.tblDispatchers WHERE strName LIKE '$($Name.replace('*','%'))'"
            If ($Filter)
            {
                $Query = "$Query AND $Filter"
            }
        }
        else
        {
            $Query = "select * from dbo.tblDispatchers"
            If ($Filter)
            {
                $Query = "$Query WHERE $Filter"
            }
        }
        
        Invoke-SQLQuery $Query -Type Dispatcher -Full:$Full
    }
}

#.ExternalHelp RESAM.Help.xml
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
        [Alias('ModuleGUID')]
        [guid]
        $GUID,

        [switch]
        $Full
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

        Invoke-SQLQuery $Query -Type Module -Full:$Full | Add-RESAMFolderName| %{
            If ($Full)
            {
                If ($_.Tasks.tasks.task)
		        {
			        $Tasks = $_.Tasks.tasks.task | ?{!$_.Hidden}
			        $ModuleTasks = @()
                    foreach ($Task in $Tasks)
			        {
				        $ModuleTask = $Task.properties | ConvertTo-RESAMObject -Type Task
				        If ($Task.Settings)
				        {
					        $Settings = $Task.settings | ConvertTo-RESAMObject -Type TaskSetting
					        $ModuleTask | Add-Member -MemberType NoteProperty -Name Settings -Value $Settings
				        }
                        $ModuleTasks += $ModuleTask
			        }
                    $_ | Add-Member -MemberType NoteProperty -Name ModuleTasks -Value $ModuleTasks -PassThru
                }
		    }
		    else
		    {
			    $_ | Add-Member -MemberType NoteProperty -Name ModuleTasks -Value "Use '-Full' parameter for details" -PassThru
		    }
        }
    }
}

#.ExternalHelp RESAM.Help.xml
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
        $GUID,

        [switch]
        $Full
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

        Invoke-SQLQuery $Query -Type Project -Full:$Full | Add-RESAMFolderName
    }
}

#.ExternalHelp RESAM.Help.xml
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
        $GUID,

        [switch]
        $Full
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

        Invoke-SQLQuery $Query -Type RunBook -Full:$Full | Add-RESAMFolderName
    }
}

#.ExternalHelp RESAM.Help.xml
function Get-RESAMResource
{
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipelineByPropertyName=$true,
                   Position = 0)]
        [Alias('strFileName')]
        [string]
        $Name,

        [Parameter(ValueFromPipelineByPropertyName=$true,
                   Position = 1)]
        [guid]
        $GUID,

        [Switch]
        $Full
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
            $Query = "select * from dbo.tblResources WHERE strFileName LIKE '$($Name.replace('*','%'))'"
        }
        else
        {
            $Query = "select * from dbo.tblResources"
        }

        Invoke-SQLQuery $Query -Type Resource -Full:$Full | Add-RESAMFolderName
    }
}

#.ExternalHelp RESAM.Help.xml
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
        [ValidateSet('DataBase','Virtualization','Mail','Directory','RemoteHosts','SmallBusiness')]
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

#.ExternalHelp RESAM.Help.xml
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
        $GUID,

        [switch]
        $Full
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

        Invoke-SQLQuery $Query -Type Console -Full:$Full | %{
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

#.ExternalHelp RESAM.Help.xml
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

#.ExternalHelp RESAM.Help.xml
function Get-RESAMMasterJob
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
        [guid]
        $MasterJobGUID,

        [Parameter(ValueFromPipelineByPropertyName=$true,
                   Position = 2)]
        [Alias('AgentName')]
        [Alias('TeamName')]
        [string]
        $Who,

        [Parameter(ValueFromPipelineByPropertyName=$true,
                   Position = 3)]
        [Alias('GUID')]
        [guid]
        $ModuleGUID,

        [Parameter(ValueFromPipelineByPropertyName=$true,
                   Position = 4)]
        [guid]
        $RunBookJobGUID,
        
        [Parameter(ValueFromPipelineByPropertyName=$false)]
        [ValidateSet('On Hold',
                    'Scheduled',
                    'Active',
                    'Aborting',
                    'Aborted',
                    'Completed',
                    'Failed',
                    'Failed Halted',
                    'Cancelled',
                    'Completed with Errors',
                    'Skipped')]
        [string]
        $Status,

        [psobject]
        $StartDate,

        [switch]
        $InvokedByRunbook,

        [int]
        $Last = 1000,

        [switch]
        $Full
    )
    begin
    {
        If (!$PSBoundParameters.ContainsKey('Last') -and !$MasterJobGUID)
        {
            Write-Warning "Only the last 1000 jobs will be displayed. If more are required use the '-Last' parameter."
        }
        If ($Last -eq 0)
        {
            $LastNr = ""
        }
        else
        {
            $LastNr = "TOP $Last"
        }
    }
    process
    {
        $Filter = @()
        
        If ($RunBookJobGUID)
        {
            Write-Verbose "Running query based on MasterJobGUID '$RunBookJobGUID'."
            $Filter += "MasterJobGUID = '$($RunBookJobGUID.tostring())'"
        }
        elseIf ($ModuleGUID -and !$MasterJobGUID)
        {
            $Filter += "ModuleGUID = '$ModuleGUID'"
        }
        if ($InvokedByRunbook)
        {
            $Filter += "lngJobInvoker = 9"
        }
        elseif (!$ModuleGUID)
        {
            $Filter += "lngJobInvoker <> 9"
        }
        If ($MasterJobGUID -and !$ModuleGUID)
        {
            Write-Verbose "Running query based on GUID $GUID."
            $Filter += "MasterJobGUID = '$($MasterJobGUID.tostring())'"
        }
        If ($Description -and !$ModuleGUID)
        {
            Write-Verbose "Running query based on description '$Description'."
            $Filter += "strDescription LIKE '$($Description.replace('*','%'))'"
        }
        If ($Who -and !$RunBookJobGUID)
        {
            If ($Who -notmatch '\*')
            {
                $Who = "*$Who*" #Jobs can have multiple agents
            }
            $Filter += "strWho LIKE '$($Who.Replace('*','%'))'"
        }
        If ($Status)
        {
            Write-Verbose "Filtering jobs on status '$Status'..."
            switch ($Status)
            {
                'On Hold'               {$StatusNr = -1}
                'Scheduled'             {$StatusNr = 0}
                'Active'                {$StatusNr = 1}
                'Aborting'              {$StatusNr = 2}
                'Aborted'               {$StatusNr = 3}
                'Completed'             {$StatusNr = 4}
                'Failed'                {$StatusNr = 5}
                'Failed Halted'         {$StatusNr = 6}
                'Cancelled'             {$StatusNr = 7}
                'Completed with Errors' {$StatusNr = 8}
                'Skipped'               {$StatusNr = 9}
            }
            $Filter += "lngStatus = $StatusNr"
        }
        else
        {
            Write-Verbose 'No status specified. Skipping active masterjobs...'
            foreach ($StatusNr in -1..2)
            {
                $Filter += "lngStatus <> $StatusNr"
            }
        }
        If ($StartDate)
        {
            $uDate = (Get-Date $StartDate -ea 1).ToUniversalTime()
            $Date1 = Get-Date $uDate.AddSeconds(-1) -Format 'yyyy-MM-dd HH:mm:ss'
            $Date2 = Get-Date $uDate.AddSeconds(1) -Format 'yyyy-MM-dd HH:mm:ss'
            $Filter += "dtmStartDateTime BETWEEN '$Date1' AND '$Date2'"
        }
        $Query = "select $LastNr * from dbo.tblMasterJob"
        If ($Filter)
        {
            $Filter = $Filter -join ' AND '
            $Query = "$Query WHERE $Filter"
        }

        $Query = "$Query order by dtmStartDateTime DESC"
        Invoke-SQLQuery $Query -Type MasterJob -Full:$Full | Optimize-RESAMJob
        If ((Get-RESAMDatabaseLevel) -ge 61)
        {
            Invoke-SQLQuery $Query.Replace('tblMasterJob','tblMasterJobHistory') -Type MasterJob -Full:$Full | Optimize-RESAMJob
        }
    }
}

#.ExternalHelp RESAM.Help.xml
function Get-RESAMJob
{
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipelineByPropertyName=$true,
                   Position = 0)]
        [Alias('WUIDAgent')]
        [Alias('AgentGUID')]
        $Agent,

        [Parameter(ValueFromPipelineByPropertyName=$true,
                   Position = 1)]
        [guid]
        $MasterJobGUID,

        [Parameter(ValueFromPipelineByPropertyName=$true,
                   Position = 3)]
        [guid]
        $JobGUID,
        
        [Parameter(ValueFromPipelineByPropertyName=$false,
                   Position = 3)]
        [ValidateSet('On Hold',
                    'Scheduled',
                    'Active',
                    'Aborting',
                    'Aborted',
                    'Completed',
                    'Failed',
                    'Failed Halted',
                    'Cancelled',
                    'Completed with Errors',
                    'Skipped')]
        [string]
        $Status,

        [int]
        $Last = 1000,

        [switch]
        $Full
    )
    begin
    {
        If (!$PSBoundParameters.ContainsKey('Last'))
        {
            Write-Warning "Only the last 1000 jobs will be displayed. If more are required use the '-Last' parameter."
        }
        If ($Last -eq 0)
        {
            $LastNr = ""
        }
        else
        {
            $LastNr = "TOP $Last"
        }
    }
    process
    {
        $Filter = @()
        If ($Agent)
        {
            If ($Agent -is [guid])
            {
                $Filter += "AgentGUID = '$Agent'"
            }
            else 
            {
                $Filter += "strAgent = '$Agent'"
            }
        }
        If ($JobGUID)
        {
            $Filter += "JobGUID = '$JobGUID'"
        }
        If ($MasterJobGUID)
        {
            Write-Verbose "Running query based on MasterJobGUID $MasterJobGUID."
            $Filter += "MasterJobGUID = '$MasterJobGUID'"
        }
        If ($Status)
        {
            Write-Verbose "Filtering jobs on status '$Status'..."
            switch ($Status)
            {
                'On Hold'               {$StatusNr = -1}
                'Scheduled'             {$StatusNr = 0}
                'Active'                {$StatusNr = 1}
                'Aborting'              {$StatusNr = 2}
                'Aborted'               {$StatusNr = 3}
                'Completed'             {$StatusNr = 4}
                'Failed'                {$StatusNr = 5}
                'Failed Halted'         {$StatusNr = 6}
                'Cancelled'             {$StatusNr = 7}
                'Completed with Errors' {$StatusNr = 8}
                'Skipped'               {$StatusNr = 9}
            }
            $Filter += "lngStatus = $StatusNr"
        }
        else
        {
            Write-Verbose 'No status specified. Skipping active jobs...'
            foreach ($StatusNr in -1..2)
            {
                $Filter += "lngStatus <> $StatusNr"
            }
        }

        $Query = "select $LastNr * from dbo.tblJobs"
        If ($Filter)
        {
            $Filter = $Filter -join ' AND '
            $Query = "$Query WHERE $Filter"
        }

        $Query = "$Query order by dtmStartDateTime DESC"
        Invoke-SQLQuery $Query -Type Job -Full:$Full | Optimize-RESAMJob
        If ((Get-RESAMDatabaseLevel) -ge 61)
        {
            Invoke-SQLQuery $Query.Replace('tblJobs','tblJobsHistory') -Type Job -Full:$Full | Optimize-RESAMJob
        }
    }
}

#.ExternalHelp RESAM.Help.xml
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
        $Last = 1000
    )
    begin
    {
        If (!$PSBoundParameters.ContainsKey('Last'))
        {
            Write-Warning "Only the last 1000 jobs will be displayed. If more are required use the '-Last' parameter."
        }
        $LastNr = "TOP $Last"
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

#NOT READY!!
function Get-RESAMLog
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True,
                   ValueFromPipelineByPropertyName=$true,
                   ParameterSetName='Job',
                   Position = 0)]
        [Alias('strAgent')]
        [guid]
        $JobGUID,

        [Parameter(Mandatory=$True,
                   ValueFromPipelineByPropertyName=$true,
                   ParameterSetName='Task',
                   Position = 0)]
        [Alias('QueryGUID')]
        [guid]
        $TaskGUID
    )
    begin
    {
    }
    process
    {
        If ($JobGUID)
        {
            Write-Verbose "Running query based on JobGUID $JobGUID."
            $Query = "select * from dbo.tblLogs WHERE JobGUID = '$JobGUID'"
        }
        ElseIf ($TaskGUID)
        {
            Write-Verbose "Running query based on TaskGUID $TaskGUID."
            $Query = "select * from dbo.tblLogs WHERE TaskGUID = '$TaskGUID'"
        }
        $Logs = Invoke-SQLQuery $Query
        foreach ($Log in $Logs)
        {
            $FileQuery = "select * from dbo.tblFiles WHERE GUID = '$($Log.FileGUID)'"
            Invoke-SQLQuery $FileQuery -Type LogFile
        }
    }
}

#.ExternalHelp RESAM.Help.xml
function New-RESAMJob
{
    [CmdletBinding()]
	param(
        [Parameter(Mandatory=$True)]
		[String]
        $Dispatcher,

        [Parameter(Mandatory=$True)]
		$Credential,

		[String]
        $Description,

		[Parameter(ValueFromPipeline=$true)]
        $Who,

        [Parameter(ParameterSetName='Module')]
        $Module,

        [Parameter(ParameterSetName='Project')]
        $Project,

        [Parameter(ParameterSetName='RunBook')]
        $RunBook,

		[DateTime]
        $Start,

        [Switch]
        $LocalTime = $true,

		[Switch]
        $UseWOL,

        [Switch]
        $UseDefaults,

		[HashTable]
        $Parameters
	)

    begin
    {
        If ($UseDefaults -and $Parameters)
        {
            throw "Illegal operation! You cannot use '-UseDefaults' and '-Parameters' together in a single command."
        }
        If ($Credential) {
            Write-Verbose "Processing credentials."
            $Message = "Please enter RES Automation Manager credentials to connect to the Dispatcher."
            switch ($Credential.GetType().Name)
            {
                'PSCredential' {}
                'String' {$Credential = Get-Credential $Credential -Message $Message}
            }
        }
        If ($Start)
        {
            $Immediate = $false
        }
        else
        {
            $Immediate = $True
            $Start = Get-Date
        }
        If ($Module)
        {
            IF ($Module.PSObject.TypeNames -contains 'RES.AutomationManager.Module')
            {
                $Task = $Module
            }
            elseIf ($Module.GetType().Name -eq 'String')
            {
                $Task = Get-RESAMModule $Module
            }
            else
            {
                Throw 'Incorrect object type for Module parameter.'
            }
            If (!$Who)
            {
                $Who = Read-Host -Prompt 'Please provide an agent for this job'
                If (!$Who)
                {
                    throw 'Unable to schedule module without an agent.'
                }
            }
            $Type = 0
        }
        If ($Project)
        {
            IF ($Project.PSObject.TypeNames -contains 'RES.AutomationManager.Project')
            {
                $Task = $Project
            }
            elseIf ($Project.GetType().Name -eq 'String')
            {
                $Task = Get-RESAMProject $Project
            }
            else
            {
                Throw 'Incorrect object type for Project parameter.'
            }
            If (!$Who)
            {
                $Who = Read-Host -Prompt 'Please provide an agent for this job'
                If (!$Who)
                {
                    throw 'Unable to schedule project without an agent.'
                }
            }
            $Type = 1
        }
        If ($RunBook)
        {
            IF ($RunBook.PSObject.TypeNames -contains 'RES.AutomationManager.RunBook')
            {
                $Task = $RunBook
            }
            elseIf ($RunBook.GetType().Name -eq 'String')
            {
                $Task = Get-RESAMRunBook $RunBook -Full
            }
            else
            {
                Throw 'Incorrect object type for RunBook parameter.'
            }
            If (!$Who -and
                $Task.Properties.properties.jobs.job.properties.whoname -contains '' -and
                $Task.Properties.properties.jobs.job.properties.use_runbookparam -eq 'no')
            {
                $Who = Read-Host -Prompt 'Please provide an agent for this job'
                If (!$Who)
                {
                    throw 'Unable to schedule runbook without an agent.'
                }
            }
            $Type = 2
        }
        If (!$Description)
        {
            $Description = $Task.Name
        }
        Write-Verbose "Getting input parameter object for '$Task'."
        $InputParameters = Get-RESAMInputParameter -Dispatcher $Dispatcher -Credential $Credential -What $Task -Raw

        If ($InputParameters)
        {
            Write-Verbose 'Required input parameters found.'
            If ($Parameters)
            {
                Write-Verbose 'Setting new parameter values...'
                foreach ($jobParam in $InputParameters.JobParameters)
                {
                    $Parameters.GetEnumerator() | %{
                        If($_.Key -eq $jobParam.Name)
                        {
                            $Value = $_.Value
                            If ($jobParam.Value2)
                            {
                                Write-Verbose 'Testing values...'
                                $Value.Split(';') | %{
                                    If ($jobParam.Value2.Split(';') -contains $_)
                                    {
                                        Write-Verbose "Value $_ is correct."
                                    }
                                    else
                                    {
                                        Throw "Incorrect value for parameter '$($jobParam.Name)'! Only the following values are allowed: '$($jobParam.Value2)'"
                                    }
                                }    
                            }
                            $jobParam.Value1 = $Value
                        }
                    }
                } # end foreach
                Write-Verbose 'All parameter values have been set.'
            }
            elseif (!$UseDefaults) # No Parameters
            {
                Write-Verbose 'Prompting for parameter values:'
                foreach ($jobParam in $InputParameters.JobParameters)
                {
                    $Correct = $True
                    $Value = Read-Host "Please provide value for parameter '$($jobParam.Name)'"
                    If ($jobParam.Value2)
                    {
                        $Value.Split(';') | %{
                            If ($jobParam.Value2.Split(';') -contains $_ -and $Correct)
                            {
                                $Correct = $True
                            }
                            else
                            {
                                Write-Verbose "Incorrect value found for parameter '$($jobParam.Name)':"
                                Write-Verbose "Faulty value is $_."
                                $Correct = $False
                            }
                        }
                        If (!$Correct)
                        {
                            Write-Verbose 'Incorrect parameter value(s) found.'
                            Do {
                                $Value = Read-Host "Allowed values are '$($jobParam.Value2)'"
                                $Correct = $True
                                $Value.Split(';') | %{
                                    If ($jobParam.Value2.Split(';') -contains $_ -and $Correct)
                                    {
                                        $Correct = $True
                                    }
                                    else
                                    {
                                        $Correct = $False
                                    }
                                }
                            }
                            until ($Correct)
                        }
                    } # end If $jobParam.Value2
                    $jobParam.Value1 = $Value
                } # end foreach
            } # end If-else $Parameters
        } # end IF $inputparameters
        $ArrWho = @()
    }
	process {
        foreach ($AMWho in $Who)
        {
            Write-Verbose "Processing target $AMWho..."
            If ($AMWho.PSObject.TypeNames -contains 'RES.AutomationManager.Agent')
            {
                $ArrWho += [pscustomobject]@{
                    ID = "{$($AMWho.WUIDAgent.ToString().ToUpper())}"
                    Type = 0
                    Name = $AMWho.Name
                }
            }
            ElseIf ($AMWho.PSObject.TypeNames -contains 'RES.AutomationManager.Team')
            {
                $ArrWho += [pscustomobject]@{
                    ID = "{$($AMWho.GUID.ToString().ToUpper())}"
                    Type = 1
                    Name = $AMWho.Name
                }
            }
            else
            {
                Write-Verbose "Determinig target type..."
                $Target = Get-RESAMAgent $AMWho
                If ($Target)
                {
                    Write-Verbose "Target $AMWho is an Agent."
                    $ArrWho += [pscustomobject]@{
                        ID = "{$($Target.WUIDAgent.ToString().ToUpper())}"
                        Type = 0
                        Name = $Target.Name
                    }
                }
                else
                {
                    $Target = Get-RESAMTeam $AMWho
                    If (!$Target)
                    {
                        Throw "Unable to find Agent/Team named $AMWho."
                    }
                    Write-Verbose "Target $AMWho is a Team."
                    $ArrWho += [pscustomobject]@{
                        ID = "{$($Target.GUID.ToString().ToUpper())}"
                        Type = 1
                        Name = $Target.Name
                    }
                }
            } # end If-elsif-else
        } # end foreach
    }
    End
    {
		$endPoint = "Dispatcher/SchedulingService/jobs"
		$uri = "http://$Dispatcher/$($endPoint)"
		
		$blob = [pscustomobject]@{
			Description = $Description
			When = @{
			    ScheduledDateTime = $Start
                Immediate = $Immediate.ToString().ToLower()
                IsLocalTime = $LocalTime.ToString().ToLower()
                UseWakeOnLAN = $UseWOL.ToString().ToLower()
			}
            What = @(
                        [pscustomobject]@{
                            ID = "{$($Task.GUID.ToString().ToUpper())}"
                            Type = $Type
                            Name = $Task.Name
                        }
                    )
            Who = $ArrWho
            Parameters = @($InputParameters)
		}
		$pREST = @{
			Uri = $Uri
			Method = "POST"
			Credential = $Credential
		}
		$Job = Invoke-RESAMRestMethod @pREST -Body (ConvertTo-Json $blob -Depth 99)
        switch ($Job.Status.JobInvoker)
        {
            'InvokeRunBook' {Get-RESAMMasterJob -MasterJobGUID $Job.JobID -InvokedByRunbook | Get-RESAMMasterJob -Full -WA 0}
            Default         {Get-RESAMMasterJob -MasterJobGUID $Job.JobID -Full -WA 0}
        }
	}
}
