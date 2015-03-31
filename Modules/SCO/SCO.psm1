#requires -Version 4

################################################################################
# COMPILE Orchestrator.svc Service
try {
    [Microsoft.SystemCenter.Orchestrator.WebService.OrchestratorContext] | Out-Null
} catch {
    $Stream         = New-Object System.IO.MemoryStream (,[Convert]::FromBase64String('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'))
    $GZipStream     = New-Object System.IO.Compression.GZipStream $Stream,([System.IO.Compression.CompressionMode]::Decompress)
    $Reader         = New-Object System.IO.StreamReader $GZipStream
    $TypeDefinition = $Reader.ReadToEnd()
    
    Add-Type -AssemblyName 'System.Data.Services.Client, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'
    Add-Type -TypeDefinition $TypeDefinition -Language CSharp -ReferencedAssemblies 'System.Data.Services.Client, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'
}

<#
.SYNOPSIS
    Connects to a specific Orchestrator 2012 R2 computer.

.EXAMPLE
    Connect-SCOService -ComputerName localhost

    Connect to localhost on the default port with the current users credentials.

.EXAMPLE
    Connect-SCOService -ComputerName localhost -Port 81

    Connect to localhost on port 81 with the current users credentials.

.EXAMPLE
    Connect-SCOService -ComputerName localhost -Secure

    Connect to localhost on the default port over https with the current user credentials.

.EXAMPLE
    Connect-SCOService -ComputerName localhost -Credential DOMAIN\Administrator

    Connect to localhost on the default port with the credentials of DOMAIN\Administrator.
#>
function Connect-SCOService {
    [OutputType([Microsoft.SystemCenter.Orchestrator.WebService.OrchestratorContext])]
    [CmdletBinding(DefaultParameterSetName='ComputerName')]
    param
    (
        [Parameter(Mandatory, ParameterSetName='ComputerName')]
        [ValidateNotNullOrEmpty()]
        [string] $ComputerName,

        [ValidateRange(1, 65535)]
        [int] $Port = 81,

        [switch] $Secure = $false,

        [Parameter(Mandatory, ParameterSetName='Uri')]
        [ValidateNotNull()]
        [Uri] $Uri = "http$(if ($Secure) {'s'})://${ComputerName}:${Port}/Orchestrator2012/Orchestrator.svc",

        [System.Management.Automation.Credential()]
        [ValidateNotNull()]
        $Credential = ([PSCredential]::Empty)
    )

    $Context = New-Object Microsoft.SystemCenter.Orchestrator.WebService.OrchestratorContext $Uri
    
    if ($Credential -eq [PSCredential]::Empty) {
        $Context.Credentials = [System.Net.CredentialCache]::DefaultCredentials
    } else {
        $Context.Credentials = $Credential.GetNetworkCredential()
    }

    $global:SCOService = $Context
}

<#
.SYNOPSIS
    Gets the runbooks from the Orchestrator server.

.EXAMPLE
    Get-SCORunbook -Path '\Path\To\Runbook'

    Get the runbook based on the absolute path within Orchestrator.

.EXAMPLE
    Get-SCORunbook -Id 3bd19297-7e6c-48cb-8413-d897409367a3

    Get the runbook with the Id 3bd19297-7e6c-48cb-8413-d897409367a3

.EXAMPLE
    Get-SCOJob -Status Running | Get-SCORunbook

    Get all the jobs that are running and then gets the associated runbook for each job.
#>
function Get-SCORunbook {
    [CmdletBinding(DefaultParameterSetName='Path')]
    param
    (
        [Parameter(Mandatory,ParameterSetName='Path')]
        [ValidateNotNullOrEmpty()]
        [string] $Path,

        [Parameter(Mandatory,ValueFromPipelineByPropertyName,ParameterSetName='RunbookId')]
        [Alias('Id')]
        [Guid] $RunbookId,

        [ValidateNotNull()]
        [Microsoft.SystemCenter.Orchestrator.WebService.OrchestratorContext] $Context = $($global:SCOService)
    )

    begin {
        if ($PSCmdlet.ParameterSetName -eq 'Path') {
            $Context.Runbooks.AddQueryOption('$filter', "Path eq '$Path'")
        }

        $Runbooks = @()
    }

    process {
        if ($PSCmdlet.ParameterSetName -eq 'RunbookId') {
            if ($Runbooks -notcontains $RunbookId) {
                $Context.Runbooks.AddQueryOption('$filter', "Id eq guid'$RunbookId'")
            }
        }
    }
}

<#
.SYNOPSIS
    Gets the parameters of the specified runbook.

.EXAMPLE
    Get-SCOJob -Status Running | Get-SCORunbook | Get-SCORunbookParameters

    Gets the parameters of all the runbooks that currently have jobs running.

.EXAMPLE
    Get-SCORunbookParameters -RunbookId 3bd19297-7e6c-48cb-8413-d897409367a3

    Gets the parameters of the runbook 3bd19297-7e6c-48cb-8413-d897409367a3.
#>
function Get-SCORunbookParameters {
    [OutputType([Microsoft.SystemCenter.Orchestrator.WebService.Runbook])]
    [CmdletBinding(DefaultParameterSetName='Runbook')]
    param
    (
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName='RunbookId')]
        [Guid] $RunbookId,

        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName='Runbook')]
        [ValidateNotNull()]
        [Microsoft.SystemCenter.Orchestrator.WebService.Runbook] $Runbook = (Get-SCORunbook -RunbookId $RunbookId),

        [ValidateNotNull()]
        [Microsoft.SystemCenter.Orchestrator.WebService.OrchestratorContext] $Context = $($global:SCOService)
    )

    $Context.RunbookParameters.AddQueryOption('$filter', "RunbookId eq guid'$($Runbook.Id)' and Direction eq 'In'").Execute()
}

<#
.SYNOPSIS
    Gets the Orchestrator jobs.

.EXAMPLE
    Get-SCOJob

.EXAMPLE
    Get-SCOJob -Status Running

.EXAMPLE
    Get-SCORunbook -RunbookId 3bd19297-7e6c-48cb-8413-d897409367a3 | Get-SCOJob -Status Running
#>
function Get-SCOJob {
    [OutputType([Microsoft.SystemCenter.Orchestrator.WebService.Job])]
    [CmdletBinding(DefaultParameterSetName='Status')]
    param
    (
        [Parameter(Mandatory, ParameterSetName='Id')]
        [Guid] $Id,

        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName='Runbook')]
        [ValidateNotNull()]
        [Microsoft.SystemCenter.Orchestrator.WebService.Runbook] $Runbook,

        [Parameter(ParameterSetName='Id')]
        [Parameter(ParameterSetName='Runbook')]
        [Parameter(Mandatory, ParameterSetName='Status')]
        [ValidateSet('Pending','Running','Failed','Cancelled','Completed')]
        [string] $Status,

        [Parameter(Mandatory, ParameterSetName='Filter')]
        [string] $Filter,

        [ValidateNotNull()]
        [Microsoft.SystemCenter.Orchestrator.WebService.OrchestratorContext] $Context = $($global:SCOService)
    )

    switch ($PSCmdlet.ParameterSetName) {
        'Id' {
            $Filter = "Id eq guid'${Id}'"
        }

        'Runbook' {
            $Filter = "RunbookId eq guid'$($Runbook.Id)'"
        }

        'Status' {
            $Filter = ''
        }
    }

    if ($Status -and $PSCmdlet.ParameterSetName -ne 'Filter') {
        if ($Filter) {
            $Filter += 'and '
        }

        $Filter += "Status eq '$Status'"
    }

    $Context.Jobs.AddQueryOption('$filter', $Filter)
}

<#
.SYNOPSIS
    Stops a running job.

.EXAMPLE
    Get-SCOJob -Status Running | Stop-SCOJob

.EXAMPLE
    Get-SCORunbook -RunbookId 3bd19297-7e6c-48cb-8413-d897409367a3 | Get-SCOJob -Status Running | Stop-SCOJob
#>
function Stop-SCOJob {
    [CmdletBinding(SupportsShouldProcess)]
    param
    (
        [Parameter(Mandatory,ValueFromPipeline)]
        [ValidateNotNull()]
        [Microsoft.SystemCenter.Orchestrator.WebService.Job] $Job,

        [ValidateNotNull()]
        [Microsoft.SystemCenter.Orchestrator.WebService.OrchestratorContext] $Context = $($global:SCOService)
    )

    process {
        if ($PSCmdlet.ShouldProcess($Job.Id, 'Stop-SCOJob')) {
            $Job.Status = 'Canceled'
            $Context.UpdateObject($Job)
        }
    }

    end {
        if ($PSCmdlet.ShouldProcess($Context.BaseUri, 'Stop-SCOJob')) {
            $Context.SaveChanges()
        }
    }
}

<#
.SYNOPSIS
    Starts a runbook job.

.EXAMPLE
    Get-SCORunbook -Path '\Path\To\Runbook' | Start-SCOJob 
    
.EXAMPLE
    Get-SCORunbook -Path '\Path\To\Runbook' | Start-SCOJob -Parameters @{ParameterName1='Value';ParameterName2='Value'}   
#>
function Start-SCOJob {
    [OutputType([Microsoft.SystemCenter.Orchestrator.WebService.Job])]
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ValidateNotNull()]
        [Microsoft.SystemCenter.Orchestrator.WebService.Runbook] $Runbook,

        [ValidateNotNull()]
        [Hashtable] $Parameters,

        [ValidateNotNull()]
        [Microsoft.SystemCenter.Orchestrator.WebService.OrchestratorContext] $Context = $($global:SCOService)
    )
    
    $Job = New-Object Microsoft.SystemCenter.Orchestrator.WebService.Job -Property @{RunbookId = $Runbook.Id}
    
    try {
        if ($Parameters -and $Parameters.Count) {
            $RunbookParameters = Get-SCORunbookParameters -Runbook $Runbook -ErrorAction Stop | Group-Object Name -AsHashTable -AsString
            $Job.Parameters = "<Data>$($RunbookParameters.GetEnumerator() | ForEach-Object {"<Parameter><ID>$($_.Value.ID.ToString('B'))</ID><Value>$($Parameters[$_.Key])</Value></Parameter>"})</Data>"
            Write-Verbose $Job.Parameters
        }

        $Context.AddToJobs($Job)
        $Context.SaveChanges()

        $Job
    } catch {
        throw
    }
}

Export-ModuleMember -Function Connect-SCOService,Get-SCORunbook,Get-SCORunbookParameters,Get-SCOJob,Start-SCOJob,Stop-SCOJob