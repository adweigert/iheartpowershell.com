#requires -Version 4

<#
.SYNOPSIS
    Create desired state configuration archive.

.DESCRIPTION
    Create an archive (zip) that contains the DSC mof file, an install script, and all the modules references by the DSC resources.

.EXAMPLE
    MyConfiguration -ConfigurationData $Config | .\Package-DscConfiguration.ps1
#>
[CmdletBinding()]
param
(
    [Parameter(Mandatory, ValueFromPipelineByPropertyName, Position=0)]
    [ValidateNotNullOrEmpty()]
    [string] $Path
)

process {
    Add-Type -Assembly System.IO.Compression.FileSystem
        
    Get-ChildItem -Path $Path -Filter *.mof -PipelineVariable MOF | ForEach-Object {
        Write-Verbose "MOF '$($MOF.FullName)' found."

        $Content = Get-Content -Path $MOF.FullName | Out-String
        $Content = $Content -replace '/\*','<#'
        $Content = $Content -replace '\*/','#>'
        $Content = $Content -replace '= (NULL|TRUE|FALSE)','= $$$1'
        $Content = $Content -replace 'instance of .+\n','[PSCustomObject][Ordered]@'
        $Content = $Content -replace '\\\\','\'

        $Instances = @(Invoke-Expression -Command $Content)
        $Modules   = @($Instances | Select-Object -Property ModuleName,ModuleVersion -Unique).Where{$_.ModuleName -and $_.ModuleVersion -and $_.ModuleName -ne 'PSDesiredStateConfiguration'}

        $Temp       = [System.IO.Path]::GetTempFileName()
        $TempMOF    = Join-Path -Path $Temp -ChildPath $MOF.Name
        $TempScript = Join-Path -Path $Temp -ChildPath "$($MOF.Name).ps1"
        $MOFArchive = "$($MOF.FullName).zip"

        if (Test-Path -Path $Temp) {
            Remove-Item -Path $Temp -Recurse -Force
        }

        New-Item -Path $Temp -ItemType Directory -ErrorAction Stop | Out-Null
        Copy-Item -Path $MOF.FullName -Destination $TempMOF -ErrorAction Stop

        $Modules.ForEach{
            $Module = Get-Module -FullyQualifiedName @{ModuleName=$_.ModuleName;ModuleVersion=$_.ModuleVersion} -ListAvailable -ErrorAction SilentlyContinue

            if (!$Module) {
                Write-Error "Could not locate module '$($_.ModuleName)' version '$($_.ModuleVersion)'." -ErrorAction Stop
            }

            Write-Verbose "Module '$($_.ModuleName)' version $($_.ModuleVersion) referenced by MOF: $($Module.ModuleBase)"
                
            $TempModule = Join-Path -Path $Temp -ChildPath $Module.Name

            Copy-Item -Path $Module.ModuleBase -Destination $TempModule -Recurse -ErrorAction Stop
        }

        "#requires -Version 4`n`n[CmdletBinding(SupportsShouldProcess)]`nparam()`n`Get-ChildItem -Path `$PSScriptRoot -Directory | Copy-Item -Destination ""`$env:ProgramFiles\WindowsPowerShell\Modules"" -Recurse -Force`nStart-DscConfiguration -Path `$PSScriptRoot -Verbose -Wait -Force" | Out-File -FilePath $TempScript -Encoding ascii

        if (Test-Path -Path $MOFArchive -PathType Leaf) {
            Remove-Item -Path $MOFArchive -Force -ErrorAction Stop
        }

        [System.IO.Compression.ZipFile]::CreateFromDirectory($Temp, $MOFArchive)

        Remove-Item -Path $Temp -Recurse -Force

        Get-Item -path $MOFArchive
    }
}