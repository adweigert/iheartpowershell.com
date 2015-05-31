#requires -Version 4

[CmdletBinding()]
param
(
    [Parameter(Mandatory)]
    [System.Management.Automation.Credential()]
    [PSCredential] $Credential,

    [ValidateNotNullOrEmpty()]
    [string] $Path = $("$($env:USERPROFILE)\Documents\WindowsPowerShell\Modules")
)

$Authorization = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$($Credential.UserName):$($Credential.GetNetworkCredential().Password)"));

@((Invoke-WebRequest -Uri https://api.github.com/users/PowerShell/repos -Headers @{Authorization=$Authorization}).Content | ConvertFrom-Json)[0].ForEach{
    Write-Verbose "Processing repository $($_.name)..."

    $FileUrl           = "$($_.html_url)/archive/master.zip"
    $DestinationFile   = "$($env:USERPROFILE)\Downloads\PowerShell\$($_.name).zip"
    $DestinationFolder = "$($env:USERPROFILE)\Downloads\PowerShell\$($_.name)"
    $ContentFolder     = "$($DestinationFolder)\$($_.name)-master"
    $ModuleFolder      = "$Path\$($_.name)"

    if (Test-path -Path $DestinationFile -PathType Leaf) {
        Remove-Item -Path $DestinationFile -Force
    }

    Start-BitsTransfer -Source $FileUrl -Destination $DestinationFile -Authentication Basic -Credential $Credential -TransferType Download -DisplayName $_.name
    
    if (Test-Path -Path $DestinationFolder -PathType Container) {
        Remove-Item -Path $DestinationFolder -Recurse -Force
    }

    Expand-Archive -Path $DestinationFile -DestinationPath $DestinationFolder -Force

    if (Get-ChildItem -Path $ContentFolder | Where-Object Extension -match '\.(psd1|psm1)') {
        Write-Verbose "Publishing '$($_.name)' PowerShell module to '$ModuleFolder'..."
        
        if (Test-Path -Path $ModuleFolder -PathType Container) {
            Remove-Item -Path $ModuleFolder -Recurse -Force
        }

        Copy-Item -Path $ContentFolder -Destination $ModuleFolder -Recurse -Force
    } else {
        Write-Warning "The repository '$($_.name)' does not contain a PowerShell module."
    }

    Remove-Item $DestinationFile -Force
}
