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

Add-Type -Assembly System.IO.Compression.FileSystem

$Authorization  = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$($Credential.UserName):$($Credential.GetNetworkCredential().Password)"));
$DownloadFolder = "$($env:USERPROFILE)\Downloads\PowerShell"

if (!(Test-Path -Path $DownloadFolder -PathType Container)) {
    New-Item -Path $DownloadFolder -ItemType Directory | Out-Null
}

$NextPage = 'https://api.github.com/users/PowerShell/repos'

while ($NextPage) {
    $Response = Invoke-WebRequest -Uri $NextPage -Headers @{Authorization=$Authorization}

    ($Response.Content | ConvertFrom-Json).ForEach{
        Write-Verbose "Processing repository $($_.name)..."

        $FileUrl           = "$($_.html_url)/archive/master.zip"
        $DestinationFile   = "$($DownloadFolder)\$($_.name).zip"
        $DestinationFolder = "$($DownloadFolder)\$($_.name)"
        $ContentFolder     = "$($DestinationFolder)\$($_.name)-master"
        $ModuleFolder      = "$Path\$($_.name)"

        if (Test-path -Path $DestinationFile -PathType Leaf) {
            Remove-Item -Path $DestinationFile -Force
        }

        Start-BitsTransfer -Source $FileUrl -Destination $DestinationFile -Authentication Basic -Credential $Credential -TransferType Download -DisplayName $_.name
    
        if (Test-Path -Path $DestinationFolder -PathType Container) {
            Remove-Item -Path $DestinationFolder -Recurse -Force
        }

        [System.IO.Compression.ZipFile]::ExtractToDirectory($DestinationFile, $DestinationFolder)

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

    $NextPage = ($Response.Headers.Link -split ', ' | Where-Object {$_ -like '* rel="next"'}) -split ';| |<|>' | Where-Object {$_} | Select-Object -First 1
}
