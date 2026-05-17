param(
    [string]$ConverterRoot
)

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

& (Join-Path $scriptDir "CreateDistributionPackage.ps1") -ConverterRoot $ConverterRoot
& (Join-Path $scriptDir "CreateDistributionPackage_Lite.ps1") -ConverterRoot $ConverterRoot

Write-Host "Created both v3 and Lite distribution packages."
