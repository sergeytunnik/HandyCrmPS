Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$userDocuments = [Environment]::GetFolderPath("mydocuments")
$userPSModulePath = $userDocuments + '\WindowsPowerShell\Modules'
$moduleName = 'Handy.Crm.Powershell.Cmdlets'
$moduleDir = $userPSModulePath + '\' + $moduleName

if (!(Test-Path -LiteralPath $moduleDir))
{
    New-Item -ItemType Directory -Path $moduleDir | Out-Null
}

$moduleFiles = @(
    'Handy.Crm.Powershell.Cmdlets.dll',
    'Handy.Crm.Powershell.Cmdlets.psd1',
    'Microsoft.Crm.Sdk.Proxy.dll',
    'Microsoft.Xrm.Client.dll',
    'Microsoft.Xrm.Sdk.dll')

foreach ($file in $moduleFiles)
{
    Copy-Item -Path $file -Destination $moduleDir -Force
}