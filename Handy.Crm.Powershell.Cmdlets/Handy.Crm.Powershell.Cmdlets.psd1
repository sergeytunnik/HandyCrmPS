@{

# Script module or binary module file associated with this manifest.
# RootModule = ''

# Version number of this module.
ModuleVersion = '0.6.0.0'

# ID used to uniquely identify this module
GUID = '921009b6-0b33-49b8-a37a-d864da65160f'

# Author of this module
Author = 'Sergey Tunnik'

# Company or vendor of this module
CompanyName = ''

# Copyright statement for this module
Copyright = '(c) 2018 Sergey Tunnik, licensed under MIT Licence.'

# Description of the functionality provided by this module
Description = 'Set of basic cmdlets for PowerShell. It helps running common routine day-to-day tasks and as part of DevOps processes in Microsoft Dynamics CRM.'

# Minimum version of the Windows PowerShell engine required by this module
PowerShellVersion = '4.0'

# Name of the Windows PowerShell host required by this module
# PowerShellHostName = ''

# Minimum version of the Windows PowerShell host required by this module
# PowerShellHostVersion = ''

# Minimum version of Microsoft .NET Framework required by this module
DotNetFrameworkVersion = '4.6.2'

# Minimum version of the common language runtime (CLR) required by this module
# CLRVersion = ''

# Processor architecture (None, X86, Amd64) required by this module
# ProcessorArchitecture = ''

# Modules that must be imported into the global environment prior to importing this module
RequiredModules = @()

# Assemblies that must be loaded prior to importing this module
RequiredAssemblies = @('Microsoft.Xrm.Sdk.dll', 'Microsoft.Crm.Sdk.Proxy.dll', 'SolutionPackager.dll')

# Script files (.ps1) that are run in the caller's environment prior to importing this module.
# ScriptsToProcess = @()

# Type files (.ps1xml) to be loaded when importing this module
# TypesToProcess = @()

# Format files (.ps1xml) to be loaded when importing this module
# FormatsToProcess = @()

# Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
NestedModules = @('Handy.Crm.Powershell.Cmdlets.dll', 'Handy.Crm.Powershell.Cmdlets.psm1')

# Functions to export from this module
FunctionsToExport = '*'

# Cmdlets to export from this module
CmdletsToExport = '*'

# Variables to export from this module
VariablesToExport = '*'

# Aliases to export from this module
AliasesToExport = '*'

# List of all modules packaged with this module
# ModuleList = @()

# List of all files packaged with this module
FileList = @('Handy.Crm.Powershell.Cmdlets.dll', 'Handy.Crm.Powershell.Cmdlets.psm1',
    'Microsoft.Crm.Sdk.Proxy.dll', 'Microsoft.Xrm.Sdk.Deployment.dll', 'Microsoft.Xrm.Sdk.dll', 'Microsoft.Xrm.Sdk.Workflow.dll', 'SolutionPackager.dll')

# Private data to pass to the module specified in RootModule/ModuleToProcess. This may also contain a PSData hashtable with additional module metadata used by PowerShell.
PrivateData = @{

    PSData = @{

        # Tags applied to this module. These help with module discovery in online galleries.
        Tags = @('PSModule', 'Dynamics', 'CRM', 'DevOps', 'XRM')

        # A URL to the license for this module.
        LicenseUri = 'https://raw.githubusercontent.com/sergeytunnik/HandyCrmPS/master/LICENSE'

        # A URL to the main website for this project.
        ProjectUri = 'https://github.com/sergeytunnik/HandyCrmPS'

        # A URL to an icon representing this module.
        # IconUri = ''

        # ReleaseNotes of this module
        # ReleaseNotes = ''

    } # End of PSData hashtable

} # End of PrivateData hashtable

# HelpInfo URI of this module
# HelpInfoURI = ''

# Default prefix for commands exported from this module. Override the default prefix using Import-Module -Prefix.
# DefaultCommandPrefix = ''

}

