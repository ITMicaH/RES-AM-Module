#
# Module manifest for module 'RESAM'
#
# Generated by: Michaja van der Zouwen
#
# Generated on: 10-6-2015
#

@{

# Script module or binary module file associated with this manifest.
RootModule = 'RESAM'

# Version number of this module.
ModuleVersion = '1.0'

# ID used to uniquely identify this module
GUID = 'dfa827eb-4e34-47f1-888d-8c2f59081afa'

# Author of this module
Author = 'Michaja van der Zouwen'

# Company or vendor of this module
CompanyName = 'ITMicaH'

# Copyright statement for this module
Copyright = '(c) 2015 Michaja van der Zouwen. All rights reserved.'

# Description of the functionality provided by this module
# Description = 'Module for RES Automation Manager.'

# Minimum version of the Windows PowerShell engine required by this module
PowerShellVersion = '3.0'

# Name of the Windows PowerShell host required by this module
# PowerShellHostName = ''

# Minimum version of the Windows PowerShell host required by this module
# PowerShellHostVersion = ''

# Minimum version of Microsoft .NET Framework required by this module
# DotNetFrameworkVersion = ''

# Minimum version of the common language runtime (CLR) required by this module
# CLRVersion = ''

# Processor architecture (None, X86, Amd64) required by this module
# ProcessorArchitecture = ''

# Modules that must be imported into the global environment prior to importing this module
# RequiredModules = @()

# Assemblies that must be loaded prior to importing this module
# RequiredAssemblies = @()

# Script files (.ps1) that are run in the caller's environment prior to importing this module.
# ScriptsToProcess = @()

# Type files (.ps1xml) to be loaded when importing this module
TypesToProcess = 'RESAM.Types.ps1xml'

# Format files (.ps1xml) to be loaded when importing this module
FormatsToProcess = 'RESAM.Format.ps1xml'

# Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
# NestedModules = 'RESAM.Types.ps1xml'

# Functions to export from this module
FunctionsToExport = @(
    'Connect-RESAMDatabase'
    'Disconnect-RESAMDatabase'
    'Get-RESAMAgent'
	'Remove-RESAMAgent'
    'Get-RESAMTeam'
    'Get-RESAMAudit'
    'Get-RESAMDispatcher'
    'Get-RESAMModule'
    'Get-RESAMProject'
    'Get-RESAMRunBook'
    'Get-RESAMResource'
    'Get-RESAMAgent'
    'Get-RESAMConnector'
    'Get-RESAMDatabaseLevel'
    'Get-RESAMConsole'
    'Get-RESAMMasterJob'
    'Get-RESAMJob'
    'Get-RESAMQueryResult'
    'Get-RESAMLog'
    'New-ResAMJob'
)

# Cmdlets to export from this module
CmdletsToExport = '*'

# Variables to export from this module
VariablesToExport = '*'

# Aliases to export from this module
AliasesToExport = '*'

# List of all modules packaged with this module
# ModuleList = @()

# List of all files packaged with this module
FileList = @(
    'RESAM.psm1'
    'RESAM.psd1'
    'RESAM.Format.ps1xml'
    'RESAM.Types.ps1xml'
	'RESAM.Help.xml'
)

# Private data to pass to the module specified in RootModule/ModuleToProcess
# PrivateData = ''

# HelpInfo URI of this module
HelpInfoURI = 'http://itmicah.wordpress.com'

# Default prefix for commands exported from this module. Override the default prefix using Import-Module -Prefix.
# DefaultCommandPrefix = ''

}

