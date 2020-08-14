#
# Module manifest for module 'User-Management'
#
# Generated by: Zachary Hilliker
#
# Generated on: 7/8/2019
#

@{

    # Script module or binary module file associated with this manifest.
    RootModule = 'User-Management.psm1'
    
    # Version number of this module.
    ModuleVersion = '1.5'
    
    # Supported PSEditions
    # CompatiblePSEditions = @()
    
    # ID used to uniquely identify this module
    GUID = 'ba8b0b59-03fe-4409-ac65-a8fc5eb96f7f'
    
    # Author of this module
    Author = 'Zachary Hilliker'
    
    # Company or vendor of this module
    CompanyName = 'RhythMedix'
    
    # Copyright statement for this module
    Copyright = '(c) 2019 Zachary Hilliker. All rights reserved.'
    
    # Description of the functionality provided by this module
    Description = '
    Add-NewUser - Accepts [String]Full Name, [String]Title, [String]Department, [String]Manager, [Boolean]Ladies 
                                - Parses name into First and Last, adds user to AD, 
                                 creates UPN, determines groups/roles based on title, applies required licenses, adds users to O365 groups
    Add-RhythmstarUser - Called by Add-NewUser if necessary, Adds user to Rhythmstar using @rhythmedix.com upn, if used with $Demo=$True parameter, will add user to Demo portal instead
    Convert-ADUserToCloudOnly - Removes user from Azure sync group, restores user from Azure AD recycle bin, removes immutable ID so AD user no longer matches Azure AD user, then removes AD User
    Sync-Azure - uses psexec to invoke an Azure AD delta sync on the DC 
    Export-DLtoCSV - Exports a distribution list to a CSV file
    Connect-O365Compliance - If one does not exist, Creates a PSsession to connect to the exo security and compliance center
    Connect-EXO - If one does not exist creates a pssession to connect to EXO
    Remove-Phishing - Run a content search in the O365 Security and compliance center to find the messages you want to remove. Note the name of the search, this 
            Takes results of the search and runs a purge with a hard delete
    Disconnect-EXO - disconnects exo PSsession
    Add-EXOMailboxPermission -grants full access for one user to a target exo mailbox
    Remove-EXOMailboxPermissoin -removes full access for one user to a target exo mailbox
    Add-CSVtoO365Group - imports a csv and adds to a specified O365 group
    Add-O365GroupUser - adds an individual user to an O365 group
    Disconnect-O365Compliance - disconnects compliance PSsession
    Update-Printers - accepts ComputerName, PrinterDepartment (Billing, Main, or Service), and Color (True or False), removes the old printers, and adds the appropriate printers
    Add-AdminUser - Adds RMXAdmin user and disables built-in Admin account
    Add-RMXVPN - sets up VPN,
    ProvisionBitlocker - turns on Bitlocker, backs up to Azure 
    Update-ProductKey - updates windows product key with key stroed in BIOS
    Add-WVDAppUser - (removes user from WVD DesktopUser group) adds user to remote review app group in Azure WVD
    Add-WVDDesktopUser - (removes user from WVDAppUser group) adds user to WVD Desktop group
    Disable-WVDSessionHost - disables new connections to session host (leech mode)
    Enable-WVDSessionHost - enables new connections to session host
    Get-WVDSession - returns all WVD sessions in rhythmstar_review, if passed a hostname parameter, it will return sessions on that host specifically
    Set-AzureComputerSync - (adds the given computer to the azure AD sync group for hybrid azure AD join - used for new device provisioning)
    Add-IPWhitelist - Adds the IP to the user whitelist in the rhythmstar portal.
    Restart-Holter - Restarts the Holter VM in Azure, based on SAMAccountName
    Invoke-WVDUserDisconnect - logs a user off of WVD based on SAMAccountName
    Get-WVDUsers - returns all provisioned remote review users
    update-associateIDs - import a properly formatted CSV and it will update with the new associate IDs for ADP SSO
    Convert-CloudUserToADSYnc - input UPN,Title,Manager,Department, creates AD user, gets AD immutableID and maps it to the correct cloud user, adds user to Azure Sync and begins a sync cycle
    Connect-WVDAccount - connects to windows virtual desktop app service

    '
    
    # Minimum version of the Windows PowerShell engine required by this module
    # PowerShellVersion = ''
    
    # Name of the Windows PowerShell host required by this module
    # PowerShellHostName = ''
    
    # Minimum version of the Windows PowerShell host required by this module
    # PowerShellHostVersion = ''
    
    # Minimum version of Microsoft .NET Framework required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
    # DotNetFrameworkVersion = ''
    
    # Minimum version of the common language runtime (CLR) required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
    # CLRVersion = ''
    
    # Processor architecture (None, X86, Amd64) required by this module
    # ProcessorArchitecture = ''
    
    # Modules that must be imported into the global environment prior to importing this module
    #RequiredModules = @('InvokePsExec', 'CredentialManager', 'MSOnline','AZ')
    
    # Assemblies that must be loaded prior to importing this module
    # RequiredAssemblies = @()
    
    # Script files (.ps1) that are run in the caller's environment prior to importing this module.
    #ScriptsToProcess = @()
    
    # Type files (.ps1xml) to be loaded when importing this module
    # TypesToProcess = @()
    
    # Format files (.ps1xml) to be loaded when importing this module
    # FormatsToProcess = @()
    
    # Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
    # NestedModules = @()
    
    # Functions to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no functions to export.
    FunctionsToExport = @('Add-NewUser','Add-RhythmstarUser','Convert-ADUserToCloudOnly','Sync-Azure','Export-DLtoCSV',
    'Connect-O365Compliance','Connect-EXO','Remove-Phishing','Disconnect-EXO','Add-EXOMailboxPermission','Remove-EXOMailboxPermission',
    'Add-CSVtoO365Group','Add-O365GroupUser','Disconnect-O365Compliance','Update-Printers','Add-AdminUser','Add-RMXVPN','ProvisionBitlocker',
    'Update-ProductKey','Add-WVDAppUser','Add-WVDDesktopUser','New-WVDRemoteApp','Disable-User','Set-AzureComputerSync','Disable-WVDSessionHost','Enable-WVDSessionHost', 
    'Get-WVDSession','Add-IPWhitelist','Restart-Holter', 'Invoke-WVDUserDisconnect','Get-WVDUsers','update-associateIDs','Convert-CloudUserToADSYnc','Connect-WVDAccount')
    
    # Cmdlets to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no cmdlets to export.
    CmdletsToExport = ''
    
    # Variables to export from this module
    VariablesToExport = ''
    
    # Aliases to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no aliases to export.
    AliasesToExport = ''
    
    # DSC resources to export from this module
    # DscResourcesToExport = @()
    
    # List of all modules packaged with this module
     ModuleList = @('InvokePsExec','CredentialManager','MSOnline','AZ')
    
    # List of all files packaged with this module
    FileList = @('Sync-Azure.ps1')
    
    # Private data to pass to the module specified in RootModule/ModuleToProcess. This may also contain a PSData hashtable with additional module metadata used by PowerShell.
    PrivateData = @{
    
        PSData = @{
    
            # Tags applied to this module. These help with module discovery in online galleries.
            # Tags = @()
    
            # A URL to the license for this module.
            # LicenseUri = ''
    
            # A URL to the main website for this project.
            # ProjectUri = ''
    
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
    
    