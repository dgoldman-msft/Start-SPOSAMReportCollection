﻿function Write-ToLog {
    <#
        .SYNOPSIS
            Save output

        .DESCRIPTION
            Overload function for Write-Output

        .PARAMETER LoggingDirectory
            Directory to save the log file to. Default is "$env:temp".

        .PARAMETER LoggingFilename
            Filename to save the log file to. Default is "SamReportingLogs.txt".

        .EXAMPLE
            None

        .NOTES
            None
    #>

    [OutputType('System.String')]
    [CmdletBinding(DefaultParameterSetName = 'Default')]
    param
    (
        [Parameter(ParameterSetName = 'Default')]
        [string]
        $LoggingDirectory,

        [string]
        $LoggingFilename,

        [Parameter(Mandatory = $True, Position = 0)]
        [string]
        $InputString
    )

    try {
        if (-NOT(Test-Path -Path $LoggingDirectory)) {
            Write-Verbose "Creating New Logging Directory"
            New-Item -Path $LoggingDirectory -ItemType Directory -ErrorAction Stop | Out-Null
        }
    }
    catch {
        Write-Output "$_"
        return
    }

    try {
        # Console and log file output
        $stringObject = "[{0:MM/dd/yy} {0:HH:mm:ss}] - {1}" -f (Get-Date), $InputString
        Add-Content -Path (Join-Path $LoggingDirectory -ChildPath $LoggingFilename) -Value $stringObject -Encoding utf8 -ErrorAction Stop
    }
    catch {
        Write-Output "$_"
        return
    }
}

function Start-SPOSAMReportCollection {
    <#
        .SYNOPSIS
            Starts the SharePoint Online Security and Access Management (SAM) report collection.

        .DESCRIPTION
            This cmdlet is used to generate DAG reports which deal with potential oversharing of sensitive data.
            These reports are present in Sharepoint admin center. Reports are currently available for the following scenarios:

            Sharing links created in last 28 days (Anyone, People-in-your-org, Specific people shared externally).
            Content shared with Everyone except external users (EEEU) in last 28 days.
            List of sites having labelled files, as of report generation time.
            List of sites having 'too-many-users', as of report generation time, to setup an oversharing baseline.

        .PARAMETER CheckSensitivityLabel
            A switch parameter that, if specified, will check for sensitivity labels on files in the site.
            This will make a connection to the Security & Compliance Center to retrieve labels.

        .PARAMETER CountOfUsersMoreThan
            Specifies the threshold of oversharing as defined by the number of users that can access the site.
            The number of users that can access the site are determined by expanding all users, groups across all
            permissions (at site level and at the level of any item with unique permissions), deduplicate and
            arrive at a unique number. Minimum value is 100.
            Default for this script is 0.

        .PARAMETER DisconnectFromSPO
            A switch parameter that, if specified, will disconnect from SharePoint Online after the report collection is completed.

        .PARAMETER LoggingDirectory
            Directory to save the log file to. Default is "$env:temp\Logging".

        .PARAMETER LoggingFilename
            Filename to save the log file to. Default is "SamReportingLogs.txt".

        .PARAMETER Privacy
            Specifies the privacy setting of the Microsoft 365 group. Relevant in case of filtering the report for group connected sites.
            Valid values are 'All', 'Private', and 'Public'.
            Default for this script is is 'All'.

        .PARAMETER ReportEntity
            Specifies the entity for which the report should be generated. Valid values are:
            - EveryoneExceptExternalUsersAtSite
            - EveryoneExceptExternalUsersForItems
            - SharingLinks_Anyone
            - SharingLinks_PeopleInYourOrg
            - SharingLinks_Guests
            - SensitivityLabelForFiles
            - PermissionedUsers

        .PARAMETER ReportType
            Specifies the time period of data based on which DAG report is generated.
            A 'Snapshot' report will have the latest data as of the report generation time.
            A 'RecentActivity' report will be based on data in the last 28 days.
            Default for this script is 'RecentActivity'.

        .PARAMETER Template
            Specifies the template of the site. Relevant in case a report should be generated for that particular template.
            Valid values are 'AllSites', 'ClassicSites', 'CommunicationSites', 'TeamSites', and 'OtherSites'.

        .PARAMETER TenantDomain
            Specifies the domain of the tenant. This parameter is mandatory.

        .PARAMETER TenantAdminUrl
            Specifies the URL of the tenant admin site. Default is "https://$TenantDomain-admin.sharepoint.com".

        .PARAMETER UserPrincipalName
            Specifies the username for authentication.

        .PARAMETER Workload
            Specifies the workload for which the report should be generated. Valid values are 'SharePoint' and 'OneDriveForBusiness'. Default is 'SharePoint'.

        .EXAMPLE
            C:\PS> Start-SPOSAMReportCollection -TenantDomain contoso -UserPrincipalName Administrator -CheckSensitivityLabel

            This example will generate reports for all entities with the default parameters as well as check the sensitivity labels.
            This will connect you to the Security and Compliance Center to read the labels in the tenant.

        .EXAMPLE
            C:\PS> Start-SPOSAMReportCollection -TenantDomain contoso -UserPrincipalName Administrator -CountOfUsersMoreThan 100

            This example will generate reports for all entities with the default parameters and a threshold of 100 users.

        .EXAMPLE
            C:\PS> Start-SPOSAMReportCollection -TenantDomain contoso -UserPrincipalName Administrator -Privacy Private

            This example will generate reports for all entities with the default parameters and filter the report for private sites.

        .EXAMPLE
            C:\PS> Start-SPOSAMReportCollection -TenantDomain contoso -UserPrincipalName Administrator -ReportEntity SharingLinks_Anyone

            This example will generate a report for the 'SharingLinks_Anyone' entity with the default parameters.

        .EXAMPLE
            C:\PS> Start-SPOSAMReportCollection -TenantDomain contoso -UserPrincipalName Administrator -ReportType Snapshot

            This example will generate a snapshot report for all entities with the default parameters.

        .EXAMPLE
            C:\PS> Start-SPOSAMReportCollection -TenantDomain contoso -UserPrincipalName Administrator -Template TeamSites

            This example will generate reports for all entities with the default parameters and filter the report for team sites.

        .EXAMPLE
            C:\PS> Start-SPOSAMReportCollection -TenantDomain contoso -UserPrincipalName Administrator -Workload OneDriveForBusiness

            This example will generate reports for all entities with the default parameters and filter the report for OneDrive for Business.

        .EXAMPLE
            C:\PS> Start-SPOSAMReportCollection -TenantDomain contoso -UserPrincipalName Administrator -DisconnectFromSPO

            This example will generate reports for all entities with the default parameters and disconnect from SharePoint Online after the report collection is completed.

        .NOTES
            For more information please see: https://learn.microsoft.com/en-us/sharepoint/data-access-governance-reports
    #>

    [OutputType('System.String')]
    [CmdletBinding(DefaultParameterSetName = 'Default')]
    [Alias('SAMR')]
    param
    (
        [Parameter(ParameterSetName = 'Default')]
        [switch]
        $CheckSensitivityLabel,

        [Parameter(ParameterSetName = 'Default')]
        [Int]
        $CountOfUsersMoreThan = 0,

        [Parameter(ParameterSetName = 'Default')]
        [switch]
        $DisconnectFromSPO,

        [Parameter(ParameterSetName = 'Default')]
        [string]
        $LoggingDirectory = (Join-Path -Path ([Environment]::GetFolderPath("MyDocuments")) -ChildPath "SamReporting"),

        [Parameter(ParameterSetName = 'Default')]
        [string]
        $LoggingFilename = "SamReportingLogs.txt",

        [Parameter(ParameterSetName = 'Default')]
        [ValidateSet('All', 'Private', 'Public')]
        [string]
        $Privacy = 'All',

        [Parameter(ParameterSetName = 'Default')]
        [ValidateSet('EveryoneExceptExternalUsersAtSite', 'EveryoneExceptExternalUsersForItems', 'SharingLinks_Anyone', 'SharingLinks_PeopleInYourOrg', 'SharingLinks_Guests', 'SensitivityLabelForFiles', 'PermissionedUsers')]
        [string]
        $ReportEntity,

        [Parameter(ParameterSetName = 'Default')]
        [ValidateSet('Snapshot', 'RecentActivity')]
        [string]
        $ReportType = 'RecentActivity',

        [Parameter(ParameterSetName = 'Default')]
        [ValidateSet('AllSites', 'ClassicSites', 'CommunicationSites', 'TeamSites', 'OtherSites')]
        [string]
        $Template,

        [Parameter(Mandatory = $true, ParameterSetName = 'Default')]
        [string]
        $TenantDomain,

        [Parameter(ParameterSetName = 'Default')]
        [string]
        $TenantAdminUrl = "https://$TenantDomain-admin.sharepoint.com",

        [Parameter(ParameterSetName = 'Default')]
        [ValidateSet('SharePoint', 'OneDriveForBusiness')]
        [string]
        $Workload = "SharePoint",

        [Parameter(ParameterSetName = 'Default')]
        [string]
        $UserPrincipalName
    )

    # Check if running as administrator
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Output "This script must be run as an administrator."
        Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "This script must be run as an administrator."
        return
    }
    else {
        Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Starting script execution as administrator."
    }

    # Save parameters to a hashtable
    $script:reportGenerated = $false
    $script:disconnectFromSCC = $false
    $script:parameters = $PSBoundParameters
    $modules = @('Microsoft.Online.SharePoint.PowerShell', 'ExchangeOnlineManagement')

    foreach ($module in $modules) {
        try {
            # Check if the module is installed
            if (-not (Get-Module -ListAvailable -Name $module)) {
                # Install the module
                Install-Module -Name $module -Force -AllowClobber -ErrorAction SilentlyContinue
                Write-Verbose "Installed $module module."
                Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Installed $module module."
            }
            else {
                Write-Verbose "$module module already installed."
                Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "$module module already installed."
            }

            # Import the module
            if ($module -eq "ExchangeOnlineManagement" -and (Get-Module -Name $module -ListAvailable).Version -eq [version]"1.0.0.0") {
                Remove-Module -Name $module -Force -ErrorAction SilentlyContinue
                Write-Verbose "Removed ExchangeOnlineManagement module version 1.0.0.0."
                Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Removed ExchangeOnlineManagement module version 1.0.0.0"
            }

            if (-not (Get-Module -Name $module)) {
                if ($PSVersionTable.PSEdition -eq "Core" -and $module -eq "Microsoft.Online.SharePoint.PowerShell") {
                    Import-Module -Name $module -UseWindowsPowerShell -ErrorAction SilentlyContinue
                    Write-Verbose "Connecting with Windows PowerShell Core Version for $module."
                    Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Connecting with Windows PowerShell Core Version for $module."
                }
                else {
                    Import-Module -Name $module -ErrorAction SilentlyContinue
                    Write-Verbose "Connecting with Windows PowerShell Desktop Version for $module."
                    Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Connecting with Windows PowerShell Desktop Version for $module."
                }
            }
            else {
                Write-Verbose "$module module already imported."
                Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "$module module already imported."
            }
        }
        catch {
            Write-Output "$_"
            return
        }
    }

    # Check connection to SharePoint Online
    try {
        Write-Verbose "Checking for prior connection to SharePoint Online."
        Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Checking for prior connection to SharePoint Online."
        $connection = Get-SPOTenant -ErrorAction SilentlyContinue
        if (-not $connection) {
            Write-Output "Not connected to SharePoint Online. Attempting to connect to SharePoint Online"
            Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Not connected to SharePoint Online. Attempting to connect to SharePoint Online"
            Connect-SPOService -Url $TenantAdminUrl -ErrorAction SilentlyContinue
            Write-Output "Connected to SharePoint Online."
            Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Connected to SharePoint Online."
        }
        else {
            Write-Verbose "Already Connected to SharePoint Online."
            Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Already Connected to SharePoint Online."
        }
    }
    catch {
        write-Output "$_"
        return
    }

    try {
        if ($ReportEntity) {
            $reportEntities = $ReportEntity
        }
        else {
            $reportEntities = @('EveryoneExceptExternalUsersAtSite', 'EveryoneExceptExternalUsersForItems', 'SharingLinks_Anyone', 'SharingLinks_PeopleInYourOrg', 'SharingLinks_Guests', 'SensitivityLabelForFiles', 'PermissionedUsers')
        }

        # Build reports
        foreach ($entity in $reportEntities) {
            if ($ReportType -eq "Snapshot" -and $entity -ne "PermissionedUsers" -and $entity -ne "SensitivityLabelForFiles") {
                Write-Output "WARNING: ReportType 'Snapshot' is only valid for 'PermissionedUsers' and 'SensitivityLabelForFiles' entities."
                Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "WARNING: ReportType 'Snapshot' is only valid for 'PermissionedUsers' and 'SensitivityLabelForFiles' entities."
                return
            }

            if ($entity -eq "SensitivityLabelForFiles" -and ($Workload -eq "SharePoint" -or $Workload -eq "OneDriveForBusiness") -and $ReportType -eq "RecentActivity") {
                Write-Output "WARNING: ReportType 'RecentActivity' with ReportEntity 'SensitivityLabelForFiles' entity with 'SharePoint' workload at this time."
                Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "WARNING: ReportType 'RecentActivity' with ReportEntity 'SensitivityLabelForFiles' entity with 'SharePoint' workload at this time."
                return
            }

            # Check for sensitivity labels if specified.
            if ($entity -eq "SensitivityLabelForFiles" -and $CheckSensitivityLabel) {
                if (-not $UserPrincipalName) {
                    $UserPrincipalName = Read-Host "UserPrincipalName is required to connect to the Security & Compliance Center.`nPlease enter the UserPrincipalName"
                    if (-not $UserPrincipalName) {
                        Write-Output "UserPrincipalName not provided. Exiting."
                        Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "UserPrincipalName not provideed. Exiting."
                        $disconnectFromSCC = $true
                        return
                    }
                }

                Write-Output "Connecting to the Security & Compliance Center."
                Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Connecting to the Security & Compliance Center."
                Connect-IPPSSession -UserPrincipalName $UserPrincipalName -ShowBanner:$False -ErrorAction Stop
                Write-Output "Connected to the Security & Compliance Center. Obtaining labels from the Security & Compliance Center."
                Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Connected to the Security & Compliance Center. Obtaining labels from the Security & Compliance Center."
                $labels = Get-Label | Select-Object DisplayName, GUID -ErrorAction Stop
                Write-Verbose "Got labels from the Security & Compliance Center."
                Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Got labels from the Security & Compliance Center."

                do {
                    $labelMenu = $labels | ForEach-Object { "$($labels.IndexOf($_) + 1). $($_.DisplayName) - $($_.GUID)" }
                    $labelMenu | ForEach-Object { Write-Output $_ }
                    $selection = Read-Host "Select a label by entering the corresponding number or type 'c' to cancel"

                    if ($selection -eq 'c') {
                        Write-Output "Operation cancelled by user."
                        Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Operation cancelled by user."
                        return
                    }
                    $selectedLabel = $labels[$selection - 1]
                } while (-not $selectedLabel)

                Write-Verbose "Selected Label: $selectedLabel.DisplayName - $selectedLabel.GUID"
                Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Selected Label: $selectedLabel.DisplayName - $selectedLabel.GUID"
                Write-Output "Report for $($entity) with ReportType: $($ReportType) - Workload: $($Workload) - CountOfUsersMoreThan of $($CountOfUsersMoreThan) - FileSensitivityLabelGUID: $($selectedLabel.GUID) - FileSensitivityLabelName: $($selectedLabel.DisplayName) has been generated."
                Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Report for $($entity) with ReportType: $($ReportType) - Workload: $($Workload) - CountOfUsersMoreThan of $($CountOfUsersMoreThan) - FileSensitivityLabelGUID: $($selectedLabel.GUID) - FileSensitivityLabelName: $($selectedLabel.DisplayName) has been generated."
                $report = Start-SPODataAccessGovernanceInsight -ReportEntity SensitivityLabelForFiles -Workload $Workload -ReportType $ReportType -FileSensitivityLabelGUID $($selectedLabel.GUID) -FileSensitivityLabelName $($selectedLabel.DisplayName)
                if ($($report.ReportID)) { $reportGenerated = $true }
            }
            else {
                if ($Template) {
                    $report = Start-SPODataAccessGovernanceInsight -Name "$entity" -ReportEntity $entity -Workload $Workload -ReportType $ReportType -Template $Template -CountOfUsersMoreThan $CountOfUsersMoreThan -ErrorAction Stop
                    Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Report for $($entity) with ReportType: $($ReportType) - Workload: $($Workload) - CountOfUsersMoreThan of $($CountOfUsersMoreThan) - Template: $($Template) has been generated."
                    if ($($report.ReportID)) { $reportGenerated = $true }
                }
                else {
                    $report = Start-SPODataAccessGovernanceInsight -Name "$entity" -ReportEntity $entity -Workload $Workload -ReportType $ReportType -CountOfUsersMoreThan $CountOfUsersMoreThan -ErrorAction Stop
                    Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Report for $($entity) with ReportType: $($ReportType) - Workload: $($Workload) - CountOfUsersMoreThan of $($CountOfUsersMoreThan) has been generated."
                    if ($($report.ReportID)) { $reportGenerated = $true }
                }
            }

            # End user notifications
            if ($reportGenerated -eq $true) {
                Write-Output "To check the status of this report please run: Get-SPODataAccessGovernanceInsight -ReportID $($report.ReportID)`nTo download this report please run: Export-SPODataAccessGovernanceInsight -ReportID $($report.ReportID)"
                Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "To check the status of this report please run: Get-SPODataAccessGovernanceInsight -ReportID $($report.ReportID)`nTo download this report please run: Export-SPODataAccessGovernanceInsight -ReportID $($report.ReportID)"
                Write-Output "Report for $($entity) with ReportType: $($ReportType) - Workload: $($Workload) - CountOfUsersMoreThan of $($CountOfUsersMoreThan) has been generated."
                Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Report for $($entity) with ReportType: $($ReportType) - Workload: $($Workload) - CountOfUsersMoreThan of $($CountOfUsersMoreThan) has been generated."
            }
        }
    }
    catch {
        Write-Output "$_"
    }
    finally {
        # Disconnect from Security & Compliance Center if connected
        if ($CheckSensitivityLabel -or $disconnectFromSCC) {
            Write-Output "Disconnecting from the Security & Compliance Center."
            Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Disconnecting from the Security & Compliance Center."
            Disconnect-ExchangeOnline
        }
        else {
            Write-Output "Not disconnecting from the Security & Compliance Center."
            Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Not disconnecting from the Security & Compliance Center."
        }

        if ($DisconnectFromSPO -eq $True) {
            Write-Output "Disconnecting from the SPOService."
            Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Disconnecting from the SPOService."
            Disconnect-SPOService
        }
        else {
            Write-Output "Not disconnecting from the SPOService."
            Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Not disconnecting from the SPOService."
        }

        Write-Output "For more information please see the logging file: $($LoggingDirectory)\$($LoggingFilename)"
        Write-Output "Script completed."
        Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Script completed."
    }
}