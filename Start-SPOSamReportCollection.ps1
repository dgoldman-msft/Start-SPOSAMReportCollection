function Write-ToLog {
    <#
        .SYNOPSIS
            Save output

        .DESCRIPTION
            Overload function for Write-Output

        .PARAMETER LoggingDirectory
            Directory to save the log file to. Default is "$env:MyDocuments".

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
        Write-Output "$($InputString)"
        $stringObject = "[{0:MM/dd/yy} {0:HH:mm:ss}] - {1}" -f (Get-Date), $InputString
        Add-Content -Path (Join-Path $LoggingDirectory -ChildPath $LoggingFilename) -Value $stringObject -Encoding utf8 -ErrorAction Stop
        Write-Verbose "Logging to $($LoggingDirectory)\$($LoggingFilename)"
    }
    catch {
        Write-Output "$_"
        return
    }
}
function Get-ReportDescription {
    <#
        .SYNOPSIS
            Retrieves the description of a specified report entity.

        .DESCRIPTION
            The Get-ReportDescription function searches through an array of descriptions and returns the description that matches the specified report entity.

        .PARAMETER reportEntity
            The name or part of the name of the report entity to search for in the description array.

        .EXAMPLE
            PS C:\> Get-ReportDescription -reportEntity "Sales"

            This command retrieves the description for the report entity that contains "Sales" in its description.

        .NOTES
            The function uses the -like operator to perform a wildcard search for the report entity within the description array.
    #>

    [OutputType('System.String')]
    [CmdletBinding(DefaultParameterSetName = 'Default')]
    [Alias('SAMR')]
    param (
        [string] $reportEntity
    )
    return $descriptionArray | Where-Object { $_ -like "*$reportEntity*" }
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
            C:\PS> samr -TenantDomain contoso -ReportEntity All or
            C:\PS> Start-SPOSAMReportCollection -TenantDomain contoso -ReportEntity All

            This example will generate all reports by specifying full function name or alias
            NOTE: The only report that will not be generated is 'SensitivityLabelForFiles' unless you run the next example.
            This is intentional as this report requires a connection to the Security & Compliance Center.

        .EXAMPLE
            C:\PS> Start-SPOSAMReportCollection -TenantDomain contoso -UserPrincipalName Administrator@tenant -CheckSensitivityLabel -ReportType Snapshot

            This example will generate reports for SensitivityLabelForFiles.
            This will connect you to the Security and Compliance Center to read the labels in the tenant.

        .EXAMPLE
            C:\PS> Start-SPOSAMReportCollection -TenantDomain contoso -ReportEntity EveryoneExceptExternalUsersAtSite -CountOfUsersMoreThan 100

            This example will generate a report for EveryoneExceptExternalUsersAtSite with a threshold of 100 users. Default is 0

        .EXAMPLE
            C:\PS> Start-SPOSAMReportCollection -TenantDomain contoso -ReportEntity SharingLinks_Guests -Privacy Private

            This example will generate reports for SharePoints sites with links for SharingLinks_Guests.

        .EXAMPLE
            C:\PS> Start-SPOSAMReportCollection -TenantDomain contoso -ReportEntity SharingLinks_Anyone

            This example will generate a report for SharePoint sites with links for 'SharingLinks_Anyone'.

        .EXAMPLE
            C:\PS> Start-SPOSAMReportCollection -TenantDomain contoso -Workload OneDriveForBusiness

            This example will generate reports for a filtered the report for OneDrive for Business.

        .EXAMPLE
            C:\PS> Start-SPOSAMReportCollection -TenantDomain contoso -LoggingDirectory "C:\Logs" -LoggingFilename "SamReportingLogs.txt"

            This example will generate reports by saving the log file to "C:\Logs\SamReportingLogs.txt".

        .NOTES
            For more information please see: https://learn.microsoft.com/en-us/sharepoint/data-access-governance-reports
    #>

    [OutputType('System.String')]
    [CmdletBinding(DefaultParameterSetName = 'Default')]
    [Alias('SAMR')]
    param
    (
        [Parameter(ParameterSetName = 'Default', HelpMessage = 'Check for sensitivity labels on files in the site. This will make a connection to the Security & Compliance Center to retrieve labels.')]
        [switch]
        $CheckSensitivityLabel,

        [Parameter(ParameterSetName = 'Default', HelpMessage = 'Specifies the threshold of oversharing as defined by the number of users that can access the site. The number of users that can access the site are determined by expanding all users, groups across all permissions (at site level and at the level of any item with unique permissions), deduplicate and arrive at a unique number. Minimum value is 0.')]
        [Int]
        $CountOfUsersMoreThan = 0,

        [Parameter(ParameterSetName = 'Default', HelpMessage = 'Disconnect from SharePoint Online after the report collection is completed. Default is $false.')]
        [switch]
        $DisconnectFromSPO,

        [Parameter(ParameterSetName = 'Default', HelpMessage = 'Specifies the directory to save the log file to. Default is $env:MyDocuments\SamReporting.')]
        [string]
        $LoggingDirectory = (Join-Path -Path ([Environment]::GetFolderPath("MyDocuments")) -ChildPath "SamReporting"),

        [Parameter(ParameterSetName = 'Default', HelpMessage = 'Specifies the filename to save the log file to. Default is SamReportingLogs.txt.')]
        [string]
        $LoggingFilename = "SamReportingLogs.txt",

        [Parameter(ParameterSetName = 'Default', HelpMessage = 'Specifies the privacy setting of the Microsoft 365 group. Relevant in case of filtering the report for group connected sites. Valid values are: All, Private and Public. Default = All.')]
        [ValidateSet('All', 'Private', 'Public')]
        $Privacy = 'All',

        [Parameter(ParameterSetName = 'Default', HelpMessage = 'Specifies the entity that could cause oversharing and hence tracked by these reports. Valid values are: EveryoneExceptExternalUsersAtSite, EveryoneExceptExternalUsersForItems, SharingLinks_Anyone, SharingLinks_PeopleInYourOrg, SharingLinks_Guests, SensitivityLabelForFiles, PermissionedUsers.')]
        [ValidateSet('All', 'EveryoneExceptExternalUsersAtSite', 'EveryoneExceptExternalUsersForItems', 'SharingLinks_Anyone', 'SharingLinks_PeopleInYourOrg', 'SharingLinks_Guests', 'SensitivityLabelForFiles', 'PermissionedUsers')]
        [string]
        $ReportEntity,

        [Parameter(ParameterSetName = 'Default', HelpMessage = 'Specifies the time period of data based on which DAG report is generated. A [Snapshot] report will have the latest data as of the report generation time. A [RecentActivity] report will be based on data in the last 28 days. Default = RecentActivity.')]
        [ValidateSet('Snapshot', 'RecentActivity')]
        [string]
        $ReportType = 'RecentActivity',

        [Parameter(ParameterSetName = 'Default', HelpMessage = 'Specifies the template of the site. Relevant in case a report should be generated for that particular template. Valid values are: AllSites, ClassicSites, CommunicationSites, TeamSites, and OtherSites.')]
        [ValidateSet('AllSites', 'ClassicSites', 'CommunicationSites', 'TeamSites', 'OtherSites')]
        [string]
        $Template,

        [Parameter(Mandatory = $true, ParameterSetName = 'Default', HelpMessage = 'Specifies the domain of the tenant. This parameter is mandatory.')]
        [string]
        $TenantDomain,

        [Parameter(ParameterSetName = 'Default')]
        [string]
        $TenantAdminUrl = "https://$TenantDomain-admin.sharepoint.com",

        [Parameter(ParameterSetName = 'Default', HelpMessage = 'Specifies the workload for which the report should be generated. Valid values are [SharePoint] and [OneDriveForBusiness]. Default = SharePoint.')]
        [ValidateSet('SharePoint', 'OneDriveForBusiness')]
        [string]
        $Workload = "SharePoint",

        [Parameter(ParameterSetName = 'Default', HelpMessage = 'Specifies the username for authentication to the Security and Compliance Center.')]
        [string]
        $UserPrincipalName
    )

    $generateAllReports = $false
    $reportGenerated = $false
    $numOfReportsGenerated = 0
    $disconnectFromSCC = $false
    $modules = @('Microsoft.Online.SharePoint.PowerShell', 'ExchangeOnlineManagement')

    $descriptionArray = @(
        "PermissionedUsers - Report for sites that have shared content with permissioned users in last 28 days.",
        "EveryoneExceptExternalUsersAtSite - Report for sites that have shared content with Everyone except external users (EEEU) in last 28 days.",
        "EveryoneExceptExternalUsersForItems - Report for files, folders or lists that have shared content with Everyone except external users (EEEU) in last 28 days.",
        "SharingLinks_Anyone - Report for sites that have shared content with Anyone in last 28 days.",
        "SharingLinks_PeopleInYourOrg - Report for sites that have shared content with People in your organization in last 28 days.",
        "SharingLinks_Guests - Report for sites that have shared content with Guests in last 28 days.",
        "SensitivityLabelForFiles - Report for sites that have files with sensitivity labels."
    )

    function Get-ReportDescription {
        param (
            [string] $reportEntity
        )
        return $descriptionArray | Where-Object { $_ -like "*$reportEntity*" }
    }

    # Check if running as administrator
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "This script must be run as an administrator."
        return
    }
    else {
        Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Starting script execution as administrator."
    }

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
            Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Error: $_"
            return
        }
    }

    # Connection to SharePoint Online
    try {
        Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Connecting to SharePoint Online."
        Connect-SPOService -Url $TenantAdminUrl -ErrorAction SilentlyContinue
        Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Connected to SharePoint Online."
    }
    catch {
        Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Error: $_"
        return
    }

    try {
        if ($ReportEntity -eq 'All') {
            $script:generateAllReports = $true
            $reportEntities = @('EveryoneExceptExternalUsersAtSite', 'EveryoneExceptExternalUsersForItems', 'SharingLinks_Anyone', 'SharingLinks_PeopleInYourOrg', 'SharingLinks_Guests', 'SensitivityLabelForFiles', 'PermissionedUsers')
        }
        else {
            $reportEntities = @($ReportEntity)
        }

        # Build reports
        foreach ($entity in $reportEntities) {
            Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "`r`nGenerating report for $(Get-ReportDescription -reportEntity $entity)"

            try {
                if ($ReportType -eq "Snapshot" -and $entity -ne "PermissionedUsers" -and $entity -ne "SensitivityLabelForFiles" -and $generateAllReports -eq $true) {
                    Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "WARNING: ReportType 'Snapshot' is only valid for 'PermissionedUsers' and 'SensitivityLabelForFiles' entities."
                    if ($generateAllReports -eq $true) { continue } else { if ($reportEntities.Count -eq 1) { return } }
                }

                if ($entity -eq "SensitivityLabelForFiles" -and ($Workload -eq "SharePoint" -or $Workload -eq "OneDriveForBusiness") -and $ReportType -eq "RecentActivity") {
                    Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "WARNING: To run ReportType 'SensitivityLabelForFiles' ReportType must be set to 'SnapShot'."
                    if ($generateAllReports -eq $true) { continue } else { if ($reportEntities.Count -eq 1) { return } }
                }

                # Check for sensitivity labels if specified.
                if ($entity -eq "SensitivityLabelForFiles" -and $CheckSensitivityLabel) {
                    if (-not $UserPrincipalName) {
                        $UserPrincipalName = Read-Host "UserPrincipalName is required to connect to the Security & Compliance Center.`nPlease enter the UserPrincipalName"
                        if (-not $UserPrincipalName) {
                            Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "UserPrincipalName not provided. Exiting."
                            $disconnectFromSCC = $true
                            if ($generateAllReports -eq $true) { continue } else { if ($reportEntities.Count -eq 1) { return } }
                        }
                    }

                    Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Connecting to the Security & Compliance Center."
                    Connect-IPPSSession -UserPrincipalName $UserPrincipalName -ShowBanner:$False -ErrorAction Stop
                    Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Connected to the Security & Compliance Center. Obtaining labels from the Security & Compliance Center."
                    Write-Verbose "Getting labels from the Security & Compliance Center."
                    $labels = Get-Label | Select-Object DisplayName, GUID -ErrorAction Stop
                    Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Retrieved labels from the Security & Compliance Center."

                    do {
                        $labelMenu = $labels | ForEach-Object { "$($labels.IndexOf($_) + 1). $($_.DisplayName) - $($_.GUID)" }
                        $labelMenu | ForEach-Object { Write-Output $_ }
                        $selection = Read-Host "Select a label by entering the corresponding number or type 'c' to cancel"

                        if ($selection -eq 'c') {
                            Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Operation cancelled by user."
                            if ($generateAllReports -eq $true) { continue } else { if ($reportEntities.Count -eq 1) { return } }
                        }
                        $selectedLabel = $labels[$selection - 1]
                    } while (-not $selectedLabel)

                    Write-Verbose "Selected Label: $selectedLabel.DisplayName - $selectedLabel.GUID"
                    Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Selected Label: $selectedLabel.DisplayName - $selectedLabel.GUID"
                    $report = Start-SPODataAccessGovernanceInsight -ReportEntity SensitivityLabelForFiles -Workload $Workload -ReportType $ReportType -FileSensitivityLabelGUID $($selectedLabel.GUID) -FileSensitivityLabelName $($selectedLabel.DisplayName) -CountOfUsersMoreThan $CountOfUsersMoreThan
                    if ($($report.ReportID)) {
                        Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Report for $($entity) with ReportID: $($report.ReportID) ReportType: $($ReportType) - Workload: $($Workload) - CountOfUsersMoreThan of $($CountOfUsersMoreThan) - FileSensitivityLabelGUID: $($selectedLabel.GUID) - FileSensitivityLabelName: $($selectedLabel.DisplayName) has been generated."
                        $reportGenerated = $true
                        $numOfReportsGenerated ++
                    }
                }
                else {
                    if ($generateAllReports -eq $true -and $entity -eq "SensitivityLabelForFiles") {
                        Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "WARNING: To generate a report for SensitivityLabels, the ReportEntity 'SensitivityLabelForFiles' and the CheckSensitivityLabel flag must be True."
                        break
                    }
                }

                # Check for template if specified.
                if ($Template) {
                    $report = Start-SPODataAccessGovernanceInsight -Name $entity -ReportEntity $entity -Workload $Workload -ReportType $ReportType -Template $Template -CountOfUsersMoreThan $CountOfUsersMoreThan -ErrorAction Stop
                    if ($($report.ReportID)) {
                        Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Report for $($entity) with ReportType: $($ReportType) - Workload: $($Workload) - CountOfUsersMoreThan of $($CountOfUsersMoreThan) - Template: $($Template) has been generated."
                        $reportGenerated = $true
                        $numOfReportsGenerated ++
                    }
                }

                # PermissionedUsers report needs to have ReportType set to Snapshot
                if ($entity -eq "PermissionedUsers") {
                    $report = Start-SPODataAccessGovernanceInsight -Name $entity -ReportEntity $entity -Workload $Workload -ReportType Snapshot -CountOfUsersMoreThan $CountOfUsersMoreThan -ErrorAction Stop
                    if ($($report.ReportID)) {
                        Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Report for $($entity) with ReportType: Snapshot - Workload: $($Workload) - CountOfUsersMoreThan of $($CountOfUsersMoreThan) has been generated."
                        $reportGenerated = $true
                        $numOfReportsGenerated ++
                    }
                }
                else {
                    $report = Start-SPODataAccessGovernanceInsight -Name $entity -ReportEntity $entity -Workload $Workload -ReportType $ReportType -CountOfUsersMoreThan $CountOfUsersMoreThan -ErrorAction Stop
                    if ($($report.ReportID)) {
                        Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Report for $($entity) with ReportType: $($ReportType) - Workload: $($Workload) - CountOfUsersMoreThan of $($CountOfUsersMoreThan) has been generated."
                        $reportGenerated = $true
                        $numOfReportsGenerated ++
                    }
                }
            }
            catch {
                Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Error generating report for $($entity). $($_.Exception.Message)"
                if ($generateAllReports -eq $true) { continue } else { if ($reportEntities.Count -eq 1) { return } }
            }
        }

        # End user notifications
        if ($reportGenerated -eq $true) {
            Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "To check the status of this report please run: Get-SPODataAccessGovernanceInsight -ReportID $($report.ReportID)`nTo download this report please run: Export-SPODataAccessGovernanceInsight -ReportID $($report.ReportID)"
            Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Report for $($entity) with ReportType: $($ReportType) - Workload: $($Workload) - CountOfUsersMoreThan of $($CountOfUsersMoreThan) has been generated."
        }
    }
    catch {
        Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Error: $_"
    }
    finally {
        # Disconnect from Security & Compliance Center if connected
        Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "`r`n-----------------------------------------"
        if ($CheckSensitivityLabel -or $disconnectFromSCC) {
            Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Disconnecting from the Security & Compliance Center."
            Disconnect-ExchangeOnline
        }
        else {
            Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Not disconnecting from the Security & Compliance Center."
        }

        if ($DisconnectFromSPO -eq $True) {
            Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Disconnecting from the SPOService."
            Disconnect-SPOService
        }
        else {
            Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Not disconnecting from the SPOService."
        }

        Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "`r`nTotal reports generated: $($numOfReportsGenerated)"
        Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "`r`nFor more information please see the logging file: $($LoggingDirectory)\$($LoggingFilename)"
        Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Script completed."
    }
}