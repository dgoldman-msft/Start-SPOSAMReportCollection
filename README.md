# Start-SPOSAMReportCollection

Starts the SharePoint Online Security and Access Management (SAM) report collection.

## DESCRIPTION

This cmdlet is used to generate DAG reports which deal with potential oversharing of sensitive data.
These reports are present in Sharepoint admin center. Reports are currently available for the following scenarios:

Sharing links created in last 28 days (Anyone, People-in-your-org, Specific people shared externally).
Content shared with Everyone except external users (EEEU) in last 28 days.
List of sites having labelled files, as of report generation time.
List of sites having 'too-many-users', as of report generation time, to setup an oversharing baseline.

## How to get started with Start-SPOSAMReportCollection

1. Download this in to a directory of your choice
2. Navigate to that directory and run run: Import-Module -Name .\Start-SPOSAMReportCollection.ps1 (this will import in to the local PowerShell session)
3. Run one of the examples below and allow for tab completion to see all of the available options

## EXAMPLE 1
    C:\PS> samr -TenantDomain contoso -ReportEntity All or
    C:\PS> Start-SPOSAMReportCollection -TenantDomain contoso -ReportEntity All

    This example will generate all reports by specifying full function name or alias
    NOTE: The only report that will not be generated is 'SensitivityLabelForFiles' unless you run the next example.
    This is intentional as this report requires a connection to the Security & Compliance Center.

## EXAMPLE 2
    C:\PS> Start-SPOSAMReportCollection -TenantDomain contoso -UserPrincipalName Administrator@tenant -CheckSensitivityLabel -ReportType Snapshot

    This example will generate reports for SensitivityLabelForFiles.
    This will connect you to the Security and Compliance Center to read the labels in the tenant.

## EXAMPLE 3
    C:\PS> Start-SPOSAMReportCollection -TenantDomain contoso -ReportEntity EveryoneExceptExternalUsersAtSite -CountOfUsersMoreThan 100

    This example will generate a report for EveryoneExceptExternalUsersAtSite with a threshold of 100 users. Default is 0

## EXAMPLE 4
    C:\PS> Start-SPOSAMReportCollection -TenantDomain contoso -ReportEntity SharingLinks_Guests -Privacy Private

    This example will generate reports for SharePoints sites with links for SharingLinks_Guests.

## EXAMPLE 5
    C:\PS> Start-SPOSAMReportCollection -TenantDomain contoso -ReportEntity SharingLinks_Anyone

    This example will generate a report for SharePoint sites with links for 'SharingLinks_Anyone'.

## EXAMPLE 6
    C:\PS> Start-SPOSAMReportCollection -TenantDomain contoso -Workload OneDriveForBusiness

    This example will generate reports for a filtered the report for OneDrive for Business.

## EXAMPLE 6
    C:\PS> Start-SPOSAMReportCollection -TenantDomain contoso -LoggingDirectory "C:\Logs" -LoggingFilename "SamReportingLogs.txt"

    This example will generate reports by saving the log file to "C:\Logs\SamReportingLogs.txt".

## NOTES
- For more information please see: https://learn.microsoft.com/en-us/sharepoint/data-access-governance-reports
- This script will install both the Microsoft.Online.SharePoint.PowerShell and ExchangeOnlineManagement modules if they are not installed.
- Default logging directory is c:\users\username\Documents\SamReportingLogs.txt