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
    C:\PS> Start-SPOSAMReportCollection -TenantDomain contoso -UserPrincipalName Administrator -CheckSensitivityLabel

    This example will generate reports for all entities with the default parameters as well as check the sensitivity labels.
    This will connect you to the Security and Compliance Center to read the labels in the tenant.

## EXAMPLE 2
    C:\PS> Start-SPOSAMReportCollection -TenantDomain contoso -UserPrincipalName Administrator -CountOfUsersMoreThan 100

    This example will generate reports for all entities with the default parameters and a threshold of 100 users.

## EXAMPLE 3
    C:\PS> Start-SPOSAMReportCollection -TenantDomain contoso -UserPrincipalName Administrator -Privacy Private

    This example will generate reports for all entities with the default parameters and filter the report for private sites.

## EXAMPLE 4
    C:\PS> Start-SPOSAMReportCollection -TenantDomain contoso -UserPrincipalName Administrator -ReportEntity SharingLinks_Anyone

    This example will generate a report for the 'SharingLinks_Anyone' entity with the default parameters.

## EXAMPLE 5
    C:\PS> Start-SPOSAMReportCollection -TenantDomain contoso -UserPrincipalName Administrator -ReportType Snapshot

    This example will generate a snapshot report for all entities with the default parameters.

## EXAMPLE 6
    C:\PS> Start-SPOSAMReportCollection -TenantDomain contoso -UserPrincipalName Administrator -Template TeamSites

    This example will generate reports for all entities with the default parameters and filter the report for team sites.

## EXAMPLE 7
    C:\PS> Start-SPOSAMReportCollection -TenantDomain contoso -UserPrincipalName Administrator -Workload OneDriveForBusiness

    This example will generate reports for all entities with the default parameters and filter the report for OneDrive for Business.

## EXAMPLE 8
    C:\PS> Start-SPOSAMReportCollection -TenantDomain contoso -UserPrincipalName Administrator -DisconnectFromSPO

    This example will generate reports for all entities with the default parameters and disconnect from SharePoint Online after the report collection is completed.

## NOTES
- For more information please see: https://learn.microsoft.com/en-us/sharepoint/data-access-governance-reports
- This script will install both the Microsoft.Online.SharePoint.PowerShell and ExchangeOnlineManagement modules if they are not installed.