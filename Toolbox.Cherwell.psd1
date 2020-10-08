@{
    ModuleVersion     = '3.5'

    RootModule        = 'Toolbox.Cherwell.psm1'
    GUID              = '771e242d-88c0-48fc-92a6-3fea84c80062'
    Author            = 'Brecht.Gijbels@heidelbergcement.com'
    CompanyName       = 'HeidelbergCement'
    Copyright         = '(c) 2019 Brecht.Gijbels@heidelbergcement.com. All rights reserved.'
    Description       = 'Interact with the Cherwell REST API in an easy and convenient way.'
    PowerShellVersion = '5.1'

    FunctionsToExport = @(
        'Add-CherwellTicketConfigItemHC', 'Add-CherwellTicketAttachmentHC',
        'Add-CherwellTicketDetailHC',
        'Add-CherwellTicketJournalNoteHC', 'Add-CherwellTicketTaskHC',
        'Get-CherwellConfigItemHC', 'Get-CherwellConfigItemTypeHC', 'Get-CherwellCustomerHC',
        'Get-CherwellLocationHC', 'Get-CherwellTeamInfoHC', 'Get-CherwellIncidentSubCategoryHC',
        'Get-CherwellTicketHC', 'Get-CherwellTicketDetailHC', 'Get-CherwellQuickCallTemplateHC',
        'Get-CherwellTicketJournalNoteHC', 'Get-CherwellTicketTaskHC',
        'Get-CherwellServiceHC', 'Get-CherwellServiceCatalogTemplateHC',
        'Get-CherwellSupplierHC', 'Get-CherwellSlaHC', 'Get-CherwellChangeStandardTemplateHC',
        'Get-CherwellSystemUserHC',
        'Get-CherwellIncidentCategoryHC',
        'New-CherwellConfigItemHC', 'New-CherwellTicketHC',
        'Remove-CherwellTicketHC',
        'Set-CherwellTicketDetailHC',
        'Update-CherwellTicketHC')

    CmdletsToExport   = @()
    VariablesToExport = $null
    AliasesToExport   = @()

    PrivateData       = @{
        PSData = @{
            ReleaseNotes = "
           Small bug fixes
           Added function Get-CherwellSystemUserHC and Get-CherwellCustomerHC
           Removed module scoped function Get-CustomerRecID and Get-UserRecID
2019/09/27 Released version 2.3
           Added Get-CherwellLocationHC, Get-CherwellServiceHC, Get-CherwellIncidentCategoryHC, Get-CherwellTeamInfoHC, Get-CherwellSupplierHC, Get-CherwellSlaHC, ...
           Improved all Get functions
           Reduced the amount of calls to the REST API
           Removed credentials from the code, they're now in a separate file 'Password.json'
           Added better and more advanced examples to the help sections
2019/10/04 Released version 3.0
           Added Get-CherwellIncidentSubCategoryHC, Get-CherwellQuickCallTemplateHC
           Added Get-CherwellServiceCatalogTemplateHC, Get-CherwellChangeStandardTemplateHC
           Removed '-IncludeAllProperties' as this is replaced by '-Property *'
2019/10/08 Released version 3.1
           Added extra tests to the import of the password file
           Improved speed when calls are made to different environments by saving data for each environment
           Simplified authentication functions
2019/10/10 Released version 3.2
           Added Add-CherwellTicketJournalNoteHC, Add-CherwellTicketTaskHC, Add-CherwellTicketDetailHC
           Added Get-CherwellTicketJournalNoteHC, Get-CherwellTicketTaskHC
           Improved all functions by only calling the API when the info is not available
           Simplified code and module manifest
           Speed up execution time by only acquiring what we need for each function
2019/11/27 Released version 3.3
2020/06/05 Changed `$null to empty string instead                  
           Set-CherwellTicketDetailHC @testParams -Ticket `$TestTicketObject -Type Specifics -KeyValuePair @{ 
                Notes = ''
            }
            Changed if setting OwnedBy you also need to set OwnedByTeam, 
            Addded OwnedBy test in case OwnedByTeam is missing
2020/10/06 Maded tests compatible with Pester 5
           Updated ehanced 'Update-CherwellTicketHC' to accept a filter for OwnedBy
"
        }
    }
}

