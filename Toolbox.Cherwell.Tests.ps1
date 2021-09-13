#Requires -Modules Pester
#Requires -Version 5.1

BeforeDiscovery {
    # used by inModuleScope
    $testModule = $PSCommandPath.Replace('.Tests.ps1', '.psm1')
    $testModuleName = $testModule.Split('\')[-1].TrimEnd('.psm1')

    Remove-Module $testModuleName -Force -Verbose:$false -EA Ignore
    Import-Module $testModule -Force -Verbose:$false
}
BeforeAll {
    $testParams = @{
        Environment = 'Stage'
        ErrorAction = 'Stop'
    }
    $testMandatoryFields = @{
        ServiceCountryCode = 'BNL'
        Service            = 'END USER WORKPLACE'
        Category           = 'Category t.b.d.'
        SubCategory        = 'Submit Incident'
        Priority           = '2'
    }
    $testConfigItems = @(
        @{
            CIType      = 'ConfigServer'
            CIStatus     = 'Active'
            AssetTag    = '807584'
            AssetType   = 'Virtual Server'
            AssetStatus = 'New'
            IPAddress   = '192.168.1.1'
            HostName    = $env:COMPUTERNAME
            Model       = 'VmWare Virtual Platform'
        Location    = $null
        }
        @{
            CIType      = 'ConfigSystem'
            CIStatus     = 'Active'
            Description = 'Not so many of these, speeds up the test'
        }
    )
    # should be in BeforeDiscovery https://github.com/pester/Pester/issues/1705
    # available in Pester 5.1.0-beta1, for now copying where needed
    $testUsers = @(
        @{
            SamAccountName     = 'gijbelsb'
            OwnedByTeam        = 'BNL INFRA'
            CherwellSystemUser = Get-CherwellSystemUserHC @testParams -Filter @{
                FieldName  = 'SamAccountName'
                Operator   = 'eq'
                FieldValue = 'gijbelsb'
            } -PassThru
            CherwellCustomer   = Get-CherwellCustomerHC @testParams -Filter @{
                FieldName  = 'SamAccountName'
                Operator   = 'eq'
                FieldValue = 'gijbelsb'
            } -PassThru
        }
        @{
            SamAccountName     = 'dverhuls'
            OwnedByTeam        = 'BNL INFRA'
            CherwellSystemUser = Get-CherwellSystemUserHC @testParams -Filter @{
                FieldName  = 'SamAccountName'
                Operator   = 'eq'
                FieldValue = 'dverhuls'
            } -PassThru
            CherwellCustomer   = Get-CherwellCustomerHC @testParams -Filter @{
                FieldName  = 'SamAccountName'
                Operator   = 'eq'
                FieldValue = 'dverhuls'
            } -PassThru
        }
    )
}
Describe 'ticket details' {
    BeforeAll {
        $TestTicketNr1 = New-CherwellTicketHC @testParams -KeyValuePair (
            @{
                IncidentType            = 'Incident'
                RequesterSamAccountName = $testUsers[0].SamAccountName
                OwnedByTeam             = $testUsers[0].OwnedByTeam
                ShortDescription        = 'Pester automated test'
                Description             = 'Get-CherwellTicketHC'
                Source                  = 'Event'
            } + $testMandatoryFields)
    }
    Describe 'Add-CherwellTicketDetailHC' {
        It 'Add-CherwellTicketJournalNoteHC' {
            $testNotesBefore = Get-CherwellTicketDetailHC @testParams -Ticket $TestTicketNr1 -Type JournalNote

            Add-CherwellTicketJournalNoteHC @testParams -Ticket $TestTicketNr1 -KeyValuePair @{
                Details = 'This is a journal note'
            }

            $testNotesAfter = Get-CherwellTicketDetailHC @testParams -Ticket $TestTicketNr1 -Type JournalNote

            $testNotesAfter | Should -HaveCount (1 + $testNotesBefore.Count)
        }
        It 'Add-CherwellTicketTaskHC' {
            $testBefore = Get-CherwellTicketDetailHC @testParams -Ticket $TestTicketNr1 -Type Task

            Add-CherwellTicketTaskHC @testParams -Ticket $TestTicketNr1 -KeyValuePair @{
                Title       = 'Task title'
                Description = 'Task description'
            }

            $testAfter = Get-CherwellTicketDetailHC @testParams -Ticket $TestTicketNr1 -Type Task

            $testAfter | Should -HaveCount (1 + $testBefore.Count)
        }
    }
    Describe 'Get-CherwellTicketDetailHC' {
        It 'Get-CherwellTicketJournalNoteHC' {
            Get-CherwellTicketJournalNoteHC @testParams -Ticket $TestTicketNr1 | Should -Not -BeNullOrEmpty
        }
        It 'Get-CherwellTicketTaskHC' {
            Get-CherwellTicketTaskHC @testParams -Ticket $TestTicketNr1 | Should -Not -BeNullOrEmpty
        }
    }
}
Describe 'Invoke-GetSearchResultsHC' {
    $TestCases = @(
        @{
            CmdLet   = 'Get-CherwellTicketHC'
            Filter   = @{
                FieldName  = 'IncidentID'
                Operator   = 'eq'
                FieldValue = '1120690'
            }
            Property = 'RecId'
        }
        @{
            CmdLet   = 'Get-CherwellCustomerHC'
            Filter   = @{
                FieldName  = 'LastName'
                Operator   = 'eq'
                FieldValue = 'Green'
            }
            Property = 'RecId'
        }
        @{
            CmdLet   = 'Get-CherwellChangeStandardTemplateHC'
            Filter   = @{
                FieldName  = 'TemplateName'
                Operator   = 'contains'
                FieldValue = 'Install'
            }
            Property = 'RecId'
        }
        @{
            CmdLet   = 'Get-CherwellIncidentCategoryHC'
            Filter   = @{
                FieldName  = 'Country'
                Operator   = 'eq'
                FieldValue = 'BNL'
            }
            Property = 'RecId'
        }
        @{
            CmdLet   = 'Get-CherwellIncidentSubCategoryHC'
            Filter   = @{
                FieldName  = 'Country'
                Operator   = 'eq'
                FieldValue = 'BNL'
            }
            Property = 'RecId'
        }
        @{
            CmdLet   = 'Get-CherwellQuickCallTemplateHC'
            Filter   = @{
                FieldName  = 'ServiceCountry'
                Operator   = 'eq'
                FieldValue = 'BNL'
            }
            Property = 'RecId'
        }
        @{
            CmdLet   = 'Get-CherwellLocationHC'
            Filter   = @{
                FieldName  = 'Country'
                Operator   = 'eq'
                FieldValue = 'BEL'
            }
            Property = 'RecId'
        }
        @{
            CmdLet   = 'Get-CherwellServiceHC'
            Filter   = @{
                FieldName  = 'Country'
                Operator   = 'eq'
                FieldValue = 'BNL'
            }
            Property = 'RecId'
        }
        @{
            CmdLet   = 'Get-CherwellServiceCatalogTemplateHC'
            Filter   = @{
                FieldName  = 'Country'
                Operator   = 'eq'
                FieldValue = 'BNL'
            }
            Property = 'RecId'
        }
        @{
            CmdLet   = 'Get-CherwellSupplierHC'
            Filter   = @{
                FieldName  = 'SupplierName'
                Operator   = 'startswith'
                FieldValue = 'Del'
            }
            Property = 'RecId'
        }
        @{
            CmdLet   = 'Get-CherwellSystemUserHC'
            Filter   = @{
                FieldName  = 'SamAccountName'
                Operator   = 'eq'
                FieldValue = 'gijbelsb'
            }
            Property = 'RecId'
        }
        @{
            CmdLet   = 'Get-CherwellTeamInfoHC'
            Filter   = @{
                FieldName  = 'Name'
                Operator   = 'startswith'
                FieldValue = 'BNL'
            }
            Property = 'RecId'
        }
        @{
            CmdLet   = 'Get-CherwellSlaHC'
            Filter   = @{
                FieldName  = 'Title'
                Operator   = 'startswith'
                FieldValue = 'DEU'
            }
            Property = 'RecId'
        }
    ) | Sort-Object { $_.CmdLet }
    Context '<CmdLet>' -Foreach $TestCases {
        It 'Filter' {
            $Actual = & $CmdLet @testParams -Filter $Filter
            $Actual | Should -Not -BeNullOrEmpty
        }
        It 'Property' {
            $Actual = & $CmdLet @testParams -Filter $Filter -Property $Property
            $Actual | Should -Not -BeNullOrEmpty
        }
        It 'Property *' {
            $Actual = & $CmdLet @testParams -Filter $Filter -Property *
            $Actual.LastModBy | Should -Not -BeNullOrEmpty
        }
        It 'PageSize' {
            $Actual = & $CmdLet @testParams -Filter $Filter -PassThru -PageSize 10
            $Actual.busObRecId | Should -Not -BeNullOrEmpty
        }
        It 'PassThru' {
            $Actual = & $CmdLet @testParams -Filter $Filter -PassThru
            $Actual.busObRecId | Should -Not -BeNullOrEmpty
        }
    }
}
Describe 'Get-CherwellConfigItemTypeHC' {
    It 'returns all found CI types' {
        $Actual = Get-CherwellConfigItemTypeHC @testParams
        $Actual | Should -Not -BeNullOrEmpty
    }
}
Describe 'New-CherwellConfigItemHC' {
    Context 'an error is thrown when' {
        It 'the Type is unknown' {
            {
                New-CherwellConfigItemHC @testParams -Type 'NotExisting' -KeyValuePair $testConfigItems[0]
            } |
            Should -Throw -PassThru |
            Should -BeLike "*Only the following CI types are supported*"
        }
    }
    It 'create a new CI' {
        $testCI = $testConfigItems[0].Clone()
        $testCI.remove('CIType')
        $testCI.FriendlyName = '{0}duplicate' -f $testCI.FriendlyName

        $Actual = New-CherwellConfigItemHC @testParams -KeyValuePair $testCI -Type $testConfigItems[0].CIType

        $Actual | Should -Not -BeNullOrEmpty
    }
} -Tag 'ci'
Describe 'Get-CherwellConfigItemHC' {
    Context 'an error is thrown when' {
        It 'the Type is unknown' {
            {
                Get-CherwellConfigItemHC @testParams -Type 'NotExisting' -Filter @{
                    FieldName  = 'HostName'
                    Operator   = 'eq'
                    FieldValue = 'Kiwi'
                }
            } |
            Should -Throw -PassThru |
            Should -BeLike "*Only the following CI types are supported*"
        }
    }
    Context 'when Type is provided and' {
        Context 'no filter is used' {
            It 'all CI objects are returned for that specific CI type' {
                $Actual = Get-CherwellConfigItemHC @testParams -Type $testConfigItems[1].CIType

                $Actual | Should -Not -BeNullOrEmpty
                $Actual.Count | Should -BeGreaterThan 1
            }
        }
        Context 'filter is used' {
            It 'a CI is returned for that CI types' {
                $Actual = Get-CherwellConfigItemHC @testParams -Type $testConfigItems[0].CIType -Filter @{
                    FieldName  = 'HostName'
                    Operator   = 'eq'
                    FieldValue = $env:COMPUTERNAME
                }

                $Actual | Should -Not -BeNullOrEmpty
                $Actual[0].HostName | Should -Be $env:COMPUTERNAME
            }
            It 'nothing is returned when there is no match' {
                $Actual = Get-CherwellConfigItemHC @testParams -Type 'ConfigPrinter' -Filter @{
                    FieldName  = 'HostName'
                    Operator   = 'eq'
                    FieldValue = 'NotExistingPrinterName'
                }

                $Actual | Should -BeNullOrEmpty
            }
        }
        Context 'InputObject is used' {
            It 'a CI is returned' {
                $testCI = $testConfigItems[0].Clone()
                $testFriendlyName = '{0}duplicate' -f $testCI.FriendlyName
                $testCI.remove('CIType')
                $testCI.FriendlyName = $testFriendlyName

                $testCI = New-CherwellConfigItemHC @testParams -KeyValuePair $testCI -Type $testConfigItems[0].CIType

                $Actual = Get-CherwellConfigItemHC @testParams -Type $testConfigItems[0].CIType -InputObject $testCI

                $Actual | Should -Not -BeNullOrEmpty
                $Actual | Should -HaveCount 1
            }
        }
    }
} -Tag 'ci'
Describe 'Get-CherwellTicketHC' {
    BeforeAll {
        $TestTicketNr1 = New-CherwellTicketHC @testParams -KeyValuePair (
            @{
                IncidentType            = 'Incident'
                RequesterSamAccountName = $testUsers[0].SamAccountName
                OwnedByTeam             = $testUsers[0].OwnedByTeam
                ShortDescription        = 'Pester automated test'
                Description             = 'Get-CherwellTicketHC'
                Source                  = 'Event'
            } + $testMandatoryFields)

        $TestTicketNr2 = New-CherwellTicketHC @testParams -KeyValuePair (
            @{
                IncidentType            = 'Incident'
                RequesterSamAccountName = $testUsers[0].SamAccountName
                OwnedByTeam             = $testUsers[0].OwnedByTeam
                ShortDescription        = 'Pester automated test'
                Description             = 'Get-CherwellTicketHC'
                Source                  = 'Event'
            } + $testMandatoryFields)
    }
    Context 'get the ticket business object for' {
        It 'one TicketNr' {
            $Actual = Get-CherwellTicketHC @testParams -TicketNr $TestTicketNr1
            $Actual | Should -HaveCount 1
        }
        It 'multiple TicketNr' {
            $Actual = Get-CherwellTicketHC @testParams -TicketNr $TestTicketNr1, $TestTicketNr2
            $Actual | Should -HaveCount 2
        }
    }
    Context "throw a terminating error when" {
        It "requesting more than 1000 tickets at the same time, the API can't handle this" {
            {
                Get-CherwellTicketHC @testParams -TicketNr  @(0..1001)
            } | 
            Should -Throw -PassThru |
            Should -BeLike "*Cannot validate argument on parameter 'TicketNr'. The number of provided arguments, (1002), exceeds the maximum number of allowed arguments (1000). Provide fewer than 1000 arguments, and then try the command again*"
        }
    }
}
Describe 'Set-CherwellTicketDetailHC' {
    BeforeAll {
        $TestTicketObject = New-CherwellTicketHC @testParams -PassThru -KeyValuePair (@{
                IncidentType              = 'Incident'
                RequesterSamAccountName   = $testUsers[0].SamAccountName
                SubmittedBySamAccountName = $testUsers[0].SamAccountName
                Status                    = 'New'
                ShortDescription          = 'Automated test'
                Description               = 'Pester test set/get additional details'
                Source                    = 'Event'
            } + $testMandatoryFields)

        $TestTicketObject | Should -Not -BeNullOrEmpty

        $testText = 'Text for Additional details'
        $testUpdatedText = 'Updated text for Additional details'
    }
    Context 'with TicketNr' {
        It "set 'Additional details'" {
            {
                Set-CherwellTicketDetailHC @testParams -Ticket $TestTicketObject -Type Specifics -KeyValuePair @{
                    Notes = $testText
                }
            } | Should -Not -Throw
        }
        It "get 'Additional details'" {
            (Get-CherwellTicketDetailHC @testParams -Ticket $TestTicketObject -Type Specifics).Notes |
            Should -Be $testText
        }
        It "update 'Additional details'" {
            Set-CherwellTicketDetailHC @testParams -Ticket $TestTicketObject -Type Specifics -KeyValuePair @{
                Notes = $testUpdatedText
            }
        }
        It "get 'Additional details' updated" {
            (Get-CherwellTicketDetailHC @testParams -Ticket $TestTicketObject -Type Specifics).Notes |
            Should -Be $testUpdatedText
        }
        It "remove 'Additional details'" {
            {
                Set-CherwellTicketDetailHC @testParams -Ticket $TestTicketObject -Type Specifics -KeyValuePair @{
                    Notes = ''
                }
            } | Should -Not -Throw
        }
        It "get 'Additional details' with no output" {
            (Get-CherwellTicketDetailHC @testParams -Ticket $TestTicketObject -Type Specifics).Notes |
            Should -BeNullOrEmpty
        }
    }
    Context 'with a piped ticket object' {
        It "set 'Additional details'" {
            {
                $TestTicketObject | Set-CherwellTicketDetailHC @testParams -Type Specifics -KeyValuePair @{
                    Notes = $testText
                }
            } | Should -Not -Throw
        }
        It "get 'Additional details'" {
            ($TestTicketObject | Get-CherwellTicketDetailHC @testParams -Type Specifics).Notes |
            Should -Be $testText
        }
    }
}
Describe 'Get-CherwellTicketDetailHC' {
    BeforeAll {
        $TestTicketObject = New-CherwellTicketHC @testParams -PassThru -KeyValuePair (@{
                IncidentType              = 'Incident'
                RequesterSamAccountName   = $testUsers[0].SamAccountName
                SubmittedBySamAccountName = $testUsers[0].SamAccountName
                Status                    = 'New'
                ShortDescription          = 'Automated test'
                Description               = 'Pester test set/get additional details'
                Source                    = 'Event'
            } + $testMandatoryFields)

        $TestTicketObject | Should -Not -BeNullOrEmpty
    }
    It 'get the Journals entries from a ticket' {
        $Actual = Get-CherwellTicketDetailHC @testParams -Ticket $TestTicketObject.busObPublicId -Type 'Journals' -PassThru
        $Actual | Should -Not -BeNullOrEmpty
    }
}
Describe 'New-CherwellTicketHC' {
    Context 'throw a non terminating error when' {
        It 'IncidentType is missing' { {
                New-CherwellTicketHC @testParams -KeyValuePair (@{
                        #IncidentType = 'Incident'
                        RequesterSamAccountName   = $testUsers[0].SamAccountName
                        SubmittedBySamAccountName = $testUsers[0].SamAccountName
                        ShortDescription          = 'Automated test'
                        Description               = 'Pester test IncidentType missing'
                        Source                    = 'Event'
                        Attachment                = $null
                    } + $testMandatoryFields) 
            } | Should -Throw
        }
        It 'RequesterSamAccountName or CustomerRecID is missing' { {
                New-CherwellTicketHC @testParams -KeyValuePair (@{
                        IncidentType     = 'Incident'
                        #RequesterSamAccountName = $testUsers[0].SamAccountName
                        Status           = 'New'
                        ShortDescription = 'Automated test'
                        Description      = 'Pester test description'
                        Source           = 'Event'
                    } + $testMandatoryFields) 
            } | 
            Should -Throw -PassThru |
            Should -BeLike "*The field 'RequesterSamAccountName' or 'CustomerRecID' is mandatory*"
        }
        It 'OwnedByTeam is missing when OwnedBy is set' { 
            {
                New-CherwellTicketHC @testParams -KeyValuePair (@{
                        IncidentType              = 'Incident'
                        RequesterSamAccountName   = $testUsers[0].SamAccountName
                        SubmittedBySamAccountName = $testUsers[0].SamAccountName
                        OwnedBy                   = $testUsers[0].SamAccountName
                        # OwnedByTeam      = $testUsers[0].OwnedByTeam
                        ShortDescription          = 'Automated test'
                        Description               = 'Pester test description without attachment'
                        Source                    = 'Event'
                        Attachment                = $null
                    } + $testMandatoryFields) 
            } | Should -Throw
        } 
        It 'Test-InvalidPropertyCombinationHC is called' {
            Mock Test-InvalidPropertyCombinationHC
            
            New-CherwellTicketHC @testParams -KeyValuePair (@{
                    IncidentType            = 'Incident'
                    RequesterSamAccountName = $testUsers[0].SamAccountName
                    Status                  = 'New'
                    ShortDescription        = 'Automated test'
                    Description             = 'Pester test incorrect properties'
                    Source                  = 'Event'
                } + $testMandatoryFields) 
            
            Should -Invoke Test-InvalidPropertyCombinationHC -Exactly -Times 1
        }
    }
    Context 'create a new ticket with' {
        Context "'Primary CI' and 'Linked Config Item' where" {
            It 'CI is an object coming from Get-CherwellConfigItemHC' {
                $testCI = Get-CherwellConfigItemHC -PassThru @testParams -Type $testConfigItems[0].CIType -Filter @{
                    FieldName  = 'HostName'
                    Operator   = 'eq'
                    FieldValue = $env:COMPUTERNAME
                }

                $testCI | Should -Not -BeNullOrEmpty

                $TicketNr = New-CherwellTicketHC @testParams -KeyValuePair (@{
                        IncidentType            = 'Incident'
                        RequesterSamAccountName = $testUsers[0].SamAccountName
                        Status                  = 'New'
                        ShortDescription        = 'Automated test primary CI as CI was provided as object by Get-CherwellConfigItemHC'
                        Description             = 'Pester test Primary CI'
                        Source                  = 'Event'
                        CI                      = $testCI[0]
                    } + $testMandatoryFields) 

                $TicketNr | Should -Not -BeNullOrEmpty

                $Actual = Get-CherwellTicketHC @testParams -TicketNr $TicketNr -Property ConfigItemDisplayName, ConfigItemRecID

                $Actual.ConfigItemRecID | Should -Not -BeNullOrEmpty
                $Actual.ConfigItemDisplayName | Should -Not -BeNullOrEmpty
                $Actual.ConfigItemDisplayName | 
                Should -Be "$env:COMPUTERNAME.$env:USERDNSDOMAIN"
            } -Tag test
            It 'CI is a hash table' {
                $TicketNr = New-CherwellTicketHC @testParams -KeyValuePair (@{
                        IncidentType            = 'Incident'
                        RequesterSamAccountName = $testUsers[0].SamAccountName
                        Status                  = 'New'
                        ShortDescription        = 'Automated test primary CI as CI was found from hash table'
                        Description             = 'Pester test Primary CI'
                        Source                  = 'Event'
                        CI                      = @{
                            Type   = $testConfigItems[0].CIType
                            Filter = @{
                                FieldName  = 'HostName'
                                Operator   = 'eq'
                                FieldValue = $env:COMPUTERNAME
                            }
                        }
                    } + $testMandatoryFields) 

                $TicketNr | Should -Not -BeNullOrEmpty

                $Actual = Get-CherwellTicketHC @testParams -TicketNr $TicketNr -Property ConfigItemDisplayName, ConfigItemRecID

                $Actual.ConfigItemRecID | Should -Not -BeNullOrEmpty
                $Actual.ConfigItemDisplayName | Should -Not -BeNullOrEmpty
                $Actual.ConfigItemDisplayName | 
                Should -Be "$env:COMPUTERNAME.$env:USERDNSDOMAIN"
            }
            It 'CI a hash table but the CI is not found, so no CI is set' {
                $TicketNr = New-CherwellTicketHC @testParams -KeyValuePair (@{
                        IncidentType            = 'Incident'
                        RequesterSamAccountName = $testUsers[0].SamAccountName
                        Status                  = 'New'
                        ShortDescription        = 'Automated test no primary CI as CI was not found from hash table'
                        Description             = 'Pester test Primary CI'
                        Source                  = 'Event'
                        CI                      = @{
                            Type   = $testConfigItems[0].CIType
                            Filter = @{
                                FieldName  = 'HostName'
                                Operator   = 'eq'
                                FieldValue = 'NotExistingCI'
                            }
                        }
                    } + $testMandatoryFields)

                $TicketNr | Should -Not -BeNullOrEmpty

                $Actual = Get-CherwellTicketHC @testParams -TicketNr $TicketNr -Property ConfigItemDisplayName, ConfigItemRecID

                $Actual.ConfigItemRecID | Should -BeNullOrEmpty
                $Actual.ConfigItemDisplayName | Should -BeNullOrEmpty
            }
        } -Tag 'ci'
        It 'CustomerRecID' {
            $ticketNr = New-CherwellTicketHC @testParams -KeyValuePair (@{
                    IncidentType              = 'Incident'
                    CustomerRecID             = $testUsers[0].CherwellCustomer.busObRecId
                    SubmittedBySamAccountName = $testUsers[0].SamAccountName
                    Status                    = 'New'
                    ShortDescription          = 'Automated test'
                    Description               = 'Pester test RequesterSamAccountName'
                    Source                    = 'Event'
                } + $testMandatoryFields)

            $ticketNr | Should -Not -BeNullOrEmpty

            $Actual = Get-CherwellTicketHC @testParams -TicketNr $ticketNr -Property CustomerRecID
            
            $Actual.CustomerRecID | 
            Should -Be $testUsers[0].CherwellCustomer.busObRecId
        }
        It 'RequesterSamAccountName > CustomerRecID' {
            $ticketNr = New-CherwellTicketHC @testParams -KeyValuePair (@{
                    IncidentType              = 'Incident'
                    RequesterSamAccountName   = $testUsers[0].SamAccountName
                    SubmittedBySamAccountName = $testUsers[0].SamAccountName
                    Status                    = 'New'
                    ShortDescription          = 'Automated test'
                    Description               = 'Pester test RequesterSamAccountName'
                    Source                    = 'Event'
                } + $testMandatoryFields)

            $ticketNr | Should -Not -BeNullOrEmpty

            $Actual = Get-CherwellTicketHC @testParams -TicketNr $ticketNr -Property CustomerRecID

            $Actual.CustomerRecID | 
            Should -Be $testUsers[0].CherwellCustomer.busObRecId
        }
        It 'Description with non English characters' {
            $TicketNr = New-CherwellTicketHC @testParams -KeyValuePair (@{
                    IncidentType              = 'Incident'
                    RequesterSamAccountName   = $testUsers[0].SamAccountName
                    SubmittedBySamAccountName = $testUsers[0].SamAccountName
                    Status                    = 'New'
                    ShortDescription          = 'Automated test'
                    Description               = 'Pester test description with non English characters žluťoučký kůň úpěl ďábelské ódy'
                    Source                    = 'Event'
                } + $testMandatoryFields)

            $TicketNr | Should -Not -BeNullOrEmpty

            $Actual = Get-CherwellTicketHC @testParams -TicketNr $TicketNr -Property SubmitOnBehalfOfID

            $Actual.SubmitOnBehalfOfID | Should -Not -BeNullOrEmpty
        }
        It 'SubmittedBySamAccountName > SubmitOnBehalfOfID' {
            $TicketNr = New-CherwellTicketHC @testParams -KeyValuePair (@{
                    IncidentType              = 'Incident'
                    RequesterSamAccountName   = $testUsers[0].SamAccountName
                    SubmittedBySamAccountName = $testUsers[1].SamAccountName
                    Status                    = 'New'
                    ShortDescription          = 'Automated test'
                    Description               = 'Pester test SubmittedBySamAccountName'
                    Source                    = 'Event'
                } + $testMandatoryFields)

            $TicketNr | Should -Not -BeNullOrEmpty

            $Actual = Get-CherwellTicketHC @testParams -TicketNr $TicketNr -Property SubmitOnBehalfOfID
            
            $Actual.SubmitOnBehalfOfID | 
            Should -Be $testUsers[1].CherwellCustomer.busObRecId
        }
        It 'SubmitOnBehalfOfID' {
            $TicketNr = New-CherwellTicketHC @testParams -KeyValuePair (@{
                    IncidentType            = 'Incident'
                    RequesterSamAccountName = $testUsers[0].SamAccountName
                    SubmitOnBehalfOfID      = $testUsers[1].CherwellCustomer.busObRecId
                    Status                  = 'New'
                    ShortDescription        = 'Automated test'
                    Description             = 'Pester test SubmitOnBehalfOfID'
                    Source                  = 'Event'
                } + $testMandatoryFields)

            $TicketNr | Should -Not -BeNullOrEmpty

            $Actual = Get-CherwellTicketHC @testParams -TicketNr $TicketNr -Property SubmitOnBehalfOfID
            
            $Actual.SubmitOnBehalfOfID | 
            Should -Be $testUsers[1].CherwellCustomer.busObRecId
        }
        It 'OwnedByTeam' {
            $TicketNr = New-CherwellTicketHC @testParams -KeyValuePair (@{
                    IncidentType            = 'Incident'
                    RequesterSamAccountName = $testUsers[0].SamAccountName
                    OwnedByTeam             = $testUsers[0].OwnedByTeam
                    Status                  = 'New'
                    ShortDescription        = 'Automated test'
                    Description             = 'Pester test description with OwnedByTeam'
                    Source                  = 'Event'
                } + $testMandatoryFields)

            $TicketNr | Should -Not -BeNullOrEmpty

            $Actual = Get-CherwellTicketHC @testParams -TicketNr $TicketNr
            $Actual.OwnedByTeam | Should -Be $testUsers[0].OwnedByTeam
        }
        It 'OwnedByTeam and OwnedBy' {
            $Expected = (@{
                    IncidentType            = 'Incident'
                    RequesterSamAccountName = $testUsers[0].SamAccountName
                    OwnedByTeam             = $testUsers[0].OwnedByTeam
                    OwnedBy                 = $testUsers[0].CherwellSystemUser.busObPublicId
                    OwnedById               = $testUsers[0].CherwellSystemUser.busObRecId
                    Status                  = 'New'
                    ShortDescription        = 'Automated test'
                    Description             = 'Pester test description with OwnedBy'
                    Source                  = 'Event'
                } + $testMandatoryFields)

            $ticketNr = New-CherwellTicketHC @testParams -KeyValuePair $Expected 
            $ticketNr | Should -Not -BeNullOrEmpty

            $Actual = Get-CherwellTicketHC @testParams -Ticket $ticketNr -Property OwnedByID, OwnedBy

            $Actual.OwnedByID | 
            Should -Be $testUsers[0].CherwellSystemUser.busObRecId
            $Actual.OwnedBy | Should -Not -BeNullOrEmpty
        }
        It 'OwnedByTeam and OwnedBySamAccountName > OwnedByID OwnedBy' {
            $ticketNr = New-CherwellTicketHC @testParams -KeyValuePair (@{
                    IncidentType            = 'Incident'
                    RequesterSamAccountName = $testUsers[0].SamAccountName
                    OwnedByTeam             = $testUsers[0].OwnedByTeam
                    OwnedBySamAccountName   = $testUsers[0].SamAccountName
                    Status                  = 'New'
                    ShortDescription        = 'Automated test'
                    Description             = 'Pester test OwnedBySamAccountName and OwnedByTeam'
                    Source                  = 'Event'
                } + $testMandatoryFields)

            $ticketNr | Should -Not -BeNullOrEmpty

            $Actual = Get-CherwellTicketHC @testParams -TicketNr $ticketNr -Property OwnedByID

            $Actual.OwnedByID | Should -Not -BeNullOrEmpty
        }
        It 'multiple Attachments' {
            $File1 = (New-Item -Path "TestDrive:/testFile1.txt" -ItemType File).FullName
            $File2 = (New-Item -Path "TestDrive:/testFile2.txt" -ItemType File).FullName

            $TicketNr = New-CherwellTicketHC @testParams -KeyValuePair (@{
                    IncidentType            = 'Incident'
                    RequesterSamAccountName = $testUsers[0].SamAccountName
                    Status                  = 'New'
                    ShortDescription        = 'Automated test'
                    Description             = 'Pester test description with multiple attachment'
                    Source                  = 'Event'
                    Attachment              = $File1, $File2
                } + $testMandatoryFields) 

            $TicketNr | Should -Not -BeNullOrEmpty
        }
        It "IncidentType 'Incident'" {
            $TicketNr = New-CherwellTicketHC @testParams -KeyValuePair (@{
                    IncidentType            = 'Incident'
                    RequesterSamAccountName = $testUsers[0].SamAccountName
                    Status                  = 'New'
                    ShortDescription        = 'Automated test'
                    Description             = 'Pester test description standard ticket'
                    Source                  = 'Event'
                } + $testMandatoryFields)

            $TicketNr | Should -Not -BeNullOrEmpty

            $Actual = Get-CherwellTicketHC @testParams -TicketNr $TicketNr -Property IncidentType

            $Actual.IncidentType | Should -Be 'Incident'
        }
        It "IncidentType 'Service Request'" {
            $TicketNr = New-CherwellTicketHC @testParams -KeyValuePair (@{
                    IncidentType            = 'Service Request'
                    RequesterSamAccountName = $testUsers[0].SamAccountName
                    Status                  = 'New'
                    ShortDescription        = 'Automated test'
                    Description             = 'Pester test description standard ticket'
                    Source                  = 'Event'
                } + $testMandatoryFields)

            $TicketNr | Should -Not -BeNullOrEmpty

            $Actual = Get-CherwellTicketHC @testParams -TicketNr $TicketNr -Property IncidentType

            $Actual.IncidentType | Should -Be 'Service Request'
        }
        It 'Multiple lines with HTML tags' {
            $Actual = New-CherwellTicketHC @testParams -KeyValuePair (@{
                    IncidentType              = 'Incident'
                    RequesterSamAccountName   = $testUsers[0].SamAccountName
                    SubmittedBySamAccountName = $testUsers[0].SamAccountName
                    Status                    = 'New'
                    ShortDescription          = 'Automated test'
                    Description               = '
                            Pester test description:<br>
                            Line1: Some stuff<br><br>
                            LIne2: Other stuff'
                    Source                    = 'Event'
                } + $testMandatoryFields)

            $Actual | Should -Not -BeNullOrEmpty
        }
        It 'blank Attachments' {
            $TicketNr = New-CherwellTicketHC @testParams -KeyValuePair (@{
                    IncidentType              = 'Incident'
                    RequesterSamAccountName   = $testUsers[0].SamAccountName
                    SubmittedBySamAccountName = $testUsers[0].SamAccountName
                    Status                    = 'New'
                    ShortDescription          = 'Automated test'
                    Description               = 'Pester test description without attachment'
                    Source                    = 'Event'
                    Attachment                = $null
                } + $testMandatoryFields)

            $TicketNr | Should -Not -BeNullOrEmpty
        }
        It 'blank SubmittedBySamAccountName' {
            $TicketNr = New-CherwellTicketHC @testParams -KeyValuePair (@{
                    IncidentType              = 'Incident'
                    RequesterSamAccountName   = $testUsers[0].SamAccountName
                    SubmittedBySamAccountName = $null
                    Status                    = 'New'
                    ShortDescription          = 'Automated test'
                    Description               = 'Pester test description without SubmittedBySamAccountName'
                    Source                    = 'Event'
                } + $testMandatoryFields) 

            $TicketNr | Should -Not -BeNullOrEmpty
        }
    }
    It 'create multiple tickets at once' {
        $Actual = New-CherwellTicketHC @testParams -KeyValuePair @(
            (@{
                    IncidentType            = 'Incident'
                    RequesterSamAccountName = $testUsers[0].SamAccountName
                    OwnedByTeam             = $testUsers[0].OwnedByTeam
                    Status                  = 'New'
                    ShortDescription        = 'Automated test'
                    Description             = 'Pester test description'
                    Source                  = 'Event'
                } + $testMandatoryFields)
            (@{
                    IncidentType            = 'Incident'
                    RequesterSamAccountName = $testUsers[0].SamAccountName
                    Status                  = 'New'
                    ShortDescription        = 'Automated test'
                    Description             = 'Pester test description'
                    Source                  = 'Event'
                } + $testMandatoryFields)
        )

        $Actual | Should -HaveCount 2
    }
    Context 'the output of creating a ticket is' {
        It 'an integer representing the ticket number' {
            $Actual = New-CherwellTicketHC @testParams -KeyValuePair (@{
                    IncidentType            = 'Incident'
                    RequesterSamAccountName = $testUsers[0].SamAccountName
                    Status                  = 'New'
                    ShortDescription        = 'Automated test'
                    Description             = 'Pester test default output is an integer'
                    Source                  = 'Event'
                } + $testMandatoryFields)

            $Actual | Should -Not -BeNullOrEmpty
            $Actual | Should -BeOfType [Int]
        }
        It "a 'PSCustomObject' when using 'PassThru'" {
            $Actual = New-CherwellTicketHC -PassThru @testParams -KeyValuePair (@{
                    IncidentType            = 'Incident'
                    RequesterSamAccountName = $testUsers[0].SamAccountName
                    Status                  = 'New'
                    ShortDescription        = 'Automated test'
                    Description             = 'Pester test PassThru output is an object'
                    Source                  = 'Event'
                } + $testMandatoryFields)

            $Actual | Should -Not -BeNullOrEmpty
            $Actual | Should -BeOfType [PSCustomObject]
        }
    }
}
Describe 'Add-CherwellTicketConfigItemHC' {
    Context 'add a CI to a ticket when' {
        It 'a ticket number is provided' {
            $testTicketNr = New-CherwellTicketHC @testParams -KeyValuePair (@{
                    IncidentType            = 'Incident'
                    RequesterSamAccountName = $testUsers[0].SamAccountName
                    Status                  = 'New'
                    ShortDescription        = 'Automated test Add-CherwellTicketConfigItemHC'
                    Description             = 'Pester test description standard ticket'
                    Source                  = 'Event'
                } + $testMandatoryFields)

            $testCI = Get-CherwellConfigItemHC -PassThru @testParams -Type $testConfigItems[0].CIType -Filter @{
                FieldName  = 'HostName'
                Operator   = 'eq'
                FieldValue = $env:COMPUTERNAME
            }

            $testCI | Should -Not -BeNullOrEmpty

            { Add-CherwellTicketConfigItemHC @testParams -Ticket $testTicketNr -ConfigItem $testCI } |
            Should -Not -Throw
        }
        It 'a ticket object is provided' {
            $testTicket = New-CherwellTicketHC @testParams -PassThru -KeyValuePair (@{
                    IncidentType            = 'Incident'
                    RequesterSamAccountName = $testUsers[0].SamAccountName
                    Status                  = 'New'
                    ShortDescription        = 'Automated test Add-CherwellTicketConfigItemHC'
                    Description             = 'Pester test description standard ticket'
                    Source                  = 'Event'
                } + $testMandatoryFields)

            $testCI = Get-CherwellConfigItemHC -PassThru @testParams -Type $testConfigItems[0].CIType -Filter @{
                FieldName  = 'HostName'
                Operator   = 'eq'
                FieldValue = $env:COMPUTERNAME
            }

            $testCI | Should -Not -BeNullOrEmpty

            { Add-CherwellTicketConfigItemHC @testParams -Ticket $testTicket -ConfigItem $testCI } |
            Should -Not -Throw
        }
    }
    It "add multiple CI's to a ticket" {
        $testTicketNr = New-CherwellTicketHC @testParams -KeyValuePair (@{
                IncidentType            = 'Incident'
                RequesterSamAccountName = $testUsers[0].SamAccountName
                Status                  = 'New'
                ShortDescription        = 'Automated test Add-CherwellTicketConfigItemHC multiple CI'
                Description             = 'Pester test description standard ticket'
                Source                  = 'Event'
            } + $testMandatoryFields)

        $testCIMultiple = Get-CherwellConfigItemHC -PassThru @testParams -Type $testConfigItems[0].CIType -Filter @{
            FieldName  = 'HostName'
            Operator   = 'contains'
            FieldValue = $env:COMPUTERNAME
        }

        $testCIMultiple | Should -Not -BeNullOrEmpty

        {
            Add-CherwellTicketConfigItemHC @testParams -Ticket $testTicketNr -ConfigItem $testCIMultiple
        } |
        Should -Not -Throw
    }
} -Tag 'ci'
Describe 'Test-InvalidPropertyCombinationHC' {
    InModuleScope $testModuleName {
        Context 'thrown an error on an incorrect combination' {
            $errorMessageOwnedBySamAccountName = "The field 'OwnedBySamAccountName' cannot be combined with the fields 'OwnedBy' or 'OwnedById'. Please use 'Get-CherwellSystemUserHC' that provides you with 'OwnedBy' and 'OwnedById' if you want to be specific."
            $errorMessageOwnedByIdOrOwnedBy = "Both the fields 'OwnedBy' and 'OwnedById' need to be specified. Please use the field 'OwnedBySamAccountName' instead or use 'Get-CherwellSystemUserHC' that provides you with 'OwnedBy' and 'OwnedById' if you want to be specific."
            $errorMessageOwnedByTeam = "When the fields 'OwnedBy', 'OwnedById' or 'OwnedBySamAccountName' are used it is mandatory to specify the field 'OwnedByTeam' too."
            $errorMessageRequester = "The field 'RequesterSamAccountName' cannot be combined with the field 'CustomerRecID'. Please use 'Get-CherwellCustomerHC' to obtain the 'CustomerRecID' or use the SamAccountName in the field 'RequesterSamAccountName'."
            $errorSubmitOnBehalf = "The field 'SubmittedBySamAccountName' cannot be combined with the field 'SubmitOnBehalfOfID'. Please use 'Get-CherwellCustomerHC' to obtain the 'SubmitOnBehalfOfID' or use the SamAccountName in the field 'SubmittedBySamAccountName'."
            $TestCases = @(
                $KeyValuePair = @{
                    RequesterSamAccountName = $null
                    CustomerRecID           = $null
                }
                @{
                    KeyValuePair = $KeyValuePair
                    ErrorMessage = $errorMessageRequester
                    Name         = $KeyValuePair.Keys -join ' + '
                }
                $KeyValuePair = @{
                    OwnedBySamAccountName = $null
                    OwnedBy               = $null
                    OwnedByTeam           = $null
                }
                @{
                    KeyValuePair = $KeyValuePair
                    ErrorMessage = $errorMessageOwnedBySamAccountName
                    Name         = $KeyValuePair.Keys -join ' + '
                }
                $KeyValuePair = @{
                    OwnedBySamAccountName = $null
                    OwnedById             = $null
                    OwnedByTeam           = $null
                }
                @{
                    KeyValuePair = $KeyValuePair
                    ErrorMessage = $errorMessageOwnedBySamAccountName
                    Name         = $KeyValuePair.Keys -join ' + '
                }
                $KeyValuePair = @{
                    OwnedById   = $null
                    OwnedByTeam = $null
                }
                @{
                    KeyValuePair = $KeyValuePair
                    ErrorMessage = $errorMessageOwnedByIdOrOwnedBy
                    Name         = $KeyValuePair.Keys -join ' + '
                }
                $KeyValuePair = @{
                    OwnedBy     = $null
                    OwnedByTeam = $null
                }
                @{
                    KeyValuePair = $KeyValuePair
                    ErrorMessage = $errorMessageOwnedByIdOrOwnedBy
                    Name         = $KeyValuePair.Keys -join ' + '
                }
                $KeyValuePair = @{
                    OwnedBy   = $null
                    OwnedById = $null
                }
                @{
                    KeyValuePair = $KeyValuePair
                    ErrorMessage = $errorMessageOwnedByTeam
                    Name         = $KeyValuePair.Keys -join ' + '
                }
                $KeyValuePair = @{
                    OwnedBySamAccountName = $null
                }
                @{
                    KeyValuePair = $KeyValuePair
                    ErrorMessage = $errorMessageOwnedByTeam
                    Name         = $KeyValuePair.Keys -join ' + '
                }
                $KeyValuePair = @{
                    SubmitOnBehalfOfID        = $null
                    SubmittedBySamAccountName = $null
                }
                @{
                    KeyValuePair = $KeyValuePair
                    ErrorMessage = $errorSubmitOnBehalf
                    Name         = $KeyValuePair.Keys -join ' + '
                }
            )
            It "<Name>" -Foreach $TestCases {
                {
                    Test-InvalidPropertyCombinationHC -KeyValuePair $KeyValuePair
                } |
                Should -Throw -PassThru |
                Should -BeLike "*$ErrorMessage*"
            }
        }
        Context 'no error is thrown on an correct combination' {
            $TestCases = @(
                $KeyValuePair = @{
                    CustomerRecID = $null
                }
                @{
                    KeyValuePair = $KeyValuePair
                    Name         = $KeyValuePair.Keys -join ' + '
                }
                $KeyValuePair = @{
                    RequesterSamAccountName = $null
                }
                @{
                    KeyValuePair = $KeyValuePair
                    Name         = $KeyValuePair.Keys -join ' + '
                }
                $KeyValuePair = @{
                    OwnedBySamAccountName = $null
                    OwnedByTeam           = $null
                }
                @{
                    KeyValuePair = $KeyValuePair
                    Name         = $KeyValuePair.Keys -join ' + '
                }
                $KeyValuePair = @{
                    OwnedBy     = $null
                    OwnedById   = $null
                    OwnedByTeam = $null
                }
                @{
                    KeyValuePair = $KeyValuePair
                    Name         = $KeyValuePair.Keys -join ' + '
                }
                $KeyValuePair = @{
                    SubmittedBySamAccountName = $null
                }
                @{
                    KeyValuePair = $KeyValuePair
                    Name         = $KeyValuePair.Keys -join ' + '
                }
                $KeyValuePair = @{
                    SubmitOnBehalfOfID = $null
                }
                @{
                    KeyValuePair = $KeyValuePair
                    Name         = $KeyValuePair.Keys -join ' + '
                }
            )
            It "<Name>" -Foreach $TestCases {
                {
                    Test-InvalidPropertyCombinationHC -KeyValuePair $KeyValuePair
                } |
                Should -Not -Throw 
            }
        }
    }
}
Describe 'New-CherwellSearchFilterHC' {
    InModuleScope $testModuleName {
        BeforeAll {
            $testSchema = @{
                FieldDefinitions = (
                    @{
                        Name                 = 'ComputerName'
                        isFullTextSearchable = $true
                        FieldID              = 5
                    },
                    @{
                        Name                 = 'ComputerName'
                        isFullTextSearchable = $true
                        FieldID              = 7
                    }
                )
            }
        }
        Context 'an error is thrown when the Filter hash table' {
            It 'is missing a mandatory key' { 
                {
                    New-CherwellSearchFilterHC -Schema $testSchema -Filter @{
                        Operator   = 'eq'
                        FieldValue = 'PC1'
                    }
                } |
                Should -Throw -PassThru |
                Should -BeLike '*Failed creating a search filter*mandatory*'
            }
            It 'contains an unknown key' { {
                    New-CherwellSearchFilterHC -Schema $testSchema -Filter @{
                        FieldName  = 'ComputerName'
                        Operator   = 'eq'
                        FieldValue = 'PC1'
                        UnknownKey = $null
                    }
                } |
                Should -Throw -PassThru |
                Should -BeLike '*Failed creating a search filter*Only the following keys are allowed*'
            }
            It 'contains an unknown operator' { {
                    New-CherwellSearchFilterHC -Schema $testSchema -Filter @{
                        FieldName  = 'ComputerName'
                        Operator   = 'NotExisting'
                        FieldValue = 'PC1'
                    }
                } |
                Should -Throw -PassThru |
                Should -BeLike '*Failed creating a search filter*operator*'
            }
        }
        Context 'no error is thrown when the Filter contains the correct keys for' {
            It 'only one hash table' { {
                    New-CherwellSearchFilterHC -Schema $testSchema -Filter @{
                        FieldName  = 'ComputerName'
                        Operator   = 'eq'
                        FieldValue = 'PC1'
                    }
                } | Should -Not -Throw
            }
            It 'multiple hash tables' { {
                    New-CherwellSearchFilterHC -Schema $testSchema -Filter (
                        @{
                            FieldName  = 'ComputerName'
                            Operator   = 'eq'
                            FieldValue = 'PC1'
                        },
                        @{
                            FieldName  = 'ComputerName'
                            Operator   = 'eq'
                            FieldValue = 'PC2'
                        }
                    )
                } | Should -Not -Throw
            }
        }
        Context 'create a new correct hash table with' {
            It "FieldID, Operator and Value" { {
                    $testHash = @{
                        FieldName  = 'ComputerName'
                        Operator   = 'eq'
                        FieldValue = 'PC1'
                    }

                    $Actual = New-CherwellSearchFilterHC -Schema $testSchema -Filter $testHash

                    $Actual.FieldID | Should -Not -BeNullOrEmpty
                    $Actual.Value | Should -Not -BeNullOrEmpty
                    $Actual.Operator | Should -Not -BeNullOrEmpty
                    $Actual.count | Should -BeExactly 3
                } | Should -Not -Throw
            }
        }
    }
} 
Describe 'Remove-CherwellTicketHC' {
    Context 'the ticket is removed when' {
        It 'a ticket number is provided' {
            $TestTicketNr = New-CherwellTicketHC @testParams -KeyValuePair (
                @{
                    IncidentType            = 'Incident'
                    RequesterSamAccountName = $testUsers[0].SamAccountName
                    OwnedByTeam             = $testUsers[0].OwnedByTeam
                    ShortDescription        = 'Pester automated test'
                    Description             = 'Update ticket fields'
                    Source                  = 'Event'
                } + $testMandatoryFields)

            { Remove-CherwellTicketHC @testParams -TicketNr $TestTicketNr } | Should -Not -Throw
        } -Skip:$true
        It 'multiple ticket numbers are provided' {
            $TestTicketNr = New-CherwellTicketHC @testParams -KeyValuePair (
                @{
                    IncidentType            = 'Incident'
                    RequesterSamAccountName = $testUsers[0].SamAccountName
                    OwnedByTeam             = $testUsers[0].OwnedByTeam
                    ShortDescription        = 'Pester automated test'
                    Description             = 'Update ticket fields'
                    Source                  = 'Event'
                } + $testMandatoryFields)

            $TestTicketNr2 = New-CherwellTicketHC @testParams -KeyValuePair (
                @{
                    IncidentType            = 'Incident'
                    RequesterSamAccountName = $testUsers[0].SamAccountName
                    OwnedByTeam             = $testUsers[0].OwnedByTeam
                    ShortDescription        = 'Pester automated test'
                    Description             = 'Update ticket fields'
                    Source                  = 'Event'
                } + $testMandatoryFields)

            { Remove-CherwellTicketHC @testParams -TicketNr $TestTicketNr, $TestTicketNr2 } |
            Should -Not -Throw
        } -Skip:$true
    }
}
Describe 'Update-CherwellTicketHC' {
    BeforeAll {
        $TestTicketObject = New-CherwellTicketHC @testParams -PassThru -KeyValuePair (
            @{
                IncidentType            = 'Incident'
                RequesterSamAccountName = $testUsers[0].SamAccountName
                OwnedByTeam             = $testUsers[0].OwnedByTeam
                ShortDescription        = 'Automated test'
                Description             = 'Update-CherwellTicketHC'
                Source                  = 'Event'
            } + $testMandatoryFields)
    }
    Context 'an error is thrown when' {
        It 'the ticket number is not found' {
            {
                Update-CherwellTicketHC @testParams -TicketNr 99999999 -KeyValuePair @{ } 
            } |
            Should -Throw -PassThru | Should -BeLike '*Ticket number * not found*'
        }
        It 'Test-InvalidPropertyCombinationHC is called' {
            Mock Test-InvalidPropertyCombinationHC
            
            Update-CherwellTicketHC @testParams -TicketNr $TestTicketObject.busObPublicId -KeyValuePair (
                @{
                    Description = 'test update'
                } + $testMandatoryFields
            ) 
            
            Should -Invoke Test-InvalidPropertyCombinationHC -Exactly -Times 1
        }
    }
    Context 'the field is updated for' {
        $TestCases = @(
            @{
                Name     = 'Source'
                Expected = @{
                    Source = 'Phone'
                }
            }
            @{
                Name     = 'ShortDescription'
                Expected = @{
                    ShortDescription = 'Updated ShortDescription'
                }
            }
            @{
                Name     = 'Priority'
                Expected = @{
                    Priority = 3
                }
            }
            @{
                Name     = 'Description'
                Expected = @{
                    Description = 'Updated description with non English characters žluťoučký kůň úpěl ďábelské ódy'
                }
            }
        )

        It "<Name>" -Foreach $TestCases {
            $TicketNr = $TestTicketObject.busObPublicId

            $Before = Get-CherwellTicketHC @testParams -TicketNr $TicketNr -Property *

            $Expected.GetEnumerator().ForEach( {
                    $Before.($_.Key) | Should -Not -Be $_.Value
                })

            Update-CherwellTicketHC @testParams -TicketNr $TicketNr -KeyValuePair $Expected

            $Actual = Get-CherwellTicketHC @testParams -TicketNr $TicketNr -Property *

            $Expected.GetEnumerator().ForEach( {
                    $Actual.($_.Key) | Should -Be $_.Value
                })
        } 
        It 'update multiple tickets with multiple fields' {
            $TestTicketNr1 = New-CherwellTicketHC @testParams -KeyValuePair (
                @{
                    IncidentType            = 'Incident'
                    RequesterSamAccountName = $testUsers[0].SamAccountName
                    OwnedByTeam             = $testUsers[0].OwnedByTeam
                    ShortDescription        = 'Automated test'
                    Description             = 'Update-CherwellTicketHC'
                    Source                  = 'Event'
                } + $testMandatoryFields)

            $TestTicketNr2 = New-CherwellTicketHC @testParams -KeyValuePair (
                @{
                    IncidentType            = 'Incident'
                    RequesterSamAccountName = $testUsers[0].SamAccountName
                    OwnedByTeam             = $testUsers[0].OwnedByTeam
                    ShortDescription        = 'Automated test'
                    Description             = 'Update-CherwellTicketHC'
                    Source                  = 'Event'
                } + $testMandatoryFields)

            $Expected = @{
                Priority    = 3
                Source      = 'Phone'
                Description = 'Update-CherwellTicketHC: Updated'
            }
            Update-CherwellTicketHC @testParams -TicketNr $TestTicketNr1, $TestTicketNr2 -KeyValuePair $Expected

            $Actual = Get-CherwellTicketHC @testParams -TicketNr $TestTicketNr1, $TestTicketNr2 -Property *

            Foreach ($A in $Actual) {
                $Expected.GetEnumerator().ForEach( {
                        $A.($_.Key) | Should -Be $_.Value
                    })
            }

        }
        It 'CustomerRecID' {
            $CherwellCustomer = Get-CherwellCustomerHC @testParams -Filter @{
                FieldName  = 'SamAccountName'
                Operator   = 'eq'
                FieldValue = $testUsers[1].SamAccountName
            } -PassThru

            $Expected = (
                @{
                    CustomerRecID = $CherwellCustomer.busObRecId
                } + $testMandatoryFields
            )

            $Before = Get-CherwellTicketHC @testParams -TicketNr $TestTicketObject.busObPublicId -Property CustomerRecID

            $Before.CustomerRecID | Should -Not -Be $Expected.CustomerRecID

            Update-CherwellTicketHC @testParams -TicketNr $TestTicketObject.busObPublicId -KeyValuePair $Expected

            $Actual = Get-CherwellTicketHC @testParams -TicketNr $TestTicketObject.busObPublicId -Property CustomerRecID

            $Actual.CustomerRecID | Should -Be $Expected.CustomerRecID
        } 
        It 'SubmittedBySamAccountName > SubmitOnBehalfOfID' {
            $Expected = (
                @{
                    SubmittedBySamAccountName = $testUsers[1].SamAccountName
                } + $testMandatoryFields
            )

            $Before = Get-CherwellTicketHC @testParams -TicketNr $TestTicketObject.busObPublicId -Property SubmitOnBehalfOfID

            $Before.SubmitOnBehalfOfID | 
            Should -Not -Be $testUsers[1].CherwellCustomer.busObRecId

            Update-CherwellTicketHC @testParams -TicketNr $TestTicketObject.busObPublicId -KeyValuePair $Expected

            $Actual = Get-CherwellTicketHC @testParams -TicketNr $TestTicketObject.busObPublicId -Property SubmitOnBehalfOfID

            $Actual.SubmitOnBehalfOfID | 
            Should -Be $testUsers[1].CherwellCustomer.busObRecId
        }
        It 'RequesterSamAccountName > CustomerRecID' {
            $Expected = (
                @{
                    RequesterSamAccountName = $testUsers[0].SamAccountName
                } + $testMandatoryFields
            )

            $Before = Get-CherwellTicketHC @testParams -TicketNr $TestTicketObject.busObPublicId -Property CustomerRecID

            $Before.CustomerRecID | 
            Should -Not -Be  $testUsers[0].CherwellCustomer.busObRecId

            Update-CherwellTicketHC @testParams -TicketNr $TestTicketObject.busObPublicId -KeyValuePair $Expected

            $Actual = Get-CherwellTicketHC @testParams -TicketNr $TestTicketObject.busObPublicId -Property CustomerRecID

            $Actual.CustomerRecID | 
            Should -Be  $testUsers[0].CherwellCustomer.busObRecId
        }
        It 'OwnedBySamAccountName and OwnedByTeam > OwnedBy, OwnedById, OwnedByTeam' {
            $Expected = (
                @{
                    OwnedBySamAccountName = $testUsers[0].SamAccountName
                    OwnedByTeam           = $testUsers[0].OwnedByTeam
                } + $testMandatoryFields
            )
        
            $Before = Get-CherwellTicketHC @testParams -TicketNr $TestTicketObject.busObPublicId -Property OwnedBy, OwnedByID

            $Before.OwnedBy | Should -Not -Be $Expected.OwnedBy
            $Before.OwnedByID | Should -Not -Be  $Expected.OwnedByID
        
            Update-CherwellTicketHC @testParams -TicketNr $TestTicketObject.busObPublicId -KeyValuePair $Expected
        
            $Actual = Get-CherwellTicketHC @testParams -TicketNr $TestTicketObject.busObPublicId -Property OwnedBy, OwnedByID, OwnedByTeam
        
            $Actual.OwnedBy | 
            Should -Be $testUsers[0].CherwellSystemUser.busObPublicId
            $Actual.OwnedByID | 
            Should -Be  $testUsers[0].CherwellSystemUser.busObRecId
            $Actual.OwnedByTeam | Should -Be  $Expected.OwnedByTeam
        }
        It "OwnedBy, OwnedById and OwnedByTeam" {
            $Expected = (
                @{
                    OwnedBy     = $testUsers[1].CherwellSystemUser.busObPublicId
                    OwnedById   = $testUsers[1].CherwellSystemUser.busObRecId
                    OwnedByTeam = $testUsers[1].OwnedByTeam
                } + $testMandatoryFields
            )
        
            $Before = Get-CherwellTicketHC @testParams -TicketNr $TestTicketObject.busObPublicId -Property OwnedBy, OwnedByID

            $Before.OwnedBy | Should -Not -Be $Expected.OwnedBy
            $Before.OwnedByID | Should -Not -Be  $Expected.OwnedByID
        
            Update-CherwellTicketHC @testParams -TicketNr $TestTicketObject.busObPublicId -KeyValuePair $Expected
        
            $Actual = Get-CherwellTicketHC @testParams -TicketNr $TestTicketObject.busObPublicId -Property OwnedBy, OwnedByID, OwnedByTeam
        
            $Actual.OwnedBy | Should -Be $Expected.OwnedBy
            $Actual.OwnedByID | Should -Be  $Expected.OwnedByID
            $Actual.OwnedByTeam | Should -Be  $Expected.OwnedByTeam
        }
        It 'a piped object' {
            $testShortDescription = 'Update-CherwellTicketHC updated ShortDescription'

            $TestTicketObject | Update-CherwellTicketHC @testParams  -KeyValuePair @{
                ShortDescription = $testShortDescription
            }

            $Actual = Get-CherwellTicketHC @testParams -TicketNr $TestTicketObject.busObPublicId -Property ShortDescription

            $Actual.ShortDescription | Should -Be $testShortDescription
        }
    } 
    It "close ticket $($TestTicketObject.busObPublicId) by setting OwnedBy, OwnedById, OwnedByTeam, Status and CloseDescription" {
        $Expected = @{
            Status           = 'Resolved'
            CloseDescription = 'Done!'
            OwnedBy          = $testUsers[0].CherwellSystemUser.busObPublicId
            OwnedById        = $testUsers[0].CherwellSystemUser.busObRecId
            OwnedByTeam      = $testUsers[0].OwnedByTeam
        }

        $Before = Get-CherwellTicketHC @testParams -TicketNr $TestTicketObject.busObPublicId -Property Status, CloseDescription

        $Before.Status | Should -Not -Be $Expected.Status
        $Before.CloseDescription | Should -Not -Be $Expected.CloseDescription

        Update-CherwellTicketHC @testParams -TicketNr $TestTicketObject.busObPublicId -KeyValuePair $Expected

        $Actual = Get-CherwellTicketHC @testParams -TicketNr $TestTicketObject.busObPublicId -Property Status, CloseDescription

        $Actual.Status | Should -Be $Expected.Status
        $Actual.CloseDescription | Should -Be $Expected.CloseDescription
    }
}
