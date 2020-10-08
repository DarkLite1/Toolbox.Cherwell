#Requires -Version 5.1
Set-StrictMode -Version Latest

<#
    .LINK
        https://github.com/DarkLite1/Toolbox.Module
 #>

$PasswordFile = "$PSScriptRoot\Passwords.json"
$Environment = $null
$GetParams = $null
$PostParams = $null
$Uri = $null

$RelationshipEnum = @{
    Attachments            = 'Incident Owns Attachments'
    Approvals              = 'Incident Owns Approvals'
    JournalCustomerRequest = 'Incident Owns Journal Customer Request'
    JournalNote            = 'Incident Owns Jounal Note'
    Journals               = 'Incident Owns Journals'
    SLMHistory             = 'Incident Owns SLM History'
    Specifics              = 'Incident Owns Specifics'
    Task                   = 'Incident Owns Tasks'
}

#region Import password file
Function ConvertTo-HashTableHC {
    [CmdletBinding()]
    [OutputType('HashTable')]
    Param (
        [Parameter(ValueFromPipeline)]
        $InputObject
    )

    Process {
        if ($null -eq $InputObject) {
            return $null
        }

        if ($InputObject -is [System.Collections.IEnumerable] -and $InputObject -isnot [String]) {
            $collection = @(
                foreach ($object in $InputObject) {
                    ConvertTo-HashtableHC -InputObject $object
                }
            )

            Write-Output -NoEnumerate $collection
        }
        elseif ($InputObject -is [PSObject]) {
            $hash = @{ }
            foreach ($property in $InputObject.PSObject.Properties) {
                $hash[$property.Name] = ConvertTo-HashtableHC -InputObject $property.Value
            }
            $hash
        }
        else {
            $InputObject
        }
    }
}

Function Test-CredentialPropertyHC {
    Param (
        [HashTable]$Name,
        [String[]]$Property = @('AuthMode', 'Uri', 'UserName', 'Password', 'KeyAPI'),
        [String[]]$AuthMode = @('Internal') # @('Internal', 'SAML', 'LDAP', 'Windows')

    )

    foreach ($N in $Name.GetEnumerator()) {
        if ($N.Name -eq '') {
            throw 'The environment name cannot be blank'
        }

        foreach ($P in $Property) {
            if (-not $N.Value.ContainsKey($P)) {
                throw "The properties '$($Property -join ', ')' are mandatory"
            }

            if ((-not $N.Value[$P]) -or ($N.Value[$P] -eq '')) {
                throw "The property '$P' cannot be empty"
            }
        }

        if ($AuthMode -notcontains $N.Value['AuthMode']) {
            throw "Authentication mode '$($N.Value['AuthMode'])' is not supported. The parameter 'AuthMode' only supports the values: $($AuthMode -join ', ')."
        }
    }
}

if (-not (Test-Path -Path $PasswordFile -PathType Leaf)) {
    throw "File 'Passwords.json' not found in the module folder. Please add your credentials to this file and save it in the folder '$PSScriptRoot'."
}

$EnvironmentList = Get-Content -Path $PasswordFile -Raw -EA Stop |
ConvertFrom-Json -EA Stop | ConvertTo-HashTableHC

foreach ($E in $EnvironmentList.GetEnumerator()) {
    $E.Value['Schema'] = @{ }
    $E.Value['Summary'] = @{ }
    $E.Value['Template'] = @{ }
}

Test-CredentialPropertyHC -Name $EnvironmentList
#endregion

Function New-DynamicParameterHC {
    Param (
        [Parameter(Mandatory)]
        [ValidateSet('Environment', 'Type')]
        [String[]]$Name
    )

    $ParamDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary

    if ($Name -contains 'Environment') {
        $EnvironmentParamName = 'Environment'
        $EnvironmentParamValue = $EnvironmentList.Keys

        $EnvironmentAttribute = New-Object System.Management.Automation.ParameterAttribute
        $EnvironmentAttribute.Position = 1
        $EnvironmentAttribute.Mandatory = $true

        $EnvironmentCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $EnvironmentCollection.Add($EnvironmentAttribute)

        $EnvironmentAttribValidateSet = New-Object System.Management.Automation.ValidateSetAttribute($EnvironmentParamValue)
        $EnvironmentCollection.Add($EnvironmentAttribValidateSet)

        $EnvironmentParam = New-Object System.Management.Automation.RuntimeDefinedParameter(
            $EnvironmentParamName, [String], $EnvironmentCollection)

        $ParamDictionary.Add($EnvironmentParamName, $EnvironmentParam)
    }

    if ($Name -contains 'Type') {
        $TypeParameterName = 'Type'
        $TypeParameterValue = $RelationshipEnum.Keys

        $TypeAttribute = New-Object System.Management.Automation.ParameterAttribute
        $TypeAttribute.Position = 2
        $TypeAttribute.Mandatory = $true

        $TypeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $TypeCollection.Add($TypeAttribute)

        $TypeAttribValidateSet = New-Object System.Management.Automation.ValidateSetAttribute($TypeParameterValue)
        $TypeCollection.Add($TypeAttribValidateSet)

        $TypeParam = New-Object System.Management.Automation.RuntimeDefinedParameter(
            $TypeParameterName, [String], $TypeCollection)

        $ParamDictionary.Add($TypeParameterName, $TypeParam)
    }

    $ParamDictionary
}

Function Get-SummaryHC {
    Param (
        [Parameter(Mandatory)]
        [String]$Name
    )

    Try {
        $EnvironmentList[$Environment].Summary.$Name
    }
    Catch {
        $Global:Error.Remove($Global:Error[0])

        if (-not ($Summary = Invoke-GetBusinessObjectSummaryHC -BusObName $Name)) {
            throw "No 'Summary' found in Cherwell with name '$Name'"
        }

        $EnvironmentList[$Environment].Summary.$Name = $Summary

        $Summary
    }
}

Function Get-SchemaHC {
    Param (
        [Parameter(Mandatory)]
        [String]$Name
    )

    Try {
        $EnvironmentList[$Environment].Schema.$Name
    }
    Catch {
        $Global:Error.Remove($Global:Error[0])

        if (-not ($Schema = Invoke-GetBusinessObjectSchemaHC -BusObId (Get-SummaryHC -Name $Name).busobid)) {
            throw "No 'Schema' found in Cherwell with name '$Name'"
        }

        $EnvironmentList[$Environment].Schema.$Name = $Schema

        $Schema
    }
}

Function Get-TemplateHC {
    Param (
        [Parameter(Mandatory)]
        [String]$Name
    )

    Try {
        $EnvironmentList[$Environment].Template.$Name
    }
    Catch {
        $Global:Error.Remove($Global:Error[0])

        if (-not ($Template = Invoke-GetBusinessObjectTemplateHC -BusObId (Get-SummaryHC -Name $Name).busobid)) {
            throw "No 'Template' found in Cherwell with name '$Name'"
        }

        $EnvironmentList[$Environment].Template.$Name = $Template

        $Template
    }
}

Function Add-CherwellTicketConfigItemHC {
    <#
    .SYNOPSIS
        Add a CI to a Cherwell ticket.

    .DESCRIPTION
        The CI can be added to a ticket by providing the ticket number or
        the ticket record ID (as retrieved by Get-CherwellTicketHC) and the
        CI record ID.

    .PARAMETER Environment
        The Cherwell system that needs to be addressed.

        All available systems are defined in the file 'Passwords.json' in the 
        module folder. Ex. Test, Prod, ...

    .PARAMETER ConfigItem
        Objects coming from 'Get-CherwellConfigItemHC -PassThru' or from the 
        API. Multiple CI's are supported and can be added at the same time.

    .PARAMETER Ticket
        Objects coming from 'Get-CherwellTicketHC -PassThru' or a ticket number.
        Multiple tickets are supported.

    .EXAMPLE
        Link all computer CI's that have the FriendlyName 'BELCL003000' to the 
        ticket in the test environment with number 5

        $CI = Get-CherwellConfigItemHC -PassThru -Environment Test -Type 'ConfigComputer' -Filter @{
            FieldName  = 'FriendlyName'
            Operator   = 'eq'
            FieldValue = 'BELCL003000'
        }

        Add-CherwellTicketConfigItemHC -Environment Test -Ticket 5 -ConfigItem $CI
 #>

    [CmdLetBinding()]
    Param (
        [Parameter(Mandatory)]
        [Object[]]$Ticket,
        [Parameter(Mandatory)]
        [PSCustomObject[]]$ConfigItem
    )
    DynamicParam {
        New-DynamicParameterHC -Name Environment
    }

    Begin {
        Try {
            $Environment = $PSBoundParameters.Environment
            Select-EnvironmentHC -Name $Environment

            $IncidentSchema = Get-SchemaHC -Name 'Incident'

            if (-not (
                    $CommonIncidentSchemaConfigItem = $IncidentSchema.relationships.Where( { $_.DisplayName -eq 'Incident Links Configuration Items' })
                )) {
                throw "Field 'Incident Links Configuration Items' not found in the Incident schema"
            }
        }
        Catch {
            throw "Failed adding CI to ticket: $_"
        }
    }

    Process {
        Try {
            foreach ($T in $Ticket) {
                $TicketObject = ConvertTo-TicketObjectHC -Item $T

                foreach ($C in $ConfigItem) {
                    $ConfigItemRecId = if ($C.PSObject.Properties['busObRecId']) {
                        $C.busObRecId
                    }

                    # Support non PassThru:
                    # $ConfigItemRecId = if ($C.PSObject.Properties['busObRecId']) {
                    #     $C.busObRecId
                    # }
                    # elseif ($C.PSObject.Properties['RecId']) {
                    #     $C.RecId
                    # }

                    if (-not $ConfigItemRecId) {
                        throw "Please provide a proper CI object as retrieved by 'Get-CherwellConfigItemHC -PassThru'."
                    }

                    #region Add config item to ticket
                    Write-Verbose "Add CI '$ConfigItemRecId' to ticket '$($TicketObject.busObPublicId)'"

                    $Params = @{
                        parentBusObId    = $IncidentSchema.busObId
                        ParentBusobRecId = $TicketObject.busObRecId
                        RelationshipId   = $CommonIncidentSchemaConfigItem.relationshipId
                        BusobId          = $CommonIncidentSchemaConfigItem.target
                        BusobRecId       = $ConfigItemRecId
                    }
                    Invoke-LinkRelatedBusinessObjectHC  @Params
                    #endregion
                }
            }
        }
        Catch {
            throw "Failed adding the CI to ticket in Cherwell '$Environment': $_"
        }
    }
}

Function Add-CherwellTicketAttachmentHC {
    <#
    .SYNOPSIS
        Add an attachment to a ticket.

    .DESCRIPTION
        One or more attachments can be added to a ticket by specifying the full 
        file name and the ticket number. When the attachment is added correctly 
        the file name is returned.

    .PARAMETER TicketNr
        The ticket number to which the files need to be added as attachments.

    .PARAMETER File
        Can be one or more files. The full name of the path is required.

    .PARAMETER Environment
        The Cherwell system that needs to be addressed.

        All available systems are defined in the file 'Passwords.json' in the 
        module folder. Ex. Test, Prod, ...

    .EXAMPLE
        Add the file 'Finance report.txt' to ticket '150150' as an attachment

        Add-CherwellTicketAttachmentHC -Environment Test -TicketNr 150150 -File 'E:\Accounting\Finance report.txt' -Verbose
 #>

    [CmdLetBinding()]
    [OutputType([String[]])]
    Param (
        [Parameter(Mandatory)]
        [Int]$TicketNr,
        [Parameter(Mandatory)]
        [ValidateScript( { Test-Path -Path $_ })]
        [String[]]$File
    )
    DynamicParam {
        New-DynamicParameterHC -Name Environment
    }

    Process {
        Try {
            $Environment = $PSBoundParameters.Environment
            Select-EnvironmentHC -Name $Environment

            $IncidentSummary = Get-SummaryHC -Name 'Incident'

            #region Get the incident RecId
            Try {
                $Params = @{
                    BusObId  = $IncidentSummary.busobid
                    PublicId = $ticketNr
                }
                $IncidentTicket = Invoke-GetBusinessObjectHC @Params
                $BusObRecId = $IncidentTicket.busObRecId
            }
            Catch {
                throw 'Ticket number not found in Cherwell.'
            }
            #endregion

            #region Upload the files
            foreach ($F in $File) {
                Try {
                    $FileItem = Get-Item -Path $F -EA Stop

                    $Params = @{
                        FileName   = $FileItem.Name
                        BusObId    = $IncidentSummary.busobId
                        BusObRecId = $BusObRecId
                        Offset     = 0
                        TotalSize  = $FileItem.Length
                        Body       = $FileItem.FullName
                    }
                    Invoke-UploadBusinessObjectAttachmentHC @Params

                    $F
                    Write-Verbose "Ticket '$TicketNr' attachment '$F' added"
                }
                Catch {
                    Write-Error "Failed adding attachment file '$F': $_"
                }
            }
            #endregion
        }
        Catch {
            throw "Failed uploading the attachment '$File' to ticket '$TicketNr' in Cherwell '$Environment': $_"
        }
    }
}

Function Add-CherwellTicketDetailHC {
    <#
    .SYNOPSIS
        Add a note, task or other object to a Cherwell ticket.

    .DESCRIPTION
        The object can be added to a ticket by providing the ticket number or
        the ticket record ID (as retrieved by Get-CherwellTicketHC) and the
        details of the object with the KeuValuePair parameter.

    .PARAMETER Environment
        The Cherwell system that needs to be addressed.

        All available systems are defined in the file 'Passwords.json' in the 
        module folder. Ex. Test, Prod, ...

    .PARAMETER Type
        The type of object to add to the ticket.

    .PARAMETER Ticket
        Objects coming from 'Get-CherwellTicketHC -PassThru' or a ticket number.
        Multiple tickets are supported.

    .PARAMETER KeyValuePair
        One or more hash tables containing key value pair combinations to 
        create new notes, tasks and other objects. The key name represents the 
        field name in Cherwell and the key value contains the Cherwell value 
        desired for that field.

    .PARAMETER PassThru
        Return the output of API calls in their raw format. The output has not 
        been converted or changed in any way and can be used by other functions 
        as input. This will allow the piping of objects from one function to 
        the next.

    .EXAMPLE
        Add a journal note to ticket 123123

        $Params = @{
            Environment  = 'Test'
            Ticket       = 123123
            Type         = 'JournalNote'
            KeyValuePair = @{
                Details = 'Please take the appropriate action'
            }
        }
        Add-CherwellTicketDetailHC @Params

    .EXAMPLE
        Add a journal note to ticket 42173 in the test environment

        Add-CherwellTicketDetailHC Environment 'Test' -Type JournalNote -Ticket 42173 -KeyValuePair @{
            Details      = 'This is a journal note'
            Priority     = 'High'
            MarkedAsRead = $false
        }

    .EXAMPLE
        Add a new task to ticket 42173 in the test environment

        Add-CherwellTicketDetailHC @testParams -Type 'Task' -Ticket 42173 -KeyValuePair @{
            Type        = 'Work Item'
            Title       = 'Add printer to server'
            Description = 'Printer BELPR0001 needs ot be added to BELSF0001'
        }
#>

    [CmdLetBinding()]
    [OutputType()]
    Param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [Alias('TicketNr')]
        [Object[]]$Ticket,
        [Parameter(Mandatory)]
        [HashTable]$KeyValuePair,
        [Switch]$PassThru
    )
    DynamicParam {
        New-DynamicParameterHC -Name Environment, Type
    }

    Begin {
        Try {
            $Environment = $PSBoundParameters.Environment
            $Type = $PSBoundParameters.Type

            Select-EnvironmentHC -Name $Environment

            $IncidentSchema = Get-SchemaHC -Name 'Incident'

            #region Get RelationshipId
            $RelationshipId = $IncidentSchema.relationships.Where( {
                    $_.DisplayName -eq $RelationshipEnum[$Type] }) |
            Select-Object -ExpandProperty 'RelationshipId'

            if (-not $RelationshipId) {
                throw "RelationshipId not found for type '$Type'."
            }
            #endregion

            $Template = Get-TemplateHC -Name $Type
        }
        Catch {
            throw "Failed adding ticket details of type '$Type' to ticket '$Ticket': $_ "
        }
    }

    Process {
        Try {
            Write-Verbose "Add ticket details of type '$Type' in environment '$Environment'"

            foreach ($T in $Ticket) {
                $TicketObject = ConvertTo-TicketObjectHC -Item $T

                #region Create a new template
                $Params = @{
                    Name         = $Template
                    KeyValuePair = $KeyValuePair
                }
                $DirtyTemplate = New-TemplateHC @Params
                #endregion

                #region Add object to ticket
                Write-Verbose "Add ticket details of type '$Type' to ticket '$($TicketObject.busObPublicId)'"

                $Params = @{
                    Uri  = ($Uri + 'api/V1/SaveRelatedBusinessObject')
                    Body = [System.Text.Encoding]::UTF8.GetBytes((@{
                                parentBusObId       = $IncidentSchema.busObId
                                parentBusObPublicId = $TicketObject.busObPublicId
                                relationshipId      = $RelationshipId
                                fields              = $DirtyTemplate.Fields
                                persist             = $true
                            } | ConvertTo-Json))
                }
                $Result = Invoke-RestMethod @PostParams @Params

                if ($PassThru) {
                    $Result
                }
                #endregion
            }
        }
        Catch {
            throw "Failed adding ticket details of type '$Type' to ticket '$Ticket' in Cherwell '$Environment': $_ "
        }
    }
}

#region Add-CherwellTicketDetailHC proxy functions
Function Add-CherwellTicketJournalNoteHC {
    <#
    .SYNOPSIS
        Add a journal note to a Cherwell ticket.

    .DESCRIPTION
        The journal note can be added to a ticket by providing the ticket 
        number or the ticket record ID (as retrieved by Get-CherwellTicketHC) 
        and the note details.

    .PARAMETER Environment
        The Cherwell system that needs to be addressed.

        All available systems are defined in the file 'Passwords.json' in the 
        module folder. Ex. Test, Prod, ...

    .PARAMETER Ticket
        Objects coming from 'Get-CherwellTicketHC -PassThru' or a ticket number.
        Multiple tickets are supported.

    .PARAMETER KeyValuePair
        One or more hash tables containing key value pair combinations to 
        create new notes. The key name represents the field name in Cherwell 
        and the key value contains the Cherwell value desired for that field.

    .PARAMETER PassThru
        Return the output of API calls in their raw format. The output has not 
        been converted or changed in any way and can be used by other functions 
        as input. This will allow the piping of objects from one function to 
        the next.

    .EXAMPLE
        Add a new journal note to ticket 123123 with the text 'Please take the 
        appropriate action in the Cherwell test environment. With the switch 
        'PassThru' the output of the API is returned, which can be used by 
        other functions.

        Add-CherwellTicketJournalNoteHC -Environment Test -Ticket 123123 -KeyValuePair @{
            Details = 'Please take the appropriate action'
        } -PassThru

    .EXAMPLE
        Add a journal note to ticket 42173 with the priority set to 'High' and 
        MarkedAsRead to 'False'. The 'Verbose' switch provides progress 
        indications to see what's happening in the background.

        Add-CherwellTicketJournalNoteHC -Environment Test -Ticket 42173 -KeyValuePair @{
            Details      = 'This is a journal note'
            Priority     = 'High'
            MarkedAsRead = $false
        } -Verbose
 #>

    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
        [Alias('TicketNr')]
        [System.Object[]]${Ticket},
        [Parameter(Mandatory = $true, Position = 2)]
        [HashTable]${KeyValuePair},
        [Switch]${PassThru}
    )
    DynamicParam {
        New-DynamicParameterHC -Name Environment
    }

    Begin {
        try {
            $outBuffer = $null
            if ($PSBoundParameters.TryGetValue('OutBuffer', [ref]$outBuffer)) {
                $PSBoundParameters['OutBuffer'] = 1
            }
            $wrappedCmd = $ExecutionContext.InvokeCommand.GetCommand(
                'Add-CherwellTicketDetailHC', [System.Management.Automation.CommandTypes]::Function
            )
            $PSBoundParameters['Type'] = 'JournalNote'
            $scriptCmd = { & $wrappedCmd @PSBoundParameters }
            $steppablePipeline = $scriptCmd.GetSteppablePipeline()
            $steppablePipeline.Begin($PSCmdlet)
        }
        catch { throw }
    }
    Process { try { $steppablePipeline.Process($_) }catch { throw } }
    End { try { $steppablePipeline.End() }catch { throw } }
}

Function Add-CherwellTicketTaskHC {
    <#
    .SYNOPSIS
        Add a task to a Cherwell ticket.

    .DESCRIPTION
        The task can be added to a ticket by providing the ticket number or
        the ticket record ID (as retrieved by Get-CherwellTicketHC) and the
        task details.

    .PARAMETER Environment
        The Cherwell system that needs to be addressed.

        All available systems are defined in the file 'Passwords.json' in the 
        module folder. Ex. Test, Prod, ...

    .PARAMETER Ticket
        Objects coming from 'Get-CherwellTicketHC -PassThru' or a ticket number.
        Multiple tickets are supported.

    .PARAMETER KeyValuePair
        One or more hash tables containing key value pair combinations to 
        create new tasks. The key name represents the field name in Cherwell 
        and the key value contains the Cherwell value desired for that field.

    .PARAMETER PassThru
        Return the output of API calls in their raw format. The output has not 
        been converted or changed in any way and can be used by other functions 
        as input. This will allow the piping of objects from one function to 
        the next.

    .EXAMPLE
        Add a work item task to ticket 42173

        Add-CherwellTicketTaskHC -Environment Test -Ticket 42173 -KeyValuePair @{
            Type        = 'Work Item'
            Title       = 'Add printer to server'
            Description = 'Printer BELPR0001 needs ot be added to BELSF0001'
        }

    .EXAMPLE
        Add an approval task to ticket 10 that is assigned to the team 'BNL'

        Add-CherwellTicketTaskHC -Environment Test -Ticket 10 -KeyValuePair @{
            Type         = 'Approval Task'
            Title        = 'Please approve access to ICP'
            Description  = 'Please provide access to the ICP, this is a new service desk agent.'
            OwnedByTeam  = 'BNL'
        }

    .EXAMPLE
        Add an approval request task to ticket 5

        Add-CherwellTicketTaskHC -Environment Test -Ticket 5 -KeyValuePair @{
            Type        = 'Approval Request'
            Title       = 'Request approval for ICP access'
            Description = 'Please provide access to the ICP, this is a new service desk agent.'
            OwnedByTeam = 'BNL INFRA'
        }
 #>

    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
        [Alias('TicketNr')]
        [System.Object[]]${Ticket},
        [Parameter(Mandatory = $true, Position = 2)]
        [HashTable]${KeyValuePair},
        [Switch]${PassThru}
    )
    DynamicParam {
        New-DynamicParameterHC -Name Environment
    }

    Begin {
        try {
            $outBuffer = $null
            if ($PSBoundParameters.TryGetValue('OutBuffer', [ref]$outBuffer)) {
                $PSBoundParameters['OutBuffer'] = 1
            }
            $wrappedCmd = $ExecutionContext.InvokeCommand.GetCommand(
                'Add-CherwellTicketDetailHC', [System.Management.Automation.CommandTypes]::Function
            )

            $PSBoundParameters['Type'] = 'Task'

            $scriptCmd = { & $wrappedCmd @PSBoundParameters }
            $steppablePipeline = $scriptCmd.GetSteppablePipeline()
            $steppablePipeline.Begin($PSCmdlet)
        }
        catch {
            throw
        }
    }
    Process { try { $steppablePipeline.Process($_) }catch { throw } }
    End { try { $steppablePipeline.End() }catch { throw } }
}
#endregion

Function Convert-ApiErrorHC {
    Param (
        $Exception
    )

    Try {
        $Json = $Exception | ConvertFrom-Json
        "API Error: {0} - {1}" -f $Json.errorCode, $Json.errorMessage
        $Global:Error.RemoveAt(0)
    }
    Catch {
        $Global:Error.RemoveAt(0)
        $Exception
    }
}

Function ConvertTo-PSCustomObjectHC {
    Param (
        [PSCustomObject[]]$BusinessObject
    )
    foreach ($B in $BusinessObject) {
        $NewObject = [Ordered]@{ }

        foreach ($F in ($B.Fields | Sort-Object -Property Name)) {
            $NewObject.($F.Name) = $F.Value
        }

        [PSCustomObject]$NewObject
    }
}

Function ConvertTo-TicketObjectHC {
    Param (
        [Parameter(Mandatory)]
        [Object]$Item
    )

    Try {
        if ($Item -is [System.Management.Automation.PSCustomObject]) {
            $Item
        }
        elseif (($Item -is [Int]) -or ($Item -is [String])) {
            if (-not ($TicketObject = Get-CherwellTicketHC -Environment $Environment -TicketNr $Item -PassThru)) {
                throw "Ticket number '$Item' not found"
            }

            $TicketObject
        }
        else {
            throw "Object type '$($Item.GetType())' not supported."
        }
    }
    Catch {
        throw "The object cannot be converted to a valid ticket object: $_."
    }
}

Function Get-CherwellConfigItemHC {
    <#
    .SYNOPSIS
        Retrieve CI's from Cherwell.

    .DESCRIPTION
        Retrieve configuration items like servers, computers, printers, ... 
        from the Cherwell CMDB. When only the 'Type' parameter is used and the 
        'Filter' is omitted, all CI's in the CMDB for that type of object will 
        be returned.

        In case the 'Filter' is used, only those CI's matching the condition of 
        that specific type will be returned.

    .PARAMETER Environment
        The Cherwell system that needs to be addressed.

        All available systems are defined in the file 'Passwords.json' in the 
        module folder. Ex. Test, Prod, ...

    .PARAMETER Type
        The type of configuration item as it is known in Cherwell. To retrieve 
        a list of possible values the function 'Get-CherwellConfigItemTypeHC' 
        can be used.

    .PARAMETER Filter
        A hash table containing the search conditions used by the API. When the 
        search conditions is met the configuration items that match these 
        conditions are returned. When no object is matching the condition, 
        nothing is returned.

        The following fields/keys are mandatory:
        @{
            FieldName  = ''
            Operator   = ''
            FieldValue = ''
        }

        The following operators are supported: 'eq', 'gt', 'lt', 'contains' and 
        'startswith'

        You can specify more than one filter. If you add multiple filters for 
        the same FieldName, the result is an OR operation between those fields. 
        If the FieldNames are different, the result is an AND operation between 
        those fields.

    .PARAMETER InputObject
        Accepts a CI object coming from 'New-CherwellCongigItemHC'. This can be 
        convenient to verify all the properties of a CI. When 'InputObject' is 
        used, all properties are returned by default.

    .PARAMETER Property
        Specifies the properties to select. When using '-Property *' all 
        properties are returned by the API. The downside is that this is slower 
        when working with large data sets.

    .PARAMETER PageSize
        Defines how many results are retrieved at once from the Cherwell 
        server. This can be useful for managing the speed of retrieving 
        results. A high PageSize number might take longer to execute but will 
        make less calls to the API. A small PageSize number might
        be faster in some cases but result in more calls.

        Experimenting with this feature can be done by using the '-Verbose' 
        switch to analyze speed and the number of retrieved objects.

    .PARAMETER PassThru
        Return the output of API calls in their raw format. The output has not 
        been converted or changed in any way and can be used by other functions 
        as input through the 'InputObject' parameter. This will allow the 
        piping of objects from one function to the next.

        When 'PassThru' is omitted the returning business object property 
        'Fields', that comes from the API, will be converted to a 
        PSCustomObject.

    .EXAMPLE
        Retrieves all computer CI's with 'FriendlyName' equal to 'BELCL003000'

        $Params = @{
            Environment = 'Test'
            Type        = 'ConfigComputer'
            Filter      = @{
                FieldName  = 'FriendlyName'
                Operator   = 'eq'
                FieldValue = 'BELCL003000'
            }
        }
        Get-CherwellConfigItemHC @Params

    .EXAMPLE
        Retrieves all server CI's known in Cherwell
        
        $Params = @{
            Environment = 'Test'
            Type        = 'ConfigServer'
            PageSize    = 1000
        }
        Get-CherwellConfigItemHC @Params

    .EXAMPLE
        Retrieve all server CI's that contain the text 'Virtual' in the property 'AssetType'
        
        $Params = @{
            Environment = 'Test'
            Type        = 'ConfigServer'
            Filter      = @{
                FieldName  = 'AssetType'
                Operator   = 'contains'
                FieldValue = 'Virtual'
            }
        }
        Get-CherwellConfigItemHC @Params

    .EXAMPLE
        Retrieve all printer CI's known in Cherwell with all their properties. 
        Note that this will take longer but the results are more complete as 
        all possible fields are returned. The '-Verbose' switch is used to show 
        the progress and the PageSize is set very low to allow the API to 
        collect and return the info faster than if we would request 1000 objects
        at once, especially when using the '-Property *'.

        $Params = @{
            Environment      = 'Test'
            Type             = 'ConfigPrinter'
            Property         = '*'
            PageSize         = 500
        }
        Get-CherwellConfigItemHC @Params -Verbose

    .EXAMPLE
        Retrieve all details about a newly created CI in the CI Printer table. 
        Because the objects are piped the parameter 'InputObject' is used. When 
        this parameter is used all details are retrieved by default and 
        '-Property *' is not needed.

        $CIKeyValuePair = @{
            CIType          = 'ConfigPrinter'
            Owner           = 'thardey'
            SupportedByTeam = 'BNL'
            FriendlyName    = 'BELPROOST14'
            SerialNumber    = '9874605849684'
            Country         = 'BNL'
            PrinterType     = 'Local printer'
            Manufacturer    = 'Konica Minolta'
            Model           = 'Bizhub 3545c'
        }
        $CI = New-CherwellConfigItemHC -Environment Stage -KeyValuePair $CIKeyValuePair

        $Params = @{
            Environment = 'Stage'
            Type        = $CIKeyValuePair.CIType
        }
        $Actual = $CI | Get-CherwellConfigItemHC @Params

    .EXAMPLE
        Retrieve all server CI's that match the hostname of the local machine 
        and return the results as business objects by using the switch 
        'PassThru'.

        $Params = @{
            Environment = 'Test'
            Type        = 'ConfigServer'
            Filter      = @{
                FieldName  = 'HostName'
                Operator   = 'eq'
                FieldValue = $env:COMPUTERNAME
            }
            PassThru    = $true
        }
        Get-CherwellConfigItemHC @Params

    .EXAMPLE
        Retrieve all computer CI's starting with the string 'BEL', 'NLD' or 
        'LUX' in the property 'Hostname' and retrieve all properties of these 
        objects. The '-Verbose' switch shows the progress

        $Params = @{
            Environment      = 'Prod'
            Type             = 'ConfigComputer'
            Property         = '*'
            Filter           = @(
                @{
                    FieldName  = 'Hostname'
                    Operator   = 'startswith'
                    FieldValue = 'BEL'
                }
                @{
                    FieldName  = 'Hostname'
                    Operator   = 'startswith'
                    FieldValue = 'NLD'
                }
                @{
                    FieldName  = 'Hostname'
                    Operator   = 'startswith'
                    FieldValue = 'LUX'
                }
            )
        }
        Get-CherwellConfigItemHC @Params -Verbose
#>

    [CmdletBinding(DefaultParameterSetName = 'Default')]
    [OutputType([PSCustomObject[]])]
    Param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [String]$Type,
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'InputObject')]
        [ValidateNotNullOrEmpty()]
        [PSCustomObject]$InputObject,
        [ValidateNotNullOrEmpty()]
        [String[]]$Property,
        [Parameter(Mandatory, ParameterSetName = 'Filter')]
        [ValidateNotNullOrEmpty()]
        [HashTable[]]$Filter,
        [ValidateRange(10, 5000)]
        [Int]$PageSize = 1000,
        [Switch]$PassThru
    )
    DynamicParam {
        New-DynamicParameterHC -Name Environment
    }

    Begin {
        Try {
            $Environment = $PSBoundParameters.Environment
            Select-EnvironmentHC -Name $Environment

            Try {
                $ConfigItemSchema = Get-SchemaHC -Name $Type
            }
            Catch {
                $CiTypes = Get-CherwellConfigItemTypeHC -Environment $Environment
                throw "$_. Only the following CI types are supported: $($CiTypes -join ', ')"
            }
        }
        Catch {
            throw "Failed retrieving the CI for type '$Type' from Cherwell '$Environment': $_"
        }
    }

    Process {
        Try {
            Write-Verbose "Get configuration item from environment '$Environment'"

            if ($InputObject) {
                Write-Verbose "Get CI input object '$($InputObject.busObRecId)'"

                $Params = @{
                    BusObId    = $ConfigItemSchema.busObId
                    BusObRecId = $InputObject.busObRecId
                }
                $Result = Invoke-GetBusinessObjectHC @Params

                if ($Result) {
                    if ($PassThru) {
                        $Result
                    }
                    else {
                        ConvertTo-PSCustomObjectHC -BusinessObject $Result
                    }
                    Write-Verbose "Retrieved CI for type '$Type'"
                }
            }
            else {
                Write-Verbose "Get CI of type '$Type'"

                if ($Filter -or $Property) {
                    $null = $PSBoundParameters.Add('Schema', $ConfigItemSchema )
                }

                $null = $PSBoundParameters.Remove('Type')
                $null = $PSBoundParameters.Remove('Environment')
                $PSBoundParameters.BusObId = $ConfigItemSchema.busobid

                Invoke-GetSearchResultsHC @PSBoundParameters
            }
        }
        Catch {
            $M = Convert-ApiErrorHC $_
            throw "Failed retrieving the CI for type '$Type' from Cherwell '$Environment': $M"
        }
    }
}

Function Get-CherwellConfigItemTypeHC {
    <#
    .SYNOPSIS
        Retrieve a list of Cherwell CI types.

    .DESCRIPTION
        This can be convenient to discover the available CI types in Cherwell. 
        These are all the configuration item types available in a specific 
        environment. When changing to other environments (Stage, Test, Prod, ...
        ) it is possible that different configuration item types are found.

        These values can be used with the configuration item 'Type' parameter 
        of the configuration item functions (Get-CherwellConfigItemHC, 
        Set-CherwellConfigItemHC, ...).

    .PARAMETER Environment
        The Cherwell system that needs to be addressed.

        All available systems are defined in the file 'Passwords.json' in the 
        module folder. Ex. Test, Prod, ...

    .EXAMPLE
        Retreive all the Cherwell configuration items in the test environment

        Get-CherwellConfigItemTypeHC -Environment Test
 #>

    [OutputType([String[]])]
    [CmdletBinding()]
    Param ()
    DynamicParam {
        New-DynamicParameterHC -Name Environment
    }

    Begin {
        Try {
            $Environment = $PSBoundParameters.Environment
            Select-EnvironmentHC -Name $Environment

        }
        Catch {
            throw "Failed retrieving the CI types from Cherwell '$Environment': $_"
        }
    }

    Process {
        Try {
            Write-Verbose "Get Cherwell configuration item type from environment '$Environment'"

            $Params = @{
                Uri = ($Uri + 'api/V1/getbusinessobjectsummaries/type/All')
            }
            $AllSummaries = Invoke-RestMethod @GetParams @Params

            $ConfigurationItem = $AllSummaries.where( { $_.Name -eq 'ConfigurationItem' })

            $ConfigurationItem.GroupSummaries.Name
        }
        Catch {
            throw "Failed retrieving the CI types from Cherwell '$Environment': $_"
        }
    }
}

Function Get-CherwellSearchResultsHC {
    <#
    .SYNOPSIS
        Retrieve Cherwell search results.

    .DESCRIPTION
        Retrieve Cherwell search results for tables like 'incident 
        subcategory', 'Location','SubCategory'...

    .PARAMETER Environment
        The Cherwell system that needs to be addressed.

        All available systems are defined in the file 'Passwords.json' in the 
        module folder. Ex. Test, Prod, ...

    .PARAMETER Filter
        A hash table containing the search conditions used by the API. When the 
        search conditions is met the configuration items that match these 
        conditions are returned. When no object is matching the condition, 
        nothing is returned.

        The following fields/keys are mandatory:
        @{
            FieldName  = ''
            Operator   = ''
            FieldValue = ''
        }

        The following operators are supported: 'eq', 'gt', 'lt', 'contains' and 
        'startswith'

        You can specify more than one filter. If you add multiple filters for 
        the same FieldName, the result is an OR operation between those fields. 
        If the FieldNames are different, the result is an AND operation between 
        those fields.

    .PARAMETER Property
        Specifies the properties to select. When using '-Property *' all 
        properties are returned by the API. The downside is that this is slower 
        when working with large data sets.

    .PARAMETER PageSize
        Defines how many results are retrieved at once from the Cherwell 
        server. This can be useful for managing the speed of retrieving 
        results. A high PageSize number might take longer to execute but will 
        make less calls to the API. A small PageSize number might
        be faster in some cases but result in more calls.

        Experimenting with this feature can be done by using the '-Verbose' 
        switch to analyze speed and the number of retrieved objects.

    .PARAMETER PassThru
        Return the output of API calls in their raw format. The output has not 
        been converted or changed in any way and can be used by other functions 
        as input through the 'InputObject' parameter. This will allow the 
        piping of objects from one function to the next.

        When 'PassThru' is omitted the returning business object property 
        'Fields', that comes from the API, will be converted to a 
        PSCustomObject.

    .EXAMPLE
        Retrieve all incident subcategories from the Cherwell 'Prod' 
        environment and display the progress by using the '-Verbose' switch

        Get-CherwellSearchResultsHC -Environment Prod -Name IncidentSubCategory -Verbose

    .EXAMPLE
        Retrieve all incident subcategories from the Cherwell 'Prod' 
        environment in batches of 100 at a time. Request all possible 
        properties and display the progress by using the '-Verbose' switch

        Get-CherwellSearchResultsHC -Environment Prod -Name IncidentSubCategory -PageSize 100 -Property * -Verbose

    .EXAMPLE
        Retrieve all incident subcategories that have the 'Country' set to 
        'BNL' from the Cherwell 'Prod' environment and select only the 
        properties 'IncidentCategory', 'Service' and 'ServiceStatus'

        Get-CherwellSearchResultsHC -Environment Prod -Name IncidentSubCategory -Filter @{
            FieldName  = 'Country'
            Operator   = 'eq'
            FieldValue = 'BNL'
        } -Property IncidentCategory, Service, ServiceStatus
#>

    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    Param (
        [Parameter(Mandatory)]
        [String]$Name,
        [ValidateNotNullOrEmpty()]
        [String[]]$Property,
        [ValidateNotNullOrEmpty()]
        [HashTable[]]$Filter,
        [ValidateRange(10, 5000)]
        [Int]$PageSize = 1000,
        [Switch]$PassThru
    )
    DynamicParam {
        New-DynamicParameterHC -Name Environment
    }

    Begin {
        Try {
            $Environment = $PSBoundParameters.Environment
            Select-EnvironmentHC -Name $Environment

        }
        Catch {
            throw "Failed retrieving '$Name': $_"
        }
    }

    Process {
        Try {
            Write-Verbose "Get '$Name' from environment '$Environment'"

            if ($Filter -or $Property) {
                $null = $PSBoundParameters.Add('Schema', (Get-SchemaHC -Name $Name) )
            }

            $null = $PSBoundParameters.Remove('Environment')
            $null = $PSBoundParameters.Remove('Name')
            $PSBoundParameters.BusObId = (Get-SummaryHC -Name $Name).busobid

            Invoke-GetSearchResultsHC @PSBoundParameters
        }
        Catch {
            throw "Failed retrieving '$Name' from Cherwell '$Environment': $_"
        }
    }
}

#region Get-CherwellSearchResultsHC proxy functions
Function Get-CherwellChangeStandardTemplateHC {
    <#
    .SYNOPSIS
        Retrieve Cherwell change standard template

    .DESCRIPTION
        Retrieve all change standard templates in the Cherwell system.

    .PARAMETER Environment
        The Cherwell system that needs to be addressed.

        All available systems are defined in the file 'Passwords.json' in the 
        module folder. Ex. Test, Prod, ...

    .PARAMETER Filter
        A hash table containing the search conditions used by the API. When the 
        search conditions is met the configuration items that match these 
        conditions are returned. When no object is matching the condition, 
        nothing is returned.

        The following fields/keys are mandatory:
        @{
            FieldName  = ''
            Operator   = ''
            FieldValue = ''
        }

        The following operators are supported: 'eq', 'gt', 'lt', 'contains' and 
        'startswith'

        You can specify more than one filter. If you add multiple filters for 
        the same FieldName, the result is an OR operation between those fields. 
        If the FieldNames are different, the result is an AND operation between 
        those fields.

    .PARAMETER Property
        Specifies the properties to select. When using '-Property *' all 
        properties are returned by the API. The downside is that this is slower 
        when working with large data sets.

    .PARAMETER PageSize
        Defines how many results are retrieved at once from the Cherwell 
        server. This can be useful for managing the speed of retrieving 
        results. A high PageSize number might take longer to execute but will 
        make less calls to the API. A small PageSize number mightbe faster in 
        some cases but result in more calls.

        Experimenting with this feature can be done by using the '-Verbose' 
        switch to analyze speed and the number of retrieved objects.

    .PARAMETER PassThru
        Return the output of API calls in their raw format. The output has not 
        been converted or changed in any way and can be used by other functions 
        as input through the 'InputObject' parameter. This will allow the 
        piping of objects from one function to the next.

        When 'PassThru' is omitted the returning business object property 
        'Fields', that comes from the API, will be converted to a 
        PSCustomObject.

    .EXAMPLE
        Retrieve all change standard templates with the 'Country' property set 
        to 'BEL' from the 'Test' environment

        Get-CherwellChangeStandardTemplateHC -Environment Test -Filter @{
            FieldName  = 'Country'
            Operator   = 'eq'
            FieldValue = 'BEL'
        }

    .EXAMPLE
        Retrieve all change standard templates in Cherwell with all possible 
        properties from the 'Prod' environment and display the progress by 
        using the '-Verbose' switch

        Get-CherwellChangeStandardTemplateHC -Environment Prod -Property * -Verbose
#>

    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    Param (
        [Parameter(Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string[]]${Property},
        [Parameter(Position = 2)]
        [ValidateNotNullOrEmpty()]
        [hashTable[]]${Filter},
        [Parameter(Position = 3)]
        [ValidateRange(10, 5000)]
        [int]${PageSize},
        [switch]${PassThru}
    )
    DynamicParam {
        New-DynamicParameterHC -Name Environment
    }

    Begin {
        try {
            $outBuffer = $null
            if ($PSBoundParameters.TryGetValue('OutBuffer', [ref]$outBuffer)) {
                $PSBoundParameters['OutBuffer'] = 1
            }
            $wrappedCmd = $ExecutionContext.InvokeCommand.GetCommand(
                'Get-CherwellSearchResultsHC', [System.Management.Automation.CommandTypes]::Function
            )
            $PSBoundParameters['Name'] = 'ChangeStandardTemplate'
            $scriptCmd = { & $wrappedCmd @PSBoundParameters }
            $steppablePipeline = $scriptCmd.GetSteppablePipeline()
            $steppablePipeline.Begin($PSCmdlet)
        }
        catch { throw }
    }
    Process { try { $steppablePipeline.Process($_) } catch { throw } }
    End { try { $steppablePipeline.End() } catch { throw } }
}

Function Get-CherwellCustomerHC {
    <#
    .SYNOPSIS
        Retrieve Cherwell customer

    .DESCRIPTION
        Retrieve all Cherwell customers. Cherwell customers are users that are 
        not allowed to work in the Cherwell tool, but are reporting issues. 
        Users found with this query function can be used in the fields
        `CustomerRecID` and 'SubmitOnBehalfOfID'.

    .PARAMETER Environment
        The Cherwell system that needs to be addressed.

        All available systems are defined in the file 'Passwords.json' in the 
        module folder. Ex. Test, Prod, ...

    .PARAMETER Filter
        A hash table containing the search conditions used by the API. When the 
        search conditions is met the configuration items that match these 
        conditions are returned. When no object ismatching the condition, 
        nothing is returned.

        The following fields/keys are mandatory:
        @{
            FieldName  = ''
            Operator   = ''
            FieldValue = ''
        }

        The following operators are supported: 'eq', 'gt', 'lt', 'contains' and 'startswith'

        You can specify more than one filter. If you add multiple filters for 
        the same FieldName, the result is an OR operation between those fields. 
        If the FieldNames are different, the result is an AND operation between 
        those fields.

    .PARAMETER Property
        Specifies the properties to select. When using '-Property *' all 
        properties are returned by the API. The downside is that this is slower 
        when working with large data sets.

    .PARAMETER PageSize
        Defines how many results are retrieved at once from the Cherwell 
        server. This can be useful for managing the speed of retrieving 
        results. A high PageSize number might take longer to execute but will 
        make less calls to the API. A small PageSize number mightbe faster in 
        some cases but result in more calls.

        Experimenting with this feature can be done by using the '-Verbose' 
        switch to analyze speed and the number of retrieved objects.

    .PARAMETER PassThru
        Return the output of API calls in their raw format. The output has not 
        been converted or changed in any way and can be used by other functions 
        as input through the 'InputObject' parameter. This will allow the 
        piping of objects from one function to the next.

        When 'PassThru' is omitted the returning business object property 
        'Fields', that comes from the API, will be converted to a 
        PSCustomObject.

    .EXAMPLE
        Get-CherwellCustomerHC -Environment Prod -Verbose

        Retrieve all customers from the Cherwell production environment and 
        display the progress by using the '-Verbose' switch.

    .EXAMPLE
        Get-CherwellCustomerHC -Environment Prod -PageSize 100 -Property * -Verbose

        Retrieve all customers from the Cherwell production environment in batches of 100 at a time. Request all possible fields and  and display 
        the progress by using the '-Verbose' switch.

    .EXAMPLE
        Retrieve all customers from the Cherwell test environment that have the
        field 'LastName' set to 'Green'

        Get-CherwellCustomerHC -Environment Test -Filter @{
            FieldName  = 'LastName'
            Operator   = 'eq'
            FieldValue = 'Green'
        }

    .EXAMPLE
        Retrieve all customers created before '8/8/2016' with all their 
        properties

        Get-CherwellCustomerHC -Environment Stage -Filter @{
            FieldName  = 'CreatedDateTime'
            Operator   = 'lt'
            FieldValue = '8/8/2018'
        } -Verbose -Property *

    .EXAMPLE
        Retrieve all customers with the field 'City' set to 'London'

        Get-CherwellCustomerHC -Environment Test -Filter @{
            FieldName  = 'City'
            Operator   = 'eq'
            FieldValue = 'London'
        }

    .EXAMPLE
        Retrieve the Cherwell customer with SamAccountName 'bmarley'

        Get-CherwellCustomerHC -Environment Test -Property * -Filter @{
            FieldName  = 'SamAccountName'
            Operator   = 'eq'
            FieldValue = 'bmarley'
        }
      
    .EXAMPLE
        Retrieve the Cherwell customer with EMail 'gmail@chucknoriss.com'

        $Requester = Get-CherwellCustomerHC -Environment Test -Property * -Filter @{
            FieldName  = 'EMail'
            Operator   = 'eq'
            FieldValue = 'gmail@chucknoriss.com'
        }
        
        $SubittedBy = Get-CherwellCustomerHC -Environment Test -Property * -Filter @{
            FieldName  = 'EMail'
            Operator   = 'eq'
            FieldValue = 'bon.marley@ja.com'
        }

        $CustomerRecID = $Requester.busObRecId
        $SubmitOnBehalfOfID = $SubittedBy.busObRecId
#>

    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    Param (
        [Parameter(Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string[]]${Property},
        [Parameter(Position = 2)]
        [ValidateNotNullOrEmpty()]
        [hashTable[]]${Filter},
        [Parameter(Position = 3)]
        [ValidateRange(10, 5000)]
        [int]${PageSize},
        [switch]${PassThru}
    )
    DynamicParam {
        New-DynamicParameterHC -Name Environment
    }

    Begin {
        try {
            $outBuffer = $null
            if ($PSBoundParameters.TryGetValue('OutBuffer', [ref]$outBuffer)) {
                $PSBoundParameters['OutBuffer'] = 1
            }
            $wrappedCmd = $ExecutionContext.InvokeCommand.GetCommand(
                'Get-CherwellSearchResultsHC', [System.Management.Automation.CommandTypes]::Function
            )
            $PSBoundParameters['Name'] = 'CustomerInternal'
            $scriptCmd = { & $wrappedCmd @PSBoundParameters }
            $steppablePipeline = $scriptCmd.GetSteppablePipeline()
            $steppablePipeline.Begin($PSCmdlet)
        }
        catch { throw }
    }
    Process { try { $steppablePipeline.Process($_) } catch { throw } }
    End { try { $steppablePipeline.End() } catch { throw } }
}

Function Get-CherwellIncidentCategoryHC {
    <#
    .SYNOPSIS
        Retrieve Cherwell incident category

    .DESCRIPTION
        Retrieve all Cherwell incident categories. Cherwell incident categories 
        are used in a ticket under the name 'Category'.

    .PARAMETER Environment
        The Cherwell system that needs to be addressed.

        All available systems are defined in the file 'Passwords.json' in the 
        module folder. Ex. Test, Prod, ...

    .PARAMETER Filter
        A hash table containing the search conditions used by the API. When the 
        search conditions is met the configuration items that match these 
        conditions are returned. When no object is matching the condition, 
        nothing is returned.

        The following fields/keys are mandatory:
        @{
            FieldName  = ''
            Operator   = ''
            FieldValue = ''
        }

        The following operators are supported: 'eq', 'gt', 'lt', 'contains' and 
        'startswith'

        You can specify more than one filter. If you add multiple filters for
        the same FieldName, the result is an OR operation between those fields. 
        If the FieldNames are different, the result is an AND operation between 
        those fields.

    .PARAMETER Property
        Specifies the properties to select. When using '-Property *' all 
        properties are returned by the API. The downside is that this is slower 
        when working with large data sets.

    .PARAMETER PageSize
        Defines how many results are retrieved at once from the Cherwell 
        server. This can be useful for managing the speed of retrieving 
        results. A high PageSize number might take longer to execute but will 
        make less calls to the API. A small PageSize number might be faster in 
        some cases but result in more calls.

        Experimenting with this feature can be done by using the '-Verbose' 
        switch to analyze speed and the number of retrieved objects.

    .PARAMETER PassThru
        Return the output of API calls in their raw format. The output has not 
        been converted or changed in any way and can be used by other functions 
        as input through the 'InputObject' parameter. This will allow the 
        piping of objects from one function to the next.

        When 'PassThru' is omitted the returning business object property 
        'Fields', that comes from the API, will be converted to a 
        PSCustomObject.

    .EXAMPLE
        Retrieve all incident categories from the Cherwell 'Prod' environment 
        and display the progress by using the '-Verbose' switch

        Get-CherwellIncidentCategoryHC -Environment Prod -Verbose

    .EXAMPLE
        Retrieve all incident categories from the Cherwell 'Prod' environment 
        in batches of 100 at a time. Request all possible properties and 
        display the progress by using the '-Verbose' switch

        Get-CherwellIncidentCategoryHC -Environment Prod -PageSize 100 -Property * -Verbose

    .EXAMPLE
        Retrieve all incident categories that have the 'Country' set to 'BNL' 
        from the Cherwell 'Prod' environment and select only the properties 
        'IncidentCategory', 'Service' and 'ServiceStatus'

        Get-CherwellIncidentCategoryHC -Environment Prod -Filter @{
            FieldName  = 'Country'
            Operator   = 'eq'
            FieldValue = 'BNL'
        } -Property IncidentCategory, Service, ServiceStatus
#>

    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    Param (
        [Parameter(Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string[]]${Property},
        [Parameter(Position = 2)]
        [ValidateNotNullOrEmpty()]
        [hashTable[]]${Filter},
        [Parameter(Position = 3)]
        [ValidateRange(10, 5000)]
        [int]${PageSize},
        [switch]${PassThru}
    )
    DynamicParam {
        New-DynamicParameterHC -Name Environment
    }

    Begin {
        try {
            $outBuffer = $null
            if ($PSBoundParameters.TryGetValue('OutBuffer', [ref]$outBuffer)) {
                $PSBoundParameters['OutBuffer'] = 1
            }
            $wrappedCmd = $ExecutionContext.InvokeCommand.GetCommand(
                'Get-CherwellSearchResultsHC', [System.Management.Automation.CommandTypes]::Function
            )
            $PSBoundParameters['Name'] = 'IncidentCategory'
            $scriptCmd = { & $wrappedCmd @PSBoundParameters }
            $steppablePipeline = $scriptCmd.GetSteppablePipeline()
            $steppablePipeline.Begin($PSCmdlet)
        }
        catch { throw }
    }
    Process { try { $steppablePipeline.Process($_) } catch { throw } }
    End { try { $steppablePipeline.End() } catch { throw } }
}

Function Get-CherwellIncidentSubCategoryHC {
    <#
    .SYNOPSIS
        Retrieve Cherwell incident subcategory

    .DESCRIPTION
        Retrieve all Cherwell incident subcategories. Cherwell incident 
        subcategories are used in a ticket under the name 'SubCategory'.

    .PARAMETER Environment
        The Cherwell system that needs to be addressed.

        All available systems are defined in the file 'Passwords.json' in the 
        module folder. Ex. Test, Prod, ...

    .PARAMETER Filter
        A hash table containing the search conditions used by the API. When the 
        search conditions is met the configuration items that match these 
        conditions are returned. When no object is matching the condition, 
        nothing is returned.

        The following fields/keys are mandatory:
        @{
            FieldName  = ''
            Operator   = ''
            FieldValue = ''
        }

        The following operators are supported: 'eq', 'gt', 'lt', 'contains' and 
        'startswith'

        You can specify more than one filter. If you add multiple filters for 
        the same FieldName, the result is an OR operation between those fields. 
        If the FieldNames are different, the result is an AND operation between 
        those fields.

    .PARAMETER Property
        Specifies the properties to select. When using '-Property *' all 
        properties are returned by the API. The downside is that this is slower 
        when working with large data sets.

    .PARAMETER PageSize
        Defines how many results are retrieved at once from the Cherwell 
        server. This can be useful for managing the speed of retrieving 
        results. A high PageSize number might take longer to execute but will 
        make less calls to the API. A small PageSize number might be faster in 
        some cases but result in more calls.

        Experimenting with this feature can be done by using the '-Verbose' 
        switch to analyze speed and the number of retrieved objects.

    .PARAMETER PassThru
        Return the output of API calls in their raw format. The output has not 
        been converted or changed in any way and can be used by other functions 
        as input through the 'InputObject' parameter. This will allow the 
        piping of objects from one function to the next.

        When 'PassThru' is omitted the returning business object property 
        'Fields', that comes from the API, will be converted to a 
        PSCustomObject.

    .EXAMPLE
        Retrieve all incident subcategories from the Cherwell 'Prod' 
        environment and display the progress by using the '-Verbose' switch

        Get-CherwellIncidentSubCategoryHC -Environment Prod -Verbose

    .EXAMPLE
        Retrieve all incident subcategories from the Cherwell 'Prod' 
        environment in batches of 100 at a time. Request all possible 
        properties and display the progress by using the '-Verbose' switch.

        Get-CherwellIncidentSubCategoryHC -Environment Prod -PageSize 100 -Property * -Verbose

    .EXAMPLE
        Retrieve all incident subcategories that have the 'Country' set to 
        'BNL' from the Cherwell 'Prod' environment and select only the 
        properties 'IncidentCategory', 'Service' and 'ServiceStatus'

        Get-CherwellIncidentSubCategoryHC -Environment Prod -Filter @{
            FieldName  = 'Country'
            Operator   = 'eq'
            FieldValue = 'BNL'
        } -Property IncidentCategory, Service, ServiceStatus
#>

    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    Param (
        [Parameter(Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string[]]${Property},
        [Parameter(Position = 2)]
        [ValidateNotNullOrEmpty()]
        [hashTable[]]${Filter},
        [Parameter(Position = 3)]
        [ValidateRange(10, 5000)]
        [int]${PageSize},
        [switch]${PassThru}
    )
    DynamicParam {
        New-DynamicParameterHC -Name Environment
    }

    Begin {
        try {
            $outBuffer = $null
            if ($PSBoundParameters.TryGetValue('OutBuffer', [ref]$outBuffer)) {
                $PSBoundParameters['OutBuffer'] = 1
            }
            $wrappedCmd = $ExecutionContext.InvokeCommand.GetCommand(
                'Get-CherwellSearchResultsHC', [System.Management.Automation.CommandTypes]::Function
            )
            $PSBoundParameters['Name'] = 'IncidentSubCategory'
            $scriptCmd = { & $wrappedCmd @PSBoundParameters }
            $steppablePipeline = $scriptCmd.GetSteppablePipeline()
            $steppablePipeline.Begin($PSCmdlet)
        }
        catch { throw }
    }
    Process { try { $steppablePipeline.Process($_) } catch { throw } }
    End { try { $steppablePipeline.End() } catch { throw } }
}

Function Get-CherwellLocationHC {
    <#
    .SYNOPSIS
        Retrieve Cherwell location

    .DESCRIPTION
        Retrieve all locations in the Cherwell system. These are office 
        locations used in the customer or hardware asset location picker.

    .PARAMETER Environment
        The Cherwell system that needs to be addressed.

        All available systems are defined in the file 'Passwords.json' in the 
        module folder. Ex. Test, Prod, ...

    .PARAMETER Filter
        A hash table containing the search conditions used by the API. When the 
        search conditions is met the configuration items that match these 
        conditions are returned. When no object is matching the condition, 
        nothing is returned.

        The following fields/keys are mandatory:
        @{
            FieldName  = ''
            Operator   = ''
            FieldValue = ''
        }

        The following operators are supported: 'eq', 'gt', 'lt', 'contains' and 
        'startswith'

        You can specify more than one filter. If you add multiple filters for 
        the same FieldName, the result is an OR operation between those fields. 
        If the FieldNames are different, the result is an AND operation between 
        those fields.

    .PARAMETER Property
        Specifies the properties to select. When using '-Property *' all 
        properties are returned by the API. The downside is that this is slower 
        when working with large data sets.

    .PARAMETER PageSize
        Defines how many results are retrieved at once from the Cherwell 
        server. This can be useful for managing the speed of retrieving 
        results. A high PageSize number might take longer to execute but will 
        make less calls to the API. A small PageSize number might be faster in 
        some cases but result in more calls.

        Experimenting with this feature can be done by using the '-Verbose' 
        switch to analyze speed and the number of retrieved objects.

    .PARAMETER PassThru
        Return the output of API calls in their raw format. The output has not 
        been converted or changed in any way and can be used by other functions 
        as input through the 'InputObject' parameter. This will allow the 
        piping of objects from one function to the next.

        When 'PassThru' is omitted the returning business object property 
        'Fields', that comes from the API, will be converted to a 
        PSCustomObject.

    .EXAMPLE
        Retrieve all locations with the 'Country' property set to 'BEL' from 
        the 'Test' environment

        Get-CherwellLocationHC -Environment Test -Filter @{
            FieldName  = 'Country'
            Operator   = 'eq'
            FieldValue = 'BEL'
        }

    .EXAMPLE
        Retrieve all locations in Cherwell with all possible properties from 
        the 'Prod' environment and display the progress by using the '-Verbose' 
        switch
        
        Get-CherwellLocationHC -Environment Prod -Property * -Verbose
#>

    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    Param (
        [Parameter(Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string[]]${Property},
        [Parameter(Position = 2)]
        [ValidateNotNullOrEmpty()]
        [hashTable[]]${Filter},
        [Parameter(Position = 3)]
        [ValidateRange(10, 5000)]
        [int]${PageSize},
        [switch]${PassThru}
    )
    DynamicParam {
        New-DynamicParameterHC -Name Environment
    }

    Begin {
        try {
            $outBuffer = $null
            if ($PSBoundParameters.TryGetValue('OutBuffer', [ref]$outBuffer)) {
                $PSBoundParameters['OutBuffer'] = 1
            }
            $wrappedCmd = $ExecutionContext.InvokeCommand.GetCommand(
                'Get-CherwellSearchResultsHC', [System.Management.Automation.CommandTypes]::Function
            )
            $PSBoundParameters['Name'] = 'location'
            $scriptCmd = { & $wrappedCmd @PSBoundParameters }
            $steppablePipeline = $scriptCmd.GetSteppablePipeline()
            $steppablePipeline.Begin($PSCmdlet)
        }
        catch { throw }
    }
    Process { try { $steppablePipeline.Process($_) } catch { throw } }
    End { try { $steppablePipeline.End() } catch { throw } }
}

Function Get-CherwellQuickCallTemplateHC {
    <#
    .SYNOPSIS
        Retrieve Cherwell quick call template

    .DESCRIPTION
        Retrieve all quick call templates in the Cherwell system.

    .PARAMETER Environment
        The Cherwell system that needs to be addressed.

        All available systems are defined in the file 'Passwords.json' in the 
        module folder. Ex. Test, Prod, ...

    .PARAMETER Filter
        A hash table containing the search conditions used by the API. When the 
        search conditions is met the configuration items that match these 
        conditions are returned. When no object is matching the condition, 
        nothing is returned.

        The following fields/keys are mandatory:
        @{
            FieldName  = ''
            Operator   = ''
            FieldValue = ''
        }

        The following operators are supported: 'eq', 'gt', 'lt', 'contains' and 
        'startswith'. You can specify more than one filter. If you add multiple 
        filters for the same FieldName, the result is an OR operation between 
        those fields. If the FieldNames are different, the result is an AND 
        operation between those fields.

    .PARAMETER Property
        Specifies the properties to select. When using '-Property *' all 
        properties are returned by the API. The downside is that this is slower 
        when working with large data sets.

    .PARAMETER PageSize
        Defines how many results are retrieved at once from the Cherwell 
        server. This can be useful for managing the speed of retrieving 
        results. A high PageSize number might take longer to execute but will 
        make less calls to the API. A small PageSize number might be faster in 
        some cases but result in more calls.

        Experimenting with this feature can be done by using the '-Verbose' 
        switch to analyze speed and the number of retrieved objects.

    .PARAMETER PassThru
        Return the output of API calls in their raw format. The output has not 
        been converted or changed in any way and can be used by other functions 
        as input through the 'InputObject' parameter. This will allow the 
        piping of objects from one function to the next.

        When 'PassThru' is omitted the returning business object property 
        'Fields', that comes from the API, will be converted to a 
        PSCustomObject.

    .EXAMPLE
        Retrieve all quick call templates with the 'ServiceCountry' property 
        set to 'BNL' from the 'Test' environment

        Get-CherwellQuickCallTemplateHC -Environment Test -Filter @{
            FieldName  = 'ServiceCountry'
            Operator   = 'eq'
            FieldValue = 'BNL'
        }

    .EXAMPLE
        Retrieve all QuickCallTemplates in Cherwell with all possible 
        properties from the 'Prod' environment and display the progress by 
        using the '-Verbose' switch

        Get-CherwellQuickCallTemplateHC -Environment Prod -Property * -Verbose
 #>

    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    Param (
        [Parameter(Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string[]]${Property},
        [Parameter(Position = 2)]
        [ValidateNotNullOrEmpty()]
        [hashTable[]]${Filter},
        [Parameter(Position = 3)]
        [ValidateRange(10, 5000)]
        [int]${PageSize},
        [switch]${PassThru}
    )
    DynamicParam {
        New-DynamicParameterHC -Name Environment
    }

    Begin {
        try {
            $outBuffer = $null
            if ($PSBoundParameters.TryGetValue('OutBuffer', [ref]$outBuffer)) {
                $PSBoundParameters['OutBuffer'] = 1
            }
            $wrappedCmd = $ExecutionContext.InvokeCommand.GetCommand(
                'Get-CherwellSearchResultsHC', [System.Management.Automation.CommandTypes]::Function
            )
            $PSBoundParameters['Name'] = 'QuickCallTemplate'
            $scriptCmd = { & $wrappedCmd @PSBoundParameters }
            $steppablePipeline = $scriptCmd.GetSteppablePipeline()
            $steppablePipeline.Begin($PSCmdlet)
        }
        catch { throw }
    }
    Process { try { $steppablePipeline.Process($_) } catch { throw } }
    End { try { $steppablePipeline.End() } catch { throw } }
}

Function Get-CherwellServiceCatalogTemplateHC {
    <#
    .SYNOPSIS
        Retrieve Cherwell service catalog template

    .DESCRIPTION
        Retrieve all Cherwell incident categories. Cherwell incident categories are used in
        a ticket under the name 'Category'.

    .PARAMETER Environment
        The Cherwell system that needs to be addressed.

        All available systems are defined in the file 'Passwords.json' in the 
        module folder. Ex. Test, Prod, ...

    .PARAMETER Filter
        A hash table containing the search conditions used by the API. When the 
        search conditions is met the configuration items that match these 
        conditions are returned. When no object is matching the condition, 
        nothing is returned.

        The following fields/keys are mandatory:
        @{
            FieldName  = ''
            Operator   = ''
            FieldValue = ''
        }

        The following operators are supported: 'eq', 'gt', 'lt', 'contains' and 
        'startswith'. You can specify more than one filter. If you add multiple 
        filters for the same FieldName, the result is an OR operation between 
        those fields. If the FieldNames are different, the result is an AND 
        operation between those fields.

    .PARAMETER Property
        Specifies the properties to select. When using '-Property *' all 
        properties are returned by the API. The downside is that this is slower 
        when working with large data sets.

    .PARAMETER PageSize
        Defines how many results are retrieved at once from the Cherwell 
        server. This can be useful for managing the speed of retrieving 
        results. A high PageSize number might take longer to execute but will 
        make less calls to the API. A small PageSize number might be faster in 
        some cases but result in more calls.

        Experimenting with this feature can be done by using the '-Verbose' 
        switch to analyze speed and the number of retrieved objects.

    .PARAMETER PassThru
        Return the output of API calls in their raw format. The output has not 
        been converted or changed in any way and can be used by other functions 
        as input through the 'InputObject' parameter. This will allow the 
        piping of objects from one function to the next.

        When 'PassThru' is omitted the returning business object property 
        'Fields', that comes from the API, will be converted to a 
        PSCustomObject.

    .EXAMPLE
        Retrieve all service catalog templates from the Cherwell 'Prod' 
        environment and display the progress by using the '-Verbose' switch.

        Get-CherwellServiceCatalogTemplateHC -Environment Prod -Verbose

    .EXAMPLE
        Retrieve all service catalog templates from the 'Prod' environment in 
        batches of 100 at a time. Request all possible properties and display 
        the progress by using the '-Verbose' switch.

        Get-CherwellServiceCatalogTemplateHC -Environment Prod -PageSize 100 -Property * -Verbose

    .EXAMPLE
        Retrieve all service catalog templates that have the 'Country' set to 
        'BNL' from the Cherwell 'Prod' environment and select only specific 
        properties.

        Get-CherwellServiceCatalogTemplateHC -Environment Prod -Filter @{
            FieldName  = 'Country'
            Operator   = 'eq'
            FieldValue = 'BNL'
        } -Property Description, Service, ServiceStatus, SubCategory, Title

    .EXAMPLE
        Retrieve all service catalog templates with the field 'Title' set to 
        'BNL - CL - Leaver' from the 'Test' environment

        Get-CherwellServiceCatalogTemplateHC -Environment Test -Property * -Filter @{
            FieldName  = 'Title'
            Operator   = 'eq'
            FieldValue = 'BNL - CL - Leaver'
        }
 #>

    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    Param (
        [Parameter(Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string[]]${Property},
        [Parameter(Position = 2)]
        [ValidateNotNullOrEmpty()]
        [hashTable[]]${Filter},
        [Parameter(Position = 3)]
        [ValidateRange(10, 5000)]
        [int]${PageSize},
        [switch]${PassThru}
    )
    DynamicParam {
        New-DynamicParameterHC -Name Environment
    }

    Begin {
        try {
            $outBuffer = $null
            if ($PSBoundParameters.TryGetValue('OutBuffer', [ref]$outBuffer)) {
                $PSBoundParameters['OutBuffer'] = 1
            }
            $wrappedCmd = $ExecutionContext.InvokeCommand.GetCommand(
                'Get-CherwellSearchResultsHC', [System.Management.Automation.CommandTypes]::Function
            )
            $PSBoundParameters['Name'] = 'ServiceCatalogTemplate'
            $scriptCmd = { & $wrappedCmd @PSBoundParameters }
            $steppablePipeline = $scriptCmd.GetSteppablePipeline()
            $steppablePipeline.Begin($PSCmdlet)
        }
        catch { throw }
    }
    Process { try { $steppablePipeline.Process($_) } catch { throw } }
    End { try { $steppablePipeline.End() } catch { throw } }
}

Function Get-CherwellServiceHC {
    <#
    .SYNOPSIS
        Retrieve Cherwell service

    .DESCRIPTION
        Retrieve all Cherwell services. Cherwell services are used for 
        classifying a ticket based on its subject. A mobile phone request might 
        have service 'Mobile' for example.

    .PARAMETER Environment
        The Cherwell system that needs to be addressed.

        All available systems are defined in the file 'Passwords.json' in the 
        module folder. Ex. Test, Prod, ...

    .PARAMETER Filter
        A hash table containing the search conditions used by the API. When the 
        search conditions is met the configuration items that match these 
        conditions are returned. When no object is matching the condition, 
        nothing is returned.

        The following fields/keys are mandatory:
        @{
            FieldName  = ''
            Operator   = ''
            FieldValue = ''
        }

        The following operators are supported: 'eq', 'gt', 'lt', 'contains' and 
        'startswith'. You can specify more than one filter. If you add multiple 
        filters for the same FieldName, the result is an OR operation between 
        those fields. If the FieldNames are different, the result is an AND 
        operation between those fields.

    .PARAMETER Property
        Specifies the properties to select. When using '-Property *' all 
        properties are returned by the API. The downside is that this is slower 
        when working with large data sets.

    .PARAMETER PageSize
        Defines how many results are retrieved at once from the Cherwell 
        server. This can be useful for managing the speed of retrieving 
        results. A high PageSize number might take longer to execute but will 
        make less calls to the API. A small PageSize number might be faster in 
        some cases but result in more calls.

        Experimenting with this feature can be done by using the '-Verbose' 
        switch to analyze speed and the number of retrieved objects.

    .PARAMETER PassThru
        Return the output of API calls in their raw format. The output has not 
        been converted or changed in any way and can be used by other functions 
        as input through the 'InputObject' parameter. This will allow the 
        piping of objects from one function to the next.

        When 'PassThru' is omitted the returning business object property 
        'Fields', that comes from the API, will be converted to a 
        PSCustomObject.

    .EXAMPLE
        Retrieve all services from the Cherwell 'Test' environment and display 
        the progress by using the '-Verbose' switch

        Get-CherwellServiceHC -Environment Test -Verbose

    .EXAMPLE
        Retrieve all services from the Cherwell 'Prod' environment with the 
        field 'Country' set to 'BNL'. Display the progress by using the 
        '-Verbose' switch. The select only the fields 'ServiceName' and 
        'ServiceDescription'

        Get-CherwellServiceHC -Environment Prod -Filter @{
            FieldName  = 'Country'
            Operator   = 'eq'
            FieldValue = 'BNL'
        } -Verbose  | Select-Object ServiceName, ServiceDescription
#>

    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    Param (
        [Parameter(Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string[]]${Property},
        [Parameter(Position = 2)]
        [ValidateNotNullOrEmpty()]
        [hashTable[]]${Filter},
        [Parameter(Position = 3)]
        [ValidateRange(10, 5000)]
        [int]${PageSize},
        [switch]${PassThru}
    )
    DynamicParam {
        New-DynamicParameterHC -Name Environment
    }

    Begin {
        try {
            $outBuffer = $null
            if ($PSBoundParameters.TryGetValue('OutBuffer', [ref]$outBuffer)) {
                $PSBoundParameters['OutBuffer'] = 1
            }
            $wrappedCmd = $ExecutionContext.InvokeCommand.GetCommand(
                'Get-CherwellSearchResultsHC', [System.Management.Automation.CommandTypes]::Function
            )
            $PSBoundParameters['Name'] = 'Service'
            $scriptCmd = { & $wrappedCmd @PSBoundParameters }
            $steppablePipeline = $scriptCmd.GetSteppablePipeline()
            $steppablePipeline.Begin($PSCmdlet)
        }
        catch { throw }
    }
    Process { try { $steppablePipeline.Process($_) } catch { throw } }
    End { try { $steppablePipeline.End() } catch { throw } }
}

Function Get-CherwellSlaHC {
    <#
    .SYNOPSIS
        Retrieve Cherwell SLA

    .DESCRIPTION
        Retrieve all Cherwell SLA's. Cherwell service level agreements are used 
        to indicate the time window in which a problem/issue needs to be 
        resolved.

    .PARAMETER Environment
        The Cherwell system that needs to be addressed.

        All available systems are defined in the file 'Passwords.json' in the 
        module folder. Ex. Test, Prod, ...

    .PARAMETER Filter
        A hash table containing the search conditions used by the API. When the 
        search conditions is met the configuration items that match these 
        conditions are returned. When no object is matching the condition, 
        nothing is returned.

        The following fields/keys are mandatory:
        @{
            FieldName  = ''
            Operator   = ''
            FieldValue = ''
        }

        The following operators are supported: 'eq', 'gt', 'lt', 'contains' and 
        'startswith'. You can specify more than one filter. If you add multiple 
        filters for the same FieldName, the result is an OR operation between 
        those fields. If the FieldNames are different, the result is an AND 
        operation between those fields.

    .PARAMETER Property
        Specifies the properties to select. When using '-Property *' all 
        properties are returned by the API. The downside is that this is slower 
        when working with large data sets.

    .PARAMETER PageSize
        Defines how many results are retrieved at once from the Cherwell 
        server. This can be useful for managing the speed of retrieving 
        results. A high PageSize number might take longer to execute but will 
        make less calls to the API. A small PageSize number might be faster in 
        some cases but result in more calls.

        Experimenting with this feature can be done by using the '-Verbose' 
        switch to analyze speed and the number of retrieved objects.

    .PARAMETER PassThru
        Return the output of API calls in their raw format. The output has not 
        been converted or changed in any way and can be used by other functions 
        as input through the 'InputObject' parameter. This will allow the 
        piping of objects from one function to the next.

        When 'PassThru' is omitted the returning business object property 
        'Fields', that comes from the API, will be converted to a 
        PSCustomObject.

    .EXAMPLE
        Retrieve all SLA's from the Cherwell 'Prod' environment in batches of 
        100 at a time. Request all possible fields and display the progress by 
        using the '-Verbose' switch

        Get-CherwellSLAHC -Environment Prod -PageSize 100 -Property * -Verbose

    .EXAMPLE
        Retrieve all SLA's that start with the value 'DEU' in the property 
        'Title' from the 'Prod' environment

        Get-CherwellSLAHC -Environment Prod -Filter @{
            FieldName  = 'Title'
            Operator   = 'startswith'
            FieldValue = 'DEU'
        }

    .EXAMPLE
        Retrieve all SLA's that have been created after '4/18/2019' from the 
        'Prod' environment

        $RecentlyCreatedSla = Get-CherwellSLAHC -Environment Prod -Filter @{
            FieldName  = 'CreatedDateTime'
            Operator   = 'gt'
            FieldValue = '4/18/2019'
        } -Property *

        $RecentlyCreatedSla | Select-Object CreatedDateTime, Title
 #>

    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    Param (
        [Parameter(Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string[]]${Property},
        [Parameter(Position = 2)]
        [ValidateNotNullOrEmpty()]
        [hashTable[]]${Filter},
        [Parameter(Position = 3)]
        [ValidateRange(10, 5000)]
        [int]${PageSize},
        [switch]${PassThru}
    )
    DynamicParam {
        New-DynamicParameterHC -Name Environment
    }

    Begin {
        try {
            $outBuffer = $null
            if ($PSBoundParameters.TryGetValue('OutBuffer', [ref]$outBuffer)) {
                $PSBoundParameters['OutBuffer'] = 1
            }
            $wrappedCmd = $ExecutionContext.InvokeCommand.GetCommand(
                'Get-CherwellSearchResultsHC', [System.Management.Automation.CommandTypes]::Function
            )
            $PSBoundParameters['Name'] = 'SLA'
            $scriptCmd = { & $wrappedCmd @PSBoundParameters }
            $steppablePipeline = $scriptCmd.GetSteppablePipeline()
            $steppablePipeline.Begin($PSCmdlet)
        }
        catch { throw }
    }
    Process { try { $steppablePipeline.Process($_) } catch { throw } }
    End { try { $steppablePipeline.End() } catch { throw } }
}

Function Get-CherwellSupplierHC {
    <#
    .SYNOPSIS
        Retrieve Cherwell supplier

    .DESCRIPTION
        Retrieve all Cherwell suppliers. Cherwell suppliers are used to 
        indicate where we bought hardware, ... .

    .PARAMETER Environment
        The Cherwell system that needs to be addressed.

        All available systems are defined in the file 'Passwords.json' in the 
        module folder. Ex. Test, Prod, ...

    .PARAMETER Filter
        A hash table containing the search conditions used by the API. When the 
        search conditions is met the configuration items that match these 
        conditions are returned. When no object is matching the condition, 
        nothing is returned.

        The following fields/keys are mandatory:
        @{
            FieldName  = ''
            Operator   = ''
            FieldValue = ''
        }

        The following operators are supported: 'eq', 'gt', 'lt', 'contains' and 
        'startswith'. You can specify more than one filter. If you add multiple 
        filters for the same FieldName, the result is an OR operation between 
        those fields. If the FieldNames are different, the result is an AND 
        operation between those fields.

    .PARAMETER Property
        Specifies the properties to select. When using '-Property *' all 
        properties are returned by the API. The downside is that this is slower 
        when working with large data sets.

    .PARAMETER PageSize
        Defines how many results are retrieved at once from the Cherwell 
        server. This can be useful for managing the speed of retrieving 
        results. A high PageSize number might take longer to execute but will 
        make less calls to the API. A small PageSize number might be faster in 
        some cases but result in more calls.

        Experimenting with this feature can be done by using the '-Verbose' 
        switch to analyze speed and the number of retrieved objects.

    .PARAMETER PassThru
        Return the output of API calls in their raw format. The output has not 
        been converted or changed in any way and can be used by other functions 
        as input through the 'InputObject' parameter. This will allow the 
        piping of objects from one function to the next.

        When 'PassThru' is omitted the returning business object property 
        'Fields', that comes from the API, will be converted to a 
        PSCustomObject.

    .EXAMPLE
        Retrieve all suppliers from the Cherwell 'Prod' environment in batches 
        of 100 at a time. Request all possible fields and display the progress 
        by using the '-Verbose' switch

        Get-CherwellSupplierHC -Environment Prod -PageSize 100 -Property * -Verbose

    .EXAMPLE
        Retrieve all suppliers that start with the value 'DE' in the property 
        'SupplierName' from the 'Prod' environment

        Get-CherwellSupplierHC -Environment Prod -Filter @{
            FieldName  = 'SupplierName'
            Operator   = 'startswith'
            FieldValue = 'DE'
        }
#>

    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    Param (
        [Parameter(Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string[]]${Property},
        [Parameter(Position = 2)]
        [ValidateNotNullOrEmpty()]
        [hashTable[]]${Filter},
        [Parameter(Position = 3)]
        [ValidateRange(10, 5000)]
        [int]${PageSize},
        [switch]${PassThru}
    )
    DynamicParam {
        New-DynamicParameterHC -Name Environment
    }

    Begin {
        try {
            $outBuffer = $null
            if ($PSBoundParameters.TryGetValue('OutBuffer', [ref]$outBuffer)) {
                $PSBoundParameters['OutBuffer'] = 1
            }
            $wrappedCmd = $ExecutionContext.InvokeCommand.GetCommand(
                'Get-CherwellSearchResultsHC', [System.Management.Automation.CommandTypes]::Function
            )
            $PSBoundParameters['Name'] = 'Supplier'
            $scriptCmd = { & $wrappedCmd @PSBoundParameters }
            $steppablePipeline = $scriptCmd.GetSteppablePipeline()
            $steppablePipeline.Begin($PSCmdlet)
        }
        catch { throw }
    }
    Process { try { $steppablePipeline.Process($_) } catch { throw } }
    End { try { $steppablePipeline.End() } catch { throw } }
}

Function Get-CherwellSystemUserHC {
    <#
    .SYNOPSIS
        Retrieve Cherwell system user

    .DESCRIPTION
        Retrieve Cherwell system user. The Cherwell system users are not by 
        default customers (Get-CherwellCustomerHC ).

    .PARAMETER Environment
        The Cherwell system that needs to be addressed.

        All available systems are defined in the file 'Passwords.json' in the 
        module folder. Ex. Test, Prod, ...

    .PARAMETER Filter
        A hash table containing the search conditions used by the API. When the 
        search conditions is met the configuration items that match these 
        conditions are returned. When no object is matching the condition, 
        nothing is returned.

        The following fields/keys are mandatory:
        @{
            FieldName  = ''
            Operator   = ''
            FieldValue = ''
        }

        The following operators are supported: 'eq', 'gt', 'lt', 'contains' and 
        'startswith'. You can specify more than one filter. If you add multiple 
        filters for the same FieldName, the result is an OR operation between 
        those fields. If the FieldNames are different, the result is an AND 
        operation between those fields.

    .PARAMETER Property
        Specifies the properties to select. When using '-Property *' all 
        properties are returned by the API. The downside is that this is slower 
        when working with large data sets.

    .PARAMETER PageSize
        Defines how many results are retrieved at once from the Cherwell 
        server. This can be useful for managing the speed of retrieving 
        results. A high PageSize number might take longer to execute but will 
        make less calls to the API. A small PageSize number might be faster in 
        some cases but result in more calls.

        Experimenting with this feature can be done by using the '-Verbose' 
        switch to analyze speed and the number of retrieved objects.

    .PARAMETER PassThru
        Return the output of API calls in their raw format. The output has not 
        been converted or changed in any way and can be used by other functions 
        as input through the 'InputObject' parameter. This will allow the 
        piping of objects from one function to the next.

        When 'PassThru' is omitted the returning business object property 
        'Fields', that comes from the API, will be converted to a 
        PSCustomObject.

    .EXAMPLE
        Retrieve all system users from the Cherwell production environment and 
        display the progress by using the '-Verbose' switch. The '-verbose' 
        switch can be useful because it can take a long time before all users 
        are retrieved.

        To further improve the speed of the query you can limit the properties 
        to return by using the parameter '-Property' or set the '-PageSize' 
        parameter to only retrieve the data in chunks.

        Get-CherwellSystemUserHC -Environment Prod -Verbose
    .EXAMPLE
        Retrieve all system users from the 'Test' environment with the property 
        'DefaultTeamName' set to 'BNL'

        Get-CherwellSystemUserHC -Environment Test -Filter @{
            FieldName  = 'DefaultTeamName'
            Operator   = 'eq'
            FieldValue = 'BNL'
        }

    .EXAMPLE
        Retrieve all system users from the Cherwell production environment in 
        batches of 100 at a time. Request all possible properties and display 
        the progress by using the '-Verbose' switch

        Get-CherwellSystemUserHC -Environment Prod -PageSize 100 -Property * -Verbose

    .EXAMPLE
        Retrieve all system users from the Cherwell test environment that have 
        the property 'LastName' set to 'strelec'

        Get-CherwellSystemUserHC -Environment Test -Filter @{
            FieldName  = 'LastName'
            Operator   = 'eq'
            FieldValue = 'strelec'
        }

    .EXAMPLE
        Get-CherwellSystemUserHC -Environment Test -Property * -Filter @{
            FieldName  = 'SamAccountName'
            Operator   = 'eq'
            FieldValue = 'thardey'
        }
 #>

    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    Param (
        [Parameter(Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string[]]${Property},
        [Parameter(Position = 2)]
        [ValidateNotNullOrEmpty()]
        [hashTable[]]${Filter},
        [Parameter(Position = 3)]
        [ValidateRange(10, 5000)]
        [int]${PageSize},
        [switch]${PassThru}
    )
    DynamicParam {
        New-DynamicParameterHC -Name Environment
    }

    Begin {
        try {
            $outBuffer = $null
            if ($PSBoundParameters.TryGetValue('OutBuffer', [ref]$outBuffer)) {
                $PSBoundParameters['OutBuffer'] = 1
            }
            $wrappedCmd = $ExecutionContext.InvokeCommand.GetCommand(
                'Get-CherwellSearchResultsHC', [System.Management.Automation.CommandTypes]::Function
            )
            $PSBoundParameters['Name'] = 'UserInfo'
            $scriptCmd = { & $wrappedCmd @PSBoundParameters }
            $steppablePipeline = $scriptCmd.GetSteppablePipeline()
            $steppablePipeline.Begin($PSCmdlet)
        }
        catch { throw }
    }
    Process { try { $steppablePipeline.Process($_) } catch { throw } }
    End { try { $steppablePipeline.End() } catch { throw } }
}

Function Get-CherwellTeamInfoHC {
    <#
    .SYNOPSIS
        Retrieve Cherwell team information

    .DESCRIPTION
        Retrieve all Cherwell teams and the team details.

    .PARAMETER Environment
        The Cherwell system that needs to be addressed.

        All available systems are defined in the file 'Passwords.json' in the 
        module folder. Ex. Test, Prod, ...

    .PARAMETER Filter
        A hash table containing the search conditions used by the API. When the 
        search conditions is met the configuration items that match these 
        conditions are returned. When no object is matching the condition, 
        nothing is returned.

        The following fields/keys are mandatory:
        @{
            FieldName  = ''
            Operator   = ''
            FieldValue = ''
        }

        The following operators are supported: 'eq', 'gt', 'lt', 'contains' and 
        'startswith'. You can specify more than one filter. If you add multiple 
        filters for the same FieldName, the result is an OR operation between 
        those fields. If the FieldNames are different, the result is an AND 
        operation between those fields.

    .PARAMETER Property
        Specifies the properties to select. When using '-Property *' all 
        properties are returned by the API. The downside is that this is slower 
        when working with large data sets.

    .PARAMETER PageSize
        Defines how many results are retrieved at once from the Cherwell 
        server. This can be useful for managing the speed of retrieving 
        results. A high PageSize number might take longer to execute but will 
        make less calls to the API. A small PageSize number might be faster in 
        some cases but result in more calls.

        Experimenting with this feature can be done by using the '-Verbose' 
        switch to analyze speed and the number of retrieved objects.

    .PARAMETER PassThru
        Return the output of API calls in their raw format. The output has not 
        been converted or changed in any way and can be used by other functions 
        as input through the 'InputObject' parameter. This will allow the 
        piping of objects from one function to the next.

        When 'PassThru' is omitted the returning business object property 
        'Fields', that comes from the API, will be converted to a 
        PSCustomObject.

    .EXAMPLE
        Retrieve all teams from the Cherwell 'Prod' environment in batches of 
        100 at a time. Request all possible fields and display the progress by 
        using the '-Verbose' switch

        Get-CherwellTeamInfoHC -Environment Prod -PageSize 100 -Property * -Verbose

    .EXAMPLE
        Retrieve all teams that start with the value 'BNL' in the property 
        'Name' from the 'Prod' environment

        Get-CherwellTeamInfoHC -Environment Prod -Filter @{
            FieldName  = 'Name'
            Operator   = 'startswith'
            FieldValue = 'BNL'
        }
 #>

    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    Param (
        [Parameter(Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string[]]${Property},
        [Parameter(Position = 2)]
        [ValidateNotNullOrEmpty()]
        [hashTable[]]${Filter},
        [Parameter(Position = 3)]
        [ValidateRange(10, 5000)]
        [int]${PageSize},
        [switch]${PassThru}
    )
    DynamicParam {
        New-DynamicParameterHC -Name Environment
    }

    Begin {
        try {
            $outBuffer = $null
            if ($PSBoundParameters.TryGetValue('OutBuffer', [ref]$outBuffer)) {
                $PSBoundParameters['OutBuffer'] = 1
            }
            $wrappedCmd = $ExecutionContext.InvokeCommand.GetCommand(
                'Get-CherwellSearchResultsHC', [System.Management.Automation.CommandTypes]::Function
            )
            $PSBoundParameters['Name'] = 'TeamInfo'
            $scriptCmd = { & $wrappedCmd @PSBoundParameters }
            $steppablePipeline = $scriptCmd.GetSteppablePipeline()
            $steppablePipeline.Begin($PSCmdlet)
        }
        catch { throw }
    }
    Process { try { $steppablePipeline.Process($_) } catch { throw } }
    End { try { $steppablePipeline.End() } catch { throw } }
}
#endregion

Function Get-CherwellTicketHC {
    <#
    .SYNOPSIS
        Retrieve a ticket

    .DESCRIPTION
        Retrieve a ticket from Cherwell

    .PARAMETER TicketNr
        The number to identify the ticket in Cherwell, also knows as the ticket 
        number.

    .PARAMETER Environment
        The Cherwell system that needs to be addressed.

        All available systems are defined in the file 'Passwords.json' in the 
        module folder. Ex. Test, Prod, ...

    .PARAMETER Filter
        A hash table containing the search conditions used by the API. When the 
        search conditions is met the configuration items that match these 
        conditions are returned. When no object is matching the condition, 
        nothing is returned.

        The following fields/keys are mandatory:
        @{
            FieldName  = ''
            Operator   = ''
            FieldValue = ''
        }

        The following operators are supported: 'eq', 'gt', 'lt', 'contains' and 
        'startswith'. You can specify more than one filter. If you add multiple 
        filters for the same FieldName, the result is an OR operation between 
        those fields. If the FieldNames are different, the result is an AND 
        operation between those fields.

    .PARAMETER Property
        Specifies the properties to select. When using '-Property *' all 
        properties are returned by the API. The downside is that this is slower 
        when working with large data sets.

    .PARAMETER PageSize
        Defines how many results are retrieved at once from the Cherwell 
        server. This can be useful for managing the speed of retrieving 
        results. A high PageSize number might take longer to execute but will 
        make less calls to the API. A small PageSize number might be faster in 
        some cases but result in more calls.

        Experimenting with this feature can be done by using the '-Verbose' 
        switch to analyze speed and the number of retrieved objects.

    .PARAMETER PassThru
        Return the output of API calls in their raw format. The output has not 
        been converted or changed in any way and can be used by other functions 
        as input through the 'InputObject' parameter. This will allow the 
        piping of objects from one function to the next.

        When 'PassThru' is omitted the returning business object property 
        'Fields', that comes from the API, will be converted to a 
        PSCustomObject.

    .EXAMPLE
        Retrieve ticket '40735' from the 'Test' environment and return the 
        values for specific properties. This is way faster than using 
        '-Property *' to retrieve all properties

        Get-CherwellTicketHC -TicketNr 40735 -Environment Test -Property 
        IncidentID, CreatedDateTime, CreatedBy, CreatedDuring, 
        LastModifiedDateTime, LastModBy, CustomerDisplayName, CustomerPhone,
        RequesterDepartment, Location, BusinessLine, ConfigItemDisplayName, 
        OwnedByTeam, SLAName, SLAResolveByDeadline, Status, RecId

    .EXAMPLE
        Retrieve all tickets from the Cherwell 'Prod' environment in batches of 
        1000 at a time. Display the progress by using the '-Verbose' switch.

        Get-CherwellTicketHC -Environment Prod -PageSize 1000 -Verbose

    .EXAMPLE
        Retrieve the ticket with IncidentID number '2290' from the 'Test' 
        environment with all its possible properties

        Get-CherwellTicketHC -Environment Test -TicketNr 2290 -Property *

    .EXAMPLE
        Retrieve all tickets that are created after October 1 2019 from the 
        'Test' environment

        Get-CherwellTicketHC -Environment Test -Filter @{
            FieldName  = 'CreatedDateTime'
            Operator   = 'gt'
            FieldValue = '10/1/2019'
        }

    .EXAMPLE
        Retrieve all tickets that are assigned to the team 'BNL' from the 
        'Test' environment where the status is not 'Closed'

        Get-CherwellTicketHC -Environment Test -Filter @{
            FieldName  = 'OwnedByTeam'
            Operator   = 'eq'
            FieldValue = 'BNL'
        } | Where-Object {$_.Status -ne 'Closed'}

    .EXAMPLE
        Retrieve all tickets with 'Requestor' or 'SubmittedBySamAccountName' 
        set to 'thardey' from the 'Test' environment. Use '-PageSize 500' to 
        retrieve the results in batches of 500 at a time

        $Customer = Get-CherwellCustomerHC -Environment Test -Filter @{
            FieldName  = 'SamAccountName'
            Operator   = 'eq'
            FieldValue = 'thardey'
        } -PassThru

        Get-CherwellTicketHC -Environment Test -Filter @(
            @{
                FieldName  = 'CustomerRecID'
                Operator   = 'eq'
                FieldValue = $Customer.busObRecId
            }
            @{
                FieldName  = 'SubmitOnBehalfOfID'
                Operator   = 'eq'
                FieldValue = $Customer.busObRecId
            }
        ) -PageSize 500 -Property IncidentID, CreatedDateTime, CreatedBy,
        CreatedDuring, LastModifiedDateTime, LastModBy, CustomerDisplayName, 
        CustomerRecID, SubmitOnBehalfOfID,  CustomerPhone, RequesterDepartment, 
        Location, BusinessLine, ConfigItemDisplayName, OwnedByTeam, SLAName, 
        SLAResolveByDeadline, Status, RecId  -Verbose

    .EXAMPLE
        Retrieve all tickets with 'CI' set to 'BELCL003000' from the 'Test' 
        environment

        $ConfigItem = Get-CherwellConfigItemHC -Environment Test -Type 'ConfigComputer' -Filter @{
            FieldName  = 'HostName'
            Operator   = 'eq'
            FieldValue = 'BELCL003000'
        } -PassThru

        Get-CherwellTicketHC -Environment Test -Filter @{
            FieldName  = 'ConfigItemRecID'
            Operator   = 'eq'
            FieldValue = $ConfigItem.busObRecId
        } -Property IncidentID, CreatedDateTime, ConfigItemDisplayName, Status, RecId
#>

    [CmdletBinding(DefaultParameterSetName = 'Default')]
    [OutputType([PSCustomObject[]])]
    Param (
        [Parameter(Mandatory, ParameterSetName = 'TicketNr')]
        [Parameter(Mandatory, ParameterSetName = 'TicketNrField')]
        [ValidateNotNullOrEmpty()]
        [Int[]]$TicketNr,
        [Parameter(Mandatory, ParameterSetName = 'Filter')]
        [Parameter(Mandatory, ParameterSetName = 'FilterField')]
        [ValidateNotNullOrEmpty()]
        [HashTable[]]$Filter,
        [Parameter(Mandatory, ParameterSetName = 'DefaultField')]
        [Parameter(Mandatory, ParameterSetName = 'TicketNrField')]
        [Parameter(Mandatory, ParameterSetName = 'FilterField')]
        [ValidateNotNullOrEmpty()]
        [String[]]$Property,
        [ValidateRange(10, 5000)]
        [Int]$PageSize = 1000,
        [Switch]$PassThru
    )
    DynamicParam {
        New-DynamicParameterHC -Name Environment
    }

    Begin {
        Try {
            $Environment = $PSBoundParameters.Environment
            Select-EnvironmentHC -Name $Environment
        }
        Catch {
            throw "Failed retrieving ticket: $_"
        }
    }

    Process {
        Try {
            Write-Verbose "Get ticket from environment '$Environment'"

            if ($TicketNr) {
                $Filter += $TicketNr.Foreach( {
                        @{
                            FieldName  = 'IncidentID'
                            Operator   = 'eq'
                            FieldValue = $_
                        }
                    })
            }

            if ($Filter) {
                if (-not $PSBoundParameters.ContainsKey('Filter')) {
                    $PSBoundParameters.Add('Filter', $null )
                }

                $PSBoundParameters.Filter = $Filter
            }

            if ($Filter -or $Property) {
                $null = $PSBoundParameters.Add('Schema', (Get-SchemaHC -Name 'Incident') )
            }

            $null = $PSBoundParameters.Remove('Environment')
            $null = $PSBoundParameters.Remove('TicketNr')
            $PSBoundParameters.BusObId = (Get-SummaryHC -Name 'Incident').busobid

            Invoke-GetSearchResultsHC @PSBoundParameters
        }
        Catch {
            throw "Failed retrieving ticket from Cherwell '$Environment': $_"
        }
    }
}

Function Get-CherwellTicketDetailHC {
    <#
    .SYNOPSIS
        Get ticket details from related objects.

    .DESCRIPTION
        Get details from objects that are related to the ticket.

    .PARAMETER Ticket
        Objects coming from 'Get-CherwellTicketHC -PassThru' or a ticket number.
        Multiple tickets are supported.

    .PARAMETER Environment
        The Cherwell system that needs to be addressed.

        All available systems are defined in the file 'Passwords.json' in the 
        module folder. Ex. Test, Prod, ...

    .PARAMETER PassThru
        Return the output of API calls in their raw format. The output has not 
        been converted or changed in any way and can be used by other functions 
        as input through the 'InputObject' parameter. This will allow the 
        piping of objects from one function to the next.

        When 'PassThru' is omitted the returning business object property 
        'Fields', that comes from the API, will be converted to a 
        PSCustomObject.

    .EXAMPLE
        Retrieve the text in the property 'Additional notes'

        $Params = @{
            Environment = 'Test'
            Ticket      = 38076
            Type        = 'Specifics'
        }
        $AdditionalDetails = (Get-CherwellTicketDetailHC @Params).Notes

    .EXAMPLE
        Retrieve all 'Tasks' from ticket 38568 and select only the important 
        stuff.

        $Params = @{
            Environment  = 'Test'
            Ticket       = 38568
            Type         = 'Tasks'
        }
        Get-CherwellTicketDetailHC @Params |
        Select-Object Title, CreatedBy, CreatedDateTime, DueDate, TaskTypeName,
        Status, StatusDescription, Description, ActionNotes

    .EXAMPLE
        Retrieve all JournalNotes from ticket '55555' in the 'Test' environment

        $Params = @{
            Environment = 'Test'
            Ticket      = 55555
            Type        = 'JournalNote'
        }
        $Notes = Get-CherwellTicketDetailHC @Params
        $Notes | Select-Object CreatedBy, QuickJournal, Details | Format-List

    .EXAMPLE
        Retrieve the Journals for ticket '37941' in the 'Test' environment

        $Params = @{
            Environment = 'Test'
            Ticket      = 37941
            Type        = 'Journals'
        }
        $Journals = Get-CherwellTicketDetailHC @Params
        $Journals | Select-Object CreatedDateTime, Details | Format-List
#>

    [OutputType([PSCustomObject[]])]
    [OutputType()]
    Param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [Alias('TicketNr')]
        [Object[]]$Ticket,
        [Switch]$PassThru
    )
    DynamicParam {
        New-DynamicParameterHC -Name Environment, Type
    }

    Begin {
        Try {
            $Environment = $PSBoundParameters.Environment
            $Type = $PSBoundParameters.Type

            Select-EnvironmentHC -Name $Environment

            $IncidentSchema = Get-SchemaHC -Name 'Incident'

            #region Get RelationshipId
            $RelationshipId = $IncidentSchema.relationships.Where( {
                    $_.DisplayName -eq $RelationshipEnum[$Type] }) |
            Select-Object -ExpandProperty 'RelationshipId'

            if (-not $RelationshipId) {
                throw "RelationshipId not found for type '$Type'."
            }
            #endregion
        }
        Catch {
            throw "Failed retrieving ticket details fpr '$Type' of ticket '$Ticket': $_ "
        }
    }

    Process {
        Try {
            foreach ($T in $Ticket) {
                $TicketObject = ConvertTo-TicketObjectHC -Item $T

                #region Get ticket related object
                Write-Verbose "Get details for ticket '$($TicketObject.busObPublicId)'"

                $Params = @{
                    parentBusObId    = $IncidentSchema.busObId
                    parentBusObRecId = $TicketObject.busObRecId
                    relationshipId   = $RelationshipId
                }
                $Result = Invoke-GetRelatedBusinessObjectHC @Params
                #endregion

                if ($PassThru) {
                    $Result
                }
                else {
                    ConvertTo-PSCustomObjectHC -BusinessObject $Result.relatedBusinessObjects
                }
            }
        }
        Catch {
            throw "Failed retrieving ticket details fpr '$Type' of ticket '$Ticket' from Cherwell '$Environment': $_ "
        }
    }
}

Function Get-CherwellTicketJournalNoteHC {
    <#
    .SYNOPSIS
        Get the journal notes from a Cherwell ticket.

    .DESCRIPTION
        Retrieve all journal notes that are connected to a specific Cherwell 
        ticket.

    .PARAMETER Environment
        The Cherwell system that needs to be addressed.

        All available systems are defined in the file 'Passwords.json' in the 
        module folder. Ex. Test, Prod, ...

    .PARAMETER Ticket
        Objects coming from 'Get-CherwellTicketHC -PassThru' or a ticket number.
        Multiple tickets are supported.

    .PARAMETER PassThru
        Return the output of API calls in their raw format. The output has not 
        been converted or changed in any way and can be used by other functions 
        as input through the 'InputObject' parameter. This will allow the 
        piping of objects from one function to the next.

        When 'PassThru' is omitted the returning business object property 
        'Fields', that comes from the API, will be converted to a 
        PSCustomObject.

    .EXAMPLE
        Get the journal notes from ticket 123123 in the Cherwell test 
        environment. With the switch 'PassThru' the output of the API is 
        returned, which can be used by other functions.

        Get-CherwellTicketJournalNoteHC -Environment Test -Ticket 123123 -PassThru

    .EXAMPLE
        Get the journal notes from ticket 42173. The 'Verbose' switch provides progress indications to see what's happening in the background

        Get-CherwellTicketJournalNoteHC Environment 'Test' -Ticket 42173 
 #>

    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
        [Alias('TicketNr')]
        [System.Object[]]${Ticket},
        [Switch]${PassThru}
    )
    DynamicParam {
        New-DynamicParameterHC -Name Environment
    }

    begin {
        try {
            $outBuffer = $null
            if ($PSBoundParameters.TryGetValue('OutBuffer', [ref]$outBuffer)) {
                $PSBoundParameters['OutBuffer'] = 1
            }
            $wrappedCmd = $ExecutionContext.InvokeCommand.GetCommand(
                'Get-CherwellTicketDetailHC', [System.Management.Automation.CommandTypes]::Function
            )

            $PSBoundParameters['Type'] = 'JournalNote'

            $scriptCmd = { & $wrappedCmd @PSBoundParameters }
            $steppablePipeline = $scriptCmd.GetSteppablePipeline()
            $steppablePipeline.Begin($PSCmdlet)
        }
        catch {
            throw
        }
    }

    process {
        try {
            $steppablePipeline.Process($_)
        }
        catch {
            throw
        }
    }

    end {
        try {
            $steppablePipeline.End()
        }
        catch {
            throw
        }
    }
}

Function Get-CherwellTicketTaskHC {
    <#
    .SYNOPSIS
        Get tasks from a Cherwell ticket.

    .DESCRIPTION
        Retrieve all tasks that are connected to a specific Cherwell ticket.

    .PARAMETER Environment
        The Cherwell system that needs to be addressed.

        All available systems are defined in the file 'Passwords.json' in the 
        module folder. Ex. Test, Prod, ...

    .PARAMETER Ticket
        Objects coming from 'Get-CherwellTicketHC -PassThru' or a ticket number.
        Multiple tickets are supported.

    .PARAMETER PassThru
        Return the output of API calls in their raw format. The output has not 
        been converted or changed in any way and can be used by other functions 
        as input through the 'InputObject' parameter. This will allow the 
        piping of objects from one function to the next.

        When 'PassThru' is omitted the returning business object property 
        'Fields', that comes from the API, will be converted to a 
        PSCustomObject.

    .EXAMPLE
        Get the tasks from ticket 42173 in the test environment

        Get-CherwellTicketTaskHC @testParams -Type Task -Ticket 42173
#>

    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
        [Alias('TicketNr')]
        [System.Object[]]${Ticket},
        [Switch]${PassThru}
    )
    DynamicParam {
        New-DynamicParameterHC -Name Environment
    }

    begin {
        try {
            $outBuffer = $null
            if ($PSBoundParameters.TryGetValue('OutBuffer', [ref]$outBuffer)) {
                $PSBoundParameters['OutBuffer'] = 1
            }
            $wrappedCmd = $ExecutionContext.InvokeCommand.GetCommand(
                'Get-CherwellTicketDetailHC', [System.Management.Automation.CommandTypes]::Function
            )

            $PSBoundParameters['Type'] = 'Task'

            $scriptCmd = { & $wrappedCmd @PSBoundParameters }
            $steppablePipeline = $scriptCmd.GetSteppablePipeline()
            $steppablePipeline.Begin($PSCmdlet)
        }
        catch {
            throw
        }
    }

    process {
        try {
            $steppablePipeline.Process($_)
        }
        catch {
            throw
        }
    }

    end {
        try {
            $steppablePipeline.End()
        }
        catch {
            throw
        }
    }
}

Function Get-FieldDefinitionHC {
    Param (
        [Parameter(Mandatory)]
        [PSCustomObject]$Schema,
        [Parameter(Mandatory)]
        [String[]]$FieldName
    )

    # The following is not including all fields that are searchable, bug in the API:
    # $SearchableFields = $Schema.FieldDefinitions.Where( { $_.isFullTextSearchable })

    $FieldDefinitions = @($Schema.FieldDefinitions)

    foreach ($F in $FieldName) {
        if (
            -not ($FieldDefinition = $FieldDefinitions.Where( { $_.Name -eq $F }, 'First'))
        ) {
            throw "Field name '$F' is not valid, only the following field names are supported: $($FieldDefinitions.Name -join ', ')."
        }

        $FieldDefinition
    }
}

Function Invoke-GetAccessTokenHC {
    Param (
        [Parameter(Mandatory)]
        [String]$Environment,
        [Parameter(Mandatory)]
        [String]$KeyAPI,
        [Parameter(Mandatory)]
        [String]$Uri,
        [Parameter(Mandatory)]
        [String]$UserName,
        [Parameter(Mandatory)]
        [String]$Password,
        [Parameter(Mandatory)]
        [String]$AuthMode
    )

    Try {
        Write-Verbose "Request a new Cherwell access token for environment '$Environment'"

        $TokenRequestParams = @{
            Method      = 'POST'
            Uri         = '{0}token?auth_mode={1}&api_key={2}' -f $Uri, $AuthMode, $KeyAPI
            Body        = @{
                Accept     = 'application/json'
                grant_type = 'password'
                client_id  = $KeyAPI
                username   = $UserName
                password   = $Password
            }
            TimeoutSec  = 90
            Verbose     = $false
            ErrorAction = 'Stop'
        }
        $Token = Invoke-RestMethod @TokenRequestParams

        if (-not ($Token -is [System.Management.Automation.PSCustomObject])) {
            throw 'The Cherwell server did not return a token in the JSON format'
        }

        if (-not $Token) {
            throw 'The Cherwell server did not return an access token'
        }

        #region Convert strings to DateTime
        $DateTimeProperties = @('.issued', '.expires')

        foreach ($D in $DateTimeProperties) {
            $Token.$D = [DateTime]$Token.$D
        }
        #endregion

        $Token
    }
    Catch {
        $ErrorDetails = ($_.ErrorDetails | Select-Object -ExpandProperty Message) -replace "`r`n"

        throw "Failed authenticating to the Cherwell API`r`n- Environment: $Environment`r`n- Uri: $Uri`r`n- KeyAPI: $KeyAPI`r`n- UserName: $UserName`r`n- Password: $Password`r`n- AuthMode: $AuthMode`r`n- ERROR MESSAGE: $($_.Exception.Message)-`r`n- ERROR DETAILS: $ErrorDetails"
    }
}

Function Invoke-GetBusinessObjectHC {
    Param (
        [Parameter(Mandatory)]
        [String]$BusObId,
        [Parameter(Mandatory, ParameterSetName = 'A')]
        [String]$BusObRecId,
        [Parameter(Mandatory, ParameterSetName = 'B')]
        [String]$PublicId
    )

    $ID = if ($BusObRecId) {
        '/busobrecid/' + $BusObRecId
    }
    else {
        '/publicid/' + $PublicId
    }

    $Params = @{
        Uri = (
            $Uri + 'api/V1/GetBusinessObject' +
            '/busobid/' + $BusObId +
            $ID
        )
    }
    Invoke-RestMethod @GetParams @Params
}

Function Invoke-GetBusinessObjectSchemaHC {
    Param (
        [Parameter(Mandatory)]
        [String]$BusObId
    )

    Write-Verbose "Get business object schema for '$BusObId'"

    $Params = @{
        Uri = (
            $Uri + 'api/V1/GetBusinessObjectSchema' +
            '/busobid/' + $BusObId +
            '?' + 'includerelationships=true'
        )
    }
    Invoke-RestMethod @GetParams @Params
}

Function Invoke-GetBusinessObjectSummaryHC {
    Param (
        [Parameter(Mandatory)]
        [String]$BusObName
    )

    Write-Verbose "Get business object summary for '$BusObName'"

    $Params = @{
        Uri = (
            $Uri + 'api/V1/GetBusinessObjectSummary' +
            '/busobname/' + $BusObName
        )
    }
    Invoke-RestMethod @GetParams @Params
}

Function Invoke-GetBusinessObjectTemplateHC {
    Param (
        [Parameter(Mandatory)]
        [String]$BusObId
    )

    Write-Verbose "Get business object template for '$BusObId'"

    $Params = @{
        Uri  = ($Uri + 'api/V1/GetBusinessObjectTemplate')
        Body = @{
            busObId         = $BusObId
            includeRequired = $true
            includeAll      = $true
        } | ConvertTo-Json
    }
    Invoke-RestMethod @PostParams @Params
}

Function Invoke-GetListOfUsersHC {
    Param ()

    # https://1itsm-test.grouphc.net/CherwellAPI/api/V1/getlistofusers?loginidfilter=Both&stoponerror=false

    $Params = @{
        Uri = (
            $Uri + 'api/V1/GetListOfUsers' +
            '?loginidfilter=Both' +
            '&stoponerror=false'
        )
    }
    Invoke-RestMethod @GetParams @Params
}

Function Invoke-GetRelatedBusinessObjectHC {
    Param (
        [Parameter(Mandatory)]
        [String]$ParentBusObId,
        [Parameter(Mandatory)]
        [String]$ParentBusObRecId,
        [Parameter(Mandatory)]
        [String]$RelationshipId
    )

    $Params = @{
        Uri  = ($Uri + "api/V1/GetRelatedBusinessObject")
        Body = @{
            AllFields        = $true
            ParentBusObId    = $ParentBusObId
            ParentBusObRecId = $ParentBusObRecId
            RelationshipId   = $RelationshipId
            UseDefaultGrid   = $true
        } | ConvertTo-Json
    }
    Invoke-RestMethod @PostParams @Params
}

Function Invoke-GetSearchItemsHC {
    <#
    .EXAMPLE
        $Test = Invoke-GetSearchItemsHC
        $test.supportedAssociations.name

        Approval Request
        Approval Task
        Change Request
        Config - Computer
        Config - Mobile Device
        Config - Network Device
        Config - Other CI
        Config - Printer
        Config - Server
        Config - Software Component
        Config - Software License
        Config - System
        Config - Telephony Equipment
        Configuration Item
        CountrySpecific
        Customer - Internal
        Event
        HC-AirWatch Import Table
        HC-CiscoAP Import Table
        HC-CiscoSNIF Import Table
        HC-itracks Import Table
        HC-KoMi Import Table
        HC-SCCM Computer Import Table
        HC-SCCM Server Import Table
        HC-SolarWinds Import Table
        HC-SolarWinds Staging Table
        HC-VMWare Server Import Table
        HC-WyseThinclients DEUHEIDWYDM01 Import Table
        HC-WyseThinclients SGPSINGWYDM01 Import Table
        HC-WyseThinclients USALEWIWYDM01 Import Table
        Journal
        Journal Attachment
        Knowledge Article
        Problem
        Product Catalog
        Quick Call Template
        Risk Assessment
        Scorecard
        SCT_WU_Join
        Service
        Service Cart
        Service Catalog Template
        Service Schedule
        SLA
        SLA Target Time
        Specific - HC Generic Incident Form Questions
        Task
        Ticket
        UserInfo
        Work Item
        Work Unit
    #>
    Param (
    )

    $Params = @{
        Uri = (
            $Uri + 'api/V2/GetSearchItems'
        )
    }
    Invoke-RestMethod @GetParams @Params
}

Function Invoke-GetSearchResultsHC {
    <#
    .SYNOPSIS
        Internal module function

    .DESCRIPTION
        Internal module function to call the API only in one place and have a 
        structured approach on how to do so in case of a search.
#>

    [CmdletBinding(DefaultParameterSetName = 'Default')]
    [OutputType([PSCustomObject[]])]
    Param (
        [Parameter(Mandatory)]
        [String]$BusObId,
        [Parameter(Mandatory, ParameterSetName = 'Field')]
        [Parameter(Mandatory, ParameterSetName = 'FieldFilter')]
        [String[]]$Property,
        [Parameter(Mandatory, ParameterSetName = 'Filter')]
        [Parameter(Mandatory, ParameterSetName = 'FieldFilter')]
        [HashTable[]]$Filter,
        [Parameter(Mandatory, ParameterSetName = 'Field')]
        [Parameter(Mandatory, ParameterSetName = 'Filter')]
        [Parameter(Mandatory, ParameterSetName = 'FieldFilter')]
        [PSCustomObject]$Schema,
        [ValidateRange(10, 5000)]
        [Int]$PageSize = 1000,
        [Switch]$PassThru
    )

    Begin {
        $Body = @{
            BusObID    = $BusObId
            PageSize   = $PageSize
            PageNumber = 0
        }

        if ($Filter) {
            $Params = @{
                Schema = $Schema
                Filter = $Filter
            }
            $Body.Filters = @(New-CherwellSearchFilterHC @Params)
        }

        if ($Property -eq '*') {
            $Body.includeAllFields = $true
        }
        elseif ($Property) {
            $Body.Fields = @((Get-FieldDefinitionHC -Schema $Schema -FieldName $Property).FieldId)
        }
    }

    Process {
        $RetrievedObjects = 0
        $Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

        Do {
            Select-EnvironmentHC -Name $Environment

            $Body.PageNumber++

            $Params = @{
                Uri        = ($Uri + 'api/V1/getsearchresults')
                Body       = [System.Text.Encoding]::UTF8.GetBytes(($Body | ConvertTo-Json))
                TimeoutSec = 300
            }
            $Result = Invoke-RestMethod @PostParams @Params

            $RetrievedObjects += @($Result.businessObjects).Count

            if ($Result.TotalRows -ne 0) {
                if (@($Result.businessObjects).Count -eq 0) {
                    throw "The REST API reports '$($Result.TotalRows)' TotalRows but there are no businessObjects present in the returned object. This must be a bug in the Cherwell REST API, see SR-152489."
                }
                if ($PassThru) {
                    $Result.businessObjects
                }
                else {
                    ConvertTo-PSCustomObjectHC -BusinessObject $Result.businessObjects
                }

                Write-Verbose "Retrieved $RetrievedObjects/$($Result.TotalRows) in '$($Stopwatch.Elapsed.ToString())'"
            }
            else {
                Write-Verbose "Nothing found"
            }
        } While ($RetrievedObjects -lt $Result.TotalRows)

        $Stopwatch.Stop()
    }
}

Function Invoke-LinkRelatedBusinessObjectHC {
    Param (
        [Parameter(Mandatory)]
        [String]$ParentBusObId,
        [Parameter(Mandatory)]
        [String]$ParentBusobRecId,
        [Parameter(Mandatory)]
        [String]$RelationshipId,
        [Parameter(Mandatory)]
        [String]$BusobId,
        [Parameter(Mandatory)]
        [String]$BusobRecId
    )

    $Params = @{
        Uri = (
            $Uri + 'api/V1/LinkRelatedBusinessObject' +
            '/parentbusobid/' + $ParentBusObId +
            '/parentbusobrecid/' + $ParentBusobRecId +
            '/relationshipid/' + $RelationshipId +
            '/busobid/' + $BusobId +
            '/busobrecid/' + $BusobRecId
        )
    }
    $null = Invoke-RestMethod @GetParams @Params
}

Function Invoke-SaveBusinessObjectHC {
    Param (
        [Parameter(Mandatory)]
        [String]$BusObId,
        [Parameter(Mandatory)]
        [PSCustomObject[]]$Property
    )

    $Params = @{
        Uri  = ($Uri + 'api/V1/SaveBusinessObject')
        Body = [System.Text.Encoding]::UTF8.GetBytes((@{
                    BusObId = $BusObId
                    Fields  = $Property
                } | ConvertTo-Json))
    }
    Invoke-RestMethod @PostParams @Params
}

Function Invoke-UploadBusinessObjectAttachmentHC {
    Param (
        [Parameter(Mandatory)]
        [String]$FileName,
        [Parameter(Mandatory)]
        [String]$BusobId,
        [Parameter(Mandatory)]
        [String]$BusobRecId,
        [Parameter(Mandatory)]
        [Int]$Offset,
        [Parameter(Mandatory)]
        [Double]$TotalSize,
        [Parameter(Mandatory)]
        [String]$Body
    )

    $Params = @{
        Uri  = (
            $Uri + 'api/V1/UploadBusinessObjectAttachment' +
            '/filename/' + $FileName +
            '/busobid/' + $BusobId +
            '/busobrecid/' + $BusobRecId +
            '/offset/' + $Offset +
            '/totalsize/' + $TotalSize
        )
        Body = [System.IO.File]::ReadAllBytes($Body)
    }
    $null = Invoke-RestMethod @PostParams @Params
}

Function New-CherwellConfigItemHC {
    <#
    .SYNOPSIS
        Create a new CI in Cherwell.

    .DESCRIPTION
        Create a new configuration item in Cherwell.

    .PARAMETER Environment
        The Cherwell system that needs to be addressed.

        All available systems are defined in the file 'Passwords.json' in the 
        module folder. Ex. Test, Prod, ...

    .PARAMETER Type
        The type of configuration item as it is known in Cherwell. To retrieve 
        a list of possible values the function 'Get-CherwellConfigItemTypeHC' 
        can be used.

    .PARAMETER KeyValuePair
        One or more hash tables containing key value pair combinations to 
        create new CI's. The key name represents the field name in Cherwell and 
        the key value contains the Cherwell value desired for that field.

    .EXAMPLE
        Create a new CI in the Cherwell test environment

        New-CherwellConfigItemHC -Environment Test -Type 'ConfigPrinter' -KeyValuePair @{
            FriendlyName    = 'BELPROOST14'
            HostName        = 'BELPROOST14'
            SerialNumber    = '9874605849684'
            Country         = 'BNL'
            Location        = "BEL\Braine L'Alleud\"
            Comment         = 'KoMi rollout 2019'
            Manufacturer    = 'Konica Minolta'
            Model           = 'Bizhub 3545c'
        }
#>

    [OutputType()]
    [CmdLetBinding()]
    Param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [String]$Type,
        [Parameter(Mandatory)]
        [HashTable[]]$KeyValuePair
    )
    DynamicParam {
        New-DynamicParameterHC -Name Environment
    }

    Begin {
        Try {
            $Environment = $PSBoundParameters.Environment
            Select-EnvironmentHC -Name $Environment

            Try {
                $ConfigItemTemplate = Get-TemplateHC -Name $Type
                $ConfigItemSummary = Get-SummaryHC -Name $Type
            }
            Catch {
                $CiTypes = Get-CherwellConfigItemTypeHC -Environment $Environment
                throw "$_. Only the following CI types are supported: $($CiTypes -join ', ')"
            }
        }
        Catch {
            throw "Failed creating a new CI in Cherwell: $_."
        }
    }

    Process {
        Try {
            foreach ($K in $KeyValuePair) {
                Try {
                    Write-Verbose "Create CI of type '$Type'"
                    $DirtyTemplate = New-TemplateHC -Name $ConfigItemTemplate -KeyValuePair $K

                    #region Create a new CI
                    $Params = @{
                        BusObId  = $ConfigItemSummary.busobid
                        Property = $DirtyTemplate.Fields
                    }
                    $createBOResponse = Invoke-SaveBusinessObjectHC @Params

                    $createBOResponse
                    #endregion

                    #region API errors
                    if ($createBOResponse.hasError) {
                        throw "Configuration item has error code '$($createBOResponse.errorCode)' with message '$($createBOResponse.errorMessage)'"
                    }
                    #endregion

                    Write-Verbose 'CI created'
                }
                Catch {
                    $M = Convert-ApiErrorHC $_
                    $KeyValueString = $K.GetEnumerator().ForEach( { '{0} = "{1}"' -f $_.Key, $_.Value }) -join "`n"

                    Write-Error "Failed creating a new CI in Cherwell '$Environment' with the following details: `n`n$KeyValueString `n`nError: $M"
                }
            }
        }
        Catch {
            throw "Failed creating a new CI in Cherwell '$Environment': $_"
        }
    }
}

Function New-CherwellSearchFilterHC {
    <#
    .SYNOPSIS
        Create a valid Cherwell search filter.

    .DESCRIPTION
        Create a valid Cherwell search filter that is accepted by the API and 
        is tested for invalid operators, search fields names, ...

    .PARAMETER Schema
        The schema of the configuration item.

    .PARAMETER Filter
        A hash table containing the search parameters used by API to search for 
        the correct configuration item in the collection defined in the 
        parameter 'Type'. The following fields/keys are mandatory: FieldName, 
        Operator and FieldValue.

        The following operators are supported: 'eq', 'gt', 'lt', 'contains' and 
        'startswith'

        You can specify more than one filter. If you add multiple filters for 
        the same FieldName, the result is an OR operation between those fields. 
        If the FieldNames are different, the result is an AND operation between 
        those fields

    .EXAMPLE
        New-CherwellSearchFilterHC -Schema (Get-SchemaHC -Name 'Incident') -Filter @{
            FieldName  = 'IncidentID'
            Operator   = 'eq'
            FieldValue = '354980'
        }
 #>

    [OutputType([HashTable[]])]
    [CmdLetBinding()]
    Param (
        [Parameter(Mandatory)]
        [PSCustomObject]$Schema,
        [Parameter(Mandatory)]
        [HashTable[]]$Filter
    )

    Begin {
        Try {
            $AcceptedOperators = @('eq', 'gt', 'lt', 'contains', 'startswith')
            $MandatoryKeys = @('FieldName', 'Operator', 'FieldValue')
        }
        Catch {
            throw "Failed creating a search filter: $_"
        }
    }

    Process {
        Try {
            $FilterString = @()

            $Result = foreach ($F in $Filter) {
                $newHash = $F.Clone()

                #region Test mandatory keys present
                foreach ($K in $MandatoryKeys) {
                    if (-not $newHash.ContainsKey($K)) {
                        throw "Key '$K' is mandatory"
                    }
                }
                #endregion

                #region Test only 3 keys in the hashtable
                if ($newHash.Count -ne $MandatoryKeys.Count) {
                    throw "Only the following keys are allowed: $($MandatoryKeys -join ', ')."
                }
                #endregion

                #region Test operator value
                if (-not $AcceptedOperators.Contains($newHash.Operator)) {
                    throw "The operator '$($newHash.Operator)' is not accepted, only the following operators are supported: $($AcceptedOperators -join ', ')."
                }
                #endregion

                $newHash.FieldId = (Get-FieldDefinitionHC -Schema $Schema -FieldName $newHash.FieldName).FieldId
                $newHash.Value = $newHash.FieldValue

                $FilterString += "'$($newHash.FieldName) $($newHash.Operator) $($newHash.FieldValue)'"

                $newHash.Remove('FieldName')
                $newHash.Remove('FieldValue')

                $newHash
            }

            $Result

            Write-Verbose "Filter $($FilterString -join ' or ')"
        }
        Catch {
            throw "Failed creating a search filter: $_"
        }
    }
}

Function New-CherwellTicketHC {
    <#
    .SYNOPSIS
        Create a new ticket in Cherwell.

    .DESCRIPTION
        This function facilitates the creation of one or more tickets in the 
        Cherwell system, by using the Cherwell REST API. By default the ticket 
        number is returned of the newly created ticket. When the the switch 
        '-PassThru' is used a Cherwell object will be returned that can be used 
        with other Cherwell CmdLets.

    .PARAMETER Environment
        The Cherwell system that needs to be addressed.

        All available systems are defined in the file 'Passwords.json' in the 
        module folder. Ex. Test, Prod, ...

    .PARAMETER KeyValuePair
        One or more hash tables containing key value pair combinations to 
        create tickets. The key represents the field name in a Cherwell ticket 
        and the value represents the value for that specific field.

        Mandatory fields:
        - 'CustomerRecID' or 'RequesterSamAccountName'
        - 'IncidentType' (Incident, Service Request, ...)
        - And others as required by the Cherwell API 
          (Service, Category, SubCategory, ...)

        Convenience fields:
        These are custom convenience parameters where the `SamAccountName` is 
        matched (by `Get-CherwellSystemUserHC` or `Get-CherwellCustomerHC`) 
        with the correct user/customer within Cherwell:

           Name                       > Cherwell field name
        - `SubmittedBySamAccountName` > `SubmitOnBehalfOfID`
        - `OwnedBySamAccountName`     > `OwnedByID` and `OwnedBy`
        - `RequesterSamAccountName`   > `CustomerRecID`

        The key 'CI' can contain an object coming from 
        Get-CherwellConfigItemHC -PassThru', representing
        a CI object that has been retrieved from Cherwell. To simplify ticket 
        creation it is also possible to provide a simple hash table that 
        contains the search criteria for the CI. When using the latter, the 
        function 'Get-CherwellConfigItemHC' is called in the background to find
        the correct CI.

        Note: when using a hash table as CI it is possible that the CI cannot
        be found. When this happens a warning is displayed but the ticket is 
        nonetheless created but without CI.

    .PARAMETER PassThru
        Return the output of API calls in their raw format. The output has not 
        been converted or changed in any way and can be used by other functions 
        as input through the 'InputObject' parameter. This will allow the 
        piping of objects from one function to the next.

        When 'PassThru' is omitted the returning business object property 
        'Fields', that comes from the API, will be converted to a 
        PSCustomObject.

    .EXAMPLE
        Create a ticket using the convenience arguments, the ones ending with 
        `*SamAccountName`. This is the quickest way to create a new ticket. The 
        module will try to find a single matching system user/customer 
        utomatically. 

        $ticketNr = New-CherwellTicketHC -Environment Stage -KeyValuePair @{
            IncidentType            = 'Incident'
            RequesterSamAccountName = 'cnorris'
            OwnedBySamAccountName   = 'bmarley'
            OwnedByTeam             = 'GLB'
            ShortDescription        = 'Server decommissioning'
            Description             = 'Decommission server'
            Priority                = '2'
            Source                  = 'Event'
            ServiceCountryCode      = 'BNL'
            Service                 = 'SERVER'
            Category                = 'Server lifecycle'
            SubCategory             = 'Submit Request'
        }

    .EXAMPLE
        create a ticket with known Cherwell fields only. This will allows more 
        control in case of errors. 

        $systemUser = Get-CherwellSystemUserHC -Environment Stage -Filter @{
            FieldName  = 'EMail'
            Operator   = 'eq'
            FieldValue = 'gmail@chucknorris.com'
        } -PassThru

        if ($systemUser.count -ne 1) { throw 'No user or multiple users found' }

        $customer = Get-CherwellCustomerHC -Environment Stage -Filter @{
            FieldName  = 'SamAccountName'
            Operator   = 'eq'
            FieldValue = 'bmarley'
        } -PassThru

        if ($customer.count -ne 1) { throw 'No customer or multiple customers found' }

        $ticketNr =  New-CherwellTicketHC -Environment Stage -KeyValuePair @{
            IncidentType            = 'Incident'
            CustomerRecID           = $customer.busObRecId
            OwnedByTeam             = 'GLB'
            OwnedBy                 = $systemUser.busObPublicId
            OwnedById               = $systemUser.busObRecId
            ShortDescription        = 'Server decommissioning'
            Description             = 'Decommission server'
            Priority                = '2'
            Source                  = 'Event'
            ServiceCountryCode      = 'BNL'
            Service                 = 'SERVER'
            Category                = 'Server lifecycle'
            SubCategory             = 'Submit Request'
        }

    .EXAMPLE
        Create two tickets.

        New-CherwellTicketHC -Environment Test -KeyValuePair @(
            @{
                IncidentType            = 'Incident'
                RequesterSamAccountName = 'norrisc'
                OwnedByTeam             = 'PRINTER TEAM'
                ShortDescription        = 'Printer offline'
                Description             = 'Check the connection to the printer'
                Priority                = '2'
                Source                  = 'Event'
                ServiceCountryCode      = 'BNL'
                Service                 = 'CLIENT'
                Category                = 'Printing'
                SubCategory             = 'Submit incident'
            }
            @{
                IncidentType              = 'Service Request'
                RequesterSamAccountName   = 'bswagger'
                SubmittedBySamAccountName = 'jbond'
                OwnedByTeam               = 'AD TEAM'
                ShortDescription          = 'Account expiring'
                Description               = "The account will expire"
                Priority                  = '3'
                Source                    = 'Event'
                ServiceCountryCode        = 'BNL'
                Service                   = 'CLIENT'
                Category                  = 'AD OBJECT MANAGEMENT'
                SubCategory               = 'Request Service'
            }
        )

    .EXAMPLE
        Create a new Cherwell ticket with a CI and return a Cherwell object 
        that can be used with other CmdLets. See the `-PassThru` argument.

        $CI = New-CherwellConfigItemHC -Environment Stage -KeyValuePair @(
            @{
                CIType          = 'ConfigPrinter'
                Owner           = 'thardey'
                SupportedByTeam = 'BNL'
                FriendlyName    = 'BELPROOST14'
                SerialNumber    = '9874605849684'
                Country         = 'BNL'
                PrinterType     = 'Local printer'
                Manufacturer    = 'Konica Minolta'
                Model           = 'Bizhub 3545c'
            }
        )

        New-CherwellTicketHC -PassThru -Environment Stage -KeyValuePair @{
            IncidentType            = 'Incident'
            RequesterSamAccountName = 'norrisc'
            OwnedByTeam             = 'PRINTER TEAM'
            ShortDescription        = 'Printer offline'
            Description             = 'Please check the connection'
            Priority                = '2'
            Source                  = 'Event'
            ServiceCountryCode      = 'BNL'
            Service                 = 'CLIENT'
            Category                = 'Printing'
            SubCategory             = 'Submit incident'
            CI                      = $CI
        }

    .EXAMPLE
        Create a new Cherwell ticket with a CI provided by a hash table and 
        return the ticket number. When the CI cannot be found the ticket will 
        simply be created without CI and a warning will be displayed.

        $CI = @{
            Type   = 'ConfigPrinter'
            Filter = @{
                FieldName  = 'FriendlyName'
                Operator   = 'eq'
                FieldValue = 'BELPROOS3'
            }
        }

        New-CherwellTicketHC -Environment Stage -KeyValuePair @{
            IncidentType            = 'Incident'
            RequesterSamAccountName = 'norrisc'
            OwnedByTeam             = 'PRINTER TEAM'
            ShortDescription        = 'Printer offline'
            Description             = 'Check the connection to the printer'
            Priority                = '2'
            Source                  = 'Event'
            ServiceCountryCode      = 'BNL'
            Service                 = 'CLIENT'
            Category                = 'Printing'
            SubCategory             = 'Submit incident'
            CI                      = $CI
        }

    .EXAMPLE
        Create a new ticket and fill in the 'Additional details' field in the 
        ticket.        
        
        New-CherwellTicketHC -Environment Stage -KeyValuePair @{
            IncidentType            = 'Request'
            RequesterSamAccountName = 'norrisc'
            OwnedByTeam             = 'PRINTER TEAM'
            ShortDescription        = 'Printer installation'
            Description             = 'Please install printer'
            Priority                = '3'
            Source                  = 'Event'
            ServiceCountryCode      = 'BNL'
            Service                 = 'CLIENT'
            Category                = 'Printing'
            SubCategory             = 'Submit request'
        } | 
        Set-CherwellTicketDetailAdditionalDetailsHC -Environment Stage -Text 'Please follow KB-707'
 #>

    [OutputType([Int])]
    [CmdLetBinding()]
    Param (
        [Parameter(Mandatory)]
        [HashTable[]]$KeyValuePair,
        [Switch]$PassThru
    )
    DynamicParam {
        New-DynamicParameterHC -Name Environment
    }

    Begin {
        Try {
            $Environment = $PSBoundParameters.Environment
            Select-EnvironmentHC -Name $Environment

            $Template = Get-TemplateHC -Name 'Incident'
        }
        Catch {
            throw "Failed creating a new ticket in Cherwell '$Environment': $_"
        }
    }

    Process {
        foreach ($K in $KeyValuePair) {
            Try {
                Write-Verbose "Create new ticket in environment '$Environment'"

                $RequestFields = $K.Clone()

                Test-InvalidPropertyCombinationHC -KeyValuePair $RequestFields

                #region RequesterSamAccountName or CustomerRecID is mandatory
                if (-not (
                        ($RequestFields.ContainsKey('RequesterSamAccountName')) -or ($RequestFields.ContainsKey('CustomerRecID')))
                ) {
                    throw "The field 'RequesterSamAccountName' or 'CustomerRecID' is mandatory"
                }
                #endregion

                #region IncidentType
                if (-not ($RequestFields.ContainsKey('IncidentType'))) {
                    throw "The field 'IncidentType' is mandatory"
                }
                #endregion

                #region CI
                if (
                    ($RequestFields.ContainsKey('ConfigItemRecID')) -or
                    ($RequestFields.ContainsKey('ConfigItemDisplayName'))
                ) {
                    throw "When you want to attach a 'Primary CI' to the ticket, please use the key 'CI' containing a CI object as retrieved by 'Get-CherwellConfigItemHC -PassThru' or a simple hashtable like @{Type = 'Computer'; Filter = @{FieldName = 'HostName'; Operator = 'eq' ; FieldValue = 'BELCL003000'}}. Other CI keys like 'ConfigItemRecID' and 'ConfigItemDisplayName' are not allowed."
                }

                $CI = $null

                if ($RequestFields.ContainsKey('CI')) {
                    if ($RequestFields.CI -is [HashTable]) {
                        $Params = @{
                            Environment = $Environment
                            Type        = $RequestFields.CI.Type
                            Filter      = $RequestFields.CI.Filter
                        }
                        $CI = @(Get-CherwellConfigItemHC @Params -PassThru)

                        if (-not $CI) {
                            Write-Warning "A CI object of type '$($RequestFields.CI.Type)' with search filter '$($RequestFields.CI.Filter.FieldName) $($RequestFields.CI.Filter.Operator) $($RequestFields.CI.Filter.FieldValue)' was not found. The ticket will be created without CI."
                        }
                    }
                    else {
                        $CI = @($RequestFields.CI)
                    }

                    $RequestFields.Remove('CI')
                }
                #endregion

                #region RequesterSamAccountName
                if ($RequestFields.ContainsKey('RequesterSamAccountName')) {

                    $RequesterSamAccountName = $RequestFields.RequesterSamAccountName

                    $CherwellCustomer = @(Get-CherwellCustomerHC -Environment $Environment -Filter @{
                            FieldName  = 'SamAccountName'
                            Operator   = 'eq'
                            FieldValue = $RequesterSamAccountName
                        } -PassThru)

                    if ($CherwellCustomer.Count -eq 1) {
                        $RequestFields.CustomerRecID = $CherwellCustomer[0].busObRecId
                        $RequestFields.Remove('RequesterSamAccountName')
                    }
                    elseif ($CherwellCustomer.Count -eq 0) {
                        throw "No customer found for field 'RequesterSamAccountName' with SamAccountName '$RequesterSamAccountName'"
                    }
                    else {
                        throw "Multiple customers found for field 'RequesterSamAccountName' with the same SamAccountName '$RequesterSamAccountName'.Please use the field 'CustomerRecID' as retrieved by 'Get-CherwellCustomerHC' to manually find the correct user."
                    }
                }
                #endregion

                #region SubmittedBySamAccountName
                if ($RequestFields.ContainsKey('SubmittedBySamAccountName')) {
                    if ($SubmittedBySamAccountName = $RequestFields.SubmittedBySamAccountName) {
                    
                        $CherwellCustomer = @(Get-CherwellCustomerHC -Environment $Environment -Filter @{
                                FieldName  = 'SamAccountName'
                                Operator   = 'eq'
                                FieldValue = $SubmittedBySamAccountName
                            } -PassThru)

                        if ($CherwellCustomer.Count -eq 1) {
                            $RequestFields.SubmitOnBehalfOfID = $CherwellCustomer[0].busObRecId                        
                        }
                        elseif ($CherwellCustomer.Count -eq 0) {
                            throw "No customer found for field 'SubmittedBySamAccountName' with SamAccountName '$($SubmittedBySamAccountName)'. Please use the field 'SubmitOnBehalfOfID' as retrieved by 'Get-CherwellCustomerHC' to manually find the correct user."
                        }
                        else {
                            throw "Multiple customers found for field 'SubmittedBySamAccountName' with the same SamAccountName '$SubmittedBySamAccountName'.Please use the field 'SubmitOnBehalfOfID' as retrieved by 'Get-CherwellCustomerHC' to manually find the correct user."
                        }
                    }
                    $RequestFields.Remove('SubmittedBySamAccountName')
                }
                #endregion

                #region OwnedBySamAccountName
                if ($RequestFields.ContainsKey('OwnedBySamAccountName')) {
                    $OwnedBySamAccountName = $RequestFields.OwnedBySamAccountName

                    $CherwellUser = @(Get-CherwellSystemUserHC -Environment $Environment -Filter @{
                            FieldName  = 'SamAccountName'
                            Operator   = 'eq'
                            FieldValue = $OwnedBySamAccountName
                        }  -PassThru)

                    if ($CherwellUser.Count -eq 1) {
                        $RequestFields.OwnedByID = $CherwellUser[0].BusObRecId
                        $RequestFields.OwnedBy = $CherwellUser[0].fields |
                        Where-Object Name -EQ 'FullName' | Select-Object -ExpandProperty Value
                     
                        $RequestFields.Remove('OwnedBySamAccountName')
                    }
                    elseif ($CherwellUser.Count -eq 0) {
                        throw "No system user found with SamAccountName '$OwnedBySamAccountName'. Please use the fields 'OwnedBy' and 'OwnedById' as retrieved by 'Get-CherwellSystemUserHC' to manually find the correct user."
                    }
                    else {
                        throw "Multiple system users found for field 'OwnedBy' with the same SamAccountName '$OwnedBySamAccountName'. Please use the fields 'OwnedBy' and 'OwnedById' as retrieved by 'Get-CherwellSystemUserHC' to manually find the correct user."
                    }
                }
                #endregion

                #region Attachment
                $Attachment = $null

                if ($RequestFields.ContainsKey('Attachment')) {
                    $Attachment = $RequestFields.Attachment
                    $RequestFields.Remove('Attachment')
                }
                #endregion

                #region Create a new incident template
                $DirtyTemplate = New-TemplateHC -Name $Template -KeyValuePair $RequestFields
                #endregion

                #region Create a new incident ticket
                $Params = @{
                    BusObId  = (Get-SummaryHC -Name 'Incident').busobid
                    Property = $DirtyTemplate.Fields
                }
                $createBOResponse = Invoke-SaveBusinessObjectHC @Params

                if ([Int]$TicketNr = $createBOResponse.busObPublicId) {
                    Write-Verbose "Ticket '$ticketNr'"
                    if ($PassThru) {
                        $createBOResponse
                    }
                    else {
                        $TicketNr
                    }
                }
                else {
                    throw 'No ticket number received from Cherwell'
                }
                #endregion

                #region Add attachment
                if ($Attachment) {
                    $Params = @{
                        Environment = $Environment
                        TicketNr    = $TicketNr
                        File        = $Attachment
                    }
                    $null = Add-CherwellTicketAttachmentHC @Params
                }
                #endregion

                #region Add CI
                if ($CI) {
                    $Params = @{
                        Environment = $Environment
                        Ticket      = $createBOResponse
                        ConfigItem  = $CI
                    }
                    $null = Add-CherwellTicketConfigItemHC @Params

                    # Set Primary CI
                    Try {
                        $Params = @{
                            Environment  = $Environment
                            TicketNr     = $TicketNr
                            KeyValuePair = @{
                                ConfigItemRecId = $CI[0].busObRecId
                            }
                        }
                        $null = Update-CherwellTicketHC @Params -EA Stop
                    }
                    Catch {
                        $CIFriendlyName = ($CI.fields.Where( { $_.name -eq 'FriendlyName' })).Value
                        Write-Warning "Failed setting the 'Primary CI' because most likely there are multiple CI objects with the same FriendlyName '$CIFriendlyName' and Cherwell can't handle this. $_"
                    }

                }
                #endregion

                #region API errors
                if ($createBOResponse.hasError) {
                    throw "Ticket '$ticketNr' has error code '$($createBOResponse.errorCode)' with message '$($createBOResponse.errorMessage)'"
                }
                #endregion
            }
            Catch {
                $M = Convert-ApiErrorHC $_
                $KeyValueString = $K.GetEnumerator().ForEach( { '{0} = "{1}"' -f $_.Key, $_.Value }) -join ";`n"
                Write-Error "Failed creating a new ticket in Cherwell '$Environment' with the following details: `n`n$KeyValueString `n`nERROR: $M"
            }
        }

    }
}

Function New-TemplateHC {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory)]
        [PSCustomObject]$Name,
        [Parameter(Mandatory)]
        [HashTable]$KeyValuePair
    )
    try {
        $NewTemplate = $Name | ConvertTo-Json -Depth 100 | ConvertFrom-Json

        $SupportedFields = $NewTemplate.Fields.Name | Sort-Object

        $KeyValuePair.GetEnumerator().ForEach( {
                $FieldName = $_.Key
                $FieldValue = $_.Value

                if (-not ($Field = $NewTemplate.Fields | Where-Object { $_.Name -eq $FieldName })) {
                    throw "Field name '$FieldName' does not exist, only the following fields are supported: $($SupportedFields -join ', ')"
                }

                $Field.Value = $FieldValue
                $Field.Dirty = $true
            })

        $NewTemplate
    }
    catch {
        throw "Failed setting the value '$FieldValue' in field '$FieldName': $_"
    }
}

Function Remove-CherwellTicketHC {
    <#
    .SYNOPSIS
        Remove a one or more tickets in Cherwell.

    .DESCRIPTION
        This function facilitates the deletion of tickets in Cherwell.  In case 
        a ticket was created by accident, it can be convenient to delete the 
        ticket.

        Delete permissions in Cherwell are required to use this function, 
        otherwise an error will be thrown similar to: 'You do not have rights 
        to delete the Ticket business object'.

    .PARAMETER TicketNr
        The number to identify the ticket in Cherwell. Multiple numbers are 
        accepted.

    .PARAMETER Environment
        The Cherwell system that needs to be addressed.

        All available systems are defined in the file 'Passwords.json' in the 
        module folder. Ex. Test, Prod, ...

    .EXAMPLE
        Delete ticket number '5' and '6' in the Cherwell ticketing tool
        
        Remove-CherwellTicketHC -Environment Test -TicketNr 5, 6
 #>

    [OutputType()]
    [CmdLetBinding()]
    Param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [Alias('TicketNr')]
        [Object[]]$Ticket
    )
    DynamicParam {
        New-DynamicParameterHC -Name Environment
    }

    Begin {
        Try {
            $Environment = $PSBoundParameters.Environment
            Select-EnvironmentHC -Name $Environment

            $IncidentSummary = Get-SummaryHC -Name 'Incident'
        }
        Catch {
            throw "Failed removing ticket '$Ticket' from Cherwell '$Environment': $_"
        }
    }

    Process {
        foreach ($T in $Ticket) {
            Try {
                Write-Verbose "Remove ticket from environment '$Environment'"

                $TicketObject = ConvertTo-TicketObjectHC -Item $T

                #region Check if the ticket exists
                Try {
                    $Params = @{
                        BusObId  = $IncidentSummary.busobid
                        PublicId = $TicketObject.publicid
                    }
                    $TicketDetails = Invoke-GetBusinessObjectHC @Params
                }
                Catch {
                    Write-Warning "Ticket number '$T' not found in Cherwell."
                }
                #endregion

                # Delete ticket
                if ($TicketDetails) {
                    $null = Invoke-RestMethod @DeleteParams -Uri (
                        $Uri + "api/V1/deletebusinessobject/busobid/" +
                        $IncidentSummary.busobId +
                        "/publicid/$($TicketObject.publicid)")
                }
                #endregion

                Write-Verbose "Removed ticket '$($TicketObject.publicid)'"
            }
            Catch {
                throw "Failed removing ticket '$T' from Cherwell '$Environment': $_"
            }
        }
    }
}

Function Select-EnvironmentHC {
    Param (
        [Parameter(Mandatory)]
        [String]$Name
    )

    Try {
        $CurrentEnvironment = $EnvironmentList[$Name]

        if ((-not $CurrentEnvironment.ContainsKey('Token')) -or
            ($CurrentEnvironment['Token'].'.expires' -le (Get-Date).AddMinutes(5))) {
            $Params = @{
                Environment = $Name
                Uri         = $CurrentEnvironment.Uri
                KeyAPI      = $CurrentEnvironment.KeyAPI
                UserName    = $CurrentEnvironment.UserName
                Password    = $CurrentEnvironment.Password
                AuthMode    = $CurrentEnvironment.AuthMode
            }
            $CurrentEnvironment.Token = Invoke-GetAccessTokenHC @Params
        }

        $Script:Environment = $Name
        $Script:Uri = $CurrentEnvironment.Uri

        $Script:GetParams = @{
            Method      = 'GET'
            ContentType = 'application/json'
            Header      = @{
                Authorization = 'Bearer {0}' -f $CurrentEnvironment.Token.access_token
            }
            Verbose     = $false
        }

        $Script:PostParams = @{
            Method      = 'POST'
            ContentType = 'application/json'
            Header      = @{
                Authorization = 'Bearer {0}' -f $CurrentEnvironment.Token.access_token
            }
            Verbose     = $false
        }
    }
    Catch {
        throw "Failed selecting the environment '$Name': $_"
    }
}

Function Set-CherwellTicketDetailHC {
    <#
    .SYNOPSIS
        Set details in tickets.

    .DESCRIPTION
        Set details in tickets by using related objects. Not all fields are 
        exposed in the standard ticket template. To be able to update all 
        fields, related ticket objects are used.

        NOTE: Only 1 related object can be updated at this time. Further 
        development of this feature is still required.

    .PARAMETER Ticket
        Objects coming from 'Get-CherwellTicketHC -PassThru' or a ticket number.
        Multiple tickets are supported.

    .PARAMETER Environment
        The Cherwell system that needs to be addressed.

        All available systems are defined in the file 'Passwords.json' in the 
        module folder. Ex. Test, Prod, ...

    .PARAMETER Type
        The related business object type.

    .PARAMETER KeyValuePair
        A hash table containing key value pair combinations of the fields to 
        update.

    .PARAMETER PassThru
        Return the output of API calls in their raw format. The output has not 
        been converted or changed in any way and can be used by other functions 
        as input through the 'InputObject' parameter. This will allow the 
        piping of objects from one function to the next.

        When 'PassThru' is omitted the returning business object property 
        'Fields', that comes from the API, will be converted to a 
        PSCustomObject.

    .EXAMPLE
        $Params = @{
            Environment  = 'Test'
            Ticket       = 38076
            Type         = 'Specifics'
            KeyValuePair = @{
                Notes = 'Please follow KB-707'
            }
        }
        Set-CherwellTicketDetailHC @Params

        Sets the text in the property 'Additional Details' for ticket 38076 in 
        the test environment

    .EXAMPLE
        Change the text in a task. This only works when there is 1 task. 
        Feature not fully developed

        $Params = @{
            Environment  = 'Test'
            Ticket       = 38568
            Type         = 'Tasks'
            KeyValuePair = @{
                Title             = 'Wok item test task updated'
                TaskTypeName      = 'Work Item'
                Status            = 'New'
                Description       = 'Test description updated'
                ActionNotes       = 'Action notes'
            }
        }
        Set-CherwellTicketDetailHC @Params
 #>

    [OutputType([PSCustomObject[]])]
    [OutputType()]
    Param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [Alias('TicketNr')]
        [Object[]]$Ticket,
        [Parameter(Mandatory)]
        [HashTable]$KeyValuePair
    )
    DynamicParam {
        New-DynamicParameterHC -Name Environment, Type
    }

    Begin {
        Try {
            $Environment = $PSBoundParameters.Environment
            $Type = $PSBoundParameters.Type

            Select-EnvironmentHC -Name $Environment

            $IncidentSchema = Get-SchemaHC -Name 'Incident'

            #region Get RelationshipId
            $RelationshipId = $IncidentSchema.relationships.Where( {
                    $_.DisplayName -eq $RelationshipEnum[$Type] }) |
            Select-Object -ExpandProperty 'RelationshipId'

            if (-not $RelationshipId) {
                throw "RelationshipId not found for type '$Type'."
            }
            #endregion
        }
        Catch {
            throw "Failed setting ticket details ticket '$Ticket': $_ "
        }
    }

    Process {
        Try {
            Write-Verbose "Set ticket details in environment '$Environment'"

            foreach ($T in $Ticket) {
                $TicketObject = ConvertTo-TicketObjectHC -Item $T

                #region Get ticket related object
                $Params = @{
                    parentBusObId    = $IncidentSchema.busObId
                    parentBusObRecId = $TicketObject.busObRecId
                    relationshipId   = $RelationshipId
                }
                $TicketRelation = Invoke-GetRelatedBusinessObjectHC @Params
                #endregion

                if ($TicketRelation.TotalRecords -ne 1) {
                    throw "Feature not supported yet: $($TicketRelation.TotalRecords) related business objects were found and currently updating only 1 is supported."
                }

                #region Create a new template
                $Params = @{
                    Name         = $TicketRelation.relatedBusinessObjects
                    KeyValuePair = $KeyValuePair
                }
                $DirtyTemplate = New-TemplateHC @Params
                #endregion

                #region Set 'Additional details' field
                Write-Verbose "Set ticket details for ticket '$($TicketObject.busObPublicId)'"

                $Params = @{
                    Uri  = ($Uri + 'api/V1/SaveRelatedBusinessObject')
                    Body = [System.Text.Encoding]::UTF8.GetBytes((@{
                                parentBusObId       = $IncidentSchema.busObId
                                parentBusObPublicId = $TicketObject.busObPublicId
                                relationshipId      = $RelationshipId
                                busObId             = $TicketRelation.relatedBusinessObjects.busObId
                                busObRecId          = $TicketRelation.relatedBusinessObjects.busObRecId
                                fields              = $DirtyTemplate.Fields
                                persist             = $true
                            } | ConvertTo-Json))
                }
                $null = Invoke-RestMethod @PostParams @Params
                #endregion
            }
        }
        Catch {
            throw "Failed setting ticket details fpr ticket '$Ticket' in Cherwell '$Environment': $_ "
        }
    }
}

Function Test-InvalidPropertyCombinationHC {
    Param (
        [Parameter(Mandatory)]
        [HashTable]$KeyValuePair
    )

    $ownedByTeamIsMissingError = "When the fields 'OwnedBy', 'OwnedById' or 'OwnedBySamAccountName' are used it is mandatory to specify the field 'OwnedByTeam' too."

    #region Test OwnedBySamAccountName is not combinable with OwnedBy or OwnedById
    if ($KeyValuePair.Keys -contains 'OwnedBySamAccountName') {
        $KeyValuePair.Keys | Where-Object {
            $_ -match '^OwnedBy$|^OwnedById$'
        } | ForEach-Object {
            throw "The field 'OwnedBySamAccountName' cannot be combined with the fields 'OwnedBy' or 'OwnedById'. Please use 'Get-CherwellSystemUserHC' that provides you wtih 'OwnedBy' and 'OwnedById' if you want to be specific."
        }
        if ($KeyValuePair.Keys -notcontains 'OwnedByTeam') {
            throw $ownedByTeamIsMissingError
        }
    }
    #endregion
    
    #region Test OwnedBy and OwnedById cannot be used independently 
    $missingProperty = @($KeyValuePair.Keys | 
        Where-Object { $_ -match '^OwnedBy$|^OwnedById$' }
    )
    if ($missingProperty.Count -eq 1) {
        throw "Both the fields 'OwnedBy' and 'OwnedById' need to be specified. Please use the field 'OwnedBySamAccountName' instead or use 'Get-CherwellSystemUserHC' that provides you wtih 'OwnedBy' and 'OwnedById' if you want to be specific."
    }
    if (($missingProperty.Count -eq 2) -and 
        ($KeyValuePair.Keys -notcontains 'OwnedByTeam') 
    ) {
        throw $ownedByTeamIsMissingError
    }
    #endregion

    #region Test RequesterSamAccountName is not combinable with CustomerRecID
    $requesterProperties = @($KeyValuePair.Keys | 
        Where-Object { $_ -match '^RequesterSamAccountName$|^CustomerRecID$' }
    )
    if ($requesterProperties.Count -eq 2) {
        throw "The field 'RequesterSamAccountName' cannot be combined with the field 'CustomerRecID'. Please use 'Get-CherwellCustomerHC' to obatain the 'CustomerRecID' or use the SamAccountName in the field 'RequesterSamAccountName'."
    }
    #endregion

    #region Test SubmitOnBehalfOfID is not combinable with SubmittedBySamAccountName
    $requesterProperties = @($KeyValuePair.Keys | 
        Where-Object { 
            $_ -match '^SubmitOnBehalfOfID$|^SubmittedBySamAccountName$' }
    )
    if ($requesterProperties.Count -eq 2) {
        throw "The field 'SubmittedBySamAccountName' cannot be combined with the field 'SubmitOnBehalfOfID'. Please use 'Get-CherwellCustomerHC' to obatain the 'SubmitOnBehalfOfID' or use the SamAccountName in the field 'SubmittedBySamAccountName'."
    }
    #endregion
}

Function Update-CherwellTicketHC {
    <#
    .SYNOPSIS
        Update one or multiple fields in an existing Cherwell ticket.

    .DESCRIPTION
        This function facilitates the editing of fields in a Cherwell ticket. 
        One or multiple fields can be edited at the same time. Only the fields 
        provided in the KeyValuePair parameter will be updated, other fields 
        will be left untouched/unchanged.

        When updating fields like 'Service', 'Category' and 'SubCategory' it is 
        required by the API to also set the 'Priority'. In cases like these the 
        API will throw an error informing you of the missing required 
        properties.

        Example of an API error:
        "The field Ticket.Priority must be filled in before the record can be 
        saved.".

        The fix would be to add the field Priority to the KeyValuePair:
        -KeyValuePair @{
            Priority    = 2
            Service     = 'CLIENT'
            Category    = 'CLIENT BIOS'
            SubCategory = 'Request Service'
        }

    .PARAMETER KeyValuePair
        One or more hash tables containing key value pair combinations of the 
        fields to update.

    .PARAMETER Ticket
        The number to identify the ticket in Cherwell. Multiple members are 
        allowed.

    .PARAMETER Environment
        The Cherwell system that needs to be addressed.

        All available systems are defined in the file 'Passwords.json' in the 
        module folder. Ex. Test, Prod, ...

    .EXAMPLE
        To close a ticket the only thing required is to updated the correct 
        Cherwell ticket fields.

        $systemUser = Get-CherwellSystemUserHC -Environment Test -Filter @{
            FieldName  = 'SamAccountName'
            Operator   = 'eq'
            FieldValue = 'cnorris'
        } -PassThru

        Update-CherwellTicketHC -Environment Test -Ticket 2020 -KeyValuePair @{
            OwnedBy          = $systemUser.busObPublicId
            OwnedById        = $systemUser.busObRecId
            OwnedByTeam      = 'BNL'
            Status           = 'Resolved'
            CloseDescription = 'We fixed it!'
        }

    .EXAMPLE
        Change the CustomerRecID of a ticket, also known as the 
        RequesterSamAccountName field

        $CherwellCustomer = Get-CherwellCustomerHC -Environment Test -Filter @{
            FieldName  = 'SamAccountName'
            Operator   = 'eq'
            FieldValue = 'cnorris'
        } -PassThru

        Update-CherwellTicketHC -Environment Test -Ticket 20 -KeyValuePair @{
            CustomerRecID = $CherwellCustomer.busObRecId
        }

    .EXAMPLE
        Update the field 'Source' to the value 'Phone'

        Update-CherwellTicketHC -Environment Test -Ticket 10 -KeyValuePair @{
            Source = 'Phone'
        }

    .EXAMPLE
        Update the field CustomerRecID also knowns as the field 
        'RequesterSamAccountName' to the customer with SamAccountName 'idrielba'

        Update-CherwellTicketHC -Environment Test -Ticket 30 -KeyValuePair @{
            RequesterSamAccountName = 'idrielba'
            ServiceCountryCode      = 'BNL'
        }

    .EXAMPLE
        Change the 'Priority' and the 'ShortDescription' for 2 tickets

        Update-CherwellTicketHC -Environment Test -Ticket 6, 7 -KeyValuePair @{
            ShortDescription = 'HIHG PRIO TICKET'
            Priority         = 1
        }

    .EXAMPLE
        $PrinterTickets = @{
            # TicketNr = PrinterName
            '133101' = 'Printer1'
            '133102' = 'Printer2'
            '133103' = 'Printer3'
        }

        $PrinterTickets.GetEnumerator().ForEach({
            Update-CherwellTicketHC -Environment Test -Ticket $_.Key -KeyValuePair @{
                ShortDescription = "Printer $($_.Value)"
                Description      = 'Please replace the printer.'
            }
        })

        $Params = @{
            Environment  = 'Test'
            Ticket       = $PrinterTickets.Keys
            Type         = 'Specifics'
            KeyValuePair = @{
                Notes = 'Please follow KB-3565'
            }
        }
        Set-CherwellTicketDetailHC @Params

        First a hash table is created with the ticket number and the printer 
        name. Then the 'ShortDescription' is set to the printer name for each 
        ticket and the 'Description' is set the same for all 3 tickets. As a 
        last step the 'Additional details' field is updated for all tickets 
        with a reference to the correct SOP.
#>

    [OutputType()]
    [CmdLetBinding()]
    Param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [Alias('TicketNr')]
        [Object[]]$Ticket,
        [Parameter(Mandatory)]
        [HashTable]$KeyValuePair
    )
    DynamicParam {
        New-DynamicParameterHC -Name Environment
    }

    Begin {
        Try {
            $Environment = $PSBoundParameters.Environment
            Select-EnvironmentHC -Name $Environment

            $IncidentSchema = Get-SchemaHC -Name 'Incident'
        }
        Catch {
            throw "Failed updating ticket '$Ticket' in Cherwell '$Environment' for $($KeyValuePair | Out-String):  $_"
        }
    }

    Process {
        Try {
            Write-Verbose "Update ticket in environment '$Environment'"

            #region Check if tickets exist
            Try {
                $TicketDetails = foreach ($T in $Ticket) {
                    ConvertTo-TicketObjectHC -Item $T
                }
            }
            Catch {
                throw "Ticket number '$T' not found in Cherwell."
            }
            #endregion

            Test-InvalidPropertyCombinationHC -KeyValuePair $KeyValuePair

            #region Create hashtable with fields to update
            $UpdateFields = @{ }

            foreach ($K in $KeyValuePair.GetEnumerator()) {
                switch ($K.Key) {
                    'RequesterSamAccountName' {
                        $RequesterSamAccountName = $K.Value

                        $CherwellCustomer = @(Get-CherwellCustomerHC -Environment $Environment -Filter @{
                                FieldName  = 'SamAccountName'
                                Operator   = 'eq'
                                FieldValue = $RequesterSamAccountName
                            } -PassThru)

                        if ($CherwellCustomer.Count -eq 1) {
                            $UpdateFields['CustomerRecID'] = $CherwellCustomer[0].busObRecId
                            Break
                        }
                        elseif ($CherwellCustomer.Count -eq 0) {
                            throw "No customer found for field 'RequesterSamAccountName' with SamAccountName '$RequesterSamAccountName'"
                        }
                        else {
                            throw "Multiple customers found for field 'RequesterSamAccountName' with the same SamAccountName '$RequesterSamAccountName'.Please use the field 'CustomerRecID' as retrieved by 'Get-CherwellCustomerHC' to manually find the correct user."
                        }
                    }
                    'SubmittedBySamAccountName' {
                        $SubmittedBySamAccountName = $K.Value

                        $CherwellCustomer = @(Get-CherwellCustomerHC -Environment $Environment -Filter @{
                                FieldName  = 'SamAccountName'
                                Operator   = 'eq'
                                FieldValue = $SubmittedBySamAccountName
                            } -PassThru)

                        if ($CherwellCustomer.Count -eq 1) {
                            $UpdateFields['SubmitOnBehalfOfID'] = $CherwellCustomer[0].busObRecId
                            Break
                        }
                        elseif ($CherwellCustomer.Count -eq 0) {
                            throw "No customer found for field 'SubmittedBySamAccountName' with SamAccountName '$SubmittedBySamAccountName'.Please use the field 'SubmitOnBehalfOfID' as retrieved by 'Get-CherwellCustomerHC' to manually find the correct user."
                        }
                        else {
                            throw "Multiple customers found for field 'SubmittedBySamAccountName' with the same SamAccountName '$SubmittedBySamAccountName'.Please use the field 'SubmitOnBehalfOfID' as retrieved by 'Get-CherwellCustomerHC' to manually find the correct user."
                        }
                    }
                    'OwnedBySamAccountName' {
                        $OwnedBySamAccountName = $K.Value

                        $CherwellUser = @(Get-CherwellSystemUserHC -Environment $Environment -Filter @{
                                FieldName  = 'SamAccountName'
                                Operator   = 'eq'
                                FieldValue = $OwnedBySamAccountName
                            }  -PassThru)

                        if ($CherwellUser.Count -eq 0) {
                            throw "No system user found with SamAccountName '$OwnedBySamAccountName'. Please use the fields 'OwnedBy' and 'OwnedById' as retrieved by 'Get-CherwellSystemUserHC' to manually find the correct user."
                        }

                        if ($CherwellUser.Count -ge 2) {
                            throw "Multiple system users found for field 'OwnedBy' with the same SamAccountName '$OwnedBySamAccountName'. Please use the fields 'OwnedBy' and 'OwnedById' as retrieved by 'Get-CherwellSystemUserHC' to manually find the correct user."
                        }

                        $UpdateFields['OwnedBy'] = $CherwellUser[0].fields |
                        Where-Object Name -EQ 'FullName' | Select-Object -ExpandProperty Value
                        $UpdateFields['OwnedByID'] = $CherwellUser[0].BusObRecId
                        Break
                    }
                    Default {
                        $UpdateFields[$_] = $K.Value
                    }
                }
            }
            #endregion

            #region Create a new incident template
            $DirtyTemplate = New-TemplateHC -Name (Get-TemplateHC -Name 'Incident') -KeyValuePair $UpdateFields
            #endregion

            #region Update each ticket
            foreach ($T in $TicketDetails) {
                Try {
                    Write-Verbose "Update ticket '$($T.busObPublicId)'"

                    $Params = @{
                        Uri  = ($Uri + 'api/V1/SaveBusinessObject')
                        Body = [System.Text.Encoding]::UTF8.GetBytes((@{
                                    busObId       = $IncidentSchema.busObId
                                    busObPublicId = $T.busObPublicId
                                    busObRecId    = $T.busObRecId
                                    fields        = $DirtyTemplate.Fields
                                    persist       = $true
                                } | ConvertTo-Json))
                    }
                    $createBOResponse = Invoke-RestMethod @PostParams @Params

                    if (-not $createBOResponse) {
                        throw 'No answer received from the API.'
                    }

                    if ($createBOResponse.hasError) {
                        throw "API error code '$($createBOResponse.errorCode)' with message '$($createBOResponse.errorMessage)'"
                    }

                    Write-Verbose 'Ticket updated'
                }
                Catch {
                    $M = Convert-ApiErrorHC $_
                    Write-Error "Failed updating ticket '$($T.busObPublicId)': $M"
                }
            }
            #endregion
        }
        Catch {
            throw "Failed updating ticket '$Ticket' in Cherwell '$Environment' for $($KeyValuePair | Out-String): $_"
        }
    }
}