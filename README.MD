
# Description

A PowerShell module designed to easily interact with the Cherwell REST API. It provides CmdLet's for the most common CRUD operations like `New-CherwellTicketHC`, `Get-CherwellTicketHC`, `Update-CherwellTicketHC`, ... .

# Examples

## Create a new ticket

### Using convenience arguments

Using the convenience arguments, the ones ending with `*SamAccountName` is the quickest way
to create a new ticket. The module will try to find a single matching system user/customer
automatically. 

```powershell
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
```

### Using Cherwell fields only

Using known Cherwell fields only allows more control in case of errors. 

```powershell
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
```

## Update a ticket

### Close a ticket

To close a ticket the only thing required is to update the correct Cherwell ticket fields.

```powershell
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
```

## Retrieve configuration items (CI)

### Retrieve the CI types

As a first step you need to now which CI types are available in your Cherwell environment.

```powershell
$allCiTypes =  Get-CherwellConfigItemTypeHC -Environment Stage
```

### Retrieve all CI's of a specific type

Now that we know what the types are, let's retrieve all CI's for the first CI type.

```powershell
Get-CherwellConfigItemHC -Environment Stage -Type $allCiTypes[0] -Verbose
```

## Retrieve Cherwell Teams

Retrieve all teams that start with the value 'BNL' in the property 'Name'.

```powershell
Get-CherwellTeamInfoHC -Environment Prod -Filter @{
    FieldName  = 'Name'
    Operator   = 'startswith'
    FieldValue = 'BNL'
}
```

# Installation

## Setup guide

1. Clone the repository into the folder `Toolbox.Cherwell`
2. Rename `Passwords-example.json` to `Passwords.json` and update the file with your environment details
3. Optionally move the module folder to one of your module folders defined in `$env:PSModulePath`

## Using the module

- Run in any script `Import-Module $modulePath` and use the CmdLet's
  (only needed when the module is not in your default module folder `$env:PSModulePath`)
- List all available CmdLet's `Get-Command -Module Toolbox.Cherwell`
- Read the help `Get-Help New-CherwellTicketHC`
- See the examples `Get-Help New-CherwellTicketHC -Examples`

# Running tests

Add your config to the beginning of the file `Toolbox.Cherwell.Tests.ps1` (overwrite the defaults as they will probably not work in our environment)

```powershell
Install-Module -Name 'Pester'
Invoke-Pester $myPesterTestFile -Output 'Detailed'
```

# Release notes version 3.5 (2020/10/08)

## Updates

- Made tests compliant with Pester 5
- Added tests for invalid arguments within `KeyValuePair`
- Improved parameter naming
  - all parameters within `KeyValuePair` are native Cherwell arguments except for the ones ending with `*SamAccountName`, these are custom convenience parameters where the `SamAccountName` is matched (by `Get-CherwellSystemUserHC` or `Get-CherwellCustomerHC`) with the correct user/customer within Cherwell:

    old name > new name > Cherwell matched field

  - `SubmittedBy` > `SubmittedBySamAccountName` > `SubmitOnBehalfOfID`
  - `OwnedBy` > `OwnedBySamAccountName` > `OwnedByID` and `OwnedBy`
  - `Requester` > `RequesterSamAccountName` > `CustomerRecID`
- Added more examples

## Breaking changes

- `KeyValuePair` arguments
  - Renamed `SubmittedBy`, `OwnedBy` and `Requester`
  - When multiple users/customers are found for one of the `*SamAccountName` parameters an error is thrown instead of selecting the first random result
