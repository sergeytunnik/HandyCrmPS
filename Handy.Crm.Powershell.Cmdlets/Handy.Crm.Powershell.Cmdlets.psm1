﻿# Enabling TLS 1.2
# More details: https://blogs.msdn.microsoft.com/crm/2017/09/28/updates-coming-to-dynamics-365-customer-engagement-connection-security/
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12

function Assert-CRMOrganizationResponse {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true,
            ValueFromPipeline=$true)]
        [Microsoft.Xrm.Sdk.Messages.ExecuteMultipleResponse]$Response
    )

    Begin {}
    Process {
        if ($Response.IsFaulted -eq $true) {
            Write-Verbose -Message "ExecuteMultiple finished with faults"

            $message = "ExecuteMultpleResponse finished with fault."
            foreach ($r in $Response.Responses) {
                if ($null -ne $r.Fault) {
                    $message += "`r`n$($r.RequestIndex): $($r.Fault.Message)"
                }
            }
            throw $message
        }
        else {
            Write-Verbose -Message "ExecuteMultiple finished without faults"
        }
    }
    End {}
}


function Get-CRMOptionSetValue {
    [CmdletBinding()]
    [OutputType([Microsoft.Xrm.Sdk.OptionSetValue])]
    Param(
        [Parameter(Mandatory=$true)]
        [int]$Value
    )

    Begin {}
    Process {
        $optionSetValue = New-Object -TypeName 'Microsoft.Xrm.Sdk.OptionSetValue' -ArgumentList $Value

        $optionSetValue
    }
    End {}
}


function Get-CRMEntityMetadata {
    [CmdletBinding()]
    [OutputType([Microsoft.Xrm.Sdk.Metadata.EntityMetadata])]
    Param(
        [Parameter(Mandatory=$true)]
        [Microsoft.Xrm.Sdk.IOrganizationService]$Connection,

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$LogicalName,

        [Parameter(Mandatory=$false)]
        [Microsoft.Xrm.Sdk.Metadata.EntityFilters]$EntityFilters = [Microsoft.Xrm.Sdk.Metadata.EntityFilters]::Default,

        [Parameter(Mandatory=$false)]
        [switch]$RetrieveAsIfPublished
    )

    Begin {}
    Process {
        $parameters = @{}

        $parameters['EntityFilters'] = $EntityFilters
        $parameters['RetrieveAsIfPublished'] = $RetrieveAsIfPublished.IsPresent

        if($PSBoundParameters.ContainsKey('LogicalName')) {
            $parameters['LogicalName'] = $LogicalName
            $parameters['MetadataId'] = [guid]::Empty

            $response = Invoke-CRMOrganizationRequest -Connection $Connection -RequestName 'RetrieveEntity' -Parameters $parameters
        }
        else {
            $response = Invoke-CRMOrganizationRequest -Connection $Connection -RequestName 'RetrieveAllEntities' -Parameters $parameters
        }

        $response['EntityMetadata']
    }
    End {}
}


function Get-CRMEntityReference {
    [CmdletBinding()]
    [OutputType([Microsoft.Xrm.Sdk.EntityReference])]
    Param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$EntityName,

        [Parameter(Mandatory=$true)]
        [guid]$Id
    )

    Begin {}
    Process {
        $entityReference = New-Object -TypeName 'Microsoft.Xrm.Sdk.EntityReference' -ArgumentList $EntityName, $Id

        $entityReference
    }
    End {}
}


function Merge-CRMAttribute {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [hashtable]$From,

        [Parameter(Mandatory=$true)]
        [hashtable]$To
    )

    foreach ($key in $From.Keys)
    {
        $To[$key] = $From[$key]
    }
}


function Add-CRMPrivilegesRole {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [Microsoft.Xrm.Sdk.IOrganizationService]$Connection,

        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [guid]$RoleId,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [Microsoft.Crm.Sdk.Messages.RolePrivilege[]]$Privileges
    )

    $parameters = @{
        RoleId = $RoleId;
        Privileges = $Privileges
    }

    $response = Invoke-CRMOrganizationRequest -Connection $Connection -RequestName 'AddPrivilegesRole' -Parameters $parameters
    # $response
}


function Remove-CRMPrivilegeRole {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [Microsoft.Xrm.Sdk.IOrganizationService]$Connection,

        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [guid]$RoleId,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [guid]$PrivilegeId
    )

    $parameters = @{
        RoleId = $RoleId;
        PrivilegeId = $PrivilegeId
    }

    $response = Invoke-CRMOrganizationRequest -Connection $Connection -RequestName 'RemovePrivilegeRole' -Parameters $parameters
    # $response
}


function Get-CRMPrivilege {
    [CmdletBinding()]
    [OutputType([Microsoft.Xrm.Sdk.Entity])]
    Param(
        [Parameter(Mandatory=$true)]
        [Microsoft.Xrm.Sdk.IOrganizationService]$Connection,

        [Parameter(Mandatory=$true,
            ParameterSetName='Id')]
        [ValidateNotNull()]
        [guid]$Id,

        [Parameter(Mandatory=$true,
            ParameterSetName='Name')]
        [ValidateNotNullOrEmpty()]
        [string]$Name
    )

    $fetchXml = @"
<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="true" count="1">
  <entity name="privilege">
    <all-attributes />
    <filter type="and">
      {0}
    </filter>
  </entity>
</fetch>
"@

    switch ($PSCmdlet.ParameterSetName) {
        'Id' {
            $condition = '<condition attribute="privilegeid" operator="eq" value="{0}" />' -f $Id
        }

        'Name' {
            $condition = '<condition attribute="name" operator="eq" value="{0}" />' -f $Name
        }
    }

    $privilege = (Get-CRMEntity -Connection $Connection -FetchXml ($fetchXml -f $condition)) | Select-Object -Index 0
    $privilege
}


function Get-CRMRole {
    <#
    .DESCRIPTION
    Without All switch it returns only root role, which is usually used for editing roles.
    With All switch present it returns all roles it can find in organization.

    .LINK
    Get-CRMConnection
    #>
    [CmdletBinding()]
    [OutputType([Microsoft.Xrm.Sdk.Entity])]
    Param(
        [Parameter(Mandatory=$true)]
        [Microsoft.Xrm.Sdk.IOrganizationService]$Connection,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Name,

        [Parameter(Mandatory=$false)]
        [switch]$All
    )

    $fetchXml = @"
<fetch version="1.0" output-format="xml-platform" mapping="logical">
  <entity name="role">
    <all-attributes />
    <filter type="and">
      <condition attribute="name" operator="eq" value="{0}" />
      <condition attribute="parentroleid" operator="null" />
    </filter>
  </entity>
</fetch>
"@

    if ($All) {
        $fetchXml = $fetchXml.Replace('<condition attribute="parentroleid" operator="null" />', '')
    }

    $roles = Get-CRMEntity -Connection $Connection -FetchXml ($fetchXml -f $Name)

    $roles
}


function Get-CRMRolePrivilege {
    [CmdletBinding()]
    [OutputType([Microsoft.Crm.Sdk.Messages.RolePrivilege])]
    Param(
        [Parameter(Mandatory=$true)]
        [Microsoft.Crm.Sdk.Messages.PrivilegeDepth]$Depth,

        [Parameter(Mandatory=$true)]
        [guid]$Id

    )

    Begin {}
    Process {
        $rolePrivilege = New-Object -TypeName 'Microsoft.Crm.Sdk.Messages.RolePrivilege' -ArgumentList $Depth, $Id

        $rolePrivilege
    }
    End {}
}


function Get-CRMSolution {
    [CmdletBinding()]
    [OutputType([Microsoft.Xrm.Sdk.Entity])]
    Param(
        [Parameter(Mandatory=$true)]
        [Microsoft.Xrm.Sdk.IOrganizationService]$Connection,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Name
    )

    $fetchXml = @"
<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="true" count="1">
  <entity name="solution">
    <all-attributes />
    <filter type="and">
      <condition attribute="uniquename" operator="eq" value="{0}" />
    </filter>
  </entity>
</fetch>
"@

    $solution = (Get-CRMEntity -Connection $Connection -FetchXml ($fetchXml -f $Name)) | Select-Object -Index 0
    $solution
}


function Set-CRMSDKStepState {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [Microsoft.Xrm.Sdk.IOrganizationService]$Connection,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [guid]$SolutionId,

        [Parameter(Mandatory=$true)]
        [ValidateSet('Enabled', 'Disabled')]
        [string]$State,

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$Include = [string]::Empty,

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$Exclude = [string]::Empty
    )

    $fetchXml = @"
<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false">
  <entity name="sdkmessageprocessingstep">
    <all-attributes />
    <filter type="and">
      <condition attribute="solutionid" operator="eq" value="{0}" />
      {1}
      {2}
    </filter>
  </entity>
</fetch>
"@

    if ($Include -ne [string]::Empty) {
        $includeCondition = "<condition attribute=`"name`" operator=`"like`" value=`"$Include`" />"
    }
    else {
        $includeCondition = [string]::Empty
    }

    if ($Exclude -ne [string]::Empty) {
        $excludeCondition = "<condition attribute=`"name`" operator=`"not-like`" value=`"$Exclude`" />"
    }
    else {
        $excludeCondition = [string]::Empty
    }

    Write-Verbose -Message "FetchXML:"
    Write-Verbose -Message ($fetchXml -f $SolutionId, $includeCondition, $excludeCondition)

    $steps = Get-CRMEntity -Connection $Connection -FetchXml ($fetchXml -f $SolutionId, $includeCondition, $excludeCondition)

    if (($null -eq $steps) -or ($steps.Count -eq 0)) {
        Write-Warning -Message "Found nothing to update"
        return
    }

    Write-Verbose -Message "Found $($steps.Count) steps"
    Write-Verbose -Message "Updating them"

    switch ($State) {
        'Enabled' {
            $response = Set-CRMState -Connection $Connection -Entity $steps -State 0 -Status 1
        }

        'Disabled' {
            $response = Set-CRMState -Connection $Connection -Entity $steps -State 1 -Status 2
        }
    }

    $response | Assert-CRMOrganizationResponse
}


function Get-CRMBusinessUnit {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [Microsoft.Xrm.Sdk.IOrganizationService]$Connection,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Name
    )

    $fetchXml = @"
<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="true">
  <entity name="businessunit">
    <all-attributes />
    <filter type="and">
      <condition attribute="name" operator="eq" value="{0}"/>
    </filter>
  </entity>
</fetch>
"@

    $bu = Get-CRMEntity -Connection $Connection -FetchXml ($fetchXml -f $Name)

    $bu
}


function New-CRMBusinessUnit {
    [CmdletBinding()]
    [OutputType([System.Guid])]
    Param(
        [Parameter(Mandatory=$true)]
        [Microsoft.Xrm.Sdk.IOrganizationService]$Connection,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Name,

        [Parameter(Mandatory=$false)]
        [ValidateNotNull()]
        [guid]$ParentBusinessUnitId = [guid]::Empty
    )

    if ($ParentBusinessUnitId -eq [guid]::Empty) {
        Write-Verbose -Message "Getting root BU"

        $fetchXml = @"
<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="true" count="1">
  <entity name="businessunit">
    <all-attributes />
    <filter type="and">
      <condition attribute="parentbusinessunitid" operator="null" />
    </filter>
  </entity>
</fetch>
"@

        $bu = (Get-CRMEntity -Connection $Connection -FetchXml $fetchXml) | Select-Object -Index 0
        $ParentBusinessUnitId = $bu.Id
    }

    $buAttributes = @{}
    $buAttributes['name'] = $Name
    $buAttributes['parentbusinessunitid'] = Get-CRMEntityReference -EntityName 'businessunit' -Id $ParentBusinessUnitId
    
    $resp = New-CRMEntity -Connection $Connection -EntityName 'businessunit' -Attributes $buAttributes -ReturnResponses

    $resp | Assert-CRMOrganizationResponse

    $resp.Responses[0].Response.id
}


function New-CRMUser {
    [CmdletBinding()]
    [OutputType([System.Guid])]
    Param(
        [Parameter(Mandatory=$true)]
        [Microsoft.Xrm.Sdk.IOrganizationService]$Connection,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$DomainName,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$FirstName,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$LastName,

        [Parameter(Mandatory=$true)]
        [guid]$BusinessUnitId,

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [hashtable]$AdditionAttributes = @{}
    )

    $userAttributes = @{}
    $userAttributes['domainname'] = $DomainName
    $userAttributes['firstname'] = $FirstName
    $userAttributes['lastname'] = $LastName
    $userAttributes['fullname'] = "$($FirstName) $($LastName)"
    $userAttributes['businessunitid'] = Get-CRMEntityReference -EntityName 'businessunit' -Id $BusinessUnitId

    Merge-CRMAttribute -From $AdditionAttributes -To $userAttributes

    $resp = New-CRMEntity -Connection $Connection -EntityName 'systemuser' -Attributes $userAttributes -ReturnResponses

    $resp | Assert-CRMOrganizationResponse

    $resp.Responses[0].Response.id
}


function Get-CRMDuplicateRule {
    [CmdletBinding()]
    [OutputType([Microsoft.Xrm.Sdk.Entity])]
    Param(
        [Parameter(Mandatory=$true)]
        [Microsoft.Xrm.Sdk.IOrganizationService]$Connection,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Name,

        [Parameter(Mandatory=$false)]
        [int]$StateCode = 0,

        [Parameter(Mandatory=$false)]
        [int]$StatusCode = 0
    )
    # statecode	statecodename	statuscode	statuscodename
    # 0         Inactive        0           Unpublished
    $fetchXml = @"
<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="true">
  <entity name="duplicaterule">
    <all-attributes />
    <filter type="and">
      <condition attribute="name" operator="eq" value="{0}" />
      <condition attribute="statecode" operator="eq" value="{1}" />
      <condition attribute="statuscode" operator="eq" value="{2}" />
    </filter>
  </entity>
</fetch>
"@

    $duplicateRule = (Get-CRMEntity -Connection $Connection -FetchXml ($fetchXml -f $Name, $StateCode, $StatusCode)) | Select-Object -Index 0

    $duplicateRule
}


function New-CRMQueue {
    [CmdletBinding()]
    [OutputType([System.Guid])]
    Param(
        [Parameter(Mandatory=$true)]
        [Microsoft.Xrm.Sdk.IOrganizationService]$Connection,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Name,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [guid]$OwnerId,

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$Email = [string]::Empty
    )

    $queueAttributes = @{}
    $queueAttributes['name'] = $Name
    $queueAttributes['ownerid'] = Get-CRMEntityReference -EntityName 'systemuser' -Id $OwnerId
    $queueAttributes['outgoingemaildeliverymethod'] = Get-CRMOptionSetValue -Value 2 # Email Router

    if ($Email -ne [string]::Empty) {
        $queueAttributes['emailaddress'] = $Email
    }

    $resp = New-CRMEntity -Connection $Connection -EntityName 'queue' -Attributes $queueAttributes -ReturnResponses

    $queueId = $resp.Responses[0].Response.id

    if ($Email -ne [string]::Empty) {
        $queueEntity = Get-CRMEntityById -Connection $Connection -EntityName 'queue' -Id $queueId -Columns 'emailrouteraccessapproval'
        $queueEntity['emailrouteraccessapproval'] = Get-CRMOptionSetValue -Value 1 # Approved
        $resp = Update-CRMEntity -Connection $Connection -Entity $queueEntity
        $resp | Assert-CRMOrganizationResponse
    }

    $queueId
}


function Set-CRMQueueForUser {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [Microsoft.Xrm.Sdk.IOrganizationService]$Connection,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [guid]$UserId,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [guid]$QueueId
    )

    $userEntity = Get-CRMEntityById -Connection $Connection -EntityName 'systemuser' -Id $UserId -Columns 'queueid'

    $userEntity['queueid'] = Get-CRMEntityReference -EntityName 'queue' -Id $QueueId

    $resp = Update-CRMEntity -Connection $Connection -Entity $userEntity -ReturnResponses

    $resp | Assert-CRMOrganizationResponse
}


function Add-CRMRoleForUser {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [Microsoft.Xrm.Sdk.IOrganizationService]$Connection,

        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [Microsoft.Xrm.Sdk.Entity]$User,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$RoleName
    )

    $fetchXml = @"
<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="true" count="1">
  <entity name="role">
    <all-attributes />
    <filter type="and">
      <condition attribute="name" operator="eq" value="{0}" />
      <condition attribute="businessunitid" operator="eq" value="{1}" />
    </filter>
  </entity>
</fetch>
"@

    $role = (Get-CRMEntity -Connection $Connection -FetchXml ($fetchXml -f $RoleName, $User['businessunitid'].Id)) | Select-Object -Index 0

    if ($null -eq $role) {
        throw "Couldn't find role $($RoleName) or something went wrong."
    }

    $reference = Get-CRMEntityReference -EntityName $role.LogicalName -Id $role.Id

    Add-CRMAssociation -Connection $Connection -EntityName $User.LogicalName -Id $User.Id -Relationship 'systemuserroles_association' -RelatedEntity $reference
}


function Remove-CRMRoleForUser {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [Microsoft.Xrm.Sdk.IOrganizationService]$Connection,

        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [Microsoft.Xrm.Sdk.Entity]$User,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$RoleName
    )

    $fetchXml = @"
<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="true" count="1">
  <entity name="role">
    <all-attributes />
    <filter type="and">
      <condition attribute="name" operator="eq" value="{0}" />
      <condition attribute="businessunitid" operator="eq" value="{1}" />
    </filter>
  </entity>
</fetch>
"@

    $role = (Get-CRMEntity -Connection $Connection -FetchXml ($fetchXml -f $RoleName, $User['businessunitid'].Id)) | Select-Object -Index 0

    if ($null -eq $role) {
        throw "Couldn't find role $($RoleName) or something went wrong."
    }

    $reference = Get-CRMEntityReference -EntityName $role.LogicalName -Id $role.Id

    Remove-CRMAssociation -Connection $Connection -EntityName $User.LogicalName -Id $User.Id -Relationship 'systemuserroles_association' -RelatedEntity $reference
}


function Get-CRMTransactionCurrency {
    [CmdletBinding()]
    [OutputType([Microsoft.Xrm.Sdk.Entity])]
    Param(
        [Parameter(Mandatory=$true)]
        [Microsoft.Xrm.Sdk.IOrganizationService]$Connection,

        [Parameter(Mandatory=$true)]
        [CurrencyCodeEnum]$CurrencyCode
    )

    $fetchXml = @"
<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="true" count="1">
  <entity name="transactioncurrency">
    <all-attributes />
    <filter type="and">
      <condition attribute="isocurrencycode" operator="eq" value="{0}" />
    </filter>
  </entity>
</fetch>
"@

    $tc = (Get-CRMEntity -Connection $Connection -FetchXml ($fetchXml -f $CurrencyCode)) | Select-Object -Index 0

    $tc
}


function New-CRMTransactionCurrency {
    [CmdletBinding()]
    [OutputType([System.Guid])]
    Param(
        [Parameter(Mandatory=$true)]
        [Microsoft.Xrm.Sdk.IOrganizationService]$Connection,
     
        [Parameter(Mandatory=$true)]
        [string]$CurrencyName,

        [Parameter(Mandatory=$true)]
        [string]$IsoCurrencyCode,

        [Parameter(Mandatory=$true)]
        [string]$CurrencySymbol,

        [Parameter(Mandatory=$false)]
        [int]$CurrencyPrecision = 2,

        [Parameter(Mandatory=$false)]
        [decimal]$ExchangeRate = [decimal]1.0,

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [hashtable]$AdditionAttributes = @{}
    )

    $currencyAttributes = @{}
    $currencyAttributes['currencyprecision'] = $CurrencyPrecision
    $currencyAttributes['exchangerate'] = $ExchangeRate
    $currencyAttributes['isocurrencycode'] = $IsoCurrencyCode

    $currencyAttributes['currencyname'] = $CurrencyName
    $currencyAttributes['currencysymbol'] = $CurrencySymbol

    Merge-CRMAttribute -From $AdditionAttributes -To $currencyAttributes

    $resp = New-CRMEntity -Connection $Connection -EntityName 'transactioncurrency' -Attributes $currencyAttributes -ReturnResponses

    $resp | Assert-CRMOrganizationResponse

    $resp.Responses[0].Response.id
}


function Set-CRMSDKStepMode {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [Microsoft.Xrm.Sdk.IOrganizationService]$Connection,

        [Parameter(Mandatory=$true)]
        [guid]$Id,

        [Parameter(Mandatory=$true)]
        [ValidateSet('Asynchronous', 'Synchronous')]
        [string]$Mode,

        [Parameter(Mandatory=$false)]
        [switch]$SetAutoDelete
    )

    $step = Get-CRMEntityById -Connection $Connection -EntityName 'sdkmessageprocessingstep' -Id $Id -Columns 'mode', 'asyncautodelete'
    # Sync - 0, Async - 1
    # Yes - 1, No - 0

    switch ($Mode) {
        'Asynchronous' {
            $step['mode'] = Get-CRMOptionSetValue -Value 1
            $step['asyncautodelete'] = $SetAutoDelete.IsPresent
            break
        }

        'Synchronous' {
            $step['mode'] = Get-CRMOptionSetValue -Value 0
            $step['asyncautodelete'] = $false
            break
        }
    }

    $resp = Update-CRMEntity -Connection $Connection -Entity $step
    $resp | Assert-CRMOrganizationResponse
}


function Enable-CRMWorkflow {
    <#
    .SYNOPSIS
    Activates workflow.

    .NOTES
    Only owner of workflow can change its state.
    There may be several workflows with same name. This cmdlet will perfom action on all of them.

    .LINK
    Get-CRMConnection
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [Microsoft.Xrm.Sdk.IOrganizationService]$Connection,

        [Parameter(Mandatory=$true,
            ParameterSetName='Name')]
        [ValidateNotNullOrEmpty()]
        [string]$Name,

        [Parameter(Mandatory=$true,
            ParameterSetName='SolutionId')]
        [ValidateNotNull()]
        [guid]$SolutionId,

        [Parameter(Mandatory=$true,
            ParameterSetName='Id')]
        [ValidateNotNull()]
        [guid]$Id
    )


    switch ($PSCmdlet.ParameterSetName) {
        'Id' {
            $fetchXmlWorkflow = @"
<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false">
  <entity name="workflow">
    <filter type="and">
      <condition attribute="workflowid" operator="eq" value="{0}" />
      <condition attribute="type" operator="eq" value="1" />
    </filter>
  </entity>
</fetch>
"@

            $workflows  = Get-CRMEntity -Connection $Connection -FetchXml ($fetchXmlWorkflow -f $Id)

            if ($workflows.Count -gt 0) {
                Write-Verbose -Message "Found $($workflows.Count) workflows with Id '$Id'"

                $response = Set-CRMState -Connection $Connection -Entity $workflows -State 1 -Status 2 -ContinueOnError

                $response | Assert-CRMOrganizationResponse
            }
            else {
                Write-Verbose -Message "No workflows with Id '$Id' were found"
            }
        }

        'Name' {
            $fetchXmlWorkflow = @"
<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false">
  <entity name="workflow">
    <filter type="and">
      <condition attribute="name" operator="eq" value="{0}" />
      <condition attribute="type" operator="eq" value="1" />
    </filter>
  </entity>
</fetch>
"@

            $workflows  = Get-CRMEntity -Connection $Connection -FetchXml ($fetchXmlWorkflow -f $Name)

            if ($workflows.Count -gt 0) {
                Write-Verbose -Message "Found $($workflows.Count) workflows with name '$Name'"

                $response = Set-CRMState -Connection $Connection -Entity $workflows -State 1 -Status 2 -ContinueOnError

                $response | Assert-CRMOrganizationResponse
            }
            else {
                Write-Verbose -Message "No workflows with name '$Name' were found"
            }
        }

        'SolutionId' {
            $fetchXmlWorkflow = @"
<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false">
  <entity name="workflow">
    <filter type="and">
      <condition attribute="solutionid" operator="eq" value="{0}" />
      <condition attribute="type" operator="eq" value="1" />
    </filter>
  </entity>
</fetch>
"@

            $workflows  = Get-CRMEntity -Connection $Connection -FetchXml ($fetchXmlWorkflow -f $SolutionId)

            if ($workflows.Count -gt 0)
            {
                Write-Verbose -Message "Found $($workflows.Count) workflows in solution with Id '$SolutionId'"

                $response = Set-CRMState -Connection $Connection -Entity $workflows -State 1 -Status 2 -ContinueOnError

                $response | Assert-CRMOrganizationResponse
            }
            else {
                Write-Verbose -Message "No workflows in solution with Id '$SolutionId' were found"
            }
        }
    }
}


function Disable-CRMWorkflow {
    <#
    .SYNOPSIS
    Deactivates workflow.

    .NOTES
    Only owner of workflow can change its state.
    There may be several workflows with same name. This cmdlet will perfom action on all of them.

    .LINK
    Get-CRMConnection
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [Microsoft.Xrm.Sdk.IOrganizationService]$Connection,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Name
    )

    $fetchXmlWorkflow = @"
<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false">
  <entity name="workflow">
    <filter type="and">
      <condition attribute="name" operator="eq" value="{0}" />
      <condition attribute="type" operator="eq" value="1" />
    </filter>
  </entity>
</fetch>
"@

    $workflows  = Get-CRMEntity -Connection $Connection -FetchXml ($fetchXmlWorkflow -f $Name)

    if ($workflows.Count -gt 0) {
        Write-Verbose -Message "Found $($workflows.Count) workflows with name '$Name'"

        $response = Set-CRMState -Connection $Connection -Entity $workflows -State 0 -Status 1 -ContinueOnError

        $response | Assert-CRMOrganizationResponse
    }
    else {
        Write-Verbose -Message "No workflows with name '$Name' were found"
        # Or throw?
    }
}


function Get-CRMSiteMap {
    [CmdletBinding()]
    [OutputType([string])]
    Param(
        [Parameter(Mandatory=$true)]
        [Microsoft.Xrm.Sdk.IOrganizationService]$Connection
    )
    
    $fetchXml = @"
<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false">
    <entity name="sitemap">            
    </entity>
</fetch>
"@

    $sitemap = (Get-CRMEntity -Connection $Connection -FetchXml $fetchXml) | Select-Object -Index 0

    $sitemap['sitemapxml']
}


function Set-CRMSiteMap {
    [CmdletBinding()]
    [OutputType([void])]
    Param(
        [Parameter(Mandatory=$true)]
        [Microsoft.Xrm.Sdk.IOrganizationService]$Connection,
        
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$SiteMapXml,
        
        [Parameter(Mandatory=$false)]
        [switch]$Publish
    )
    
    $fetchXml = @"
<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false">
    <entity name="sitemap">            
    </entity>
</fetch>
"@

    Write-Verbose -Message "Getting sitemap"
    $sitemap = (Get-CRMEntity -Connection $Connection -FetchXml $fetchXml) | Select-Object -Index 0
    $sitemap['sitemapxml'] = $SiteMapXml
    
    Write-Verbose -Message "Updating sitemap"
    Update-CRMEntity -Connection $Connection -Entity $sitemap | Assert-CRMOrganizationResponse
    
    if ($Publish) {
        Write-Verbose -Message "Publishing (only sitemap)"
        Publish-CRMXml -Connection $Connection -ParameterXml '<importexportxml><sitemaps><sitemap></sitemap></sitemaps></importexportxml>'
    }
}


function Publish-CRMXml {
    <#
    .SYNOPSIS
    Publishes only specific solution components.

    .LINK
    https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.publishxmlrequest.parameterxml.aspx
    Get-CRMConnection
    #>
    [CmdletBinding()]
    [OutputType([void])]
    Param(
        [Parameter(Mandatory=$true)]
        [Microsoft.Xrm.Sdk.IOrganizationService]$Connection,
        
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$ParameterXml
    )
    
    $parameters = @{}
    $parameters['ParameterXml'] = $ParameterXml
    
    $response = Invoke-CRMOrganizationRequest -Connection $Connection -RequestName 'PublishXml' -Parameters $parameters
}


function Get-CRMVersion {
    <#
    .SYNOPSIS
    Returns CRM version.
    #>
    [CmdletBinding()]
    [OutputType([System.Version])]
    Param(
        [Parameter(Mandatory=$true)]
        [Microsoft.Xrm.Sdk.IOrganizationService]$Connection
    )
    
    $response = Invoke-CRMOrganizationRequest -Connection $Connection -RequestName 'RetrieveVersion'
    
    $version = [System.Version]::Parse($response['Version'])
    
    $version
}


function Get-CRMRelationship {
    [CmdletBinding()]
    [OutputType([Microsoft.Xrm.Sdk.Metadata.RelationshipMetadataBase])]
    Param(
        [Parameter(Mandatory=$true)]
        [Microsoft.Xrm.Sdk.IOrganizationService]$Connection,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Name,

        [Parameter(Mandatory=$false)]
        [switch]$RetrieveAsIfPublished
    )

    $parameters = @{}

    $parameters['MetadataId'] = [guid]::Empty
    $parameters['RetrieveAsIfPublished'] = $RetrieveAsIfPublished.IsPresent
    $parameters['Name'] = $Name

    $response = Invoke-CRMOrganizationRequest -Connection $Connection -RequestName 'RetrieveRelationship' -Parameters $parameters

    $response['RelationshipMetadata']    
}


function Update-CRMRelationship {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [Microsoft.Xrm.Sdk.IOrganizationService]$Connection,

        [Parameter(Mandatory=$true)]
        [Microsoft.Xrm.Sdk.Metadata.RelationshipMetadataBase]$Relationship,

        [Parameter(Mandatory=$false)]
        [switch]$MergeLabels
    )

    $parameters = @{}

    $parameters['MergeLabels'] = $MergeLabels.IsPresent
    $parameters['Relationship'] = $Relationship

    $response = Invoke-CRMOrganizationRequest -Connection $Connection -RequestName 'UpdateRelationship' -Parameters $parameters
}


function Wait-CRMAsyncJob {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [Microsoft.Xrm.Sdk.IOrganizationService]$Connection,

        [Parameter(Mandatory=$true)]
        [guid]$AsyncJobId,

        [Parameter(Mandatory=$false)]
        [int]$CheckPeriod = 5
    )
    
    do {
        try {
            Write-Verbose -Message "Checking AsyncJob $($AsyncJobId)"
            $asyncjob = Get-CRMEntityById -Connection $Connection -EntityName 'asyncoperation' -Id $AsyncJobId -AllColumns

            Write-Verbose "AsyncJob $($AsyncJobId) state and status: $($asyncjob.FormattedValues['statecode']), $($asyncjob.FormattedValues['statuscode'])"
        }
        catch {
            if ($_.Exception.Message -eq 'SQL timeout expired.') {
                Write-Warning -Message $_.Exception.Message 
            } else {
                throw $_
            }
        }

        Start-Sleep -Seconds $CheckPeriod
    } until ($asyncjob['statecode'] -eq 3) # 3 - Completed

    Write-Verbose -Message "AsyncJob $($AsyncJobId) completed in $($asyncjob['executiontimespan']) seconds."

    if ($asyncjob['statuscode'] -eq 31) {
        Write-Warning -Message "AsyncJob $($AsyncJobId) failed"
    }
}


function Save-CRMFormattedImportJobResult {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [Microsoft.Xrm.Sdk.IOrganizationService]$Connection,

        [Parameter(Mandatory=$true)]
        [guid]$ImportJobId,

        [Parameter(Mandatory=$false)]
        [string]$FilePath
    )

    $parameters = @{'ImportJobId' = $ImportJobId}

    $response = Invoke-CRMOrganizationRequest -Connection $Connection -RequestName 'RetrieveFormattedImportJobResults' -Parameters $parameters

    Out-File -FilePath $FilePath -InputObject $response['FormattedResults']
}


function Get-CRMAuditDetail {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [Microsoft.Xrm.Sdk.IOrganizationService]$Connection,

        [Parameter(Mandatory=$true)]
        [guid]$AuditId
    )

    $parameters = @{'AuditId' = $AuditId}

    $response = Invoke-CRMOrganizationRequest -Connection $Connection -RequestName 'RetrieveAuditDetails' -Parameters $parameters

    $response['AuditDetail']
}


function Expand-CRMSolution {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ZipPath,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$Folder = '.',

        [Parameter(Mandatory = $false)]
        [Microsoft.Crm.Tools.SolutionPackager.SolutionPackageType]$PackageType = [Microsoft.Crm.Tools.SolutionPackager.SolutionPackageType]::Both,

        [Parameter(Mandatory = $false)]
        [ValidateSet('NONE', 'PLUGIN', 'WEBRESOURCE', 'WORKFLOW')]
        [string]$SingleComponent = $null,

        [Parameter(Mandatory = $false)]
        [Microsoft.Crm.Tools.SolutionPackager.AllowDelete]$AllowDeletes = [Microsoft.Crm.Tools.SolutionPackager.AllowDelete]::Yes,

        [Parameter(Mandatory = $false)]
        [Microsoft.Crm.Tools.SolutionPackager.AllowWrite]$AllowWrites = [Microsoft.Crm.Tools.SolutionPackager.AllowWrite]::Yes,

        [Parameter(Mandatory = $false)]
        [switch]$Localize,

        [Parameter(Mandatory = $false)]
        [string]$LocaleTemplate = $null,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$MappingFile = $null,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$LogFile = $null,

        [Parameter(Mandatory = $false)]
        [System.Diagnostics.TraceLevel]$ErrorLevel = [System.Diagnostics.TraceLevel]::Info
    )

    $packagerArguments = New-Object -TypeName 'Microsoft.Crm.Tools.SolutionPackager.PackagerArguments'
    $packagerArguments.Action = [Microsoft.Crm.Tools.SolutionPackager.CommandAction]::Extract
    $packagerArguments.PathToZipFile = $PSCmdlet.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ZipPath)
    $packagerArguments.PackageType = $PackageType
    $packagerArguments.Folder = $PSCmdlet.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Folder)
    $packagerArguments.SingleComponent = $SingleComponent
    $packagerArguments.AllowDeletes = $AllowDeletes
    $packagerArguments.AllowWrites = $AllowWrites
    $packagerArguments.Localize = $Localize
    $packagerArguments.LocaleTemplate = $LocaleTemplate
    if ($MappingFile) {
        $packagerArguments.MappingFile = $PSCmdlet.SessionState.Path.GetUnresolvedProviderPathFromPSPath($MappingFile)
    }
    if($LogFile) {
        $packagerArguments.LogFile = $PSCmdlet.SessionState.Path.GetUnresolvedProviderPathFromPSPath($LogFile)
    }
    $packagerArguments.ErrorLevel = $ErrorLevel

    $solutionPackager = New-Object -TypeName 'Microsoft.Crm.Tools.SolutionPackager.SolutionPackager' -ArgumentList $packagerArguments
    Write-Verbose -Message "Extracting: $($packagerArguments.PathToZipFile) -> $($packagerArguments.Folder)"
    $solutionPackager.Run()
}


function Compress-CRMSolution {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ZipPath,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$Folder = '.',

        [Parameter(Mandatory = $false)]
        [Microsoft.Crm.Tools.SolutionPackager.SolutionPackageType]$PackageType = [Microsoft.Crm.Tools.SolutionPackager.SolutionPackageType]::Both,

        [Parameter(Mandatory = $false)]
        [string]$SingleComponent = [string]::Empty,

        [Parameter(Mandatory = $false)]
        [Microsoft.Crm.Tools.SolutionPackager.AllowDelete]$AllowDeletes = [Microsoft.Crm.Tools.SolutionPackager.AllowDelete]::No,

        [Parameter(Mandatory = $false)]
        [Microsoft.Crm.Tools.SolutionPackager.AllowWrite]$AllowWrites = [Microsoft.Crm.Tools.SolutionPackager.AllowWrite]::Yes,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [switch]$Localize,

        [Parameter(Mandatory = $false)]
        [string]$LocaleTemplate = $null,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$MappingFile = $null,

        [Parameter(Mandatory = $false)]
        [string]$LogFile = $null,

        [Parameter(Mandatory = $false)]
        [System.Diagnostics.TraceLevel]$ErrorLevel = [System.Diagnostics.TraceLevel]::Info
    )

    $packagerArguments = New-Object -TypeName 'Microsoft.Crm.Tools.SolutionPackager.PackagerArguments'
    $packagerArguments.Action = [Microsoft.Crm.Tools.SolutionPackager.CommandAction]::Pack
    $packagerArguments.PathToZipFile = $PSCmdlet.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ZipPath)
    $packagerArguments.PackageType = $PackageType
    $packagerArguments.Folder = $PSCmdlet.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Folder)
    $packagerArguments.SingleComponent = $SingleComponent
    $packagerArguments.AllowDeletes = $AllowDeletes
    $packagerArguments.AllowWrites = $AllowWrites
    $packagerArguments.Localize = $Localize
    $packagerArguments.LocaleTemplate = $LocaleTemplate
    if ($MappingFile) {
        $packagerArguments.MappingFile = $PSCmdlet.SessionState.Path.GetUnresolvedProviderPathFromPSPath($MappingFile)
    }
    if($LogFile) {
        $packagerArguments.LogFile = $PSCmdlet.SessionState.Path.GetUnresolvedProviderPathFromPSPath($LogFile)
    }
    $packagerArguments.ErrorLevel = $ErrorLevel

    $solutionPackager = New-Object -TypeName 'Microsoft.Crm.Tools.SolutionPackager.SolutionPackager' -ArgumentList $packagerArguments
    Write-Verbose -Message "Packing: $($packagerArguments.Folder) -> $($packagerArguments.PathToZipFile)"
    $solutionPackager.Run()
}


function Get-CRMCurrentOrganization {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [Microsoft.Xrm.Sdk.IOrganizationService]$Connection
    )

    $parameters = @{}
    $parameters['AccessType'] = [Microsoft.Xrm.Sdk.Organization.EndpointAccessType]::Default
    
    $response = Invoke-CRMOrganizationRequest -Connection $Connection -RequestName 'RetrieveCurrentOrganization' -Parameters $parameters

    $response['Detail']
}


function Invoke-CRMCloneAsPatch {
    <#
    .SYNOPSIS
    Creates a solution patch from a managed or unmanaged solution

    .LINK
    https://msdn.microsoft.com/en-us/library/mt593040.aspx
    Get-CRMConnection
    #>
    [CmdletBinding()]
    [OutputType([guid])]
    Param(
        [Parameter(Mandatory = $true)]
        [Microsoft.Xrm.Sdk.IOrganizationService]$Connection,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ParentSolutionUniqueName,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$VersionNumber,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$DisplayName
    )

    $parameters = @{
        'ParentSolutionUniqueName' = $ParentSolutionUniqueName;
        'VersionNumber'            = $VersionNumber;
        'DisplayName'              = $DisplayName;
    }

    $response = Invoke-CRMOrganizationRequest -Connection $Connection -RequestName 'CloneAsPatch' -Parameters $parameters

    $response['SolutionId']
}


function Invoke-CRMCloneAsSolution {
    <#
    .SYNOPSIS
    Creates a new copy of an unmanaged solution that contains the original solution plus all of its patches.

    .LINK
    https://msdn.microsoft.com/en-us/library/mt593040.aspx
    Get-CRMConnection
    #>
    [CmdletBinding()]
    [OutputType([guid])]
    Param(
        [Parameter(Mandatory = $true)]
        [Microsoft.Xrm.Sdk.IOrganizationService]$Connection,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ParentSolutionUniqueName,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$VersionNumber,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$DisplayName
    )

    $parameters = @{
        'ParentSolutionUniqueName' = $ParentSolutionUniqueName;
        'VersionNumber'            = $VersionNumber;
        'DisplayName'              = $DisplayName;
    }

    $response = Invoke-CRMOrganizationRequest -Connection $Connection -RequestName 'CloneAsSolution' -Parameters $parameters

    $response['SolutionId']
}


function Invoke-CRMDeleteAndPromote {
    <#
    .SYNOPSIS
    Deletes the base solution along with all of its patches and renames the holding solution to the same name as the base solution.

    .LINK
    https://msdn.microsoft.com/en-us/library/mt593040.aspx
    Get-CRMConnection
    #>
    [CmdletBinding()]
    [OutputType([guid])]
    Param(
        [Parameter(Mandatory = $true)]
        [Microsoft.Xrm.Sdk.IOrganizationService]$Connection,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$UniqueName
    )

    $parameters = @{ 'UniqueName' = $UniqueName; }

    $response = Invoke-CRMOrganizationRequest -Connection $Connection -RequestName 'DeleteAndPromote' -Parameters $parameters

    $response['SolutionId']
}


Set-Alias -Name 'Activate-CRMWorkflow' -Value 'Enable-CRMWorkflow'
Set-Alias -Name 'Deactivate-CRMWorkflow' -Value 'Disable-CRMWorkflow'
