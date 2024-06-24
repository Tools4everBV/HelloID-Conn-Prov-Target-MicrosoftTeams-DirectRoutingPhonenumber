#################################################
# HelloID-Conn-Prov-Target-Teams-Voip-Create
# PowerShell V2
#################################################

# Enable TLS1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12

# Set debug logging
switch ($actionContext.Configuration.isDebug) {
    $true { $VerbosePreference = "Continue" }
    $false { $VerbosePreference = "SilentlyContinue" }
}

#region functions
function Resolve-MicrosoftGraphAPIError {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [object]
        $ErrorObject
    )
    process {
        $httpErrorObj = [PSCustomObject]@{
            ScriptLineNumber = $ErrorObject.InvocationInfo.ScriptLineNumber
            Line             = $ErrorObject.InvocationInfo.Line
            ErrorDetails     = $ErrorObject.Exception.Message
            FriendlyMessage  = $ErrorObject.Exception.Message
        }
        if (-not [string]::IsNullOrEmpty($ErrorObject.ErrorDetails.Message)) {
            $httpErrorObj.ErrorDetails = $ErrorObject.ErrorDetails.Message
        }
        elseif ($ErrorObject.Exception.GetType().FullName -eq 'System.Net.WebException') {
            if ($null -ne $ErrorObject.Exception.Response) {
                $streamReaderResponse = [System.IO.StreamReader]::new($ErrorObject.Exception.Response.GetResponseStream()).ReadToEnd()
                if (-not [string]::IsNullOrEmpty($streamReaderResponse)) {
                    $httpErrorObj.ErrorDetails = $streamReaderResponse
                }
            }
        }
        try {
            $errorObjectConverted = $ErrorObject | ConvertFrom-Json -ErrorAction Stop

            if ($null -ne $errorObjectConverted.error_description) {
                $httpErrorObj.FriendlyMessage = $errorObjectConverted.error_description
            }
            elseif ($null -ne $errorObjectConverted.error) {
                if ($null -ne $errorObjectConverted.error.message) {
                    $httpErrorObj.FriendlyMessage = $errorObjectConverted.error.message
                    if ($null -ne $errorObjectConverted.error.code) { 
                        $httpErrorObj.FriendlyMessage = $httpErrorObj.FriendlyMessage + " Error code: $($errorObjectConverted.error.code)"
                    }
                }
                else {
                    $httpErrorObj.FriendlyMessage = $errorObjectConverted.error
                }
            }
            else {
                $httpErrorObj.FriendlyMessage = $ErrorObject
            }
        }
        catch {
            $httpErrorObj.FriendlyMessage = $httpErrorObj.ErrorDetails
        }
        Write-Output $httpErrorObj
    }
}

function New-TeamsSession {
    [CmdletBinding()]
    [OutputType([System.Collections.Generic.Dictionary[[String], [String]]])]
    param(
        [parameter(Mandatory)]
        [string]
        $TenantId,

        [parameter(Mandatory)]
        [string]
        $ClientId,

        [parameter(Mandatory)]
        [string]
        $ClientSecret
    )
    try {
        $baseUri = "https://login.microsoftonline.com/"
        $authUri = $baseUri + "$TenantId/oauth2/v2.0/token"

        Write-Verbose "Creating Graph Access Token"
        $bodyGraph = @{
            grant_type    = "client_credentials"
            client_id     = "$ClientId"
            client_secret = "$ClientSecret"
            scope      = "https://graph.microsoft.com/.default"
        }
        $responseGraph = Invoke-RestMethod -Method POST -Uri $authUri -Body $bodyGraph -ContentType 'application/x-www-form-urlencoded' -Verbose:$false -UseBasicParsing:$true -ErrorAction "Stop"
        $accessTokenGraph = $responseGraph.access_token
        Write-Verbose "Succesfully created Graph Access Token"

        Write-Verbose "Creating MS Teams Access Token"
        $bodyTeams = @{
            grant_type    = "client_credentials"
            client_id     = "$ClientId"
            client_secret = "$ClientSecret"
            scope      = "48ac35b8-9aa8-4d74-927d-1f4a14a0b239/.default"
        }

        $responseTeams= Invoke-RestMethod -Method POST -Uri $authUri -Body $bodyTeams -ContentType 'application/x-www-form-urlencoded'-Verbose:$false -UseBasicParsing:$true -ErrorAction "Stop" 
        $accessTokenTeams = $responseTeams.access_token
        Write-Verbose "Succesfully created MS Teams Access Token"

        Write-Verbose "Connecting to Microsoft Teams"

        $connectTeamsSplatParams = @{
            AccessTokens = @("$accessTokenGraph", "$accessTokenTeams")
            Verbose      = $false
            ErrorAction  = "Stop"
        }
        $teamsSession = Connect-MicrosoftTeams @connectTeamsSplatParams
        Write-Verbose "Successfully connected to Microsoft Teams"

        Write-Output $teamsSession
    }
    catch {
        $ex = $PSItem
        $errorMessage = Resolve-MicrosoftGraphAPIError -ErrorObject $ex
        write-verbose ($errorMessage | ConvertTo-Json)
        throw "Error connecting to Microsoft Teams. Error Message: $($errorMessage.AuditErrorMessage)"
    }
}

function Resolve-HTTPError {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory,
            ValueFromPipeline
        )]
        [object]$ErrorObject
    )
    process {
        $httpErrorObj = [PSCustomObject]@{
            FullyQualifiedErrorId = $ErrorObject.FullyQualifiedErrorId
            MyCommand             = $ErrorObject.InvocationInfo.MyCommand
            RequestUri            = $ErrorObject.TargetObject.RequestUri
            ScriptStackTrace      = $ErrorObject.ScriptStackTrace
            ErrorMessage          = ''
        }
        if ($ErrorObject.Exception.GetType().FullName -eq 'Microsoft.Powershell.Commands.HttpResponseException') {
            $httpErrorObj.ErrorMessage = $ErrorObject.ErrorDetails.Message
        }
        elseif ($ErrorObject.Exception.GetType().FullName -eq 'System.Net.WebException') {
            $httpErrorObj.ErrorMessage = [System.IO.StreamReader]::new($ErrorObject.Exception.Response.GetResponseStream()).ReadToEnd()
        }
        Write-Output $httpErrorObj
    }
}
#endregion functions

# PowerShell commands to import
$commands = @(
    "Get-CsPhoneNumberAssignment"
    "Get-CsOnlineUser"
)

# Define correlation
$correlationField = $actionContext.CorrelationConfiguration.accountField
$correlationValue = $actionContext.CorrelationConfiguration.accountFieldValue

# Define account object
$account = [PSCustomObject]$actionContext.Data

try {
    # Import module
    $moduleName = "MicrosoftTeams"

    # If module is imported say that and do nothing
    if (Get-Module -Verbose:$false | Where-Object { $_.Name -eq $ModuleName }) {
        Write-Verbose "Module [$ModuleName] is already imported."
    }
    else {
        # If module is not imported, but available on disk then import
        if (Get-Module -ListAvailable -Verbose:$false | Where-Object { $_.Name -eq $ModuleName }) {
            $module = Import-Module $ModuleName -Cmdlet $commands -Verbose:$false
            Write-Verbose "Imported module [$ModuleName]"
        }
        else {
            # If the module is not imported, not available and not in the online gallery then abort
            throw "Module [$ModuleName] is not available. Please install the module using: Install-Module -Name [$ModuleName] -Force"
        }
    }

    $actionMessage = "Creating session"

    $teamsSessionSplatParams = @{
        TenantId     = $actionContext.Configuration.TenantID
        ClientId     = $actionContext.Configuration.AppId
        ClientSecret = $actionContext.Configuration.AppSecret
    }

    $teamsSession = New-TeamsSession @teamsSessionSplatParams


    $actionMessage = "verifying correlation configuration and properties"

    if ($actionContext.CorrelationConfiguration.Enabled -eq $true) {
        if ([string]::IsNullOrEmpty($correlationField)) {
            throw "Correlation is enabled but not configured correctly."
        }

        if ([string]::IsNullOrEmpty($correlationValue)) {
            throw "The correlation value for [$correlationField] is empty. This is likely a mapping issue."
        }

        if ([string]::IsNullOrEmpty($account.LocationCode)) {
            Throw "The Location Code is empty for person"
        }

        #region Get Microsoft Teams Voip assignment
        Write-Verbose "Querying MS Teams Phonenumber Assignment where [AssignedPstnTargetId] = [$($account.userPrincipalName)] and [NumberType] = [$($account.PhoneNumberType)]"

        $getPhonenumberAssignmentSplatParams = @{
            AssignedPstnTargetId = $account.userPrincipalName
            NumberType           = $account.PhoneNumberType
            Verbose              = $false
            ErrorAction          = "Stop"
        }

        $currentPhonenumberAssignment = Get-CsPhoneNumberAssignment @getPhonenumberAssignmentSplatParams

        Write-Verbose "Successfully queried MS Teams Phonenumber Assignment where [AssignedPstnTargetId] = [$($account.userPrincipalName)] and [NumberType] = [$($account.PhoneNumberType)]. Result count: $(($currentPhonenumberAssignment | Measure-Object).Count)" # Result: $($currentPhonenumberAssignment | ConvertTo-Json)"

        if (($currentPhonenumberAssignment | Measure-Object).count -eq 0) {
            # No phone number assignment found, get Teams user for correlation
            $getTeamsUserSplatParams = @{
                Identity          = $account.userPrincipalName
                Verbose           = $false
                ErrorAction       = "Stop"
            }

            $currentTeamsUser = Get-CsOnlineUser @getTeamsUserSplatParams
            $returnObjectCount = $currentTeamsUser
            $returnObject = [PSCustomObject]@{
                Id = $currentTeamsUser.Identity
                PhoneNumber = $null
            }
        } else {
            $returnObjectCount = $currentPhonenumberAssignment
            $returnObject = [PSCustomObject]@{
                Id = $currentPhonenumberAssignment.AssignedPstnTargetId
                PhoneNumber = $currentPhonenumberAssignment.TelephoneNumber
            }
        }
    }
    else {
        if ($actionContext.Configuration.correlateOnly -eq $true) {
            throw "Correlation is disabled while configuration option [correlateOnly] is toggled."
        }
        else {
            Write-Warning "Correlation is disabled."
        }
    }

    $actionMessage = "calculating action"
    if (($returnObjectCount | Measure-Object).count -eq 0) {
            $actionAccount = "NotFound"
    }
    elseif (($returnObjectCount | Measure-Object).count -eq 1) {
        $actionAccount = "Correlate"
    }
    elseif (($returnObjectCount | Measure-Object).count -gt 1) {
        $actionAccount = "MultipleFound"
    }

    switch ($actionAccount) {
        "Correlate" {
            $actionMessage = "correlating to account"

            $outputContext.AccountReference = [PSCustomObject]@{
                Id                  = $returnObject.Id
                PhoneNumber         = $returnObject.PhoneNumber
                UserPrincipalName   = $account.UserPrincipalName
            }

            $outputContext.Data = $account
            $outputContext.Data.id = $returnObject.Id
            $outputContext.Data.PhoneNumber = $returnObject.PhoneNumber

            if ([String]::IsNullOrEmpty($returnObject.PhoneNumber)) {
                $text = ""
            } else {
                $text = "and phone number [$($returnObject.PhoneNumber)]"
            }
            $outputContext.AuditLogs.Add([PSCustomObject]@{
                    Action  = "CorrelateAccount"
                    Message = "Correlated to account with [$($correlationField)] = [$($correlationValue)] $text"
                    IsError = $false
                })

            $outputContext.AccountCorrelated = $true

            break
        }

        "MultipleFound" {
            $actionMessage = "correlating to account"

            # Throw terminal error
            throw "Multiple accounts found where [$($correlationField)] = [$($correlationValue)]. Please correct this so the persons are unique."

            break
        }

        "NotFound" {
            $actionMessage = "correlating to account"

            # Throw terminal error
            throw "No account found where [$($correlationField)] = [$($correlationValue)]"

            break
        }
    }
}
catch {
    $ex = $PSItem
    if ($($ex.Exception.GetType().FullName -eq 'Microsoft.PowerShell.Commands.HttpResponseException') -or
        $($ex.Exception.GetType().FullName -eq 'System.Net.WebException')) {
        $errorObj = Resolve-MicrosoftGraphAPIError -ErrorObject $ex
        $auditMessage = "Error $($actionMessage). Error: $($errorObj.FriendlyMessage)"
        Write-Warning "Error at Line [$($errorObj.ScriptLineNumber)]: $($errorObj.Line). Error: $($errorObj.ErrorDetails)"
    }
    else {
        $auditMessage = "Error $($actionMessage). Error: $($ex.Exception.Message)"
        Write-Warning "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception.Message)"
    }

    $outputContext.AuditLogs.Add([PSCustomObject]@{
            # Action  = "" # Optional
            Message = $auditMessage
            IsError = $true
        })
}
finally {
    # Check if auditLogs contains errors, if no errors are found, set success to true
    if ($outputContext.AuditLogs.IsError -contains $true) {
        $outputContext.Success = $false
    }
    else {
        $outputContext.Success = $true
    }

    # Check if accountreference is set, if not set, set this with default value as this must contain a value
    if ([String]::IsNullOrEmpty($outputContext.AccountReference) -and $actionContext.DryRun -eq $true) {
        $outputContext.AccountReference = "DryRun: Currently not available"
    }
}