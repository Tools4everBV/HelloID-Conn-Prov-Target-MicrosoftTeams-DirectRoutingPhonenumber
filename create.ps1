#####################################################
# HelloID-Conn-Prov-Target-MicrosoftTeams-DirectRoutingPhonenumber-Create
#
# Version: 1.0.0
#####################################################
# Initialize default values
$c = $configuration | ConvertFrom-Json
$p = $person | ConvertFrom-Json
$success = $false # Set to false at start, at the end, only when no error occurs it is set to true
$auditLogs = [System.Collections.Generic.List[PSCustomObject]]::new()

# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

# Set debug logging
switch ($($c.isDebug)) {
    $true { $VerbosePreference = "Continue" }
    $false { $VerbosePreference = "SilentlyContinue" }
}
$InformationPreference = "Continue"
$WarningPreference = "Continue"

# Used to connect to Microsoft Teams in an unattended scripting scenario using an App ID and App Secret to create an Access Token.
$microsoftEntraIDTenantId = $c.MicrosoftEntraIDTenantId
$microsoftEntraIDAppID = $c.MicrosoftEntraIDAppId
$microsoftEntraIDAppSecret = $c.MicrosoftEntraIDAppSecret
$OnlySetPhoneNumberWhenEmpty = $c.OnlySetPhoneNumberWhenEmpty

# PowerShell commands to import
$commands = @(
    "Get-CsPhoneNumberAssignment"
    , "Set-CsPhoneNumberAssignment"
)

#region Change mapping here
# The available account properties are linked to the available properties of the command "Set-CsPhoneNumberAssignment": https://learn.microsoft.com/en-us/powershell/module/teams/set-csphonenumberassignment?view=teams-ps command "https://learn.microsoft.com/en-us/powershell/module/teams/set-csphonenumberassignment?view=teams-ps", 
# Phone numbers use the format "+<country code> <number>x<extension>", with extension optional.
# For example, +1 5555551234 or +1 5555551234x123 are valid. Numbers are rejected when creating/updating if they do not match the required format. 
# Phone numbers use the format "+<country code> <number>x<extension>", with extension optional.
# For example, +1 5555551234 or +1 5555551234x123 are valid. Numbers are rejected when creating/updating if they do not match the required format. 
$phoneNumber = $p.Contact.Business.Phone.Mobile
if(-not($phoneNumber.StartsWith("+31"))){
    $phoneNumber = "+31" + $phoneNumber
}
$account = [PSCustomObject]@{
    Identity        = $p.Accounts.MicrosoftAzureAD.userPrincipalName
    PhoneNumber     = $phoneNumber
    PhoneNumberType = "DirectRouting"
}

# # Correlation values - Outcommented, as there is no correlation as there is no command to get the Teams User
# $correlationProperty = "" # Has to match the name of the property that contains unique identifier
# $correlationValue = "" # Has to match the value of the unique identifier property

# Define account properties to update
$updateAccountFields = @("PhoneNumber")

# Define account properties to store in account data
$storeAccountFields = @("PhoneNumber", "PhoneNumberType")
#endregion Change mapping here

#region functions
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
            ErrorMessage          = ""
        }
        if ($ErrorObject.Exception.GetType().FullName -eq "Microsoft.PowerShell.Commands.HttpResponseException") {
            $httpErrorObj.ErrorMessage = $ErrorObject.ErrorDetails.Message
        }
        elseif ($ErrorObject.Exception.GetType().FullName -eq "System.Net.WebException") {
            $httpErrorObj.ErrorMessage = [System.IO.StreamReader]::new($ErrorObject.Exception.Response.GetResponseStream()).ReadToEnd()
        }
        Write-Output $httpErrorObj
    }
}

function Get-ErrorMessage {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory,
            ValueFromPipeline
        )]
        [object]$ErrorObject
    )
    process {
        $errorMessage = [PSCustomObject]@{
            VerboseErrorMessage = $null
            AuditErrorMessage   = $null
        }

        if ( $($ErrorObject.Exception.GetType().FullName -eq "Microsoft.PowerShell.Commands.HttpResponseException") -or $($ErrorObject.Exception.GetType().FullName -eq "System.Net.WebException")) {
            $httpErrorObject = Resolve-HTTPError -Error $ErrorObject

            $errorMessage.VerboseErrorMessage = $httpErrorObject.ErrorMessage

            $errorMessage.AuditErrorMessage = $httpErrorObject.ErrorMessage
        }

        # If error message empty, fall back on $ex.Exception.Message
        if ([String]::IsNullOrEmpty($errorMessage.VerboseErrorMessage)) {
            $errorMessage.VerboseErrorMessage = $ErrorObject.Exception.Message
        }
        if ([String]::IsNullOrEmpty($errorMessage.AuditErrorMessage)) {
            $errorMessage.AuditErrorMessage = $ErrorObject.Exception.Message
        }

        Write-Output $errorMessage
    }
}
#endregion functions

try {
    # Set aRef object for use in futher actions - Since there is no correlation as there is no command to get the Teams User, we use the Identity from the account object
    $aRef = $account.Identity

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
    }
    catch {
        $ex = $PSItem
        $errorMessage = Get-ErrorMessage -ErrorObject $ex

        Write-Verbose "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($errorMessage.VerboseErrorMessage)"
        $auditLogs.Add([PSCustomObject]@{
                # Action  = "" # Optional
                Message = "Error importing module [$ModuleName]. Error Message: $($errorMessage.AuditErrorMessage)"
                IsError = $True
            })

        # Skip further actions, as this is a critical error
        continue
    }

    # Connect to Microsoft Teams. More info on Microsoft docs: https://learn.microsoft.com/en-us/MicrosoftTeams/teams-powershell-application-authentication#:~:text=Connect%20using%20Access%20Tokens%3A
    try {
        # Create MS Graph access token
        Write-Verbose "Creating MS Graph Access Token"

        $baseUri = "https://login.microsoftonline.com/"
        $authUri = $baseUri + "$microsoftEntraIDTenantId/oauth2/v2.0/token"
        
        $body = @{
            grant_type    = "client_credentials"
            client_id     = "$microsoftEntraIDAppID"
            client_secret = "$microsoftEntraIDAppSecret"
            scope         = "https://graph.microsoft.com/.default"
        }

        $graphTokenSplatParams = @{
            Method          = "POST"
            Uri             = $authUri
            Body            = $body
            ContentType     = "application/x-www-form-urlencoded"
            UseBasicParsing = $true
            Verbose         = $false
            ErrorAction     = "Stop"
        }
        
        $graphTokenResponse = Invoke-RestMethod @graphTokenSplatParams
        $graphToken = $graphTokenResponse.access_token

        Write-Verbose "Successfully created MS Graph Access Token"

        # Create Skype and Teams Tenant Admin API access token
        Write-Verbose "Creating Skype and Teams Tenant Admin API Access Token"

        $baseUri = "https://login.microsoftonline.com/"
        $authUri = $baseUri + "$microsoftEntraIDTenantId/oauth2/v2.0/token"
        
        $body = @{
            grant_type    = "client_credentials"
            client_id     = "$microsoftEntraIDAppID"
            client_secret = "$microsoftEntraIDAppSecret"
            scope         = "48ac35b8-9aa8-4d74-927d-1f4a14a0b239/.default"
        }
        
        $teamsTokenSplatParams = @{
            Method          = "POST"
            Uri             = $authUri
            Body            = $body
            ContentType     = "application/x-www-form-urlencoded"
            UseBasicParsing = $true
            Verbose         = $false
            ErrorAction     = "Stop"
        }

        $teamsTokenResponse = Invoke-RestMethod @teamsTokenSplatParams
        $teamsToken = $teamsTokenResponse.access_token

        Write-Verbose "Successfully created Skype and Teams Tenant Admin API Access Token"

        # Connect to Microsoft Teams in an unattended scripting scenario using an access token.
        Write-Verbose "Connecting to Microsoft Teams"

        $connectTeamsSplatParams = @{
            AccessTokens = @("$graphToken", "$teamsToken")
            Verbose      = $false
            ErrorAction  = "Stop"
        }

        $teamsSession = Connect-MicrosoftTeams @connectTeamsSplatParams
        
        Write-Verbose "Successfully connected to Microsoft Teams"
    }
    catch {
        $ex = $PSItem
        $errorMessage = Get-ErrorMessage -ErrorObject $ex

        Write-Verbose "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($errorMessage.VerboseErrorMessage)"
        $auditLogs.Add([PSCustomObject]@{
                # Action  = "" # Optional
                Message = "Error connecting to Microsoft Teams. Error Message: $($errorMessage.AuditErrorMessage)"
                IsError = $True
            })

        # Skip further actions, as this is a critical error
        continue
    }
    
    # Get Current Phone Number Assignment of Microsoft Teams User. More info on Microsoft docs: https://learn.microsoft.com/en-us/powershell/module/teams/get-csphonenumberassignment?view=teams-ps
    try {
        Write-Verbose "Querying MS Teams Phonenumber Assignment where [AssignedPstnTargetId] = [$($account.Identity)] and [NumberType] = [$($account.PhoneNumberType)]"
        
        $getPhonenumberAssignmentSplatParams = @{
            AssignedPstnTargetId = $account.Identity
            NumberType           = $account.PhoneNumberType
            Verbose              = $false
            ErrorAction          = "Stop"
        } 

        $currentPhonenumberAssignment = Get-CsPhoneNumberAssignment @getPhonenumberAssignmentSplatParams

        if (($currentPhonenumberAssignment | Measure-Object).Count -eq 0) {
            Write-Verbose "No MS Teams Phonenumber Assignment found where [AssignedPstnTargetId] = [$($account.Identity)] and [NumberType] = [$($account.PhoneNumberType)]" 
        }

        Write-Verbose "Successfully queried MS Teams Phonenumber Assignment where [AssignedPstnTargetId] = [$($account.Identity)] and [NumberType] = [$($account.PhoneNumberType)]. Result count: $(($currentPhonenumberAssignment | Measure-Object).Count)"
    }
    catch { 
        $ex = $PSItem
        $errorMessage = Get-ErrorMessage -ErrorObject $ex

        Write-Verbose "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($errorMessage.VerboseErrorMessage)"
        $auditLogs.Add([PSCustomObject]@{
                # Action  = "" # Optional
                Message = "Error querying MS Teams Phonenumber Assignment where [AssignedPstnTargetId] = [$($account.Identity)] and [NumberType] = [$($account.PhoneNumberType)]. Error Message: $($errorMessage.AuditErrorMessage)"
                IsError = $True
            })
    }

    # Check if update is required
    try {
        Write-Verbose "Calculating changes"

        # Create previous account object to compare current data with specified account data
        $previousAccount = [PSCustomObject]@{
            'PhoneNumber' = $currentPhonenumberAssignment.TelephoneNumber
        }
        
        # Calculate changes between current data and provided data
        $splatCompareProperties = @{
            ReferenceObject  = @($previousAccount.PSObject.Properties | Where-Object { $_.Name -in $updateAccountFields }) # Only select the properties to update
            DifferenceObject = @($account.PSObject.Properties | Where-Object { $_.Name -in $updateAccountFields }) # Only select the properties to update
        }
        $changedProperties = $null
        $changedProperties = (Compare-Object @splatCompareProperties -PassThru)
        $oldProperties = $changedProperties.Where( { $_.SideIndicator -eq '<=' })
        $newProperties = $changedProperties.Where( { $_.SideIndicator -eq '=>' })

        if (($newProperties | Measure-Object).Count -ge 1) {
            Write-Verbose "Changed properties: $($changedProperties | ConvertTo-Json)"

            $updateAction = 'Update'
        }
        else {
            Write-Verbose "No changed properties"

            $updateAction = 'NoChanges'
        }

        Write-Verbose "Successfully calculated changes"
    }
    catch {
        $ex = $PSItem
        $errorMessage = Get-ErrorMessage -ErrorObject $ex

        Write-Verbose "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($errorMessage.VerboseErrorMessage)"
        $auditLogs.Add([PSCustomObject]@{
                # Action  = "" # Optional
                Message = "Error calculating changes. Error Message: $($errorMessage.AuditErrorMessage)"
                IsError = $True
            })

        # Skip further actions, as this is a critical error
        continue
    }

    switch ($updateAction) {
        'Update' {
            if (-not[String]::IsNullOrEmpty($currentPhonenumberAssignment.TelephoneNumber) -and $OnlySetPhoneNumberWhenEmpty -eq $true) {
                $auditLogs.Add([PSCustomObject]@{
                        # Action  = "" # Optional
                        Message = "Skipped updating MS Teams Phonenumber Assignment where [NumberType] = [$($account.PhoneNumberType)] for [$($account.Identity)]. Reason: Configured to only update MS Teams Phonenumber Assignment when empty. Old value: [$($previousAccount.PhoneNumber)]. New value: [$($account.PhoneNumber)]"
                        IsError = $false
                    })
                
                break
            }
            else {
                try {
                    if (-not($dryRun -eq $true)) {
                        Write-Verbose "Updating MS Teams Phonenumber Assignment where [NumberType] = [$($account.PhoneNumberType)] for [$($account.Identity)]. Old value: [$($previousAccount.PhoneNumber)]. New value: [$($account.PhoneNumber)]"
                    
                        $setPhonenumberAssignmentSplatParams = @{
                            Identity        = $account.Identity
                            PhoneNumber     = $account.PhoneNumber
                            PhoneNumberType = $account.PhoneNumberType
                            Verbose         = $false
                            ErrorAction     = "Stop"
                        }
            
                        $updatePhonenumberAssignment = Set-CsPhoneNumberAssignment @setPhonenumberAssignmentSplatParams

                        $auditLogs.Add([PSCustomObject]@{
                                # Action  = "" # Optional
                                Message = "Successfully updated MS Teams Phonenumber Assignment where [NumberType] = [$($account.PhoneNumberType)] for [$($account.Identity)]. Old value: [$($previousAccount.PhoneNumber)]. New value: [$($account.PhoneNumber)]"
                                IsError = $false
                            })
                    }
                    else {
                        $auditLogs.Add([PSCustomObject]@{
                                # Action  = "" # Optional
                                Message = "DryRun: Would update MS Teams Phonenumber Assignment where [NumberType] = [$($account.PhoneNumberType)] for [$($account.Identity)]. Old value: [$($previousAccount.PhoneNumber)]. New value: [$($account.PhoneNumber)]"
                                IsError = $false
                            })
                    }
                }
                catch { 
                    $ex = $PSItem
                    $errorMessage = Get-ErrorMessage -ErrorObject $ex
        
                    Write-Verbose "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($errorMessage.VerboseErrorMessage)"
                    $auditLogs.Add([PSCustomObject]@{
                            # Action  = "" # Optional
                            Message = "Error updating MS Teams Phonenumber Assignment where [NumberType] = [$($account.PhoneNumberType)] for [$($account.Identity)]. Old value: [$($previousAccount.PhoneNumber)]. New value: [$($account.PhoneNumber)]. Error Message: $($errorMessage.AuditErrorMessage)"
                            IsError = $True
                        })
                }

                break
            }
        }
        'NoChanges' {
            $auditLogs.Add([PSCustomObject]@{
                    # Action  = "" # Optional
                    Message = "Skipped updating MS Teams Phonenumber Assignment where [NumberType] = [$($account.PhoneNumberType)] for [$($account.Identity)]. Reason: No changes. Old value: [$($previousAccount.PhoneNumber)]. New value: [$($account.PhoneNumber)]"
                    IsError = $false
                })
        
            break
        }
    }

    # Define ExportData with account fields and correlation property 
    $exportData = $account.PsObject.Copy() | Select-Object $storeAccountFields
    # # Add correlation property to exportdata - Outcommented, as there is no correlation as there is no command to get the Teams User
    # $exportData | Add-Member -MemberType NoteProperty -Name $correlationProperty -Value $correlationValue -Force
    # Add aRef to exportdata
    $exportData | Add-Member -MemberType NoteProperty -Name "AccountReference" -Value $aRef -Force
}
finally {
    # Check if auditLogs contains errors, if no errors are found, set success to true
    if (-NOT($auditLogs.IsError -contains $true)) {
        $success = $true
    }

    # Send results
    $result = [PSCustomObject]@{
        Success          = $success
        AccountReference = $aRef
        AuditLogs        = $auditLogs
        Account          = $account

        # Optionally return data for use in other systems
        ExportData       = $exportData
    }

    Write-Output ($result | ConvertTo-Json -Depth 10)
}