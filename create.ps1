#####################################################
# HelloID-Conn-Prov-Target-MicrosoftTeams-DirectRoutingPhonenumber-Create
#
# Version: 2.0.0
# RJ: Converting connector to PSv2 connector
#####################################################

# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

# Set debug logging
switch ($($actionContext.Configuration.isDebug)) {
    $true { $VerbosePreference = "Continue" }
    $false { $VerbosePreference = "SilentlyContinue" }
}

# Used to connect to Microsoft Teams in an unattended scripting scenario using an App ID and App Secret to create an Access Token.
$AADTenantId = $actionContext.Configuration.MicrosoftEntraIDTenantId
$AADAppID = $actionContext.Configuration.MicrosoftEntraIDAppId
$AADAppSecret = $actionContext.Configuration.MicrosoftEntraIDAppSecret
$OnlySetPhoneNumberWhenEmpty = $actionContext.Configuration.OnlySetPhoneNumberWhenEmpty

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

# RJ: MOVE SCRIPT MAPPING TO FIELD MAPPING 
# $phoneNumber = $p.Contact.Business.Phone.Mobile # Has to be picked from a list
# if(-not($phoneNumber.StartsWith("+31"))){
#     $phoneNumber = "+31" + $phoneNumber
# }
# $account = [PSCustomObject]@{
#     Identity        = $p.Accounts.MicrosoftAzureAD.userPrincipalName
#     PhoneNumber     = $phoneNumber
#     PhoneNumberType = "DirectRouting"
# }

# # Define account properties to update
# $updateAccountFields = @("PhoneNumber")

# # Define account properties to store in account data
# $storeAccountFields = @("PhoneNumber", "PhoneNumberType")
#endregion Change mapping here

#region functions
# function Convert-StringToBoolean($obj) {
#     if ($obj -is [PSCustomObject]) {
#         foreach ($property in $obj.PSObject.Properties) {
#             $value = $property.Value
#             if ($value -is [string]) {
#                 $lowercaseValue = $value.ToLower()
#                 if ($lowercaseValue -eq "true") {
#                     $obj.$($property.Name) = $true
#                 }
#                 elseif ($lowercaseValue -eq "false") {
#                     $obj.$($property.Name) = $false
#                 }
#             }
#             elseif ($value -is [PSCustomObject] -or $value -is [System.Collections.IDictionary]) {
#                 $obj.$($property.Name) = Convert-StringToBoolean $value
#             }
#             elseif ($value -is [System.Collections.IList]) {
#                 for ($i = 0; $i -lt $value.Count; $i++) {
#                     $value[$i] = Convert-StringToBoolean $value[$i]
#                 }
#                 $obj.$($property.Name) = $value
#             }
#         }
#     }
#     return $obj
# }

# function Resolve-MicrosoftGraphAPIError {
#     [CmdletBinding()]
#     param (
#         [Parameter(Mandatory)]
#         [object]
#         $ErrorObject
#     )
#     process {
#         $httpErrorObj = [PSCustomObject]@{
#             ScriptLineNumber = $ErrorObject.InvocationInfo.ScriptLineNumber
#             Line             = $ErrorObject.InvocationInfo.Line
#             ErrorDetails     = $ErrorObject.Exception.Message
#             FriendlyMessage  = $ErrorObject.Exception.Message
#         }
#         if (-not [string]::IsNullOrEmpty($ErrorObject.ErrorDetails.Message)) {
#             $httpErrorObj.ErrorDetails = $ErrorObject.ErrorDetails.Message
#         }
#         elseif ($ErrorObject.Exception.GetType().FullName -eq 'System.Net.WebException') {
#             if ($null -ne $ErrorObject.Exception.Response) {
#                 $streamReaderResponse = [System.IO.StreamReader]::new($ErrorObject.Exception.Response.GetResponseStream()).ReadToEnd()
#                 if (-not [string]::IsNullOrEmpty($streamReaderResponse)) {
#                     $httpErrorObj.ErrorDetails = $streamReaderResponse
#                 }
#             }
#         }
#         try {
#             $errorObjectConverted = $ErrorObject | ConvertFrom-Json -ErrorAction Stop

#             if ($null -ne $errorObjectConverted.error_description) {
#                 $httpErrorObj.FriendlyMessage = $errorObjectConverted.error_description
#             }
#             elseif ($null -ne $errorObjectConverted.error) {
#                 if ($null -ne $errorObjectConverted.error.message) {
#                     $httpErrorObj.FriendlyMessage = $errorObjectConverted.error.message
#                     if ($null -ne $errorObjectConverted.error.code) { 
#                         $httpErrorObj.FriendlyMessage = $httpErrorObj.FriendlyMessage + " Error code: $($errorObjectConverted.error.code)"
#                     }
#                 }
#                 else {
#                     $httpErrorObj.FriendlyMessage = $errorObjectConverted.error
#                 }
#             }
#             else {
#                 $httpErrorObj.FriendlyMessage = $ErrorObject
#             }
#         }
#         catch {
#             $httpErrorObj.FriendlyMessage = $httpErrorObj.ErrorDetails
#         }
#         Write-Output $httpErrorObj
#     }
# }

# function New-AuthorizationHeaders {
#     [CmdletBinding()]
#     [OutputType([System.Collections.Generic.Dictionary[[String], [String]]])]
#     param(
#         [parameter(Mandatory)]
#         [string]
#         $TenantId,

#         [parameter(Mandatory)]
#         [string]
#         $ClientId,

#         [parameter(Mandatory)]
#         [string]
#         $ClientSecret
#     )
#     try {
#         Write-Verbose "Creating Access Token"
#         $baseUri = "https://login.microsoftonline.com/"
#         $authUri = $baseUri + "$TenantId/oauth2/token"
    
#         $body = @{
#             grant_type    = "client_credentials"
#             client_id     = "$ClientId"
#             client_secret = "$ClientSecret"
#             resource      = "https://graph.microsoft.com"
#         }
    
#         $Response = Invoke-RestMethod -Method POST -Uri $authUri -Body $body -ContentType 'application/x-www-form-urlencoded'
#         $accessToken = $Response.access_token
    
#         #Add the authorization header to the request
#         Write-Verbose 'Adding Authorization headers'

#         $headers = [System.Collections.Generic.Dictionary[[String], [String]]]::new()
#         $headers.Add('Authorization', "Bearer $accesstoken")
#         $headers.Add('Accept', 'application/json')
#         $headers.Add('Content-Type', 'application/json')
#         # Needed to filter on specific attributes (https://docs.microsoft.com/en-us/graph/aad-advanced-queries)
#         $headers.Add('ConsistencyLevel', 'eventual')

#         Write-Output $headers  
#     }
#     catch {
#         throw $_
#     }
# }

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
# function getFreePhoneNumber(
    
# )
# {
#     #Prepare
#     # Connect to database (do this in function)
#     # Get number (do this in function)
#     # Query system x to see if number X is free
#     # Use number
# }
#endregion functions

try {
    # Validation
    try {
        # Check if required fields are available in configuration object
        $incompleteConfiguration = $false
        foreach ($requiredConfigurationField in $requiredConfigurationFields) {
            if ($requiredConfigurationField -notin $c.PsObject.Properties.Name) {
                $incompleteConfiguration = $true
                Write-Warning "Required configuration object field [$requiredConfigurationField] is missing"
            }
            elseif ([String]::IsNullOrEmpty($c.$requiredConfigurationField)) {
                $incompleteConfiguration = $true
                Write-Warning "Required configuration object field [$requiredConfigurationField] has a null or empty value"
            }
        }

        if ($incompleteConfiguration -eq $true) {
            throw "Configuration object incomplete, cannot continue."
        }

        if ($actionContext.CorrelationConfiguration.Enabled) {
            $correlationProperty = $actionContext.CorrelationConfiguration.accountField
            $correlationValue = $actionContext.CorrelationConfiguration.accountFieldValue
    
            if ([string]::IsNullOrEmpty($correlationProperty)) {
                Write-Warning "Correlation is enabled but not configured correctly."
                Throw "Correlation is enabled but not configured correctly."
            }
    
            if ([string]::IsNullOrEmpty($correlationValue)) {
                Write-Warning "The correlation value for [$correlationProperty] is empty. This is likely a scripting issue."
                Throw "The correlation value for [$correlationProperty] is empty. This is likely a scripting issue."
            }
        }
        else {
            $outputContext.AuditLogs.Add([PSCustomObject]@{
                    Message = "Configuration of correlation is madatory."
                    IsError = $true
                })
            Throw "Configuration of correlation is madatory."
        }

        $account = $actionContext.Data
    }
    catch {
        
        $ex = $PSItem
        Write-Verbose -Verbose ($ex | ConvertTo-Json)

        $outputContext.AuditLogs.Add([PSCustomObject]@{
                Action  = "CreateAccount" # Should be correlateaccount, update done in update script
                Message = "$($ex.Exception.Message)"
                IsError = $true
            })

        throw $_
    }

    # Teams module
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
        $authUri = $baseUri + "$AADTenantId/oauth2/v2.0/token"
        
        $body = @{
            grant_type    = "client_credentials"
            client_id     = "$AADAppID"
            client_secret = "$AADAppSecret"
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
        $authUri = $baseUri + "$AADTenantId/oauth2/v2.0/token"
        
        $body = @{
            grant_type    = "client_credentials"
            client_id     = "$AADAppID"
            client_secret = "$AADAppSecret"
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
    
    # Setup DB connection

    # Get Current Phone Number Assignment of Microsoft Teams User. More info on Microsoft docs: https://learn.microsoft.com/en-us/powershell/module/teams/get-csphonenumberassignment?view=teams-ps
    try {
        Write-Verbose "Querying MS Teams Phonenumber Assignment where [AssignedPstnTargetId] = [$($account.userPrincipalName)] and [NumberType] = [$($account.PhoneNumberType)]"
        
        $getPhonenumberAssignmentSplatParams = @{
            AssignedPstnTargetId = $account.userPrincipalName
            NumberType           = $account.PhoneNumberType
            Verbose              = $false
            ErrorAction          = "Stop"
        } 

        $currentPhonenumberAssignment = Get-CsPhoneNumberAssignment @getPhonenumberAssignmentSplatParams

        if (($currentPhonenumberAssignment | Measure-Object).Count -eq 0) {
            Write-Verbose "No MS Teams Phonenumber Assignment found where [AssignedPstnTargetId] = [$($account.userPrincipalName)] and [NumberType] = [$($account.PhoneNumberType)]" 
        }

        Write-Verbose "Successfully queried MS Teams Phonenumber Assignment where [AssignedPstnTargetId] = [$($account.userPrincipalName)] and [NumberType] = [$($account.PhoneNumberType)]. Result count: $(($currentPhonenumberAssignment | Measure-Object).Count)"
    }
    catch { 
        $ex = $PSItem
        $errorMessage = Get-ErrorMessage -ErrorObject $ex

        Write-Verbose "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($errorMessage.VerboseErrorMessage)"
        $auditLogs.Add([PSCustomObject]@{
                # Action  = "" # Optional
                Message = "Error querying MS Teams Phonenumber Assignment where [AssignedPstnTargetId] = [$($account.userPrincipalName)] and [NumberType] = [$($account.PhoneNumberType)]. Error Message: $($errorMessage.AuditErrorMessage)"
                IsError = $True
            })
    }

    # Check if update is required
    try {
        Write-Verbose "Calculating changes"

        # Vanaf hier moeten er nog wijzigingen plaats vinden, maar eerst DB connectie opzetten
        
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
                        Message = "Skipped updating MS Teams Phonenumber Assignment where [NumberType] = [$($account.PhoneNumberType)] for [$($account.userPrincipalName)]. Reason: Configured to only update MS Teams Phonenumber Assignment when empty. Old value: [$($previousAccount.PhoneNumber)]. New value: [$($account.PhoneNumber)]"
                        IsError = $false
                    })
                
                break
            }
            else {
                try {
                    if (-not($dryRun -eq $true)) {
                        Write-Verbose "Updating MS Teams Phonenumber Assignment where [NumberType] = [$($account.PhoneNumberType)] for [$($account.userPrincipalName)]. Old value: [$($previousAccount.PhoneNumber)]. New value: [$($account.PhoneNumber)]"
                    
                        $setPhonenumberAssignmentSplatParams = @{
                            Identity        = $account.userPrincipalName
                            PhoneNumber     = $account.PhoneNumber
                            PhoneNumberType = $account.PhoneNumberType
                            Verbose         = $false
                            ErrorAction     = "Stop"
                        }
            
                        $updatePhonenumberAssignment = Set-CsPhoneNumberAssignment @setPhonenumberAssignmentSplatParams

                        $auditLogs.Add([PSCustomObject]@{
                                # Action  = "" # Optional
                                Message = "Successfully updated MS Teams Phonenumber Assignment where [NumberType] = [$($account.PhoneNumberType)] for [$($account.userPrincipalName)]. Old value: [$($previousAccount.PhoneNumber)]. New value: [$($account.PhoneNumber)]"
                                IsError = $false
                            })
                    }
                    else {
                        $auditLogs.Add([PSCustomObject]@{
                                # Action  = "" # Optional
                                Message = "DryRun: Would update MS Teams Phonenumber Assignment where [NumberType] = [$($account.PhoneNumberType)] for [$($account.userPrincipalName)]. Old value: [$($previousAccount.PhoneNumber)]. New value: [$($account.PhoneNumber)]"
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
                            Message = "Error updating MS Teams Phonenumber Assignment where [NumberType] = [$($account.PhoneNumberType)] for [$($account.userPrincipalName)]. Old value: [$($previousAccount.PhoneNumber)]. New value: [$($account.PhoneNumber)]. Error Message: $($errorMessage.AuditErrorMessage)"
                            IsError = $True
                        })
                }

                break
            }
        }
        'NoChanges' {
            $auditLogs.Add([PSCustomObject]@{
                    # Action  = "" # Optional
                    Message = "Skipped updating MS Teams Phonenumber Assignment where [NumberType] = [$($account.PhoneNumberType)] for [$($account.userPrincipalName)]. Reason: No changes. Old value: [$($previousAccount.PhoneNumber)]. New value: [$($account.PhoneNumber)]"
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