
#################################################
# HelloID-Conn-Prov-Target-Teams-Voip-Update
# PowerShell V2
#################################################

# Var only used here
$OnlineVoiceRoutingPolicyName = 'policy name'

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
        # Have to test this error handling by filling in incorrect password etc
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

function Invoke-SQLQuery {
    param(
        [parameter(Mandatory = $true)]
        $Server,
        [parameter(Mandatory = $true)]
        $Database,
        [parameter(Mandatory = $false)]
        $Username,

        [parameter(Mandatory = $false)]
        $Password,

        [parameter(Mandatory = $true)]
        $SqlQuery,

        [parameter(Mandatory = $true)]
        [ref]$Data
    )
    try {
        $Data.value = $null

        # Initialize connection and execute query
        if (-not[String]::IsNullOrEmpty($Username) -and -not[String]::IsNullOrEmpty($Password)) {
            # First create the PSCredential object
            $securePassword = ConvertTo-SecureString -String $Password -AsPlainText -Force
            $credential = [System.Management.Automation.PSCredential]::new($Username, $securePassword)
 
            # Set the password as read only
            $credential.Password.MakeReadOnly()
 
            # Create the SqlCredential object
            $sqlCredential = [System.Data.SqlClient.SqlCredential]::new($credential.username, $credential.password)
        }
        # Connect to the SQL server
        $ConnectionString = "Server=$Server;Database=$Database;Integrated Security=False;"
        $SqlConnection = [System.Data.SqlClient.SqlConnection]::new()

        $SqlConnection.ConnectionString = $ConnectionString

        $SqlConnection.Credential = $sqlCredential

        $SqlConnection.Open()
        Write-Verbose "Successfully connected to SQL database"

        # Set the query
        $SqlCmd = [System.Data.SqlClient.SqlCommand]::new()
        $SqlCmd.Connection = $SqlConnection
        $SqlCmd.CommandText = $SqlQuery

        # Set the data adapter
        $SqlAdapter = [System.Data.SqlClient.SqlDataAdapter]::new()
        $SqlAdapter.SelectCommand = $SqlCmd

        # Set the output with returned data
        $DataSet = [System.Data.DataSet]::new()
        $null = $SqlAdapter.Fill($DataSet)

        # Set the output with returned data
        $Data.value = $DataSet.Tables[0] | Select-Object -Property * -ExcludeProperty RowError, RowState, Table, ItemArray, HasErrors
    }
    catch {
        $Data.Value = $null
        throw $_
    }
    finally {
        if ($SqlConnection.State -eq "Open") {
            $SqlConnection.close()
            Write-Verbose "Successfully disconnected from SQL database"
        }
    }
}
#endregion

try {
    # Verify if [aRef] has a value
    if ([string]::IsNullOrEmpty($($actionContext.References.Account.id))) {
        throw 'The account reference could not be found'
    }

    # Enrich fields with None mapping
    $actionContext.Data.id                   = $actionContext.References.Account.id
    $actionContext.Data.userPrincipalName    = $actionContext.References.Account.userPrincipalName
    $actionContext.Data.PhoneNumber          = $actionContext.References.Account.PhoneNumber

    # Construct correlatedAccount based on field mapping and account reference
    $correlatedAccount = $actionContext.Data.PSObject.Copy()   

    Write-Verbose "correlatedAccount: $($correlatedAccount | ConvertTo-Json)"

    if (-Not [string]::IsNullOrEmpty($correlatedAccount.PhoneNumber) -or -Not $actionContext.AccountCorrelated) {
        $action = "NoChanges"
    } else {
        # Generate phone number
        Write-Verbose "Generating phone number"

        # Import module
        $moduleName = "MicrosoftTeams"

        # PowerShell commands to import
        $commands = @(
            "Get-CsPhoneNumberAssignment"
            "Get-CsOnlineUser"
            "Set-CsPhoneNumberAssignment"
            "Grant-CsOnlineVoiceRoutingPolicy"
        )

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

        Write-Verbose "correlatedAccount: $($correlatedAccount | ConvertTo-Json)"

        $teamsSessionSplatParams = @{
            TenantId     = $actionContext.Configuration.TenantID
            ClientId     = $actionContext.Configuration.AppId
            ClientSecret = $actionContext.Configuration.AppSecret
        }

        $teamsSession = New-TeamsSession @teamsSessionSplatParams

        # Get all numbers from SQL table
        $table = 'HelloID'
        $querySelect = "SELECT TelefoonNummer,LocatieCode FROM $table WHERE LocatieCode='$($correlatedAccount.LocationCode)'"

        Write-verbose "Querying data from table [$($table)]. Query: $($querySelect)"

        $querySelectSplatParams = @{
            Server      = $actionContext.Configuration.DbServer
            Database    = $actionContext.Configuration.DbName
            Username    = $actionContext.Configuration.DbUsername
            Password    = $actionContext.Configuration.DbPassword
            SqlQuery    = $querySelect
            ErrorAction = "Stop"
        }

        $querySelectResult = [System.Collections.ArrayList]::new()
        Invoke-SQLQuery @querySelectSplatParams -Data ([ref]$querySelectResult) -verbose:$false
        $selectRowCount = ($querySelectResult | measure-object).count
        if ($selectRowCount -eq 0) {
            throw "No numbers found in phone number database for location code [$($correlatedAccount.LocationCode)]"
        } 
        Write-Verbose "Successfully queried data from table [$($table)]. Query: $($querySelect). Returned rows: $selectRowCount"      

        Write-Verbose "Retrieving all assigned numbers from MS Teams..."
        # Get all assigned phone numbers in MS Teams
        # https://learn.microsoft.com/en-us/microsoftteams/see-a-list-of-phone-numbers-in-your-organization
        $getPhonenumberAssignmentSplatParams = @{
            NumberType  = $correlatedAccount.PhoneNumberType
            Top         = 10000
            Verbose     = $false
            ErrorAction = "Stop"
        }

        $allAssigedPhoneNumbers = $(Get-CsPhoneNumberAssignment @getPhonenumberAssignmentSplatParams | Select-Object TelephoneNumber).TelephoneNumber
        Write-Verbose "Retrieved all assigned numbers from MS Teams: $($allAssigedPhoneNumbers.Count)"

        Write-Verbose "Finding free number"
        $allPhoneNumbersForLocation = $querySelectResult.TelefoonNummer
        $correlatedAccount.PhoneNumber = $allPhoneNumbersForLocation | Where-Object { $_ -notin $allAssigedPhoneNumbers } | Select-Object -First 1
        Write-Verbose "Succesfully found a free phone number for location [$($correlatedAccount.LocationCode)]: [$($correlatedAccount.PhoneNumber)]"
    }

    # Always compare the account against the current account in target system
    if ($null -ne $correlatedAccount) {
        $splatCompareProperties = @{
            ReferenceObject  = @($correlatedAccount.PSObject.Properties)
            DifferenceObject = @($actionContext.Data.PSObject.Properties)
        }
        $propertiesChanged = Compare-Object @splatCompareProperties -PassThru | Where-Object { $_.SideIndicator -eq '=>' }
        #write-verbose ($correlatedAccount | ConvertTo-Json -Depth 10)
        #write-verbose ($actionContext.Data | ConvertTo-Json -Depth 10)

        if ($propertiesChanged) {
            $action = 'UpdateAccount'
            $dryRunMessage = "[DryRun] Account property(s) required to update: $($propertiesChanged.Name -join ', ')" # Nothing happens with this variable
        } else {
            $action = 'NoChanges'
            $dryRunMessage = '[DryRun] No changes will be made to the account during enforcement' # Nothing happens with this variable
        }
    } else {
        $action = 'NotFound'
        $dryRunMessage = "[DryRun] Correlated MS Teams account for: [$($personContext.Person.DisplayName)] not found" # Nothing happens with this variable
    }

    # Process
    switch ($action) {
        'UpdateAccount' {
            Write-Information "Updating MS Teams account with accountReference: [$($actionContext.References.Account.id)]"
            if (-not($actionContext.dryRun -eq $true)) {
                Write-Verbose "Updating MS Teams Phonenumber Assignment where [NumberType] = [$($correlatedAccount.PhoneNumberType)] for [$($correlatedAccount.userPrincipalName)]. Old value: [$($correlatedAccount.PhoneNumber)]. New value: [$($correlatedAccount.PhoneNumber)]"

                $setPhonenumberAssignmentSplatParams = @{
                    Identity        = $correlatedAccount.id
                    PhoneNumber     = $correlatedAccount.PhoneNumber
                    PhoneNumberType = $correlatedAccount.PhoneNumberType
                    Verbose         = $false
                    ErrorAction     = "Stop"
                }
                $updatePhonenumberAssignment = Set-CsPhoneNumberAssignment @setPhonenumberAssignmentSplatParams

                $outputContext.auditLogs.Add([PSCustomObject]@{
                        Action  = "UpdateAccount"
                        Message = "Successfully updated MS Teams Phonenumber Assignment where [NumberType] = [$($correlatedAccount.PhoneNumberType)] for [$($correlatedAccount.userPrincipalName)] to [$($correlatedAccount.PhoneNumber)]. (Old value: [$($actionContext.Data.PhoneNumber)])"
                        IsError = $false
                    })
    	        
                $setOnlineVoiceRoutingPolicySplatParams = @{
                    Identity        = $correlatedAccount.id
                    PolicyName      = $OnlineVoiceRoutingPolicyName
                    Verbose         = $false
                    ErrorAction     = "Stop"
                }

                Grant-CsOnlineVoiceRoutingPolicy @setOnlineVoiceRoutingPolicySplatParams
                
                $outputContext.auditLogs.Add([PSCustomObject]@{
                        Action  = "GrantPermission"
                        Message = "Successfully updated MS Teams online voice routing policy for [$($correlatedAccount.userPrincipalName)] to [$($OnlineVoiceRoutingPolicyName)]"
                        IsError = $false
                    })
            }
            else {
                $outputContext.auditLogs.Add([PSCustomObject]@{
                        # Action  = "" # Optional
                        Message = "[DryRun] Would update MS Teams Phonenumber Assignment where [NumberType] = [$($correlatedAccount.PhoneNumberType)] for [$($correlatedAccount.userPrincipalName)]. Old value: [$($actionContext.Data.PhoneNumber)]. New value: [$($correlatedAccount.PhoneNumber)]"
                        IsError = $false
                    })
                $outputContext.auditLogs.Add([PSCustomObject]@{
                        Action  = "GrantPermission"
                        Message = "Successfully updated MS Teams online voice routing policy for [$($correlatedAccount.userPrincipalName)] to [$($OnlineVoiceRoutingPolicyName)]"
                        IsError = $false
                    })

            }
            $outputContext.Data = $correlatedAccount
            $outputContext.AccountReference.PhoneNumber = $correlatedAccount.PhoneNumber
            $outputContext.Success = $true
            break
        }

        'NoChanges' {
            Write-Information "No changes for MS Teams account with accountReference: [$($actionContext.References.Account.id)]"
            $outputContext.Data = $correlatedAccount
            $outputContext.Success = $true
            if (-Not $actionContext.AccountCorrelated) {
                $outputContext.AuditLogs.Add([PSCustomObject]@{
                    Message = 'No changes are  made to the MS Teams account'
                    IsError = $false
                })
            }
            break
        }

        'NotFound' {
            $outputContext.Success  = $false
            $outputContext.AuditLogs.Add([PSCustomObject]@{
                Message = "MS Teams account with accountReference: [$($actionContext.References.Account.id)] could not be found, possibly indicating that it could be deleted, or the account is not correlated"
                IsError = $true
            })
            break
        }
    }
    Write-Verbose "OuputContext at end of script: $($outputContext | ConvertTo-Json -Depth 10)"
} catch {
    $outputContext.Success = $false
    $ex = $PSItem
    if ($($ex.Exception.GetType().FullName -eq 'Microsoft.PowerShell.Commands.HttpResponseException') -or
        $($ex.Exception.GetType().FullName -eq 'System.Net.WebException')) {
        $errorObj = Resolve-MicrosoftGraphAPIError -ErrorObject $ex
        $auditMessage = "Could not update MS Teams account. Error: $($errorObj.FriendlyMessage)"
        Write-Warning "Error at Line '$($errorObj.ScriptLineNumber)': $($errorObj.Line). Error: $($errorObj.ErrorDetails)"
    } else {
        $auditMessage = "Could not update MS Teams account. Error: $($ex.Exception.Message)"
        Write-Warning "Error at Line '$($ex.InvocationInfo.ScriptLineNumber)': $($ex.InvocationInfo.Line). Error: $($ex.Exception.Message)"
    }
    $outputContext.AuditLogs.Add([PSCustomObject]@{
            Message = $auditMessage
            IsError = $true
        })
}