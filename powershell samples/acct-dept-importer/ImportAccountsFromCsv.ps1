# Params
Param(
	[Parameter(Mandatory = $true)]
	[System.IO.FileInfo]$fileLocation,
    [Parameter(Mandatory = $true)]
    [string]$apiBaseUri,
    [Parameter(Mandatory = $true)]
    [string]$apiWSBeid,
    [Parameter(Mandatory = $true)]
    [string]$apiWSKey
)

# Global Variables
$scriptPath = Split-Path -parent $MyInvocation.MyCommand.Definition
$scriptName = ($MyInvocation.MyCommand.Name -split '\.')[0]
$logFile = "$scriptPath\$scriptName.log"
$processingLoopSeparator = "--------------------------------------------------"
$customAttrColPrefix = "customattribute-"

#region PS Functions
#############################################################
##                    Helper Functions                     ##
#############################################################

function Write-Log {
    
    param (
    
        [ValidateSet('ERROR', 'INFO', 'VERBOSE', 'WARN')]
        [Parameter(Mandatory = $true)]
        [string]$level,

        [Parameter(Mandatory = $true)]
        [string]$string

    )
    
    $logString = (Get-Date).toString("yyyy-MM-dd HH:mm:ss") + " [$level] $string"
    #Add-Content -Path $logFile -Value $logString -Force
	[System.IO.File]::AppendAllText($logFile, $logString + "`r`n")
	
    $foregroundColor = $host.ui.RawUI.ForegroundColor
    $backgroundColor = $host.ui.RawUI.BackgroundColor

    Switch ($level) {
    
        {$_ -eq 'VERBOSE' -or $_ -eq 'INFO'} {
            
            Out-Host -InputObject "$logString"
            
        }

        {$_ -eq 'ERROR'} {

            $host.ui.RawUI.ForegroundColor = "Red"
            $host.ui.RawUI.BackgroundColor = "Black"

            Out-Host -InputObject "$logString"
    
            $host.ui.RawUI.ForegroundColor = $foregroundColor
            $host.UI.RawUI.BackgroundColor = $backgroundColor

        }

        {$_ -eq 'WARN'} {
    
            $host.ui.RawUI.ForegroundColor = "Yellow"
            $host.ui.RawUI.BackgroundColor = "Black"

            Out-Host -InputObject "$logString"
    
            $host.ui.RawUI.ForegroundColor = $foregroundColor
            $host.UI.RawUI.BackgroundColor = $backgroundColor

        }
    
    }
    
}

function PrintAsciiArtAndCredits {

Write-Host " _____________   __   ___           _       _______           _   
|_   _|  _  \ \ / /  / _ \         | |     / /  _  \         | |  
  | | | | | |\ V /  / /_\ \ ___ ___| |_   / /| | | |___ _ __ | |_ 
  | | | | | |/   \  |  _  |/ __/ __| __| / / | | | / _ \ '_ \| __|
  | | | |/ // /^\ \ | | | | (_| (__| |_ / /  | |/ /  __/ |_) | |_ 
 _\_/_|___/ \/   \/ \_| |_/\___\___|\__/_/   |___/ \___| .__/ \__|
|_   _|                         | |                    | |        
  | | _ __ ___  _ __   ___  _ __| |_ ___ _ __          |_|        
  | || '_ `` _ \| '_ \ / _ \| '__| __/ _ \ '__|                    
 _| || | | | | | |_) | (_) | |  | ||  __/ |                       
 \___/_| |_| |_| .__/ \___/|_|   \__\___|_|                       
               | |                                                
               |_|
"

Write-Host '       ,@@@@@g      g@@@@@g                         ,@@@@@@@@g,
      ]@@@@@@@@    $@@@@@@@K                 ,ggggg@@"      `%@@,
      ''@@@@@@@@    ]@@@@@@@P               g@@P""""*           $@w
        "*NNP"      ''*NN*"                @@"                  ]@@
    g@@@@@g $@@,  ,@@@ g@@@@@@           ]@@       g@@g        ]@@g
   $@@@@@@@@ @@@  $@@P]@@@@@@@@       ,@@@NN     g@@@@@@@,       *B@g
   ]@@@@@@@F-*" ,, **M"@@@@@@@P      @@P-      4RRR@@@@RRRN        ]@@
    ,*MRM",  ,@@@@@@g  ,"MRM",      @@P            @@@@             @@
  ,@@@@@@@@@ @@@@@@@@P]@@@@@@@@w    $@C            @@@@            ,@@
  $@@@@@@@@@ %@@@@@@@ @@@@@@@@@@    "@@,           PPPP           g@@-
   `"******`  ,***",, ''"*****"`       %@@Ngggggggggggggggggggggg@@@"
            g@@@@@@@@@                  ''""*******************^""
            @@@@@@@@@@P             
             `"^""*""
'

	Write-Host "  Written by Matt Sayers."
	Write-Host "  Version 10.2.0"
	Write-Host " "

}

function GetNewAccountObject {
    
    # Create a fresh account object.
    # This will need to be updated should the Account object ever change.
    $newAcct = [PSCustomObject]@{
		ID=0;
        Name="";
        IsActive=$true;
        Address1="";
        Address2="";
        Address3="";
        Address4="";
        City="";
        StateAbbr="";
        PostalCode="";
        Country="";
        Phone="";
        Fax="";
        Url="";
        Notes="";
        Code="";
        IndustryID=0;
        ManagerUID=[GUID]::Empty;
        Attributes=@()
    }

    # Return the account object.
    return $newAcct

}

function DoesCsvFileContainColumn {
    Param(
        [Array]$csvColumnHeaders,
        [string]$columnToFind
    )

    $columnFound = (($csvColumnHeaders | Where-Object { $_ -eq $columnToFind } | Select-Object -First 1).Count -eq 1)
    return $columnFound

}

function IsChoiceBasedCustomAttribute {
	param (
		[string]$fieldType
	)

	$customAttributeChoiceFieldTypes = @("dropdown", "multiselect", "hradio", "vradio", "checkboxlist")
	$isChoiceBased = $customAttributeChoiceFieldTypes.Contains($fieldType.ToLower())
	return $isChoiceBased

}

function UpdateAccountFromCsv {
    Param (
		[Int32]$rowIndex,
		$userData,
		$customAttributeData,
        [Array]$csvColumnHeaders,
        [Array]$csvRecord,
        [PSCustomObject]$acctToImport        
    )

	# Name and AccountCode (always required)
	$acctToImport.Name = $csvRecord.Name.Trim()
	$acctToImport.Code = $csvRecord.AccountCode.Trim()

	# Now process optional fields.
	# IsActive
	if((DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "IsActive")) {
		
		if($csvRecord.IsActive.Trim() -eq "true") {
			$acctToImport.IsActive = $true
		} else {
			$acctToImport.IsActive = $false
		}

	}

	# Address1
	if((DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "Address1")) {
		$acctToImport.Address1 = $csvRecord.Address1.Trim()
	}

	# Address2
	if((DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "Address2")) {
		$acctToImport.Address2 = $csvRecord.Address2.Trim()
	}

	# Address3
	if((DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "Address3")) {
		$acctToImport.Address3 = $csvRecord.Address3.Trim()
	}

	# Address4
	if((DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "Address4")) {
		$acctToImport.Address4 = $csvRecord.Address4.Trim()
	}

	# City
	if((DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "City")) {
		$acctToImport.City = $csvRecord.City.Trim()
	}

	# StateAbbr
	if((DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "StateAbbr")) {
		$acctToImport.StateAbbr = $csvRecord.StateAbbr.Trim()
	}

	# PostalCode
	if((DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "PostalCode")) {
		$acctToImport.PostalCode = $csvRecord.PostalCode.Trim()
	}

	# Country
	if((DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "Country")) {
		$acctToImport.Country = $csvRecord.Country.Trim()
	}

	# Phone
	if((DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "Phone")) {
		$acctToImport.Phone = $csvRecord.Phone.Trim()
	}

	# Fax
	if((DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "Fax")) {
		$acctToImport.Fax = $csvRecord.Fax.Trim()
	}

	# Url
	if((DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "Url")) {
		$acctToImport.Url = $csvRecord.Url.Trim()
	}

	# Notes
	if((DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "Notes")) {
		$acctToImport.Notes = $csvRecord.Notes.Trim()
	}

	# IndustryID
	if((DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "IndustryID")) {
		
		# Parse this out to an integer. If it doesn't parse or is less than zero, use zero.
		[Int32]$parsedId = $null
		$parseSuccess = [Int32]::TryParse($csvRecord.IndustryID, [ref]$parsedId)
		if($parseSuccess -and $parsedId -gt 0) {
			$acctToImport.IndustryID = $parsedId
		} else {
			$acctToImport.IndustryID = 0
		}

	}

	# ManagerUsername
	$acctToImport = SetAccountManager `
		-rowIndex $rowIndex `
		-userData $userData `
		-csvColumnHeaders $csvColumnHeaders `
		-csvRecord $csvRecord `
		-acctToImport $acctToImport

	# Custom Attributes
	$acctToImport = SetItemCustomAttributeValues `
		-rowIndex $rowIndex `
		-customAttributeData $customAttributeData `
		-csvColumnHeaders $csvColumnHeaders `
		-csvRecord $csvRecord `
		-acctToImport $acctToImport

	# Return the updated account.
    return $acctToImport

}

function SetAccountManager {
	param (
		[Int32]$rowIndex,
		$userData,
		[Array]$csvColumnHeaders,
		[Array]$csvRecord,
        [PSCustomObject]$acctToImport  
	)

	# ManagerUsername
	if((DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "ManagerUsername")) {
	
		# Determine how to properly set manager UID.
		if([string]::IsNullOrWhiteSpace($csvRecord.ManagerUsername)) {
			
			# The manager username column is blank. Send empty GUID to represent no manager.
			$acctToImport.ManagerUID = [GUID]::Empty

		} else {

			# Try to find the manager by username.
			$managerUid = ($userData | Where-Object { $_.Username -eq $csvRecord.ManagerUsername.Trim() } | Select-Object -ExpandProperty UID -First 1)
			if(!$managerUid -or [string]::IsNullOrWhiteSpace($managerUid)) {
				
				# If we had a manager username but could find no match for the new value, leave
				# this field untouched. Instead log a warning and move on.
				Write-Log -level WARN -string ("Row $($rowIndex) - ManagerUsername value of `"$($csvRecord.ManagerUsername.Trim())`"" +
					" is not a valid TeamDynamix username. The account manager value will be left unchanged when the account is saved" +
					" to the server.")

			} else {

				# A user match was found for manager. Set the manager UID properly as a GUID.
				$acctToImport.ManagerUID = [GUID]::Parse($managerUid)
				
			}
		}
	}

	# Return the updated account.
	return $acctToImport

}

function SetItemCustomAttributeValues {
	param (
		[Int32]$rowIndex,
		$customAttributeData,
        [Array]$csvColumnHeaders,
        [Array]$csvRecord,
        [PSCustomObject]$acctToImport  
	)

	# Custom Attributes
	# If there are any columns starting with CustomAttribute-, loop through them and update
	# the custom attribute values on the acct/dept if possible.
	$customAttributeCols = ($csvColumnHeaders | Where-Object { $_.StartsWith($customAttrColPrefix, "OrdinalIgnoreCase") })
	if($customAttributeCols -and @($customAttributeCols).Count -gt 0) {
	
		# Loop through each custom attribute column.
		foreach($customAttributeCol in $customAttributeCols) {

			# First attempt to parse out the actual custom attribute ID.
			$attID = $customAttributeCol.Split("-", 2)[1]
			$parsedAttId = 0
			$parsedSuccessfully = [Int32]::TryParse($attId, [ref]$parsedAttId)
			if($parsedSuccessfully -and $parsedAttId -gt 0) {

				# If we found an attribute ID, validate that attribute exists in list of all acct/dept attributes.
				# If the attribute ID is not valid we cannot use this column.
				$attribute = ($customAttributeData | Where-Object { $_.ID -eq $parsedAttId} | Select-Object -First 1)
				if($attribute) {

					# Get the new attribute value.
					$newAttrVal = ($csvRecord | Select-Object -ExpandProperty $customAttributeCol)

					# If this is a choice-based attribute and the new value is not empty (meaning to clear it out),
					# further processing on the value is necessary.
					$isChoiceBased = IsChoiceBasedCustomAttribute -fieldType $attribute.FieldType
					$allChoicesInvalid = $false
					if($isChoiceBased -and ![string]::IsNullOrWhiteSpace($newAttrVal)) {

						# Split the column value on | to get all possible choice names to map.
						$choiceNames = $newAttrVal.Split("|", [System.StringSplitOptions]::RemoveEmptyEntries)
						if($choiceNames -and @($choiceNames).Count -gt 0) {

							# Instantiate a variable to store the choice IDs in.
							$selectedChoiceIds = ""

							# Loop through all choice names specified and try to find their choice IDs.
							# Only use the valid choice names.
							foreach($choiceName in $choiceNames) {

								$choice = ($attribute.Choices | Where-Object { $_.Name -eq $choiceName } | Select-Object -First 1)
								if($choice) {
									$selectedChoiceIds += "$($choice.ID),"
								}

							}

							# Set the new attribute value to the list of choice IDs, removing the trailing comma.
							$newAttrVal = $selectedChoiceIds.TrimEnd(",")

							# Now detect if all the choices specified were invalid. We can assume an empty
							# new attribute value is invalid at this point.
							$allChoicesInvalid = [string]::IsNullOrWhiteSpace($newAttrVal)

						}

					}

					# Set the account attribute value now. Remove any leading/trailing spaces on the new value first.
					# If it is a choice attribute and all choice values were invalid though, log a message and
					# do *not* change the attributes current value.
					$newAttrVal = $newAttrVal.Trim()
					if($isChoiceBased -and $allChoicesInvalid) {
						
						Write-Log -level WARN -string ("Row $($rowIndex) - Could not set value for custom attribute ID $($attribute.ID)." + 
							" This is a choice-based attribute and all choices specified were invalid. The attribute value will" +
							" be left unchanged when the account is saved to the server.")

					} else {

						$attributeToUpdate = ($acctToImport.Attributes | Where-Object { $_.ID -eq $parsedId } | Select-Object -First 1)
						if($attributeToUpdate) {
						
							# The attribute already exists on this account. Just update its value.
							$attributeToUpdate.Value = $newAttrVal

						} else {
						
							# The attribute does not exist on this account. Create it and set its value properly.
							$acctToImport.Attributes += @{
								ID=$attribute.ID;
								Value=$newAttrVal;
							}

						}					

					}					

				}

			}

		}

	}

	# Return the updated account record.
	return $acctToImport

}

#endregion

#region API Functions
#############################################################
##                    TDX API Functions                    ##
#############################################################

function GetRateLimitWaitPeriodMs {
    param (
        $apiCallResponse
    )

    # Get the rate limit period reset.
    # Be sure to convert the reset date back to universal time because PS conversions will go to machine local.
    $rateLimitReset = ([DateTime]$apiCallResponse.Headers["X-RateLimit-Reset"]).ToUniversalTime()

    # Calculate the actual rate limit period in milliseconds.
    # Add 5 seconds to the period for clock skew just to be safe.
	$duration = New-TimeSpan -Start ((Get-Date).ToUniversalTime()) -End $rateLimitReset
    $rateLimitMsPeriod = $duration.TotalMilliseconds + 5000

    # Return the millisecond rate limit wait.    
    return $rateLimitMsPeriod

}

function ApiAuthenticateAndBuildAuthHeaders {
	param (
		[string]$apiBaseUri,
        [string]$apiWSBeid,
        [string]$apiWSKey
	)
	
	# Set the admin authentication URI and create an authentication JSON body.
	$authUri = $apiBaseUri + "api/auth/loginadmin"
    $authBody = [PSCustomObject]@{
        BEID=$apiWSBeid;
	    WebServicesKey=$apiWSKey;
	} | ConvertTo-Json
	
	# Call the admin login API method and store the returned token.
	# If this part fails, display errors and exit the entire script.
	# We cannot proceed without authentication.
	$authToken = try {
		Invoke-RestMethod -Method Post -Uri $authUri -Body $authBody -ContentType "application/json"
	} catch {

		# Display errors and exit script.
		Write-Log -level ERROR -string "API authentication failed:"
		Write-Log -level ERROR -string ("Status Code - " + $_.Exception.Response.StatusCode.value__)
		Write-Log -level ERROR -string ("Status Description - " + $_.Exception.Response.StatusDescription)
		Write-Log -level ERROR -string ("Error Message - " + $_.ErrorDetails.Message)
		Write-Log -level INFO -string " "
		Write-Log -level ERROR -string "The import cannot proceed when API authentication fails. Please check your authentication settings and try again."
		Write-Log -level INFO -string " "
		Write-Log -level INFO -string "Exiting."
		Write-Log -level INFO -string $processingLoopSeparator
		Exit(1)
		
	}

	# Create an API header object containing an Authorization header with a
	# value of "Bearer {tokenReturnedFromAuthCall}".
	$apiHeadersInternal = @{"Authorization"="Bearer " + $authToken}

	# Return the API headers.
	return $apiHeadersInternal
	
}

function RetrieveAllUsersForOrganization {
	param (
        [System.Collections.Hashtable]$apiHeaders,
		[string]$apiBaseUri
    )

    # Build URI to get all users for the organization.
    $getUserListUri = $apiBaseUri + "/api/people/userlist?isActive=&isEmployee=&userType=User"

    # Get the data.
	$userDataInternal = $null
    try {
		
		# Specifically use Invoke-WebRequest here so that we get response headers.
		# We might need them to deal with rate-limiting.
		$resp = Invoke-WebRequest -Method Get -Headers $apiHeaders -Uri $getUserListUri -ContentType "application/json"
		$userDataInternal = ($resp | ConvertFrom-Json)

	} catch {

		# If we got rate limited, try again after waiting for the reset period to pass.
		$statusCode = $_.Exception.Response.StatusCode.value__
		if($statusCode -eq 429) {

			# Get the amount of time we need to wait to retry in milliseconds.
			$resetWaitInMs = GetRateLimitWaitPeriodMs -apiCallResponse $_.Exception.Response
			Write-Log -level INFO -string "Waiting $(($resetWaitInMs / 1000.0).ToString("N2")) seconds to rety API call due to rate-limiting."

			# Wait to retry now.
			Start-Sleep -Milliseconds $resetWaitInMs

			# Now retry.
			Write-Log -level INFO -string "Retrying API call to retrieve all User typed user data for the organization."
			$userDataInternal = RetrieveAllUsersForOrganization -apiHeaders $apiHeaders -apiBaseUri $apiBaseUri

		} else {

			# Display errors and exit script.
			Write-Log -level ERROR -string "Retrieving all users for the organization failed:"
			Write-Log -level ERROR -string ("Status Code - " + $statusCode)
			Write-Log -level ERROR -string ("Status Description - " + $_.Exception.Response.StatusDescription)
			Write-Log -level ERROR -string ("Error Message - " + $_.ErrorDetails.Message)
			Write-Log -level INFO -string " "
			Write-Log -level ERROR -string "The import cannot proceed when retrieving user data fails. Please check your authentication settings and try again."
			Write-Log -level INFO -string " "
			Write-Log -level INFO -string "Exiting."
			Write-Log -level INFO -string $processingLoopSeparator
			Exit(1)

		}		
		
    }
	
	# Return the user data.
    return $userDataInternal

}

function RetrieveAllAcctDeptsForOrganization {
	param (
        [System.Collections.Hashtable]$apiHeaders,
		[string]$apiBaseUri
    )

        # Build URI to get all acct/dept records for the organization.
        $getAcctsUri = $apiBaseUri + "/api/accounts/search"

        # Set the user authentication URI and create an authentication JSON body.
        $acctSearchBody = [PSCustomObject]@{
            IsActive=$null;
            MaxResults=$null;
        } | ConvertTo-Json

        # Get the data.
        $acctDataInternal = $null
		try {
            
			# Specifically use Invoke-WebRequest here so that we get response headers.
			# We might need them to deal with rate-limiting.
			$resp = Invoke-WebRequest -Method Post -Headers $apiHeaders -Uri $getAcctsUri -Body $acctSearchBody -ContentType "application/json"
			$acctDataInternal = ($resp | ConvertFrom-Json)

        } catch {

			# If we got rate limited, try again after waiting for the reset period to pass.
			$statusCode = $_.Exception.Response.StatusCode.value__
			if($statusCode -eq 429) {

				# Get the amount of time we need to wait to retry in milliseconds.
				$resetWaitInMs = GetRateLimitWaitPeriodMs -apiCallResponse $_.Exception.Response
				Write-Log -level INFO -string "Waiting $(($resetWaitInMs / 1000.0).ToString("N2")) seconds to rety API call due to rate-limiting."

				# Wait to retry now.
				Start-Sleep -Milliseconds $resetWaitInMs

				# Now retry.
				Write-Log -level INFO -string "Retrying API call to retrieve all acct/dept data for the organization."
				$acctDataInternal = RetrieveAllAcctDeptsForOrganization -apiHeaders $apiHeaders -apiBaseUri $apiBaseUri

			} else {

				# Display errors and exit script.
				Write-Log -level ERROR -string "Retrieving all acct/depts for the organization failed:"
				Write-Log -level ERROR -string ("Status Code - " + $statusCode)
				Write-Log -level ERROR -string ("Status Description - " + $_.Exception.Response.StatusDescription)
				Write-Log -level ERROR -string ("Error Message - " + $_.ErrorDetails.Message)
				Write-Log -level INFO -string " "
				Write-Log -level ERROR -string "The import cannot proceed when retrieving acct/dept data fails. Please check your authentication settings and try again."
				Write-Log -level INFO -string " "
				Write-Log -level INFO -string "Exiting."
				Write-Log -level INFO -string $processingLoopSeparator
				Exit(1)
			}            
        }

		# Return the acct/dept data.
        return $acctDataInternal
}

function RetrieveAllAcctAttributesForOrganization {
	param (
        [System.Collections.Hashtable]$apiHeaders,
		[string]$apiBaseUri
    )

    # Build URI to get all acct/dept custom attributes for the organization.
    $getAcctAttributesUri = $apiBaseUri + "/api/attributes/custom?componentId=14"

    # Get the data.
    $attrDataInternal = $null
	try {
		
		# Specifically use Invoke-WebRequest here so that we get response headers.
		# We might need them to deal with rate-limiting.
		$resp = Invoke-WebRequest -Method Get -Headers $apiHeaders -Uri $getAcctAttributesUri -ContentType "application/json"
		$attrDataInternal = ($resp | ConvertFrom-Json)

	} catch {

		# If we got rate limited, try again after waiting for the reset period to pass.
		$statusCode = $_.Exception.Response.StatusCode.value__
		if($statusCode -eq 429) {

			# Get the amount of time we need to wait to retry in milliseconds.
			$resetWaitInMs = GetRateLimitWaitPeriodMs -apiCallResponse $_.Exception.Response
			Write-Log -level INFO -string "Waiting $(($resetWaitInMs / 1000.0).ToString("N2")) seconds to rety API call due to rate-limiting."

			# Wait to retry now.
			Start-Sleep -Milliseconds $resetWaitInMs

			# Now retry.
			Write-Log -level INFO -string "Retrying API call to retrieve all acct/dept custom attribute data for the organization."
			$attrDataInternal = RetrieveAllAcctAttributesForOrganization -apiHeaders $apiHeaders -apiBaseUri $apiBaseUri

		} else {

			# Display errors and exit script.
			Write-Log -level ERROR -string "Retrieving all acct/dept custom attributes for the organization failed:"
			Write-Log -level ERROR -string ("Status Code - " + $statusCode)
			Write-Log -level ERROR -string ("Status Description - " + $_.Exception.Response.StatusDescription)
			Write-Log -level ERROR -string ("Error Message - " + $_.ErrorDetails.Message)
			Write-Log -level INFO -string " "
			Write-Log -level ERROR -string "The import cannot proceed when retrieving acct/dept custom attributes data fails. Please check your authentication settings and try again."
			Write-Log -level INFO -string " "
			Write-Log -level INFO -string "Exiting."
			Write-Log -level INFO -string $processingLoopSeparator
			Exit(1)

		}

    }

	# Return the acct/dept custom attribute data.
    return $attrDataInternal

}

function RetrieveFullAccountDetails {
	param (
		[System.Collections.Hashtable]$apiHeaders,
		[string]$apiBaseUri,
		[Int32]$accountId
	)

	# Build URI to get the full account details.
    $getAcctUri = $apiBaseUri + "/api/accounts/$($accountId)"

	# Get the data.
	$fullAccountInternal = $null
	try {

		# Specifically use Invoke-WebRequest here so that we get response headers.
		# We might need them to deal with rate-limiting.
		$resp = Invoke-WebRequest -Method Get -Headers $apiHeaders -Uri $getAcctUri -ContentType "application/json"
		Write-Log -level INFO -string "Account retrieved successfully."
		$fullAccountInternal = ($resp | ConvertFrom-Json)

	} catch {

		# If we got rate limited, try again after waiting for the reset period to pass.
		$statusCode = $_.Exception.Response.StatusCode.value__
		if($statusCode -eq 429) {

			# Get the amount of time we need to wait to retry in milliseconds.
			$resetWaitInMs = GetRateLimitWaitPeriodMs -apiCallResponse $_.Exception.Response
			Write-Log -level INFO -string "Waiting $(($resetWaitInMs / 1000.0).ToString("N2")) seconds to rety API call due to rate-limiting."

			# Wait to retry now.
			Start-Sleep -Milliseconds $resetWaitInMs

			# Now retry.
			Write-Log -level INFO -string "Retrying API call to retrieve the full acct/dept details for account ID $($accountId)."
			$fullAccountInternal = RetrieveFullAccountDetails -apiHeaders $apiHeaders -apiBaseUri $apiBaseUri -accountId $accountId

		} else {

			# Display errors and continue
			Write-Log -level ERROR -string "Retrieving full acct/dept details for account ID $($accountId) failed:"
			Write-Log -level ERROR -string ("Status Code - " + $statusCode)
			Write-Log -level ERROR -string ("Status Description - " + $_.Exception.Response.StatusDescription)
			Write-Log -level ERROR -string ("Error Message - " + $_.ErrorDetails.Message)
			$fullAccountInternal = $null

		}

	}

	# Return the full acct/dept record.
    return $fullAccountInternal

}

function SaveAccountToApi {
	param (
		[System.Collections.Hashtable]$apiHeaders,
		[string]$apiBaseUri,
		[PSCustomObject]$acctToImport
	)

	# Build URI to save the account. Assume it has to be
	# created by default.
    $saveAcctUri = $apiBaseUri + "/api/accounts"
	$method = "Post"

	# If the account has an ID greater than zero, indicating
	# an existing account, add the ID to the URI and change
	# the save method to a Put.
	if($acctToImport.ID -gt 0) {

		$saveAcctUri += "/$($acctToImport.ID)"
		$method = "Put"

	}

	# Instantiate a save successful variable and
	# a JSON request body representing the account.
	$saveSuccessful = $false
	$acctToImportJson = ($acctToImport | ConvertTo-Json)

	try {

		Invoke-WebRequest -Method $method -Headers $apiHeaders -Uri $saveAcctUri -Body $acctToImportJson -ContentType "application/json"
		Write-Log -level INFO -string "Account saved successfully."
		$saveSuccessful = $true

	} catch {

		# If we got rate limited, try again after waiting for the reset period to pass.
		$statusCode = $_.Exception.Response.StatusCode.value__
		if($statusCode -eq 429) {

			# Get the amount of time we need to wait to retry in milliseconds.
			$resetWaitInMs = GetRateLimitWaitPeriodMs -apiCallResponse $_.Exception.Response
			Write-Log -level INFO -string "Waiting $(($resetWaitInMs / 1000.0).ToString("N2")) seconds to rety API call due to rate-limiting."

			# Wait to retry now.
			Start-Sleep -Milliseconds $resetWaitInMs

			# Now retry the save.
			Write-Log -level INFO -string "Retrying API call to save account."
			$saveSuccessful = SaveAccountToApi -apiHeaders $apiHeaders -apiBaseUri $apiBaseUri -acctToImport $acctToImport

		} else {

			# Display errors and continue on.
			Write-Log -level ERROR -string "Saving the account record to the API failed. See the following log messages for more details."
			Write-Log -level ERROR -string ("Status Code - " + $_.Exception.Response.StatusCode.value__)
			Write-Log -level ERROR -string ("Status Description - " + $_.Exception.Response.StatusDescription)
			Write-Log -level ERROR -string ("Error Message - " + $_.ErrorDetails.Message)

		}

	}

	# Return whether or not the save was successful.
	return $saveSuccessful

}

#endregion

# Log credits.
PrintAsciiArtAndCredits

# Log starting.
Write-Log -level INFO -string $processingLoopSeparator
Write-Log -level INFO -string "TeamDynamix Acct/Dept import process starting."
Write-Log -level INFO -string "Processing file $($fileLocation)."
Write-Log -level INFO -string " "

# Validate that the file exists.
if (-Not ($fileLocation | Test-Path)) {
	
	Write-Log -level ERROR -string "The specified file location is invalid."
	Write-Log -level INFO -string " "
	Write-Log -level INFO -string "Exiting."
	Write-Log -level INFO -string $processingLoopSeparator
	Exit(1)
	
}

# 1. Read in the contents of the CSV file and count items to process.
$acctCsvData = try {
    Import-Csv $fileLocation
} catch {
    
    Write-Log -level ERROR -string "Error importing CSV: $($_.Exception.Message)"
    Write-Log -level ERROR -string ("If the error message was 'The member [columnName] is already present, your CSV file has duplicate column headers. " +
        "Please remove all duplicate columns and try the import again.")
    Write-Log -level ERROR -string "The import cannot continue if the input CSV data cannot be read. Please correct all errors and try again."
    Write-Log -level INFO -string " "
	Write-Log -level INFO -string "Exiting."
	Write-Log -level INFO -string $processingLoopSeparator
	Exit(1)

}

# Store how many total items to import there are.
$totalItemsCount = @($acctCsvData).Count

# Exit out if no data is found to import.
if($totalItemsCount -le 0) {
	
	Write-Log -level INFO -string "No items detected for processing."
	Write-Log -level INFO -string " "
	Write-Log -level INFO -string "Exiting."
	Write-Log -level INFO -string $processingLoopSeparator
	Exit(0)
	
}

# Now validate that we at least have a Name and an AccountCode column mapped.
# These are the minimum columns required to create an account and support duplicate matching.
$csvColumnHeaders = $acctCsvData[0].PSObject.Properties.Name
if(!(DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "Name")) {
    
    Write-Log -level ERROR -string "The file does not contain the required Name column. A Name column is required for the acct/dept import."
	Write-Log -level INFO -string " "
	Write-Log -level INFO -string "Exiting."
	Write-Log -level INFO -string $processingLoopSeparator
    Exit(1)
    
}

if (!(DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "AccountCode")) {
    
    Write-Log -level ERROR -string "The file does not contain the required AccountCode column. An AccountCode column is required for the acct/dept import."
	Write-Log -level INFO -string " "
	Write-Log -level INFO -string "Exiting."
	Write-Log -level INFO -string $processingLoopSeparator
    Exit(1)

}

# We found data. Proceed.
Write-Log -level INFO -string "Found $($totalItemsCount) item(s) to process."

# 2. Authenticate to the API with BEID and Web Service Key and get an auth token. 
#	 If API authentication fails, display error and exit.
#	 If API authentication succeeds, store the token in a Headers object.
Write-Log -level INFO -string "Authenticating to the TeamDynamix Web API with a base URL of $($apiBaseUri)."
$apiHeaders = ApiAuthenticateAndBuildAuthHeaders -apiBaseUri $apiBaseUri -apiWSBeid $apiWSBeid -apiWSKey $apiWSKey
Write-Log -level INFO -string "Authentication successful."

# 3. Retrieve all acct/dept records for this organization for dupe matching.
Write-Log -level INFO -string "Retrieving all acct/dept data for the organization (for dupe matching) from the TeamDynamix Web API."
$acctData = RetrieveAllAcctDeptsForOrganization -apiHeaders $apiHeaders -apiBaseUri $apiBaseUri
$acctCount = 0
if($acctData -and @($acctData).Count -gt 0) {
    $acctCount = @($acctData).Count
}
Write-Log -level INFO -string "Found $($acctCount) acct/dept record(s) for this organization."

# 4. Retrieve all acct/dept custom attribute data for the organization if we have any custom attribute columns mapped.
$acctCustAttrData = @()
if(($csvColumnHeaders | Where-Object { $_.StartsWith($customAttrColPrefix, "OrdinalIgnoreCase") }).Count -gt 0) {

	Write-Log -level INFO -string "Retrieving all acct/dept custom attribute data for the organization from the TeamDynamix Web API."
	$acctCustAttrData = RetrieveAllAcctAttributesForOrganization -apiHeaders $apiHeaders -apiBaseUri $apiBaseUri
	$attrCount = 0
	if($acctCustAttrData -and @($acctCustAttrData).Count -gt 0) {
		$attrCount = @($acctCustAttrData).Count
	}
	Write-Log -level INFO -string "Found $($attrCount) acct/dept custom attribute(s) for this organization."

}

# 5. Retrieve all user records for the organization for manager matching if the ManagerUsername column is mapped.
$userData = @()
if((DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "ManagerUsername")) {

	Write-Log -level INFO -string "Retrieving all User typed user data for the organization (for manager matching) from the TeamDynamix Web API."
	$userData = RetrieveAllUsersForOrganization -apiHeaders $apiHeaders -apiBaseUri $apiBaseUri
	$userCount = 0
	if($userData -and @($userData).Count -gt 0) {
		$userCount = @($userData).Count
	}
	Write-Log -level INFO -string "Found $($userCount) user(s) for this organization."

}

# 7. Now loop through the CSV data and save it.
Write-Log -level INFO -string "Starting processing of CSV data."
$rowIndex = 2
$createCount = 0
$updateCount = 0
$successCount = 0
$skipCount = 0
$failCount = 0
foreach($acctCsvRecord in $acctCsvData) {

	# Log information about the row we are processing.
	$acctName = $acctCsvRecord.Name
	$acctCode = $acctCsvRecord.AccountCode
	Write-Log -level INFO -string "Processing row $($rowIndex) with a Name value of `"$($acctName)`" and an AccountCode value of `"$($acctCode)`"."

    # Validate that we have a Name value.
	# If this is a null or empty string, skip this row.
    if([string]::IsNullOrWhiteSpace($acctName)) {

        Write-Log -level ERROR -string "Row $($rowIndex) - Required field Name has no value. This row will be skipped."
        $failCount += 1
        $rowIndex += 1
        continue

    }

    # Validate the AccountCode value for duplicate matching.
    # If this is a null or empty string, skip this row.
    if([string]::IsNullOrWhiteSpace($acctCode)) {
        
        Write-Log -level ERROR -string "Row $($rowIndex) - Required field AccountCode has no value. This row will be skipped."
        $failCount += 1
        $rowIndex += 1
        continue

    }

    # Now see if we are dealing with an existing account or not.
    # If this is a new account, new up a fresh object.
    $acctToImport = ($acctData | Where-Object { $_.Code -eq $acctCode} | Select-Object -First 1)
    if(!$acctToImport) {
        $acctToImport = GetNewAccountObject
    } else {

		# See if we need to load up the full details for the existing account or not.
		# An ID greater than zero indicates we are updating an existing account.
		# AN ID of 0 or less means we already created this account earlier in the script.
		if($acctToImport.ID -gt 0) {

			# Get the full acct/dept information here. If this fails, the row has to be skipped.
			# We cannot risk saving over an existing account and potentially erasing data.
			Write-Log -level INFO -string "Detected an existing account (ID: $($acctToImport.ID)) on the server to update. Retrieving full account details from the API before updating values."
			$acctId = $acctToImport.ID
			$acctToImport = RetrieveFullAccountDetails -apiHeaders $apiHeaders -apiBaseUri $apiBaseUri -accountId $acctId
			if(!$acctToImport) {

				# Log that we did not find the account and skip this row.
				Write-Log -level ERROR -string ("Row $($rowIndex) - Detected existing account (ID: $($acctId)) to update but the account" +
					" could not be retrieved from the API. This row will be skipped since saving the import data might" +
					" unintentially clear other fields on the account server-side.")
				$failCount += 1
				$rowIndex += 1
				continue
			}

		} else {

			# If we did find the account, but the ID is 0, this means we already created an account
			# for this account code in the same sheet. Log a warning and skip this row.
			Write-Log -level WARN -string ("Row $($rowIndex) - An acount with Name `"$($acctName)`" and AccountCode" +
				" `"$($acctCode)`" already had to be created earlier in this processing loop. To not create duplicate" +
				" records on the server, this row will be skipped.")
			$skipCount += 1
			$rowIndex += 1
			continue

		}		

	}

    # Update the account from the CSV data as needed.
    $acctToImport = UpdateAccountFromCsv `
		-rowIndex $rowIndex `
		-userData $userData `
		-customAttributeData $acctCustAttrData `
		-csvColumnHeaders $csvColumnHeaders `
		-csvRecord $acctCsvRecord `
		-acctToImport $acctToImport

    # Save account to API here
	Write-Log -level INFO -string "Saving account with Name `"$($acctName)`" and AccountCode `"$($acctCode)`" to the TeamDynamix API."
	$saveSuccess = SaveAccountToApi -apiHeaders $apiHeaders -apiBaseUri $apiBaseUri -acctToImport $acctToImport

    # If the save was successful, increment the success counter and add the new choice to list of existing choices.
	# If the save failed, increment the fail counter and spit out a message.
	if($saveSuccess) {
		
		# Add the new acct/dept to the collection of existing acct/depts if it was newly created.
		if($acctToImport.ID -le 0) {
            
			$acctData += $acctToImport
			$createCount += 1

        } else {
			$updateCount += 1
		}

		# Increment the success counter.
		$successCount += 1

	} else {
		
		# Log an error and increment the fail counter.
		Write-Log -level ERROR -string ("Row $($rowIndex) - The account with a Name value of `"$($acctName)`" and an AccountCode value of" +
			" `"$($acctCode)`" failed to save successfully. See previous errors for details.")
		$failCount += 1

    }

	# Always increment the row counter.
	$rowIndex += 1

}

# Log completion stats now.
Write-Log -level INFO -string "Processing complete."
Write-Log -level INFO -string " "

# Log failures first.
if($failCount -gt 0) {
	Write-Log -level ERROR -string "Failed to saved $($failCount) out of $($totalItemsCount) account(s). See the previous log messages for more details."
}

# Log skips second.
if($skipCount -gt 0) {
	Write-Log -level WARN -string "Skipped $($skipCount) out of $($totalItemsCount) account(s) due to duplicate matching."
}

# Log successes and total stats last.
Write-Log -level INFO -string "Successfully saved $($successCount) out of $($totalItemsCount) account(s)."
Write-Log -level INFO -string "Created $($createCount) account(s) and updated $($updateCount) account(s)."
Write-Log -level INFO -string $processingLoopSeparator