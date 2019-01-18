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

Write-Host " _____________   __  _                     _   _                           
|_   _|  _  \ \ / / | |                   | | (_)                          
  | | | | | |\ V /  | |     ___   ___ __ _| |_ _  ___  _ __                
  | | | | | |/   \  | |    / _ \ / __/ _`` | __| |/ _ \| '_ \               
  | | | |/ // /^\ \ | |___| (_) | (_| (_| | |_| | (_) | | | |              
__\_/_|___/ \/   \/ \_____/\___/_\___\__,_|\__|_|\___/|_| |_| _            
| ___ \                      |_   _|                         | |           
| |_/ /___   ___  _ __ ___     | | _ __ ___  _ __   ___  _ __| |_ ___ _ __ 
|    // _ \ / _ \| '_ `` _ \    | || '_ `` _ \| '_ \ / _ \| '__| __/ _ \ '__|
| |\ \ (_) | (_) | | | | | |  _| || | | | | | |_) | (_) | |  | ||  __/ |   
\_| \_\___/ \___/|_| |_| |_|  \___/_| |_| |_| .__/ \___/|_|   \__\___|_|   
                                            | |                            
                                            |_|
"

Write-Host '   BBBBBBBBBBBBBBBBB                    ,@@@@@@@@g,
   @@@@@gg,       ]@             ,ggggg@@"      `%@@,
   @@@@@@@@@@gg   ]@            g@@P""""*           $@w
   @@@@@@@@@@@@   ]@           @@"                  ]@@
   @@@@@@@@@@@@   ]@         ]@@       g@@g        ]@@g
   @@@@@@@@@@@@   ]@      ,@@@NN     g@@@@@@@,       *B@g
   @@@@@@@@ .@@   ]@     @@P-      4RRR@@@@RRRN        ]@@
   @@@@@@@@ @@@   ]@    @@P            @@@@             @@
   @@@@@@@@ @@@   ]@    $@C            @@@@            ,@@
   @@@@@@@@ *@@   ]@    "@@,           PPPP           g@@-
   @@@@@@@@@@@@   ]@       %@@Ngggggggggggggggggggggg@@@"
   RB@@@@@@@@@@*****        ''""*******************^""
      `"N@@@@@@       
           "*N@
'

	Write-Host "  Written by Matt Sayers."
	Write-Host "  Version 10.2.0"
	Write-Host " "

}

function GetNewRoomObject {
    
    # Create a fresh location room object.
    # This will need to be updated should the LocationRoom object ever change.
    $newRoom = [PSCustomObject]@{
		ID=0;
        Name="";
        Description="";
		ExternalID="";
		Floor="";
		Capacity=$null;
        Attributes=@()
    }

    # Return the location room object.
    return $newRoom

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

function UpdateRoomFromCsv {
    Param (
		[Int32]$rowIndex,
		$customAttributeData,
        [Array]$csvColumnHeaders,
        [Array]$csvRecord,
        [PSCustomObject]$roomToImport        
    )

	# Name and ExternalID (always required)
	$roomToImport.Name = $csvRecord.Name.Trim()
	$roomToImport.ExternalID = $csvRecord.ExternalID.Trim()

	# Now process optional fields.
	# Description
	if((DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "Description")) {
		$roomToImport.Description = $csvRecord.Description.Trim()
	}

	# Floor
	if((DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "Floor")) {
		$roomToImport.Floor = ($csvRecord | Select-Object -ExpandProperty "Floor").Trim()
	}

	# Capacity
	if((DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "Capacity")) {
		
		# Only parse out the capacity if the column has value in it. Otherwise clear it.		
		if([string]::IsNullOrWhiteSpace($csvRecord.Capacity)) {
			$roomToImport.Capacity = $null
		} else {

			# Parse this out to a double. If it doesn't parse or is not between -90.0 to 90.0, use null.
			[Int32]$parsedCapacity = $null
			$parseSuccess = [Int32]::TryParse($csvRecord.Capacity, [ref]$parsedCapacity)
			if($parseSuccess -and $parsedCapacity -ge 0) {
				$roomToImport.Capacity = $parsedCapacity
			} else {

				Write-Log -level WARN -string ("Row $($rowIndex) - An invalid value was detected for Capacity." +
					" Capacity values must be integer values between 0 and 1,000,000. Capacity will be cleared for this row.")
				$roomToImport.Capacity = $null

			}

		}	

	}

	# Custom Attributes
	$roomToImport = SetItemCustomAttributeValues `
		-rowIndex $rowIndex `
		-customAttributeData $customAttributeData `
		-csvColumnHeaders $csvColumnHeaders `
		-csvRecord $csvRecord `
		-roomToImport $roomToImport

	# Return the updated location room.
    return $roomToImport

}

function SetItemCustomAttributeValues {
	param (
		[Int32]$rowIndex,
		$customAttributeData,
        [Array]$csvColumnHeaders,
        [Array]$csvRecord,
        [PSCustomObject]$roomToImport  
	)

	# Custom Attributes
	# If there are any columns starting with CustomAttribute-, loop through them and update
	# the custom attribute values on the location room if possible.
	$customAttributeCols = ($csvColumnHeaders | Where-Object { $_.StartsWith($customAttrColPrefix, "OrdinalIgnoreCase") })
	if($customAttributeCols -and @($customAttributeCols).Count -gt 0) {
	
		# Loop through each custom attribute column.
		foreach($customAttributeCol in $customAttributeCols) {

			# First attempt to parse out the actual custom attribute ID.
			$attID = $customAttributeCol.Split("-", 2)[1]
			$parsedAttId = 0
			$parsedSuccessfully = [Int32]::TryParse($attId, [ref]$parsedAttId)
			if($parsedSuccessfully -and $parsedAttId -gt 0) {

				# If we found an attribute ID, validate that attribute exists in list of all location room attributes.
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

					# Set the location room attribute value now. Remove any leading/trailing spaces on the new value first.
					# If it is a choice attribute and all choice values were invalid though, log a message and
					# do *not* change the attributes current value.
					$newAttrVal = $newAttrVal.Trim()
					if($isChoiceBased -and $allChoicesInvalid) {
						
						Write-Log -level WARN -string ("Row $($rowIndex) - Could not set value for custom attribute ID $($attribute.ID)." + 
							" This is a choice-based attribute and all choices specified were invalid. The attribute value will" +
							" be left unchanged when the location room is saved to the server.")

					} else {

						$attributeToUpdate = ($roomToImport.Attributes | Where-Object { $_.ID -eq $parsedId } | Select-Object -First 1)
						if($attributeToUpdate) {
						
							# The attribute already exists on this location room. Just update its value.
							$attributeToUpdate.Value = $newAttrVal

						} else {
						
							# The attribute does not exist on this location room. Create it and set its value properly.
							$roomToImport.Attributes += @{
								ID=$attribute.ID;
								Value=$newAttrVal;
							}

						}					

					}					

				}

			}

		}

	}

	# Return the updated location room record.
	return $roomToImport

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

function RetrieveAllLocationsAndRoomsForOrganization {
	param (
        [System.Collections.Hashtable]$apiHeaders,
		[string]$apiBaseUri
    )

        # Build URI to get all location and room records for the organization.
        $getLocationsUri = $apiBaseUri + "/api/locations/search"

        # Set the user authentication URI and create an authentication JSON body.
        $locationSearchBody = [PSCustomObject]@{
            IsActive=$null;
			IsRoomRequired=$null;
			ReturnRooms=$true;
            MaxResults=$null;
        } | ConvertTo-Json

        # Get the data.
        $locRoomDataInternal = $null
		try {
            
			# Specifically use Invoke-WebRequest here so that we get response headers.
			# We might need them to deal with rate-limiting.
			$resp = Invoke-WebRequest -Method Post -Headers $apiHeaders -Uri $getLocationsUri -Body $locationSearchBody -ContentType "application/json"
			$locRoomDataInternal = ($resp | ConvertFrom-Json)

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
				Write-Log -level INFO -string "Retrying API call to retrieve all location and room data for the organization."
				$locRoomDataInternal = RetrieveAllLocationsAndRoomsForOrganization -apiHeaders $apiHeaders -apiBaseUri $apiBaseUri

			} else {

				# Display errors and exit script.
				Write-Log -level ERROR -string "Retrieving all locations and rooms for the organization failed:"
				Write-Log -level ERROR -string ("Status Code - " + $statusCode)
				Write-Log -level ERROR -string ("Status Description - " + $_.Exception.Response.StatusDescription)
				Write-Log -level ERROR -string ("Error Message - " + $_.ErrorDetails.Message)
				Write-Log -level INFO -string " "
				Write-Log -level ERROR -string "The import cannot proceed when retrieving location and room data fails. Please check your authentication settings and try again."
				Write-Log -level INFO -string " "
				Write-Log -level INFO -string "Exiting."
				Write-Log -level INFO -string $processingLoopSeparator
				Exit(1)
			}            
        }

		# Return the location and room data.
        return $locRoomDataInternal
}

function RetrieveAllLocationRoomAttributesForOrganization {
	param (
        [System.Collections.Hashtable]$apiHeaders,
		[string]$apiBaseUri
    )

    # Build URI to get all location room custom attributes for the organization.
    $getLocationRoomAttributesUri = $apiBaseUri + "/api/attributes/custom?componentId=80"

    # Get the data.
    $attrDataInternal = $null
	try {
		
		# Specifically use Invoke-WebRequest here so that we get response headers.
		# We might need them to deal with rate-limiting.
		$resp = Invoke-WebRequest -Method Get -Headers $apiHeaders -Uri $getLocationRoomAttributesUri -ContentType "application/json"
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
			Write-Log -level INFO -string "Retrying API call to retrieve all location room custom attribute data for the organization."
			$attrDataInternal = RetrieveAllLocationRoomAttributesForOrganization -apiHeaders $apiHeaders -apiBaseUri $apiBaseUri

		} else {

			# Display errors and exit script.
			Write-Log -level ERROR -string "Retrieving all location room custom attributes for the organization failed:"
			Write-Log -level ERROR -string ("Status Code - " + $statusCode)
			Write-Log -level ERROR -string ("Status Description - " + $_.Exception.Response.StatusDescription)
			Write-Log -level ERROR -string ("Error Message - " + $_.ErrorDetails.Message)
			Write-Log -level INFO -string " "
			Write-Log -level ERROR -string "The import cannot proceed when retrieving location room custom attributes data fails. Please check your authentication settings and try again."
			Write-Log -level INFO -string " "
			Write-Log -level INFO -string "Exiting."
			Write-Log -level INFO -string $processingLoopSeparator
			Exit(1)

		}

    }

	# Return the location room custom attribute data.
    return $attrDataInternal

}

function RetrieveFullLocationRoomDetails {
	param (
		[System.Collections.Hashtable]$apiHeaders,
		[string]$apiBaseUri,
		[Int32]$locationId,
		[Int32]$roomId
	)

	# Build URI to get the full location room details.
    $getRoomUri = $apiBaseUri + "/api/locations/$($locationId)/rooms/$($roomId)"

	# Get the data.
	$fullRoomInternal = $null
	try {

		# Specifically use Invoke-WebRequest here so that we get response headers.
		# We might need them to deal with rate-limiting.
		$resp = Invoke-WebRequest -Method Get -Headers $apiHeaders -Uri $getRoomUri -ContentType "application/json"
		Write-Log -level INFO -string "Location room retrieved successfully."
		$fullRoomInternal = ($resp | ConvertFrom-Json)

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
			Write-Log -level INFO -string "Retrying API call to retrieve the full location room details for location ID $($locationId) and room ID $($roomId)."
			$fullRoomInternal = RetrieveFullLocationRoomDetails -apiHeaders $apiHeaders -apiBaseUri $apiBaseUri -locationId $locationId -roomId $roomId

		} else {

			# Display errors and continue
			Write-Log -level ERROR -string "Retrieving full location room details for location ID $($locationId) and room ID $($roomId) failed:"
			Write-Log -level ERROR -string ("Status Code - " + $statusCode)
			Write-Log -level ERROR -string ("Status Description - " + $_.Exception.Response.StatusDescription)
			Write-Log -level ERROR -string ("Error Message - " + $_.ErrorDetails.Message)
			$fullLocationInternal = $null

		}

	}

	# Return the full location record.
    return $fullRoomInternal

}

function SaveLocationRoomToApi {
	param (
		[System.Collections.Hashtable]$apiHeaders,
		[string]$apiBaseUri,
		[Int32]$locationId,
		[PSCustomObject]$roomToImport
	)

	# Build URI to save the location room. Assume it has to be
	# created by default.
    $saveRoomUri = $apiBaseUri + "/api/locations/" + $locationId + "/rooms"
	$method = "Post"

	# If the location room has an ID greater than zero, indicating
	# an existing room, add the ID to the URI and change
	# the save method to a Put.
	if($roomToImport.ID -gt 0) {

		$saveRoomUri += "/$($roomToImport.ID)"
		$method = "Put"

	}

	# Instantiate a save successful variable and
	# a JSON request body representing the location room.
	$saveSuccessful = $false
	$roomToImportJson = ($roomToImport | ConvertTo-Json)

	try {

		Invoke-WebRequest -Method $method -Headers $apiHeaders -Uri $saveRoomUri -Body $roomToImportJson -ContentType "application/json"
		Write-Log -level INFO -string "Location room saved successfully."
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
			Write-Log -level INFO -string "Retrying API call to save location room."
			$saveSuccessful = SaveLocationRoomToApi -apiHeaders $apiHeaders -apiBaseUri $apiBaseUri -locationId $locationId -roomToImport $roomToImport

		} else {

			# Display errors and continue on.
			Write-Log -level ERROR -string "Saving the location room record to the API failed. See the following log messages for more details."
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
Write-Log -level INFO -string "TeamDynamix Location Room import process starting."
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
$locRoomCsvData = try {
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
$totalItemsCount = @($locRoomCsvData).Count

# Exit out if no data is found to import.
if($totalItemsCount -le 0) {
	
	Write-Log -level INFO -string "No items detected for processing."
	Write-Log -level INFO -string " "
	Write-Log -level INFO -string "Exiting."
	Write-Log -level INFO -string $processingLoopSeparator
	Exit(0)
	
}

# Now validate that we at least have LocationExternalID, Name and ExternalID columns mapped.
# These are the minimum columns required to create a location room and support duplicate matching.
$csvColumnHeaders = $locRoomCsvData[0].PSObject.Properties.Name
if (!(DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "LocationExternalID")) {
    
    Write-Log -level ERROR -string "The file does not contain the required LocationExternalID column. A LocationExternalID column is required for the location room import."
	Write-Log -level INFO -string " "
	Write-Log -level INFO -string "Exiting."
	Write-Log -level INFO -string $processingLoopSeparator
    Exit(1)

}

if(!(DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "Name")) {
    
    Write-Log -level ERROR -string "The file does not contain the required Name column. A Name column is required for the location room import."
	Write-Log -level INFO -string " "
	Write-Log -level INFO -string "Exiting."
	Write-Log -level INFO -string $processingLoopSeparator
    Exit(1)
    
}

if (!(DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "ExternalID")) {
    
    Write-Log -level ERROR -string "The file does not contain the required ExternalID column. An ExternalID column is required for the location room import."
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

# 3. Retrieve all location and room records for this organization for dupe matching.
Write-Log -level INFO -string "Retrieving all location and room data for the organization (for dupe matching) from the TeamDynamix Web API."
$locRoomData = RetrieveAllLocationsAndRoomsForOrganization -apiHeaders $apiHeaders -apiBaseUri $apiBaseUri
$locationCount = 0
$roomCount = 0
if($locRoomData -and @($locRoomData).Count -gt 0) {
	
	$locationCount = @($locRoomData).Count
	$roomCount = ($locRoomData | Measure-Object -Property "RoomsCount" -Sum).Sum

}
Write-Log -level INFO -string "Found $($locationCount) location record(s) and $($roomCount) room record(s) for this organization."

# 4. Retrieve all location room custom attribute data for the organization if we have any custom attribute columns mapped.
$locRoomCustAttrData = @()
if(($csvColumnHeaders | Where-Object { $_.StartsWith($customAttrColPrefix, "OrdinalIgnoreCase") }).Count -gt 0) {

	Write-Log -level INFO -string "Retrieving all location custom attribute data for the organization from the TeamDynamix Web API."
	$locRoomCustAttrData = RetrieveAllLocationRoomAttributesForOrganization -apiHeaders $apiHeaders -apiBaseUri $apiBaseUri
	$attrCount = 0
	if($locRoomCustAttrData -and @($locRoomCustAttrData).Count -gt 0) {
		$attrCount = @($locRoomCustAttrData).Count
	}
	Write-Log -level INFO -string "Found $($attrCount) location room custom attribute(s) for this organization."

}

# 5. Now loop through the CSV data and save it.
Write-Log -level INFO -string "Starting processing of CSV data."
$rowIndex = 2
$createCount = 0
$updateCount = 0
$successCount = 0
$skipCount = 0
$failCount = 0
foreach($roomCsvRecord in $locRoomCsvData) {

	# Log information about the row we are processing.
	$locationExtId = $roomCsvRecord.LocationExternalID
	$roomName = $roomCsvRecord.Name
	$roomExtId = $roomCsvRecord.ExternalID	
	Write-Log -level INFO -string "Processing row $($rowIndex) with a Name value of `"$($roomName)`" and an ExternalID value of `"$($roomExtId)`"."

	# Validate that we have a LocationExternalID value to identify the parent location.
    # If this is a null or empty string, skip this row.
    if([string]::IsNullOrWhiteSpace($locationExtId)) {
        
        Write-Log -level ERROR -string "Row $($rowIndex) - Required field LocationExternalID has no value. This row will be skipped."
        $failCount += 1
        $rowIndex += 1
        continue

	}

    # Validate that we have a Name value.
	# If this is a null or empty string, skip this row.
    if([string]::IsNullOrWhiteSpace($roomName)) {

        Write-Log -level ERROR -string "Row $($rowIndex) - Required field Name has no value. This row will be skipped."
        $failCount += 1
        $rowIndex += 1
        continue

    }

    # Validate the ExternalID value for duplicate matching.
    # If this is a null or empty string, skip this row.
    if([string]::IsNullOrWhiteSpace($roomExtId)) {
        
        Write-Log -level ERROR -string "Row $($rowIndex) - Required field ExternalID has no value. This row will be skipped."
        $failCount += 1
        $rowIndex += 1
        continue

    }
		
	# First attempt to find the location this room will be added to. If the parent location
	# record cannot be found, this row cannot be processed.
	$parentLocation = ($locRoomData | Where-Object { $_.ExternalID -eq $locationExtId} | Select-Object -First 1)
	if(!$parentLocation) {
		
		Write-Log -level ERROR -string ("Row $($rowIndex) - The specified LocationExternalID value of `"$($locationExtId)`"" +
			" does not match the ExternalID of any existing location for this organization. As a room must belong to a valid" +
			" location, this row will be skipped.")
        $failCount += 1
        $rowIndex += 1
        continue

	}
	$locationId = $parentLocation.ID

	# Now see if we are dealing with an existing room or not.
	# If this is a new room, new up a fresh object.
	$roomToImport = ($parentLocation | Select-Object -ExpandProperty Rooms | Where-Object { $_.ExternalID -eq $roomExtId } | Select-Object -First 1)
    if(!$roomToImport) {
        $roomToImport = GetNewRoomObject
    } else {

		# See if we need to load up the full details for the existing location room or not.
		# An ID greater than zero indicates we are updating an existing location room.
		# AN ID of 0 or less means we already created this location room earlier in the script.
		if($roomToImport.ID -gt 0) {

			# Get the full location room information here. If this fails, the row has to be skipped.
			# We cannot risk saving over an existing location and potentially erasing data.
			Write-Log -level INFO -string ("Detected an existing location (ID: $($locationId)) and" +
				" location room (ID: $($roomToImport.ID)) on the server to update. Retrieving full location and room details from the API before updating values.")
			$roomId = $roomToImport.ID
			$roomToImport = RetrieveFullLocationRoomDetails -apiHeaders $apiHeaders -apiBaseUri $apiBaseUri -locationId $locationId -roomId $roomId
			if(!$roomToImport) {

				# Log that we did not find the room and skip this row.
				Write-Log -level ERROR -string ("Row $($rowIndex) - Detected existing location (ID: $($locationId))" +
					" and location room (ID: $($roomId)) to update but the location and room information could not be retrieved" +
					" from the API. This row will be skipped since saving the import data might" +
					" unintentially clear other fields on the room server-side.")
				$failCount += 1
				$rowIndex += 1
				continue
			}

		} else {

			# If we did find the location room, but the ID is 0, this means we already created a room
			# for this room external ID in the same sheet. Log a warning and skip this row.
			Write-Log -level WARN -string ("Row $($rowIndex) - A location room with Name `"$($roomName)`" and ExternalID" +
				" `"$($roomExtId)`" already had to be created for location with ID $($locationId) earlier in this processing loop. To not create duplicate" +
				" records on the server, this row will be skipped.")
			$skipCount += 1
			$rowIndex += 1
			continue

		}		

	}

    # Update the location room from the CSV data as needed.
    $roomToImport = UpdateRoomFromCsv `
		-rowIndex $rowIndex `
		-customAttributeData $locRoomCustAttrData `
		-csvColumnHeaders $csvColumnHeaders `
		-csvRecord $roomCsvRecord `
		-roomToImport $roomToImport

    # Save location room to API here
	Write-Log -level INFO -string "Saving location room with Name `"$($roomName)`" and ExternalID `"$($roomExtId)`" for location with ID $($locationId) to the TeamDynamix API."
	$saveSuccess = SaveLocationRoomToApi -apiHeaders $apiHeaders -apiBaseUri $apiBaseUri -locationId $locationId -roomToImport $roomToImport

    # If the save was successful, increment the success counter and add the new choice to list of existing choices.
	# If the save failed, increment the fail counter and spit out a message.
	if($saveSuccess) {
		
		# Add the new location room to the parent location's room collection if it was newly created.
		if($roomToImport.ID -le 0) {
            
			$parentLocation.Rooms += $roomToImport
			$createCount += 1

        } else {
			$updateCount += 1
		}

		# Increment the success counter.
		$successCount += 1

	} else {
		
		# Log an error and increment the fail counter.
		Write-Log -level ERROR -string ("Row $($rowIndex) - The location room with a Name value of `"$($roomName)`" and an ExternalID value of" +
			" `"$($roomExtId)`" for location with ID $($locationId) failed to save successfully. See previous errors for details.")
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
	Write-Log -level ERROR -string "Failed to saved $($failCount) out of $($totalItemsCount) location room(s). See the previous log messages for more details."
}

# Log skips second.
if($skipCount -gt 0) {
	Write-Log -level WARN -string "Skipped $($skipCount) out of $($totalItemsCount) location room(s) due to duplicate matching."
}

# Log successes and total stats last.
Write-Log -level INFO -string "Successfully saved $($successCount) out of $($totalItemsCount) location room(s)."
Write-Log -level INFO -string "Created $($createCount) location room(s) and updated $($updateCount) location room(s)."
Write-Log -level INFO -string $processingLoopSeparator