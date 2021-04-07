# Params
Param(
	[Parameter(Mandatory = $true)]
	[System.IO.FileInfo]$fileLocation,
    [Parameter(Mandatory = $true)]
    [string]$apiBaseUri,
    [Parameter(Mandatory = $true)]
    [string]$apiWSBeid,
    [Parameter(Mandatory = $true)]
    [string]$apiWSKey,
    [Parameter(Mandatory = $true)]
    [int]$attributeId
)

# Force TLS 1.2 Usage
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Global Variables
$scriptPath = Split-Path -parent $MyInvocation.MyCommand.Definition
$scriptName = ($MyInvocation.MyCommand.Name -split '\.')[0]
$logFile = "$scriptPath\$scriptName.log"
$processingLoopSeparator = "--------------------------------------------------"
$choiceColumnName = "ChoiceName"

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

function ApiAuthenticateAndBuildAuthHeaders {
	param (
		[string]$apiBaseUri,
        [string]$apiWSBeid,
        [string]$apiWSKey
	)
	
	# Set the user authentication URI and create an authentication JSON body.
	$authUri = $apiBaseUri + "api/auth/loginadmin"
	$authBody = @{ 
		BEID=$apiWSBeid; 
		WebServicesKey=$apiWSKey
	} | ConvertTo-Json
	
	# Call the user login API method and store the returned tokenn.
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
	$apiHeaders = @{"Authorization"="Bearer " + $authToken}

	# Return the API headers.
	return $apiHeaders
	
}

function GetCurrentAttributeChoices {
	param (
		[System.Collections.Hashtable]$apiHeaders, 
		[string]$apiBaseUri,
		[int]$attributeId
	)
	
	# Set the attribute choice retrieval URI.
	$attributeChoiceUri = $apiBaseUri + "api/attributes/" + $attributeId + "/choices"
	
	# Call the attribute choice retrieval URI.
	# If this part fails, display errors and exit the entire script.
	# We cannot proceed without knowing the existing choices.
	$attributeChoices = try {
		Invoke-RestMethod -Method Get -Headers $apiHeaders -Uri $attributeChoiceUri -ContentType "application/json"
	} catch {

		# Display errors and exit script.
		Write-Log -level ERROR -string "Attribute Choice Retrieval Failed:"
		Write-Log -level ERROR -string ("Status Code - " + $_.Exception.Response.StatusCode.value__)
		Write-Log -level ERROR -string ("Status Description - " + $_.Exception.Response.StatusDescription)
		Write-Log -level ERROR -string ("Error Message - " + $_.ErrorDetails.Message)
		Write-Log -level INFO -string " "
		Write-Log -level ERROR -string "The import cannot proceed without the list of attribute choices. Please check your authentication settings and try again."
		Write-Log -level INFO -string " "
		Write-Log -level INFO -string "Exiting."
		Write-Log -level INFO -string $processingLoopSeparator
		Exit(1)
		
	}

	# Return the attribute choices.
	return $attributeChoices

}

function CreateAttributeChoice {
	param (
		[System.Collections.Hashtable]$apiHeaders, 
		[string]$apiBaseUri,
		[int]$attributeId,
		[PSCustomObject]$newChoice
	)

	# Initialize a return variable.
	$saveSuccessfull = $true

	# Set the attribute choice save URI.
	$createAttributeUri = $apiBaseUri + "api/attributes/" + $attributeId + "/choices"
	$newAttributeJson = $newChoice | ConvertTo-Json

	# Now attempt to save the choice.
	try {
		
		# Save the choice.
		Invoke-RestMethod -Method Post -Headers $apiHeaders -Uri $createAttributeUri -Body $newAttributeJson -ContentType "application/json"        
		
		# Log successful save.
		Write-Log -level INFO -string "New choice saved successfully."

    } catch {

		# Display errors and exit script.
		Write-Log -level ERROR -string "Error saving new attribute choice with Name `"$($newChoice.Name)`":"
		Write-Log -level ERROR -string ("Status Code - " + $_.Exception.Response.StatusCode.value__)
		Write-Log -level ERROR -string ("Status Description - " + $_.Exception.Response.StatusDescription)
		Write-Log -level ERROR -string ("Error Message - " + $_.ErrorDetails.Message)
        Write-Log -level INFO -string " "
        $saveSuccessfull = $false
                
	}	

	# Return whether or not the save was successful.
	return $saveSuccessfull
}

#endregion

Write-Log -level INFO -string $processingLoopSeparator
Write-Log -level INFO -string "Importing process starting."
Write-Log -level INFO -string "Processing file $($fileLocation)."
Write-Log -level INFO -string " "

# Validate that the file location.
if (-Not ($fileLocation | Test-Path)) {
	
	Write-Log -level ERROR -string "The specified file location is invalid."
	Write-Log -level INFO -string " "
	Write-Log -level INFO -string "Exiting."
	Write-Log -level INFO -string $processingLoopSeparator
	Exit(1)
	
}

# Validate that category ID is greater than 0.
if ($attributeId -le 0) {

    Write-Log -level ERROR -string "The specified attribute ID is invalid. Attribute ID must be greater than zero."
	Write-Log -level INFO -string " "
	Write-Log -level INFO -string "Exiting."
	Write-Log -level INFO -string $processingLoopSeparator
    Exit(1)
    
}

# 1. Read in contents of CSV file and count items to process.
$csvData = Import-Csv $fileLocation
$totalItemsCount = @($csvData).Count

# Exit out if no data is found to import.
if($totalItemsCount -le 0) {
	
	Write-Log -level INFO -string "No items detected for processing."
	Write-Log -level INFO -string " "
	Write-Log -level INFO -string "Exiting."
	Write-Log -level INFO -string $processingLoopSeparator
	Exit(0)
	
}

# We found data. Proceed.
Write-Log -level INFO -string "Found $($totalItemsCount) attribute choices to process."

# 2. Authenticate to the API with BEID and Web Service Key and get an auth token. 
#	 If API authentication fails, display error and exit.
#	 If API authentication succeeds, store the token in a Headers object.
Write-Log -level INFO -string "Authenticating to the TeamDynamix Web API with a base URL of $($apiBaseUri)."
$apiHeaders = ApiAuthenticateAndBuildAuthHeaders -apiBaseUri $apiBaseUri -apiWSBeid $apiWSBeid -apiWSKey $apiWSKey

# 3. Get all current choices for this attribute.
$currentChoicesArray = @()
$currentChoicesResp = GetCurrentAttributeChoices -apiHeaders $apiHeaders -apiBaseUri $apiBaseUri -attributeId $attributeId
if(!$currentChoicesResp -or $null -eq $currentChoicesResp) {
	Write-Log -level INFO -string "Found no existing choices for this attribute."
} else {
	
	Write-Log -level INFO -string "Found $(@($currentChoicesResp).Count) existing choice(s) for this attribute."
	foreach($choice in $currentChoicesResp) {
		$currentChoicesArray += $choice.Name
	}
}

# 4. For each row in the CSV file, create a new attribute choice for the specified attribute ID.
#    Only create new choices if they do not already exist.
$rowIndex = 2
$successCount = 0
$skipCount = 0
$failCount = 0
foreach($csvRow in $csvData) {

	# Get the new choice name from the CSV data using the specific choice column name.
	$newChoiceName = ($csvRow | Select-Object -ExpandProperty $choiceColumnName)
	Write-Log -level INFO -string "Processing row $($rowIndex) value of `"$($newChoiceName)`"."

	# Validate that we have an actual value. If there is no value,
	# consider this an error, log a message, and go to the next item in the loop.
	if(!$newChoiceName -or [string]::IsNullOrWhiteSpace(($newChoiceName))) {
		
		Write-Log -level ERROR -string "Row $($rowIndex) choice value is invalid. Choices must have a non-empty and non-whitespace value. This row will be skipped."
		$failCount +=1
		$rowIndex += 1
		continue

	}

	# Determine if the choice to add already exists or not.
	# If it does, skip over it.
	$choiceExists = (($currentChoicesArray | Where-Object { $_ -eq $newChoiceName}).Count -gt 0)
	if($choiceExists) {
		
		# Log a warning, increment the skip and row index counters and go to the next item in the loop.
		Write-Log -level WARN -string "Row $($rowIndex) value of `"$($newChoiceName)`" already exists for this attribute. This row will be skipped."
		$skipCount += 1
		$rowIndex += 1
		continue

	}

	# If we got this far, we need to create the attribute choice. Initialize a new choice object to use.
	$newChoice = [PSCustomObject]@{
		Name=$newChoiceName;
		IsActive=$true;
		Order=0;
	}

	# Save the data!
	Write-Log -level INFO -string "Saving row $($rowIndex) value of `"$($newChoiceName)`"."
	$successful = CreateAttributeChoice -apiHeaders $apiHeaders -apiBaseUri $apiBaseUri -attributeId $attributeId -newChoice $newChoice

	# If the save was successful, increment the success counter and add the new choice to list of existing choices.
	# If the save failed, increment the fail counter and spit out a message.
	if($successful) {
		
		# Add the new choice to the collection of existing choices.
		$currentChoicesArray += $newChoice.Name

		# Increment the success counter.
		$successCount += 1

	} else {
		
		# Log an error and increment the fail counter.
		Write-Log -level ERROR -string "Row $($rowIndex) value of `"$($newChoiceName)`" failed to save successfully. See previous errors for details."
		$failCount += 1

	}

	# Wait 1 second to respect rate limits.
	Start-Sleep -m 1000
	
	# Always increment the row counter.
	$rowIndex += 1

}

# 5. Log processing complete and stats.
Write-Log -level INFO -string " "
Write-Log -level INFO -string "Successfully processed $($successCount + $skipCount) out of $($totalItemsCount) attribute choice(s)."
Write-Log -level INFO -string "Skipped over $($skipCount) out of $($totalItemsCount) attribute choice(s) because they already exist."
Write-Log -level INFO -string "Successfully saved $($successCount) out of $($totalItemsCount) attribute choice(s)."
Write-Log -level INFO -string "Processing complete."
Write-Log -level INFO -string $processingLoopSeparator