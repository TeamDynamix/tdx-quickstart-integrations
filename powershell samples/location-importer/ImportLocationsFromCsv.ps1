# Script Version: 2.0.0

# Params
Param(
	[Parameter(Mandatory = $true)]
	[System.IO.DirectoryInfo]$monitorFolder,
	[Parameter(Mandatory = $true)]
	[System.IO.DirectoryInfo]$processedFolder,
    [Parameter(Mandatory = $true)]
    [string]$apiBaseUri,
    [Parameter(Mandatory = $true)]
    [string]$apiWSBeid,
    [Parameter(Mandatory = $true)]
    [string]$apiWSKey,
	[Parameter(Mandatory = $false)]
	[int]$maxJsonObjectSizeInBytes = 104857600,
	[Parameter(Mandatory = $false)]
	[switch]$verboseLog = $false	
)
# Load System.Web.Extensions Assembly for JSON serializer.
# Requires .NET Framework.
[System.Reflection.Assembly]::LoadWithPartialName("System.Web.Extensions")

# Force TLS 1.2 Usage
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Define a custom JSON deserializer for large datasets.
# Relies on .NET Framework being installed.
$jsonDeserializer = New-Object -TypeName System.Web.Script.Serialization.JavaScriptSerializer
$jsonDeserializer.MaxJsonLength = $maxJsonObjectSizeInBytes #100mb as bytes, default is 2-4mb.

# Global Variables
$scriptPath = Split-Path -parent $MyInvocation.MyCommand.Definition
$scriptName = ($MyInvocation.MyCommand.Name -split '\.')[0]
$logFile = "$scriptPath\$scriptName.log"
$processingLoopSeparator = "--------------------------------------------------"
$processingFileStartLine = "~~~~~ FILE START ~~~~~"
$processingFileEndLine = "~~~~~ FILE END ~~~~~"
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
    
	# First roll log if over 10MB.
	if($logFile | Test-Path) {
		
		# Get the full info for the current log file.
		$currentLog = Get-Item $logFile
		
		# Get log size in MB (from B)
		$currentLogLength = ([double]$currentLog.Length / 1024.00 / 1024.00)
		
		# If the log file exceeds 10mb, roll it by renaming with a timestamped name.
		if($currentLogLength -ge 10) {
			
			$newLogFileName = (Get-Date).ToString("yyyy-MM-dd HHmmssfff") + " " + $currentLog.Name
			Rename-Item -Path $currentLog.FullName -NewName $newLogFileName -Force
			
		}
		
	}
	
	# Next remove old log files if there are more than 10.
	Get-ChildItem -Path $scriptPath -Filter *.log* | `
		Sort CreationTime -Desc | `
		Select -Skip 10 | `
		Remove-Item -Force

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
 _\_/_|___/ \/   \/ \_____/\___/ \___\__,_|\__|_|\___/|_| |_|
|_   _|                         | |                          
  | | _ __ ___  _ __   ___  _ __| |_ ___ _ __                
  | || '_ `` _ \| '_ \ / _ \| '__| __/ _ \ '__|               
 _| || | | | | | |_) | (_) | |  | ||  __/ |                  
 \___/_| |_| |_| .__/ \___/|_|   \__\___|_|                  
               | |                                           
               |_|    
"

Write-Host '         , ,,            ,ggp                      ,@@@@@@@@g,
     ,g@@@ ]@@@gg    ,g@@@@@P               ,ggggg@@"      `%@@,
  ]@@@@@@@ ]@@@@@@P $@@@@@@@P              g@@P""""*           $@w
  ]@@@@@@@ ]@@@@@M` *N@@@@@@P             @@"                  ]@@
  ]@@@@@@@ ]@@P ,ggNgg `%@@@K           ]@@       g@@g        ]@@g
  ]@@@@@@@ ]@` @@@@@@@@K ]@@K        ,@@@NN     g@@@@@@@,       *B@g
  ]@@@@@@@ ]@ ]@@@@@@@@@  @@K       @@P-      4RRR@@@@RRRN        ]@@
  ]@@@@@@@ ]@  B@@@@@@@P ]@@K      @@P            @@@@             @@
  ]@@@@@@@ ]@@g "MNNR*-  %@@K      $@C            @@@@            ,@@
  ]@@@@@@@ ]@@@@@gp gg@@W "%@      "@@,           PPPP           g@@-
  ]@@@@@@P ]N@@@@@P $@@@@@b.gg        %@@Ngggggggggggggggggggggg@@@"
  ]@@P*        "RBP $@P"''    %@        ''''""*******************^""
'

	Write-Host "  Written by Matt Sayers."
	Write-Host "  Version 2.0.0"
	Write-Host " "

}

function ValidateScriptFolders {
	
	# Validate that the monitor folder exists.
	if (-Not ($monitorFolder | Test-Path)) {
		
		Write-Log -level ERROR -string "The specified monitor folder is invalid."
		Write-Log -level INFO -string " "
		Write-Log -level INFO -string "Exiting."
		Write-Log -level INFO -string $processingLoopSeparator
		Exit(1)
		
	}
	
	# Validate that the processed folder exists.
	if (-Not ($processedFolder | Test-Path)) {
		
		Write-Log -level ERROR -string "The specified processed folder is invalid."
		Write-Log -level INFO -string " "
		Write-Log -level INFO -string "Exiting."
		Write-Log -level INFO -string $processingLoopSeparator
		Exit(1)
		
	}
	
}

function RetrievePendingFiles {

	# Get the all .csv files in the monitor folder.
	$retrievalPath = $monitorFolder.FullName + "\*"
	$filesInternal = Get-ChildItem $retrievalPath -File -Include *.csv
	
	# If there are no files, exit gracefully.
	if(!$filesInternal -or @($filesInternal).Count -le 0) {
	
		Write-Log -level INFO -string "No location CSV files were detected for processing."
		Write-Log -level INFO -string " "
		Write-Log -level INFO -string "Exiting."
		Write-Log -level INFO -string $processingLoopSeparator
		Exit(0)		
	
	}
	
	# Return the files.
	return $filesInternal
	
}

function LoadAndValidateCsvFile {
	Param(
		$file
	)
	
	# Initialize a return object.
	$csvValidationResultInternal = [PSCustomObject]@{
		IsValid=$true;
		IsEmpty=$true;
		Data=@();
	}
	
	# Read in the contents of the CSV file and count items to process.
	$locationCsvDataInternal = try {
		Import-Csv $file.FullName
	}
	catch {
		
		Write-Log -level ERROR -string "Error importing CSV: $($_.Exception.Message)"
		Write-Log -level ERROR -string ("If the error message was 'The member [columnName] is already present, your CSV file has duplicate column headers. " +
			"Please remove all duplicate columns and try the import again.")
		Write-Log -level ERROR -string "The import cannot continue if the input CSV data cannot be read. Please correct all errors and try again."
		$csvValidationResultInternal.IsValid = $false
		return $csvValidationResultInternal
		
	}

	# Store how many total items to import there are.
	$totalItemsCount = @($locationCsvDataInternal).Count

	# Exit out if no data is found to import.
	if ($totalItemsCount -le 0) {
		
		Write-Log -level INFO -string "No items detected for processing."
		$csvValidationResultInternal.IsEmpty = $true
		return $csvValidationResultInternal
		
	}
	
	# Now validate that we at least have a Name and an ExternalID column mapped.
	# These are the minimum columns required to create a location and support duplicate matching.
	$csvColumnHeaders = $locationCsvDataInternal[0].PSObject.Properties.Name
	if(!(DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "Name")) {
		
		Write-Log -level ERROR -string "The file does not contain the required Name column. A Name column is required for the location import."
		$csvValidationResultInternal.IsValid = $false
		return $csvValidationResultInternal
		
	}

	if (!(DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "ExternalID")) {
		
		Write-Log -level ERROR -string "The file does not contain the required ExternalID column. An ExternalID column is required for the location import."
		$csvValidationResultInternal.IsValid = $false
		return $csvValidationResultInternal

	}

	# If we got here, the file is valid. Set the data and return the results.
	$csvValidationResultInternal.IsEmpty = $false
	$csvValidationResultInternal.Data = $locationCsvDataInternal
	return $csvValidationResultInternal	
	
}

function GetProcessedModulusForPctCompLogging {
	Param (
		[int]$totalItemsCount
	)

	# Set a modulus count based on total item count, defaulting to 1.
	# >10 and <=50 = mod 5
	# >50 and <=100 = mod 10
	# >100 and <=500 = mod 25
	# >500 and <=1000 = mod 100
	# >1000 and <=10000 = mod 250
	# >10000 and <=100000 = mod 2500
	# >100000 = mod 20000 (!)
	$modCountInternal = 1
	if($totalItemsCount -gt 10 -and $totalItemsCount -le 50) {
		$modCountInternal = 5
	} elseif($totalItemsCount -gt 50 -and $totalItemsCount -le 100) {
		$modCountInternal = 10
	} elseif($totalItemsCount -gt 100 -and $totalItemsCount -le 500) {
		$modCountInternal = 25
	} elseif($totalItemsCount -gt 500 -and $totalItemsCount -le 1000) {
		$modCountInternal = 100
	} elseif($totalItemsCount -gt 1000 -and $totalItemsCount -le 10000) {
		$modCountInternal = 250
	} elseif($totalItemsCount -gt 10000 -and $totalItemsCount -le 100000) {
		$modCountInternal = 2500
	} elseif($totalItemsCount -gt 100000) {
		$modCountInternal = 20000
	}

	# Return the modulus.
	return $modCountInternal
}

function ProcessAllCsvFileData {
	Param(
		$csvRecords
	)

	# Initialize a return object.
	$results = [PSCustomObject]@{
		CreateCount=0;
		UpdateCount=0;
		SuccessCount=0;
		SkipCount=0;
		FailCount=0;
	}

	# Get the total items count.
	$totalCsvRowsInternal = @($csvRecords).Count

	# Set a percent complete modulus count based on total item count, defaulting to 1.
	$modCount = GetProcessedModulusForPctCompLogging -totalItemsCount $totalCsvRowsInternal

	# Process the data.
	$rowIndex = 2
	$csvColumnHeaders = $csvRecords[0].PSObject.Properties.Name
	foreach ($locationCsvRecord in $csvRecords) {

		# Log information about the row we are processing.
		$locationName = $locationCsvRecord.Name
		$locationExtId = $locationCsvRecord.ExternalID

		if($verboseLog) {
			Write-Log -level INFO -string "Processing row $($rowIndex) with a Name value of `"$($locationName)`" and an ExternalID value of `"$($locationExtId)`"."
		}

		# Validate that we have a Name value.
		# If this is a null or empty string, skip this row.
		if([string]::IsNullOrWhiteSpace($locationName)) {

			# Log errors, increment failure count and row index.
			Write-Log -level ERROR -string "Row $($rowIndex) - Required field Name has no value. This row will be skipped."
			$results.FailCount += 1
			$rowIndex += 1

			# Log rows processed.
			$processedCount = ($rowIndex - 2)		
			LogPercentageFileProcessed -modCount $modCount -processedCount $processedCount -totalItemsCount $totalCsvRowsInternal

			# Continue the loop.
			continue

		}

		# Validate the ExternalID value for duplicate matching.
		# If this is a null or empty string, skip this row.
		if([string]::IsNullOrWhiteSpace($locationExtId)) {
			
			# Log errors, increment failure count and row index.
			Write-Log -level ERROR -string "Row $($rowIndex) - Required field ExternalID has no value. This row will be skipped."
			$results.FailCount += 1
			$rowIndex += 1

			# Log rows processed.
			$processedCount = ($rowIndex - 2)		
			LogPercentageFileProcessed -modCount $modCount -processedCount $processedCount -totalItemsCount $totalCsvRowsInternal

			# Continue the loop.
			continue

		}

		# Now see if we are dealing with an existing location or not.
		# If this is a new location, new up a fresh object.
		$locationToImport = ($locationDataG | Where-Object { $_.ExternalID -eq $locationExtId} | Select-Object -First 1)
		if(!$locationToImport) {
			$locationToImport = GetNewLocationObject
		} else {

			# See if we need to load up the full details for the existing location or not.
			# An ID greater than zero indicates we are updating an existing location.
			# AN ID of 0 or less means we already created this location earlier in the script.
			if($locationToImport.ID -gt 0) {

				# Get the full location information here. If this fails, the row has to be skipped.
				# We cannot risk saving over an existing location and potentially erasing data.
				if($verboseLog) {
					Write-Log -level INFO -string "Detected an existing location (ID: $($locationToImport.ID)) on the server to update. Retrieving full location details from the API before updating values."
				}

				$locationId = $locationToImport.ID
				$locationToImport = RetrieveFullLocationDetails -locationId $locationId
				if(!$locationToImport) {

					# Log that we did not find the location and skip this row.
					Write-Log -level ERROR -string ("Row $($rowIndex) - Detected existing location (ID: $($locationId)) to update but the location" +
						" could not be retrieved from the API. This row will be skipped since saving the import data might" +
						" unintentially clear other fields on the location server-side.")
					$results.FailCount += 1
					$rowIndex += 1
					continue

				}

			} else {

				# If we did find the location, but the ID is 0, this means we already created a location
				# for this location code in the same sheet. Log a warning and skip this row.
				Write-Log -level WARN -string ("Row $($rowIndex) - A location with Name `"$($locationName)`" and ExternalID" +
					" `"$($locationExtId)`" already had to be created earlier in this processing loop. To not create duplicate" +
					" records on the server, this row will be skipped.")
				$results.SkipCount += 1
				$rowIndex += 1
				continue

			}		

		}

		# Update the location from the CSV data as needed.
		$locationToImport = UpdateLocationFromCsv `
			-rowIndex $rowIndex `
			-csvColumnHeaders $csvColumnHeaders `
			-csvRecord $locationCsvRecord `
			-locationToImport $locationToImport

		# Store whether or not this will be a newly created record or not.
		$isCreating = ($locationToImport.ID -le 0) 

		# Save location to API here
		if($verboseLog) {
			Write-Log -level INFO -string "Saving location with Name `"$($locationName)`" and ExternalID `"$($locationExtId)`" to the TeamDynamix API."
		}
		$saveSuccess = SaveLocationToApi -apiHeaders $apiHeaders -apiBaseUri $apiBaseUri -locationToImport $locationToImport

		# If the save was successful, increment the success counter and add the new choice to list of existing choices.
		# If the save failed, increment the fail counter and spit out a message.
		if($saveSuccess) {
			
			# Add the new location to the collection of existing locations if it was newly created.
			if($isCreating) {
				
				# Now add the newly created location to the local collection of all location data.
				$locationDataG.Add($locationToImport)
				
				# Increment the creation counter.
				$results.CreateCount += 1				

			} else {
				
				# Increment the update counter.
				$results.UpdateCount += 1

			}

			# Increment the success counter.
			$results.SuccessCount += 1

		} else {
			
			# Log an error and increment the fail counter.
			Write-Log -level ERROR -string ("Row $($rowIndex) - The location with a Name value of `"$($locationName)`" and an ExternalID value of" +
				" `"$($locationExtId)`" failed to save successfully. See previous errors for details.")
			$results.FailCount += 1

		}

		# Always increment the row counter.
		$rowIndex += 1

		# List out numbers of records processed and percent complete.
		$processedCount = ($rowIndex - 2)		
		LogPercentageFileProcessed -modCount $modCount -processedCount $processedCount -totalItemsCount $totalCsvRowsInternal

	}

	# Return the results.
	return $results

}

function LogPercentageFileProcessed {
	Param(
		[int]$modCount,
		[int]$processedCount,
		[int]$totalItemCount
	)

	if($processedCount -ge $totalItemsCount) {
		Write-Log -level INFO -string "Processed $($processedCount)/$($totalItemsCount) record(s) (100.00%)."
	} elseif($processedCount % $modCount -eq 0) {

		$currentPctComp = ([double]$processedCount / [double]$totalItemsCount)
		Write-Log -level INFO -string "Processed $($processedCount)/$($totalItemsCount) record(s) ($($currentPctComp.ToString("P2")))."

	}

}

function GetNewLocationObject {
    
    # Create a fresh location object.
    # This will need to be updated should the Location object ever change.
    $newLocation = [PSCustomObject]@{
		ID=0;
        Name="";
        Description="";
		ExternalID="";
		IsActive=$true;
		Address="";
		City="";
		State="";
		PostalCode="";
		Country="";
		IsRoomRequired=$false;
		Latitude=$null;
		Longitude=$null;
        Attributes=@()
    }

    # Return the location object.
    return $newLocation

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

function UpdateLocationFromCsv {
    Param (
		[Int32]$rowIndex,
        [Array]$csvColumnHeaders,
        [Array]$csvRecord,
        [PSCustomObject]$locationToImport        
    )

	# Name and ExternalID (always required)
	$locationToImport.Name = $csvRecord.Name.Trim()
	$locationToImport.ExternalID = $csvRecord.ExternalID.Trim()

	# Now process optional fields.
	# IsActive
	if((DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "IsActive")) {
		
		if($csvRecord.IsActive.Trim() -eq "true") {
			$locationToImport.IsActive = $true
		} else {
			$locationToImport.IsActive = $false
		}

	}

	# Description
	if((DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "Description")) {
		$locationToImport.Description = $csvRecord.Description.Trim()
	}

	# Address (handle with Select-Object since Address is already a property on the CSV data record custom object.
	if((DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "Address")) {
		$locationToImport.Address = ($csvRecord | Select-Object -ExpandProperty "Address").Trim()
	}

	# City
	if((DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "City")) {
		$locationToImport.City = $csvRecord.City.Trim()
	}

	# StateAbbr
	if((DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "StateAbbr")) {
		$locationToImport.State = $csvRecord.StateAbbr.Trim()
	}

	# PostalCode
	if((DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "PostalCode")) {
		$locationToImport.PostalCode = $csvRecord.PostalCode.Trim()
	}

	# Country
	if((DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "Country")) {
		$locationToImport.Country = $csvRecord.Country.Trim()
	}

	# IsRoomRequired
	if((DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "IsRoomRequired")) {
		
		if($csvRecord.IsRoomRequired.Trim() -eq "true") {
			$locationToImport.IsRoomRequired = $true
		} else {
			$locationToImport.IsRoomRequired = $false
		}

	}

	# Handle latitude and longitude now.
	$latitudeHasValue = $false
	$longitudeHasValue = $false

	# Latitude
	if((DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "Latitude")) {
		
		# Only parse out the latitude if the column has value in it. Otherwise clear it.		
		if([string]::IsNullOrWhiteSpace($csvRecord.Latitude)) {
			$locationToImport.Latitude = $null
		} else {

			# Parse this out to a double. If it doesn't parse or is not between -90.0 to 90.0, use null.
			[double]$parsedLatitude = $null
			$parseSuccess = [double]::TryParse($csvRecord.Latitude, [ref]$parsedLatitude)
			if($parseSuccess -and ($parsedLatitude -ge [double]-90.0) -and ($parsedLatitude -le [double]90.0)) {
				
				$locationToImport.Latitude = $parsedLatitude
				$latitudeHasValue = $true

			} else {

				Write-Log -level WARN -string ("Row $($rowIndex) - An invalid value was detected for Latitude." +
					" Latitude values must be decimal values between -90.0 and 90.0. Latitude will be cleared for this row.")
				$locationToImport.Latitude = $null

			}

		}		

	}

	# Longitude
	if((DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "Longitude")) {
		
		# Only parse out the longitude if the column has value in it. Otherwise clear it.		
		if([string]::IsNullOrWhiteSpace($csvRecord.Longitude)) {
			$locationToImport.Longitude = $null
		} else {

			# Parse this out to a double. If it doesn't parse or is not between -180.0 to 180.0, use null.
			[double]$parsedLongitude = $null
			$parseSuccess = [double]::TryParse($csvRecord.Longitude, [ref]$parsedLongitude)
			if($parseSuccess -and ($parsedLongitude -ge [double]-180.0) -and ($parsedLongitude -le [double]180.0)) {
				
				$locationToImport.Longitude = $parsedLongitude
				$longitudeHasValue = $true

			} else {

				Write-Log -level WARN -string ("Row $($rowIndex) - An invalid value was detected for Longitude." +
					" Longitude values must be decimal values between -180.0 and 180.0. Longitude will be cleared for this row.")				
				$locationToImport.Longitude = $null

			}

		}

	}

	# If the latitude and longitude values were valid but both fields are *not* set,
	# write out a message. You must specify either both values or neither.
	if (($latitudeHasValue -and !$longitudeHasValue) -or ($longitudeHasValue -and !$latitudeHasValue)) {
		
		Write-Log -level WARN -string ("Row $($rowIndex) - Incomplete Latitude and Longitude values were detected." +
			" Either values for both Latitude and Longitude must be specified or both values should be blank." +
			" Latitude and Longitude will be cleared for this row.")				
		$locationToImport.Latitude = $null
		$locationToImport.Longitude = $null

	}

	# Custom Attributes
	$locationToImport = SetItemCustomAttributeValues `
		-rowIndex $rowIndex `
		-csvColumnHeaders $csvColumnHeaders `
		-csvRecord $csvRecord `
		-locationToImport $locationToImport

	# Return the updated location.
    return $locationToImport

}

function SetItemCustomAttributeValues {
	param (
		[Int32]$rowIndex,
        [Array]$csvColumnHeaders,
        [Array]$csvRecord,
        [PSCustomObject]$locationToImport  
	)

	# Custom Attributes
	# If there are any columns starting with CustomAttribute-, loop through them and update
	# the custom attribute values on the location if possible.
	$customAttributeCols = ($csvColumnHeaders | Where-Object { $_.StartsWith($customAttrColPrefix, "OrdinalIgnoreCase") })
	if($customAttributeCols -and @($customAttributeCols).Count -gt 0) {
	
		# Loop through each custom attribute column.
		foreach($customAttributeCol in $customAttributeCols) {

			# First attempt to parse out the actual custom attribute ID.
			$attID = $customAttributeCol.Split("-", 2)[1]
			$parsedAttId = 0
			$parsedSuccessfully = [Int32]::TryParse($attId, [ref]$parsedAttId)
			if($parsedSuccessfully -and $parsedAttId -gt 0) {

				# If we found an attribute ID, validate that attribute exists in list of all location attributes.
				# If the attribute ID is not valid we cannot use this column.
				$attribute = ($locationCustAttrDataG | Where-Object { $_.ID -eq $parsedAttId} | Select-Object -First 1)
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

					# Set the location attribute value now. Remove any leading/trailing spaces on the new value first.
					# If it is a choice attribute and all choice values were invalid though, log a message and
					# do *not* change the attributes current value.
					$newAttrVal = $newAttrVal.Trim()
					if($isChoiceBased -and $allChoicesInvalid) {
						
						Write-Log -level WARN -string ("Row $($rowIndex) - Could not set value for custom attribute ID $($attribute.ID)." + 
							" This is a choice-based attribute and all choices specified were invalid. The attribute value will" +
							" be left unchanged when the location is saved to the server.")

					} else {

						$attributeToUpdate = ($locationToImport.Attributes | Where-Object { $_.ID -eq $parsedId } | Select-Object -First 1)
						if($attributeToUpdate) {
						
							# The attribute already exists on this location. Just update its value.
							$attributeToUpdate.Value = $newAttrVal

						} else {
						
							# The attribute does not exist on this location. Create it and set its value properly.
							$locationToImport.Attributes += @{
								ID=$attribute.ID;
								Value=$newAttrVal;
							}

						}					

					}					

				}

			}

		}

	}

	# Return the updated location record.
	return $locationToImport

}

function MoveFileToProcessed {
	Param(
		$file
	)

	# Create a destination file name with timestamp to uniqueify it.
	$dateNowString = (Get-Date).ToString("yyyy-MM-dd HHmmssffff")
	$destFileName = $dateNowString + " " + $file.Name 
	
	# Create the full destination file path.
	$destFilePath = [System.IO.Path]::Combine($processedFolder.FullName, $destFileName)
	
	try {
	
		Write-Log -level INFO -string "Moving file $($file.Name) to processed folder."
		Move-Item -Path $file.FullName -Destination $destFilePath
		Write-Log -level INFO -string "File $($file.Name) moved successfully."
		
	} 
	catch {
		Write-Log -level ERROR -string "Error moving file $($file.Name) to processed folder:`n$_.Exception.Message"
	}
	
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

function RetrieveAllLocationsForOrganization {

	# Build URI to get all location records for the organization.
	$getLocationsUri = $apiBaseUri + "/api/locations/search"

	# Set the user authentication URI and create an authentication JSON body.
	$locationSearchBody = [PSCustomObject]@{
		IsActive=$null;
		IsRoomRequired=$null;
		MaxResults=$null;
	} | ConvertTo-Json

	# Get the data.
	$locationDataArrayList = [System.Collections.ArrayList]@()
	try {
		
		# Specifically use Invoke-WebRequest here so that we get response headers.
		# We might need them to deal with rate-limiting.
		$resp = Invoke-WebRequest -Method Post -Headers $apiHeaders -Uri $getLocationsUri -Body $locationSearchBody -ContentType "application/json"

		# Use the custom JSON deserializer in case the organization has a large set of locations.
		# Normally we would use this: $locationDataInternal = ($resp | ConvertFrom-Json)
		$locationDataInternal = $jsonDeserializer.Deserialize($resp, [System.Object])

		# Convert the location data to an array list for faster querying.
		if($locationDataInternal -and @($locationDataInternal).Count -gt 0) {
			$locationDataArrayList.AddRange($locationDataInternal)
		}
				
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
			Write-Log -level INFO -string "Retrying API call to retrieve all location data for the organization."
			$locationDataArrayList = RetrieveAllLocationsForOrganization -apiHeaders $apiHeaders -apiBaseUri $apiBaseUri

		} else {

			# Display errors and exit script.
			Write-Log -level ERROR -string "Retrieving all locations for the organization failed:"
			Write-Log -level ERROR -string ("Status Code - " + $statusCode)
			Write-Log -level ERROR -string ("Status Description - " + $_.Exception.Response.StatusDescription)
			Write-Log -level ERROR -string ("Error Message - " + $_.ErrorDetails.Message)
			Write-Log -level INFO -string " "
			Write-Log -level ERROR -string "The import cannot proceed when retrieving location data fails. Please check your authentication settings and try again."
			Write-Log -level INFO -string " "
			Write-Log -level INFO -string "Exiting."
			Write-Log -level INFO -string $processingLoopSeparator
			Exit(1)
		}            
	}

	# Return the location data.
	return $locationDataArrayList

}

function RetrieveAllLocationAttributesForOrganization {

    # Build URI to get all location custom attributes for the organization.
    $getLocationAttributesUri = $apiBaseUri + "/api/attributes/custom?componentId=71"

    # Get the data.
    $attrDataArrayList = [System.Collections.ArrayList]@()
	try {
		
		# Specifically use Invoke-WebRequest here so that we get response headers.
		# We might need them to deal with rate-limiting.
		$resp = Invoke-WebRequest -Method Get -Headers $apiHeaders -Uri $getLocationAttributesUri -ContentType "application/json"

		# Use the custom JSON deserializer in case the organization has a large set of location attributes.
		# Normally we would use this: $attrDataInternal = ($resp | ConvertFrom-Json)
		$attrDataInternal = $jsonDeserializer.Deserialize($resp, [System.Object])

		# Convert the location attribute data to an array list for faster querying.
		if($attrDataInternal -and @($attrDataInternal).Count -gt 0) {
			$attrDataArrayList.AddRange($attrDataInternal)
		}

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
			Write-Log -level INFO -string "Retrying API call to retrieve all location custom attribute data for the organization."
			$attrDataArrayList = RetrieveAllLocationAttributesForOrganization -apiHeaders $apiHeaders -apiBaseUri $apiBaseUri

		} else {

			# Display errors and exit script.
			Write-Log -level ERROR -string "Retrieving all location custom attributes for the organization failed:"
			Write-Log -level ERROR -string ("Status Code - " + $statusCode)
			Write-Log -level ERROR -string ("Status Description - " + $_.Exception.Response.StatusDescription)
			Write-Log -level ERROR -string ("Error Message - " + $_.ErrorDetails.Message)
			Write-Log -level INFO -string " "
			Write-Log -level ERROR -string "The import cannot proceed when retrieving location custom attributes data fails. Please check your authentication settings and try again."
			Write-Log -level INFO -string " "
			Write-Log -level INFO -string "Exiting."
			Write-Log -level INFO -string $processingLoopSeparator
			Exit(1)

		}

    }

	# Return the location custom attribute data.
    return $attrDataArrayList

}

function RetrieveFullLocationDetails {
	param (
		[Int32]$locationId
	)

	# Build URI to get the full location details.
    $getLocationUri = $apiBaseUri + "/api/locations/$($locationId)"

	# Get the data.
	$fullLocationInternal = $null
	try {

		# Specifically use Invoke-WebRequest here so that we get response headers.
		# We might need them to deal with rate-limiting.
		$resp = Invoke-WebRequest -Method Get -Headers $apiHeaders -Uri $getLocationUri -ContentType "application/json"

		if($verboseLog) {
			Write-Log -level INFO -string "Location retrieved successfully."
		}

		$fullLocationInternal = ($resp | ConvertFrom-Json)

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
			Write-Log -level INFO -string "Retrying API call to retrieve the full location details for location ID $($locationId)."
			$fullLocationInternal = RetrieveFullLocationDetails -apiHeaders $apiHeaders -apiBaseUri $apiBaseUri -locationId $locationId

		} else {

			# Display errors and continue
			Write-Log -level ERROR -string "Retrieving full location details for location ID $($locationId) failed:"
			Write-Log -level ERROR -string ("Status Code - " + $statusCode)
			Write-Log -level ERROR -string ("Status Description - " + $_.Exception.Response.StatusDescription)
			Write-Log -level ERROR -string ("Error Message - " + $_.ErrorDetails.Message)
			$fullLocationInternal = $null

		}

	}

	# Return the full location record.
    return $fullLocationInternal

}

function SaveLocationToApi {
	param (
		[System.Collections.Hashtable]$apiHeaders,
		[string]$apiBaseUri,
		[PSCustomObject]$locationToImport
	)

	# Build URI to save the location. Assume it has to be
	# created by default.
    $saveLocationUri = $apiBaseUri + "/api/locations"
	$method = "Post"

	# If the location has an ID greater than zero, indicating
	# an existing location, add the ID to the URI and change
	# the save method to a Put.
	if($locationToImport.ID -gt 0) {

		$saveLocationUri += "/$($locationToImport.ID)"
		$method = "Put"

	}

	# Instantiate a save successful variable and
	# a JSON request body representing the location.
	$saveSuccessful = $false
	$locationToImportJson = ($locationToImport | ConvertTo-Json)

	try {

		$savedLocResp = Invoke-WebRequest -Method $method -Headers $apiHeaders -Uri $saveLocationUri -Body $locationToImportJson -ContentType "application/json"

		if($verboseLog) {
			Write-Log -level INFO -string "Location saved successfully."
		}

		$saveSuccessful = $true

		# Reset the local acct/dept variable ID with the response from the web server.
		$locationToImport.ID = ($savedLocResp | ConvertFrom-Json).ID

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
			Write-Log -level INFO -string "Retrying API call to save location."
			$saveSuccessful = SaveLocationToApi -apiHeaders $apiHeaders -apiBaseUri $apiBaseUri -locationToImport $locationToImport

		} else {

			# Display errors and continue on.
			Write-Log -level ERROR -string "Saving the location record to the API failed. See the following log messages for more details."
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
Write-Log -level INFO -string "TeamDynamix Location import process starting."
Write-Log -level INFO -string "Detected monitor folder: $($monitorFolder)"
Write-Log -level INFO -string "Detected processed folder: $($processedFolder)"
Write-Log -level INFO -string " "

# Validate that the required script folders exist.
ValidateScriptFolders

# Resolve the full monitor and processed folder paths.
[System.IO.DirectoryInfo]$monitorFolder = (Resolve-Path $monitorFolder).Path
[System.IO.DirectoryInfo]$processedFolder = (Resolve-Path $processedFolder).Path

# 1. Retrieve the files from the monitor folder.
$files = RetrievePendingFiles
Write-Log -level INFO -string "Found $(@($files).Count) location CSV file(s) to process."

# 2. Authenticate to the API with BEID and Web Service Key and get an auth token. 
#	 If API authentication fails, display error and exit.
#	 If API authentication succeeds, store the token in a Headers object.
Write-Log -level INFO -string "Authenticating to the TeamDynamix Web API with a base URL of $($apiBaseUri)."
$global:apiHeaders = ApiAuthenticateAndBuildAuthHeaders
Write-Log -level INFO -string "Authentication successful."

# 3. Retrieve all location records for this organization for dupe matching.
Write-Log -level INFO -string "Retrieving all location data for the organization (for dupe matching) from the TeamDynamix Web API."
$global:locationDataG = [System.Collections.ArrayList](RetrieveAllLocationsForOrganization)
Write-Log -level INFO -string "Found $(@($locationDataG).Count) locations record(s) for this organization."

# 4. Retrieve all location custom attribute data for the organization if we have any custom attribute columns mapped.
Write-Log -level INFO -string "Retrieving all location custom attribute data for the organization from the TeamDynamix Web API."
$global:locationCustAttrDataG = [System.Collections.ArrayList](RetrieveAllLocationAttributesForOrganization)
Write-Log -level INFO -string "Found $(@($locationCustAttrDataG).Count) location custom attribute(s) for this organization."

# 5. Process the files.
Write-Log -level INFO -string " "
Write-Log -level INFO -string "Starting processing of location CSV file(s)."
$fileCounter = 1
$filesWithErrors = 0
$filesWithRowsSkipped = 0
foreach($file in $files) {

	# Log file start separator line.
	Write-Log -level INFO -string " "
	Write-Log -level INFO -string "Processing file $($fileCounter) out of $(@($files).Count)."
	Write-Log -level INFO -string " "

	Write-Log -level INFO -string $processingFileStartLine

	# Validate and load the CSV file. If errors are encountered, exit out instead.
	$fileValidationResult = LoadAndValidateCsvFile -file $file
	if($fileValidationResult.IsEmpty -or !$fileValidationResult.IsValid) {
		
		Write-Log -level INFO -string " "
		Write-Log -level INFO -string "Skipping empty or invalid file $($file.Name)."
		
		# Move the file to the processed folder.
		MoveFileToProcessed -file $file
		Write-Log -level INFO -string $processingFileEndLine
		
		# If it was simply an empty file, move to the next file without
		# incrementing the files with errors count.
		if($fileValidationResult.IsEmpty) {
			continue
		}
		
		# The file was invalid. Increment the counter and move to the next file.
		$filesWithErrors += 1
		continue
		
	}
	$locationCsvData = $fileValidationResult.Data

	# Store how many total items to import there are.
	$totalItemsCount = @($locationCsvData).Count

	# We found data. Proceed.
	Write-Log -level INFO -string "Found $($totalItemsCount) item(s) to process."

	# Now loop through the CSV data and save it.
	Write-Log -level INFO -string "Starting processing of CSV data."
	$processingResults = ProcessAllCsvFileData -csvRecords $locationCsvData

	# Log completion stats now.
	Write-Log -level INFO -string "Processing of file $($file.Name) complete."
	Write-Log -level INFO -string " "

	# Log failures first.
	if ($processingResults.FailCount -gt 0) {
		
		Write-Log -level ERROR -string "Failed to saved $($processingResults.FailCount) out of $($totalItemsCount) location(s). See the previous log messages for more details."
		$filesWithErrors += 1
		
	}

	# Log skips second.
	if ($processingResults.SkipCount -gt 0) {
		
		Write-Log -level WARN -string "Skipped $($processingResults.SkipCount) out of $($totalItemsCount) location(s) due to duplicate matching."
		$filesWithRowsSkipped += 1
		
	}

	# Log successes and total stats last.
	Write-Log -level INFO -string "Successfully saved $($processingResults.SuccessCount) out of $($totalItemsCount) location(s)."
	Write-Log -level INFO -string "Created $($processingResults.CreateCount) location(s) and updated $($processingResults.UpdateCount) location(s)."
	
	# Move the file to processed and log file end separator line.
	MoveFileToProcessed -file $file
	Write-Log -level INFO -string $processingFileEndLine

	# Increment file counter.
	$fileCounter += 1

}

Write-Log -level INFO -string " "

# Log how many files had errors.
if($filesWithErrors -gt 0) {
	Write-Log -level ERROR -string "Encountered $($filesWithErrors) file(s) with processing errors. See the previous log messages for more details."	
}

# Log how many files had skipped rows.
if ($filesWithRowsSkipped -gt 0) {
	Write-Log -level ERROR -string "Encountered $($filesWithRowsSkipped) file(s) with skipped rows (due to duplicate matching). See the previous log messages for more details."	
}

Write-Log -level INFO -string "Processing of all files is complete."

# Log processing complete separator.
Write-Log -level INFO -string $processingLoopSeparator