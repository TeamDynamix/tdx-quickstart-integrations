# Script Version: 3.0.0

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
    
		{ $_ -eq 'VERBOSE' -or $_ -eq 'INFO' } {
            
			Out-Host -InputObject "$logString"
            
		}

		{ $_ -eq 'ERROR' } {

			$host.ui.RawUI.ForegroundColor = "Red"
			$host.ui.RawUI.BackgroundColor = "Black"

			Out-Host -InputObject "$logString"
    
			$host.ui.RawUI.ForegroundColor = $foregroundColor
			$host.UI.RawUI.BackgroundColor = $backgroundColor

		}

		{ $_ -eq 'WARN' } {
    
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
	Write-Host "  Version 3.0.0"
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
	
		Write-Log -level INFO -string "No account CSV files were detected for processing."
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
	$acctCsvDataInternal = try {
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
	$totalItemsCount = @($acctCsvDataInternal).Count

	# Exit out if no data is found to import.
	if ($totalItemsCount -le 0) {
		
		Write-Log -level INFO -string "No items detected for processing."
		$csvValidationResultInternal.IsEmpty = $true
		return $csvValidationResultInternal
		
	} 
	
	# Now validate that we at least have a Name and an AccountCode column mapped.
	# These are the minimum columns required to create an account and support duplicate matching.
	$csvColumnHeaders = $acctCsvDataInternal[0].PSObject.Properties.Name
	if (!(DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "Name")) {
		
		Write-Log -level ERROR -string "The file does not contain the required Name column. A Name column is required for the acct/dept import."
		$csvValidationResultInternal.IsValid = $false
		return $csvValidationResultInternal
		
	}

	if (!(DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "AccountCode")) {
		
		Write-Log -level ERROR -string "The file does not contain the required AccountCode column. An AccountCode column is required for the acct/dept import."
		$csvValidationResultInternal.IsValid = $false
		return $csvValidationResultInternal
		
	}
	
	# If we got here, the file is valid. Set the data and return the results.
	$csvValidationResultInternal.IsEmpty = $false
	$csvValidationResultInternal.Data = $acctCsvDataInternal
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
	if($totalCsvRowsInternal -gt 10 -and $totalCsvRowsInternal -le 50) {

		$modCountInternal = 5
		if($totalCsvRowsInternal -gt 50 -and $totalCsvRowsInternal -le 100) {

			$modCountInternal = 10
			if($totalCsvRowsInternal -gt 100 -and $totalCsvRowsInternal -le 500) {
				
				$modCountInternal = 25
				if($totalCsvRowsInternal -gt 500 -and $totalCsvRowsInternal -le 1000) {
					
					$modCountInternal = 100
					if($totalCsvRowsInternal -gt 1000 -and $totalCsvRowsInternal -le 10000) {

						$modCountInternal = 250
						if($totalCsvRowsInternal -gt 10000 -and $totalCsvRowsInternal -le 100000) {
							
							$modCountInternal = 2500
							if($totalCsvRowsInternal -gt 100000) {
								$modCountInternal = 20000
							}

						}

					}

				}
			}
		}
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
	foreach ($acctCsvRecord in $csvRecords) {

		# Log information about the row we are processing if verbose.
		$acctName = $acctCsvRecord.Name
		$acctCode = $acctCsvRecord.AccountCode

		if($verboseLog) {
			Write-Log -level INFO -string "Processing row $($rowIndex) with a Name value of `"$($acctName)`" and an AccountCode value of `"$($acctCode)`"."
		}
		

		# Validate that we have a Name value.
		# If this is a null or empty string, skip this row.
		if ([string]::IsNullOrWhiteSpace($acctName)) {

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

		# Validate the AccountCode value for duplicate matching.
		# If this is a null or empty string, skip this row.
		if ([string]::IsNullOrWhiteSpace($acctCode)) {
			
			# Log errors, increment failure count and row index.
			Write-Log -level ERROR -string "Row $($rowIndex) - Required field AccountCode has no value. This row will be skipped."
			$results.FailCount += 1
			$rowIndex += 1

			# Log rows processed.
			$processedCount = ($rowIndex - 2)		
			LogPercentageFileProcessed -modCount $modCount -processedCount $processedCount -totalItemsCount $totalCsvRowsInternal

			# Continue the loop.			
			continue

		}

		# Now see if we are dealing with an existing account or not.
		# If this is a new account, new up a fresh object.
		$acctToImport = ($acctDataG | Where-Object { $_.Code -eq $acctCode } | Select-Object -First 1)
		if (!$acctToImport) {
			$acctToImport = GetNewAccountObject
		}
		else {

			# See if we need to load up the full details for the existing account or not.
			# An ID greater than zero indicates we are updating an existing account.
			# AN ID of 0 or less means we already created this account earlier in the script.
			if ($acctToImport.ID -gt 0) {

				# Get the full acct/dept information here. If this fails, the row has to be skipped.
				# We cannot risk saving over an existing account and potentially erasing data.
				if($verboseLog) {
					Write-Log -level INFO -string "Detected an existing account (ID: $($acctToImport.ID)) on the server to update. Retrieving full account details from the API before updating values."
				}

				$acctId = $acctToImport.ID
				$acctToImport = RetrieveFullAccountDetails -accountId $acctId
				if (!$acctToImport) {

					# Log that we did not find the account and skip this row.
					Write-Log -level ERROR -string ("Row $($rowIndex) - Detected existing account (ID: $($acctId)) to update but the account" +
						" could not be retrieved from the API. This row will be skipped since saving the import data might" +
						" unintentionally clear other fields on the account server-side.")
					$results.FailCount += 1
					$rowIndex += 1
					continue

				}

			}
			else {

				# If we did find the account, but the ID is 0, this means we already created an account
				# for this account code in the same sheet. Log a warning and skip this row.
				Write-Log -level WARN -string ("Row $($rowIndex) - An account with Name `"$($acctName)`" and AccountCode" +
					" `"$($acctCode)`" already had to be created earlier in this processing loop. To not create duplicate" +
					" records on the server, this row will be skipped.")
				$results.SkipCount += 1
				$rowIndex += 1
				continue

			}		

		}

		# Update the account from the CSV data as needed.
		$acctToImport = UpdateAccountFromCsv `
			-rowIndex $rowIndex `
			-csvColumnHeaders $csvColumnHeaders `
			-csvRecord $acctCsvRecord `
			-acctToImport $acctToImport

		# Store whether or not this will be a newly created record or not.
		$isCreating = ($acctToImport.ID -le 0) 

		# Save account to API here
		if($verboseLog) {
			Write-Log -level INFO -string "Saving account with Name `"$($acctName)`" and AccountCode `"$($acctCode)`" to the TeamDynamix API."
		}
		$saveSuccess = SaveAccountToApi -acctToImport $acctToImport

		# If the save was successful, increment the success counter and add the new choice to list of existing choices.
		# If the save failed, increment the fail counter and spit out a message.
		if ($saveSuccess) {
			
			# Add the new acct/dept to the collection of existing acct/depts if it was newly created.
			if ($isCreating) {
				
				# Now add the newly created account to the local collection of all acct/dept data.
				$acctDataG.Add($acctToImport)
				
				# Increment the creation counter.
				$results.CreateCount += 1

			}
			else {
				
				# Increment the update counter.
				$results.UpdateCount += 1

			}

			# Increment the success counter.
			$results.SuccessCount += 1

		}
		else {
			
			# Log an error and increment the fail counter.
			Write-Log -level ERROR -string ("Row $($rowIndex) - The account with a Name value of `"$($acctName)`" and an AccountCode value of" +
				" `"$($acctCode)`" failed to save successfully. See previous errors for details.")
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

function GetNewAccountObject {
    
	# Create a fresh account object.
	# This will need to be updated should the Account object ever change.
	$newAcct = [PSCustomObject]@{
		ID=0;
		Name="";
		ParentID=$null;
		IsActive= $true;
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
		[Array]$csvColumnHeaders,
		[Array]$csvRecord,
		[PSCustomObject]$acctToImport        
	)

	# Name and AccountCode (always required)
	$acctToImport.Name = $csvRecord.Name.Trim()
	$acctToImport.Code = $csvRecord.AccountCode.Trim()

	# Now process optional fields.
	# IsActive
	if ((DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "IsActive")) {
		
		if ($csvRecord.IsActive.Trim() -eq "true") {
			$acctToImport.IsActive = $true
		}
		else {
			$acctToImport.IsActive = $false
		}

	}

	# Address1
	if ((DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "Address1")) {
		$acctToImport.Address1 = $csvRecord.Address1.Trim()
	}

	# Address2
	if ((DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "Address2")) {
		$acctToImport.Address2 = $csvRecord.Address2.Trim()
	}

	# Address3
	if ((DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "Address3")) {
		$acctToImport.Address3 = $csvRecord.Address3.Trim()
	}

	# Address4
	if ((DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "Address4")) {
		$acctToImport.Address4 = $csvRecord.Address4.Trim()
	}

	# City
	if ((DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "City")) {
		$acctToImport.City = $csvRecord.City.Trim()
	}

	# StateAbbr
	if ((DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "StateAbbr")) {
		$acctToImport.StateAbbr = $csvRecord.StateAbbr.Trim()
	}

	# PostalCode
	if ((DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "PostalCode")) {
		$acctToImport.PostalCode = $csvRecord.PostalCode.Trim()
	}

	# Country
	if ((DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "Country")) {
		$acctToImport.Country = $csvRecord.Country.Trim()
	}

	# Phone
	if ((DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "Phone")) {
		$acctToImport.Phone = $csvRecord.Phone.Trim()
	}

	# Fax
	if ((DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "Fax")) {
		$acctToImport.Fax = $csvRecord.Fax.Trim()
	}

	# Url
	if ((DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "Url")) {
		$acctToImport.Url = $csvRecord.Url.Trim()
	}

	# Notes
	if ((DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "Notes")) {
		$acctToImport.Notes = $csvRecord.Notes.Trim()
	}

	# IndustryID
	if ((DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "IndustryID")) {
		
		# Parse this out to an integer. If it doesn't parse or is less than zero, use zero.
		[Int32]$parsedId = $null
		$parseSuccess = [Int32]::TryParse($csvRecord.IndustryID, [ref]$parsedId)
		if ($parseSuccess -and $parsedId -gt 0) {
			$acctToImport.IndustryID = $parsedId
		}
		else {
			$acctToImport.IndustryID = 0
		}

	}

	# ParentAccountCode
	$acctToImport = SetAccountParent `
		-rowIndex $rowIndex `
		-csvColumnHeaders $csvColumnHeaders `
		-csvRecord $csvRecord `
		-acctToImport $acctToImport

	# ManagerUsername
	$acctToImport = SetAccountManager `
		-rowIndex $rowIndex `
		-csvColumnHeaders $csvColumnHeaders `
		-csvRecord $csvRecord `
		-acctToImport $acctToImport

	# Custom Attributes
	$acctToImport = SetItemCustomAttributeValues `
		-rowIndex $rowIndex `
		-csvColumnHeaders $csvColumnHeaders `
		-csvRecord $csvRecord `
		-acctToImport $acctToImport

	# Return the updated account.
	return $acctToImport

}

function SetAccountManager {
	param (
		[Int32]$rowIndex,
		[Array]$csvColumnHeaders,
		[Array]$csvRecord,
		[PSCustomObject]$acctToImport  
	)

	# ManagerUsername
	if ((DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "ManagerUsername")) {
	
		# Determine how to properly set manager UID.
		if ([string]::IsNullOrWhiteSpace($csvRecord.ManagerUsername)) {
			
			# The manager username column is blank. Send empty GUID to represent no manager.
			$acctToImport.ManagerUID = [GUID]::Empty

		}
		else {

			# Try to find the manager by username.
			$managerUser = ($userDataG | Where-Object { $_.UserName -eq $csvRecord.ManagerUsername.Trim() } | Select-Object -First 1)
			if(!$managerUser) {

				# If we had a manager username but could find no match for the new value, leave
				# this field untouched. Instead log a warning and move on.
				Write-Log -level WARN -string ("Row $($rowIndex) - ManagerUsername value of `"$($csvRecord.ManagerUsername.Trim())`"" +
					" is not a valid TeamDynamix username. The account manager value will be left unchanged when the account is saved" +
					" to the server.")

			} else {
				
				$managerUid = $managerUser.UID
				if (!$managerUid -or [string]::IsNullOrWhiteSpace($managerUid)) {
					
					# If we had a manager username but could find no match for the new value, leave
					# this field untouched. Instead log a warning and move on.
					Write-Log -level WARN -string ("Row $($rowIndex) - ManagerUsername value of `"$($csvRecord.ManagerUsername.Trim())`"" +
						" is not a valid TeamDynamix username. The account manager value will be left unchanged when the account is saved" +
						" to the server.")

				}
				else {

					# A user match was found for manager. Set the manager UID properly as a GUID.
					$acctToImport.ManagerUID = [GUID]::Parse($managerUid)
					
				}			
			
			}			
		}
	}

	# Return the updated account.
	return $acctToImport

}

function SetAccountParent {
	param (
		[Int32]$rowIndex,
		[Array]$csvColumnHeaders,
		[Array]$csvRecord,
		[PSCustomObject]$acctToImport  
	)

	# ParentAccountCode
	if ((DoesCsvFileContainColumn -csvColumnHeaders $csvColumnHeaders -columnToFind "ParentAccountCode")) {

		# Determine how to properly set parent ID.
		if ([string]::IsNullOrWhiteSpace($csvRecord.ParentAccountCode)) {
	
			# No parent account code was detected. Set the parent ID to null so that it blanks out.
			$acctToImport.ParentID = $null

		}
		else {

			# Try to find the parent account by account code value.
			$parentCodeToFind = $csvRecord.ParentAccountCode.Trim()
			$parentAcct = ($acctDataG | Where-Object { $_.Code -eq $parentCodeToFind } | Select-Object -First 1)
			if (!$parentAcct) {

				# If we had a parent account code but could find no match for the new value, leave
				# this field untouched. Instead log a warning and move on.
				Write-Log -level WARN -string ("Row $($rowIndex) - ParentAccountCode value of `"$($parentCodeToFind)`"" +
					" is not a valid TeamDynamix acct/dept code. The account parent value will be left unchanged when the account is saved" +
					" to the server.")

			}
			else {

				# An account match was found for parent account. Set the parent account with the found account's ID.
				$acctToImport.ParentID = $parentAcct.ID

			}

		}
	}

	# Return the updated account.
	return $acctToImport

}

function SetItemCustomAttributeValues {
	param (
		[Int32]$rowIndex,
		[Array]$csvColumnHeaders,
		[Array]$csvRecord,
		[PSCustomObject]$acctToImport  
	)

	# Custom Attributes
	# If there are any columns starting with CustomAttribute-, loop through them and update
	# the custom attribute values on the acct/dept if possible.
	$customAttributeCols = ($csvColumnHeaders | Where-Object { $_.StartsWith($customAttrColPrefix, "OrdinalIgnoreCase") })
	if ($customAttributeCols -and @($customAttributeCols).Count -gt 0) {
	
		# Loop through each custom attribute column.
		foreach ($customAttributeCol in $customAttributeCols) {

			# First attempt to parse out the actual custom attribute ID.
			$attID = $customAttributeCol.Split("-", 2)[1]
			$parsedAttId = 0
			$parsedSuccessfully = [Int32]::TryParse($attId, [ref]$parsedAttId)
			if ($parsedSuccessfully -and $parsedAttId -gt 0) {

				# If we found an attribute ID, validate that attribute exists in list of all acct/dept attributes.
				# If the attribute ID is not valid we cannot use this column.
				$attribute = ($acctCustAttrDataG | Where-Object { $_.ID -eq $parsedAttId } | Select-Object -First 1)
				if ($attribute) {

					# Get the new attribute value.
					$newAttrVal = ($csvRecord | Select-Object -ExpandProperty $customAttributeCol)

					# If this is a choice-based attribute and the new value is not empty (meaning to clear it out),
					# further processing on the value is necessary.
					$isChoiceBased = IsChoiceBasedCustomAttribute -fieldType $attribute.FieldType
					$allChoicesInvalid = $false
					if ($isChoiceBased -and ![string]::IsNullOrWhiteSpace($newAttrVal)) {

						# Split the column value on | to get all possible choice names to map.
						$choiceNames = $newAttrVal.Split("|", [System.StringSplitOptions]::RemoveEmptyEntries)
						if ($choiceNames -and @($choiceNames).Count -gt 0) {

							# Instantiate a variable to store the choice IDs in.
							$selectedChoiceIds = ""

							# Loop through all choice names specified and try to find their choice IDs.
							# Only use the valid choice names.
							foreach ($choiceName in $choiceNames) {

								$choice = ($attribute.Choices | Where-Object { $_.Name -eq $choiceName } | Select-Object -First 1)
								if ($choice) {
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
					if ($isChoiceBased -and $allChoicesInvalid) {
						
						Write-Log -level WARN -string ("Row $($rowIndex) - Could not set value for custom attribute ID $($attribute.ID)." + 
							" This is a choice-based attribute and all choices specified were invalid. The attribute value will" +
							" be left unchanged when the account is saved to the server.")

					}
					else {

						$attributeToUpdate = ($acctToImport.Attributes | Where-Object { $_.ID -eq $parsedId } | Select-Object -First 1)
						if ($attributeToUpdate) {
						
							# The attribute already exists on this account. Just update its value.
							$attributeToUpdate.Value = $newAttrVal

						}
						else {
						
							# The attribute does not exist on this account. Create it and set its value properly.
							$acctToImport.Attributes += @{
								ID    = $attribute.ID;
								Value = $newAttrVal;
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
	}
 	catch {

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
	$apiHeadersInternal = @{"Authorization" = "Bearer " + $authToken }

	# Return the API headers.
	return $apiHeadersInternal
	
}

function RetrieveAllUsersForOrganization {

	# Build URI to get all users for the organization.
	$getUserListUri = $apiBaseUri + "/api/people/userlist?isActive=&isEmployee=&userType=User"

	# Get the data.
	$userDataArrayList = [System.Collections.ArrayList]@()
	try {
		
		# Specifically use Invoke-WebRequest here so that we get response headers.
		# We might need them to deal with rate-limiting.
		$resp = Invoke-WebRequest -Method Get -Headers $apiHeaders -Uri $getUserListUri -ContentType "application/json"
			
		# Use the custom JSON deserializer in case the organization has a large user base.
		# Normally we would use this: $userDataInternal = ($resp | ConvertFrom-Json)
		$userDataInternal = $jsonDeserializer.Deserialize($resp, [System.Object])
		
		# Convert the user data to an array list for faster querying.		
		if($userDataInternal -and @($userDataInternal).Count -gt 0) {
			$userDataArrayList.AddRange($userDataInternal)
		}
		
	}
 	catch {

		# If we got rate limited, try again after waiting for the reset period to pass.
		$statusCode = $_.Exception.Response.StatusCode.value__
		if ($statusCode -eq 429) {

			# Get the amount of time we need to wait to retry in milliseconds.
			$resetWaitInMs = GetRateLimitWaitPeriodMs -apiCallResponse $_.Exception.Response
			Write-Log -level INFO -string "Waiting $(($resetWaitInMs / 1000.0).ToString("N2")) seconds to rety API call due to rate-limiting."

			# Wait to retry now.
			Start-Sleep -Milliseconds $resetWaitInMs

			# Now retry.
			Write-Log -level INFO -string "Retrying API call to retrieve all User typed user data for the organization."
			$userDataArrayList = RetrieveAllUsersForOrganization

		}
		else {

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
	return $userDataArrayList

}

function RetrieveAllAcctDeptsForOrganization {

	# Build URI to get all acct/dept records for the organization.
	$getAcctsUri = $apiBaseUri + "/api/accounts/search"

	# Set the user authentication URI and create an authentication JSON body.
	$acctSearchBody = [PSCustomObject]@{
		IsActive=$null;
		MaxResults=$null;
	} | ConvertTo-Json

	# Get the data.
	$acctDataArrayList = [System.Collections.ArrayList]@()
	try {
            
		# Specifically use Invoke-WebRequest here so that we get response headers.
		# We might need them to deal with rate-limiting.
		$resp = Invoke-WebRequest -Method Post -Headers $apiHeaders -Uri $getAcctsUri -Body $acctSearchBody -ContentType "application/json"
			
		# Use the custom JSON deserializer in case the organization has a large set of accounts.
		# Normally we would use this: $acctDataInternal = ($resp | ConvertFrom-Json)
		$acctDataInternal = $jsonDeserializer.Deserialize($resp, [System.Object])

		# Convert the acct data to an array list for faster querying.
		if($acctDataInternal -and @($acctDataInternal).Count -gt 0) {
			$acctDataArrayList.AddRange($acctDataInternal)
		}
		
	}
 	catch {

		# If we got rate limited, try again after waiting for the reset period to pass.
		$statusCode = $_.Exception.Response.StatusCode.value__
		if ($statusCode -eq 429) {

			# Get the amount of time we need to wait to retry in milliseconds.
			$resetWaitInMs = GetRateLimitWaitPeriodMs -apiCallResponse $_.Exception.Response
			Write-Log -level INFO -string "Waiting $(($resetWaitInMs / 1000.0).ToString("N2")) seconds to rety API call due to rate-limiting."

			# Wait to retry now.
			Start-Sleep -Milliseconds $resetWaitInMs

			# Now retry.
			Write-Log -level INFO -string "Retrying API call to retrieve all acct/dept data for the organization."
			$acctDataArrayList = RetrieveAllAcctDeptsForOrganization

		}
		else {

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
	return $acctDataArrayList
}

function RetrieveAllAcctAttributesForOrganization {

	# Build URI to get all acct/dept custom attributes for the organization.
	$getAcctAttributesUri = $apiBaseUri + "/api/attributes/custom?componentId=14"

	# Get the data.
	$attrDataArrayList = [System.Collections.ArrayList]@()
	try {
		
		# Specifically use Invoke-WebRequest here so that we get response headers.
		# We might need them to deal with rate-limiting.
		$resp = Invoke-WebRequest -Method Get -Headers $apiHeaders -Uri $getAcctAttributesUri -ContentType "application/json"

		# Use the custom JSON deserializer in case the organization has a large set of account attributes.
		# Normally we would use this: $attrDataInternal = ($resp | ConvertFrom-Json)
		$attrDataInternal = $jsonDeserializer.Deserialize($resp, [System.Object])
		
		# Convert the acct attribute data to an array list for faster querying.
		if($attrDataInternal -and @($attrDataInternal).Count -gt 0) {
			$attrDataArrayList.AddRange($attrDataInternal)
		}
		
	}
 	catch {

		# If we got rate limited, try again after waiting for the reset period to pass.
		$statusCode = $_.Exception.Response.StatusCode.value__
		if ($statusCode -eq 429) {

			# Get the amount of time we need to wait to retry in milliseconds.
			$resetWaitInMs = GetRateLimitWaitPeriodMs -apiCallResponse $_.Exception.Response
			Write-Log -level INFO -string "Waiting $(($resetWaitInMs / 1000.0).ToString("N2")) seconds to rety API call due to rate-limiting."

			# Wait to retry now.
			Start-Sleep -Milliseconds $resetWaitInMs

			# Now retry.
			Write-Log -level INFO -string "Retrying API call to retrieve all acct/dept custom attribute data for the organization."
			$attrDataArrayList = RetrieveAllAcctAttributesForOrganization

		}
		else {

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
	return $attrDataArrayList

}

function RetrieveFullAccountDetails {
	param (
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
		
		if($verboseLog) {
			Write-Log -level INFO -string "Account retrieved successfully."
		}

		$fullAccountInternal = ($resp | ConvertFrom-Json)

	}
 	catch {

		# If we got rate limited, try again after waiting for the reset period to pass.
		$statusCode = $_.Exception.Response.StatusCode.value__
		if ($statusCode -eq 429) {

			# Get the amount of time we need to wait to retry in milliseconds.
			$resetWaitInMs = GetRateLimitWaitPeriodMs -apiCallResponse $_.Exception.Response
			Write-Log -level INFO -string "Waiting $(($resetWaitInMs / 1000.0).ToString("N2")) seconds to rety API call due to rate-limiting."

			# Wait to retry now.
			Start-Sleep -Milliseconds $resetWaitInMs

			# Now retry.
			Write-Log -level INFO -string "Retrying API call to retrieve the full acct/dept details for account ID $($accountId)."
			$fullAccountInternal = RetrieveFullAccountDetails -accountId $accountId

		}
		else {

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
		[PSCustomObject]$acctToImport
	)

	# Build URI to save the account. Assume it has to be
	# created by default.
	$saveAcctUri = $apiBaseUri + "/api/accounts"
	$method = "Post"

	# If the account has an ID greater than zero, indicating
	# an existing account, add the ID to the URI and change
	# the save method to a Put.
	if ($acctToImport.ID -gt 0) {

		$saveAcctUri += "/$($acctToImport.ID)"
		$method = "Put"

	}

	# Instantiate a save successful variable and
	# a JSON request body representing the account.
	$saveSuccessful = $false
	$acctToImportJson = ($acctToImport | ConvertTo-Json)

	try {

		$savedAcctResp = Invoke-WebRequest -Method $method -Headers $apiHeaders -Uri $saveAcctUri -Body $acctToImportJson -ContentType "application/json"
		
		if($verboseLog) {
			Write-Log -level INFO -string "Account saved successfully."
		}

		$saveSuccessful = $true
		
		# Reset the local acct/dept variable ID with the response from the web server.
		$acctToImport.ID = ($savedAcctResp | ConvertFrom-Json).ID
		
	}
 	catch {

		# If we got rate limited, try again after waiting for the reset period to pass.
		$statusCode = $_.Exception.Response.StatusCode.value__
		if ($statusCode -eq 429) {

			# Get the amount of time we need to wait to retry in milliseconds.
			$resetWaitInMs = GetRateLimitWaitPeriodMs -apiCallResponse $_.Exception.Response
			Write-Log -level INFO -string "Waiting $(($resetWaitInMs / 1000.0).ToString("N2")) seconds to rety API call due to rate-limiting."

			# Wait to retry now.
			Start-Sleep -Milliseconds $resetWaitInMs

			# Now retry the save.
			Write-Log -level INFO -string "Retrying API call to save account."
			$saveSuccessful = SaveAccountToApi -acctToImport $acctToImport

		}
		else {

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
Write-Log -level INFO -string "Found $(@($files).Count) account CSV file(s) to process."

# 2. Authenticate to the API with BEID and Web Service Key and get an auth token. 
#	 If API authentication fails, display error and exit.
#	 If API authentication succeeds, store the token in a Headers object.
Write-Log -level INFO -string "Authenticating to the TeamDynamix Web API with a base URL of $($apiBaseUri)."
$global:apiHeaders = ApiAuthenticateAndBuildAuthHeaders
Write-Log -level INFO -string "Authentication successful."

# 3. Retrieve all acct/dept records for this organization for dupe matching.
Write-Log -level INFO -string "Retrieving all acct/dept data for the organization (for dupe matching) from the TeamDynamix Web API."
$global:acctDataG = [System.Collections.ArrayList](RetrieveAllAcctDeptsForOrganization)
Write-Log -level INFO -string "Found $(@($acctDataG).Count) acct/dept record(s) for this organization."

# 4. Retrieve all acct/dept custom attribute data for the organization.
Write-Log -level INFO -string "Retrieving all acct/dept custom attribute data for the organization from the TeamDynamix Web API."
$global:acctCustAttrDataG = [System.Collections.ArrayList](RetrieveAllAcctAttributesForOrganization)
Write-Log -level INFO -string "Found $(@($acctCustAttrDataG).Count) acct/dept custom attribute(s) for this organization."

# 5. Retrieve all user records for the organization for manager matching.
Write-Log -level INFO -string "Retrieving all User typed user data for the organization (for manager matching) from the TeamDynamix Web API."
$global:userDataG = [System.Collections.ArrayList](RetrieveAllUsersForOrganization)
Write-Log -level INFO -string "Found $(@($userDataG).Count) user(s) for this organization."

# 6. Process the files.
Write-Log -level INFO -string " "
Write-Log -level INFO -string "Starting processing of account CSV file(s)."
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
	$acctCsvData = $fileValidationResult.Data

	# Store how many total items to import there are.
	$totalItemsCount = @($acctCsvData).Count

	# We found data. Proceed.
	Write-Log -level INFO -string "Found $($totalItemsCount) item(s) to process."

	# Now loop through the CSV data and save it.
	Write-Log -level INFO -string "Starting processing of CSV data."
	$processingResults = ProcessAllCsvFileData -csvRecords $acctCsvData

	# Log completion stats now.
	Write-Log -level INFO -string "Processing of file $($file.Name) complete."
	Write-Log -level INFO -string " "

	# Log failures first.
	if ($processingResults.FailCount -gt 0) {
		
		Write-Log -level ERROR -string "Failed to saved $($processingResults.FailCount) out of $($totalItemsCount) account(s). See the previous log messages for more details."
		$filesWithErrors += 1
		
	}

	# Log skips second.
	if ($processingResults.SkipCount -gt 0) {
		
		Write-Log -level WARN -string "Skipped $($processingResults.SkipCount) out of $($totalItemsCount) account(s) due to duplicate matching."
		$filesWithRowsSkipped += 1
		
	}

	# Log successes and total stats last.
	Write-Log -level INFO -string "Successfully saved $($processingResults.SuccessCount) out of $($totalItemsCount) account(s)."
	Write-Log -level INFO -string "Created $($processingResults.CreateCount) account(s) and updated $($processingResults.UpdateCount) account(s)."
	
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