# TDX Location Importer 2.0.0

## Overview ##
This PowerShell script will read location data from all .CSV files in a monitored folder, connect to the TeamDynamix API and then save the new/updated data to the API. Processed files will then be moved to a processed folder. The script duplicate matches on the **ExternalID** column value and will not create duplicate entries. 

As external ID is the identifying field for a location, this script cannot be used to update existing external ID values. The script would simply create new location records, leaving the old records intact.

**CSV files must be in comma-separated format. Semicolons, pipes or any other delimiter are not supported.**

## TeamDynamix Version Support ##
This script will work on all TeamDynamix instances on **version 10.2 or higher**.

If you are an installed (on-prem) customer or need an earlier version of this script, use the list below for previous versions.
- <a href="legacy/v1" target="_blank">Version 11.0+: legacy/v1/</a>

## Usage Requirements ##
This script requires **.NET Framework 4.0 or higher be installed.**

## Available CSV File Columns ##
A sample template file is included in this folder named **importdata.csv**.  The file contains the following columns:

- **Name** - *Required*
- **ExternalID** - *Required* - Used for duplicate matching.
- **Description**
- **IsActive** - Accepted values are `true` or `false`.
- **Address**
- **City**
- **StateAbbr** - The two letter state abbreviation.
- **PostalCode**
- **Country**
- **IsRoomRequired** - Does the location require that a room be set when selected in a form? Accepted values are `true` or `false`.
- **Latitude** - A decimal value between `-90.0` and `90.0`. If this value is specified, a value **must** be specified for Longitude as well.
- **Longitude** - A decimal value between `-180.0` and `180.0`. If this value is specified, a value **must** be specified for Latitude as well.
- **CustomAttribute-id** - The format for any custom attribute columns. The column should start with `CustomAttribute-` after which you place the actual custom attribute ID. For instance, a column for a custom attribute with an ID of 100 would be `CustomAttribute-100`.

### Custom Attribute Values ###
For custom attributes which are *not* choice-based, simply enter the value desired. For date and datetime custom attributes, see the following KB for the proper format of date values:  
https://solutions.teamdynamix.com/TDClient/KB/ArticleDet?ID=62569

For custom attributes which *are* choice-based, enter choice names as the column values. For attributes which support multiple choices, use the `|` character to separate choice values. An example of a valid value for multiple choices might be `choiceName1|choiceName2|choiceName3`.

### Not Mapping Values vs. Clearing Values ###
It is important to note that this process will only attempt to modify values for fields which are included in the import CSV file and can be mapped to a TeamDynamix field. If a column *is mapped*, blank values *will* clear the data for that field when the record is saved (unless otherwise noted).

For instance, if you do not provide the location Address field, that value will not be mapped and thus will not be changed in any way during the save process.

If you want to clear values, be sure that you provide a column for the field you want to clear and leave the cell values blank.

## Script Parameters ##
This script requires all of following parameters to be set unless otherwise marked as optional.

**-monitorFolder**  
*Data Type: String*  
The relative or absolute path to the directory containg CSV files of location data to be imported. All files **must** contain columns named **Name** and **ExternalID**. These are the required fields to create and duplicate match records with.

**-processedFolder**  
*Data Type: String*  
The relative or absolute path to the directory for processed files to be moved into after the import completes. Processed files will be prefixed with a `yyyy-MM-dd HHmmssffff ` timestamp.

**-apiBaseUri**  
*Data Type: String*  
The TeamDynamix Web API base URL.  
For SaaS production this will be in the format of https://yourTeamDynamixDomain/TDWebApi/  
For SaaS sandbox this will be in the format of https://yourTeamDynamixDomain/SBTDWebApi/  
For Installed (On-Prem) customers this will be in the format of https://yourTeamDynamixDomainAndPath/TDWebApi/

**-apiWSBeid**  
*Data Type: String*  
The TeamDynamix Web Services BEID value. This is found in the TDAdmin application organization details page. In this page, there is a **Security** box or Tab which shows the Web Services BEID value if you have the Admin permission to **Add BE Administrators**.

**-apiWSKey**  
*Data Type: String*  
The TeamDynamix Web Services Key value. This is found in the TDAdmin application organization details page. In this page, there is a **Security** box or Tab which shows the **Admin Service Accounts (and their associated web services keys)** if you have the Admin permission to **Add BE Administrators**. You will need to create at least one **Admin Service Account** to get a web services key value.

**-maxJsonObjectSizeInBytes**  
*Data Type: Integer*  
*Optional (default value is 104857600)*  
The maximum size in bytes for downloaded JSON payloads. The default is 104857600 which is roughly equivalent to 100 megabytes (MB). You likely won't need to increase this, but if your user base data (for manager matching) or account data is especially large, this gives you an option to.

**-verboseLog**  
*Data Type: Switch*  
*Optional*  
Whether or not to use verbose logging to get a more detailed look at each row being saved as files are processed.

## Usage Example ##
**Without verbose logging:**  
```powershell
.\ImportLocationsFromCsv.ps1 -fileLocation "pathToImportData\importData.csv" -apiBaseUri "https://yourTeamDynamixDomain/TDWebApi/" -apiWSBeid "apiWSBeidFromTDAdmin" -apiWSKey "apiWSKeyFromTDAdmin"
```

**With verbose logging:**  
```powershell
.\ImportLocationsFromCsv.ps1 -fileLocation "pathToImportData\importData.csv" -apiBaseUri "https://yourTeamDynamixDomain/TDWebApi/" -apiWSBeid "apiWSBeidFromTDAdmin" -apiWSKey "apiWSKeyFromTDAdmin" -verboseLog
```

## Logging ##
A set of rolling log files are generated in the directory that the script is located. Logging will keep a maximum of ten rolling 10MB files at all times. To prevent logs from rolling excessively, it is recommended to leave verbose logging off unless you are experiencing errors with the import.