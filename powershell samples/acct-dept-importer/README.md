# TDX Account/Department Importer

## Overview ##
This PowerShell script will read account/department data from a .CSV file, connect to the TeamDynamix API and then save the new/updated data to the API. The script duplicate matches on the **AccountCode** column value and will not create duplicate entries. 

As account code is the identifying field for an account/department, this script cannot be used to update existing account code values. The script would simply create new account/department records, leaving the old records intact.

**CSV files must be in comma-separated format. Semicolons, pipes or any other delimiter are not supported.**

## TeamDynamix Version Support ##
This script will work on all TeamDynamix instances on **version 10.2 or higher**.

## Available CSV File Columns ##
A sample template file is included in this folder named **importdata.csv**.  The file contains the following columns:

- **Name** - *Required*
- **AccountCode** - *Required* - Used for duplicate matching.
- **IsActive** - Accepted values are `true` or `false`. During creation, all accounts are created as active. Only edits to existing accounts can deactivate an account.
- **Address1**
- **Address2**
- **Address3**
- **Address4**
- **City**
- **StateAbbr** - The two letter state abbreviation.
- **PostalCode**
- **Country**
- **Phone**
- **Fax**
- **Url**
- **Notes**
- **IndustryID** - The TDX industry ID of the account. *Note that account industry cannot be cleared via the API, only set or changed.*
- **ManagerUsername** - The globally unique (true) username of user who is the account manager. This value is shown in the Admin > Users & Roles > Users list page, Username column.
- **CustomAttribute-id** - The format for any custom attribute columns. The column should start with `CustomAttribute-` after which you place the actual custom attribute ID. For instance, a column for a custom attribute with an ID of 100 would be `CustomAttribute-100`.

### Custom Attribute Values ###
For custom attributes which are *not* choice-based, simply enter the value desired. For date and datetime custom attributes, see the following KB for the proper format of date values:  
https://solutions.teamdynamix.com/TDClient/KB/ArticleDet?ID=62569

For custom attributes which *are* choice-based, enter choice names as the column values. For attributes which support multiple choices, use the `|` character to separate choice values. An example of a valid value for multiple choices might be `choiceName1|choiceName2|choiceName3`.

### Not Mapping Values vs. Clearing Values ###
It is important to note that this process will only attempt to modify values for fields which are included in the import CSV file and can be mapped to a TeamDynamix field. If a column *is mapped*, blank values *will* clear the data for that field when the record is saved (unless otherwise noted).

For instance, if you do not provide the account/department Address4 field, that value will not be mapped and thus will not be changed in any way during the save process.

If you want to clear values, be sure that you provide a column for the field you want to clear and leave the cell values blank.

## Script Paramters ##
This script requires all of following parameters to be set.

**-fileLocation**  
*Data Type: String*  
The relative or absolute file path to the CSV file of account/department data to be imported. The file **must** contain columns named **Name** and **AccountCode**. These are the required fields to create and duplicate match records with.

**-apiBaseUri**  
*Data Type: String*  
The TeamDynamix Web API base URL.  
For SaaS production this will be in the format of https://yourTeamDynamixDomain/TDWebApi/  
For SaaS sandbox this will be in the format of https://yourTeamDynamixDomain/SBTDWebApi/  
For Installed (On-Prem) customers this will be in the format of https://yourTeamDynamixDomainAndPath/TDWebApi/

**-apiWSBeid**  
*Data Type: String*  
The TeamDynamix Web Services BEID value. This is found in the TDAdmin application organization details page. In this page, there is a **Security** box or Tab which shows the Web Services BEID value if you have the Admin permission to **Add BE Administrators**. You will need to generate a web services key and enabled key-based services for this value to appear.

**-apiWSKey**  
*Data Type: String*  
The TeamDynamix Web Services Key value. This is found in the TDAdmin application organization details page. In this page, there is a **Security** box or Tab which shows the Web Services Key value if you have the Admin permission to **Add BE Administrators**. You will need to generate a web services key and enabled key-based services for this value to appear.

## Usage Example ##
```powershell
.\ImportAccountsFromCsv.ps1 -fileLocation "pathToImportData\importData.csv" -apiBaseUri "https://yourTeamDynamixDomain/TDWebApi/" -apiWSBeid "apiWSBeidFromTDAdmin" -apiWSKey "apiWSKeyFromTDAdmin"
```