# TDX Custom Attribute Choice Importer

## Overview ##
This PowerShell script will read choice data from a .CSV file, connect to the TeamDynamix API and then import new choices for the specified (choice-based) custom attribute. The script duplicate matches on choice name and will not create duplicate entries. 

As choice name is the only identifying field for a custom attribute choice, this script cannot be used to update existing choice names. The script would simply create new choices for the updated names, leaving the old choices intact.

**CSV files must be in comma-separated format. Semicolons, pipes or any other delimiter are not supported.**

## TeamDynamix Version Support ##
This script will work on all TeamDynamix instances on **version 9.5 or higher**.

## Script Paramters ##
This script requires all of following parameters to be set.

**-fileLocation**
Data Type: String
The relative or absolute file path to the CSV file of choice data to be imported. The file **must** contain a column named **ChoiceName**. This is the column choice names are read out of.

-apiBaseUri
Data Type: String
The TeamDynamix Web API base URL. 
For SaaS production this will be in the format of https://yourTeamDynamixDomain/TDWebApi/
For SaaS sandbox this will be in the format of https://yourTeamDynamixDomain/SBTDWebApi/
For Installed (On-Prem) customers this will be in the format of https://yourTeamDynamixDomainAndPath/TDWebApi/

-apiWSBeid
Data Type: String
The TeamDynamix Web Services BEID value. This is found in the TDAdmin application organization details page. In this page, there is a **Security** box or Tab which shows the Web Services BEID value if you have the Admin permission to **Add BE Administrators**. You will need to generate a web services key and enabled key-based services for this value to appear.

-apiWSKey
Data Type: String
The TeamDynamix Web Services Key value. This is found in the TDAdmin application organization details page. In this page, there is a **Security** box or Tab which shows the Web Services Key value if you have the Admin permission to **Add BE Administrators**. You will need to generate a web services key and enabled key-based services for this value to appear.

-attributeId
Data Type: Integer
The ID of the custom attribute to import choices for.