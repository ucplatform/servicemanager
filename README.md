# Documentation for PowerShell Scripts

This document describes the high-level usage of two PowerShell functions, `Get-ZoneFromPhoneNumbers` and `Get-TenantData`, and a usage script.

## Prerequisites

1. PowerShell version 5.1 or above.
2. The Teams PowerShell Module. If you don't have this installed, you can install it with the command `Install-Module -Name MicrosoftTeams -Force`.
3. `CountryData.csv` file placed at `C:\Temp\` directory. This file should contain the country dialling codes, country names, and their corresponding team zones. Make sure this path exists and is accessible. **Note:** This path can be changed on Line 129, as it is a parameter.
4. Connection to Microsoft Teams. Make sure that you have the required permissions to access and manipulate Teams data.

## Function Definitions

### Get-ZoneFromPhoneNumbers Function

This function retrieves the zone for a given set of phone numbers based on their dialling codes.

#### Input Parameters

- `PhoneNumbers` (Mandatory): An array of phone numbers for which you want to get the zone.
- `Countries` (Mandatory): A list of `PSCustomObject` objects that hold information about country names and their corresponding dialling codes and team zones.

#### Output

The function returns a list of `PSCustomObject` objects with properties - PhoneNumber, CountryName, CountryDiallingCode, and TeamZone.

### Get-TenantData Function

This function retrieves data for tenants including active users, voice routing policies, and other relevant information.

#### Input Parameters

- `Countries` (Mandatory): A list of `PSCustomObject` objects that hold information about country names and their corresponding dialling codes and team zones.

#### Output

The function returns a list of `PSCustomObject` objects with properties - CustomerId, TenantActiveUsers, SipcomPlatformUsers, TenantCallQueues, and zone-wise user counts (ZoneOne, ZoneTwo, ZoneThree, ZoneFour, ZoneFive, ZoneSix).

## Usage Script

This script imports data from the `CountryData.csv` file, uses this data with the `Get-TenantData` function, and exports the result into `TenantData-Test.csv` file.

## Running the Script

Run the script using PowerShell. It will:

1. Connect to Microsoft Teams.
2. Import country data from `C:\Temp\CountryData.csv`.
3. Retrieve tenant data using `Get-TenantData` function with the imported country data.
4. Export the tenant data to `C:\Temp\TenantData-Test.csv`. **Note:** This path can be changed on Line 131, as it is a parameter.

Before running the script, ensure that you're logged into an account that has the necessary permissions to perform these tasks. If you're not logged in or don't have sufficient permissions, the script will fail.

## Error Handling

In case the `Get-ZoneFromPhoneNumbers` function doesn't find a matching dialling code for a phone number, it will print a warning message.

## Conclusion

This script provides an effective way to collate and analyze information about your Microsoft Teams tenants, including their geographical distribution based on the team zones. It uses a CSV file as input, making it easy to handle a large amount of data. It exports the processed data to a CSV file for further analysis.
