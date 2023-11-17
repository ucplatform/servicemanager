# Documentation for ServiceManagerReport.ps1 Script

This document describes the high-level usage of two PowerShell functions, `Get-ZoneFromPhoneNumbers` and `Get-TenantData`, and a usage script.

## Prerequisites

1. PowerShell version 5.1 or above.
2. The Teams PowerShell Module. If you don't have this installed, you can install it with the command `Install-Module -Name MicrosoftTeams -Force`.
3. The az.storage PowerShell Module. If you don't have this installed, you can install it with the command `Install-Module -Name Az.Storage -Force`.
5. Connection to Microsoft Teams. Make sure that you have the required permissions to access and manipulate Teams data.

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

This script collects and uses the data with the `Get-TenantData` function, and exports the result into `TenantData-Test.csv` file.

## Running the Script

Run the script using PowerShell. It will:

1. Connect to Microsoft Teams.
2. Enter the required SAS Token: Provided by your Partner in the email
3. Retrieve tenant data using `Get-TenantData` function with the imported country data.
4. Export the tenant data to blob storage, where it is processed and added to your customer record

Before running the script, ensure that you're logged into an account that has the necessary permissions to perform these tasks. If you're not logged in or don't have sufficient permissions, the script will fail.

## Error Handling

In case the `Get-ZoneFromPhoneNumbers` function doesn't find a matching dialling code for a phone number, it will print a warning message.

## Conclusion

This script provides an effective way to collate and analyze information about your Microsoft Teams tenants, including their geographical distribution based on the team zones. It uses a CSV file as input, making it easy to handle a large amount of data. It exports the processed data to a CSV file for further analysis.
