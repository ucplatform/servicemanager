function Get-ZoneFromPhoneNumbers {
    [CmdletBinding()]
    param (
      [Parameter(Mandatory = $true)]
      [string[]]$PhoneNumbers,
  
      [Parameter(Mandatory = $true)]
      [System.Collections.Generic.List[PSCustomObject]]$Countries
    )
  
    begin {
      $output = [System.Collections.Generic.List[PSCustomObject]]::new()
    }
  
    process {
      $diallingCodes = $Countries | ForEach-Object { $_.CountryDiallingCode }
  
      $sortedDiallingCodes = $diallingCodes | Sort-Object { $_.Length } -Descending
  
      foreach ($phoneNumber in $PhoneNumbers) {
        $matchedCountry = $null
  
        foreach ($code in $sortedDiallingCodes) {
          if ($phoneNumber -like ("+{0}*" -f $code)) {
            $matchedCountry = $Countries | Where-Object { $_.CountryDiallingCode -eq $code }
            break
          }
        }
  
        if ($null -eq $matchedCountry) {
          Write-Warning "No country dialling code found for the phone number $phoneNumber"
        }
        else {
          $output.Add([PSCustomObject]@{
              "PhoneNumber"         = $phoneNumber
              "CountryName"         = $matchedCountry.CountryName
              "CountryDiallingCode" = $matchedCountry.CountryDiallingCode
              "TeamZone"            = $matchedCountry.TeamZone
            })
        }
      }
    }
  
    end {
      return $output
    }
  }
  
  function Get-TenantData {
    [CmdletBinding()]
    param (
      [Parameter(Mandatory = $true)]
      [System.Collections.Generic.List[PSCustomObject]]$Countries
    )
  
    begin {
      $output = [System.Collections.Generic.List[PSCustomObject]]::new()
  
      $verifiedDomain = (Get-CsTenant | Select-Object -ExpandProperty VerifiedDomains | Where-Object { ($_.Name -like "*.onmicrosoft.com") -and ($_.Name -notlike "*.mail.onmicrosoft.com") }).Name
      $activeUsers = Get-CsOnlineUser | Where-Object { $_.AccountEnabled -eq $true } | Select-Object AccountEnabled
      $activeLicensedUsersWithVoiceRoutingPolicy = [System.Collections.Generic.List[PSCustomObject]]::new()
      $Groups = @()
      $Members = @()
      $voiceRoutes = Get-CsOnlineVoiceRoute
      $pstnUsages = [System.Collections.Generic.List[PSCustomObject]]::new()
      $sipcomPolicies = [System.Collections.Generic.List[PSCustomObject]]::new()
      $customerId = ""
      foreach ($voiceRoute in $voiceRoutes) {
          if ($voiceRoute.OnlinePstnGatewayList -like "*halo.sipcom.cloud") {
          [string]$customerId = $voiceRoute.OnlinePstnGatewayList
          $pstnUsages.Add($voiceRoute.OnlinePstnUsages)
          $customerId = $customerId.Substring(3,6)
        }
      }

      $GlobalPSTN = ""
      $GlobalVRP = "0"
      $voiceRoutingPolicyGlobal = Get-CsOnlineVoiceRoutingPolicy -Identity "Global"
      $pstnUsages | ForEach-Object {
          if ($voiceRoutingPolicyGlobal.OnlinePstnUsages -contains $_) {
          $GlobalVRP = "1"
            foreach ($a in $voiceRoutingPolicyGlobal.OnlinePstnUsages){
                    $GlobalPSTN = $GlobalPSTN + $a + " "
            }
          }
          }

  
      $voiceRoutingPolicy = Get-CsOnlineVoiceRoutingPolicy | Select-Object Identity, OnlinePstnUsages
  
      foreach ($voiceRoutingPolicy in $voiceRoutingPolicy) {
        $pstnUsages | ForEach-Object {
          if ($voiceRoutingPolicy.OnlinePstnUsages -contains $_) {
            $sipcomPolicies.Add($voiceRoutingPolicy.Identity)
          }
        }
      }
  
      $sipcomPolicies = $sipcomPolicies | Select-Object -Unique
      $sipcomPolicies =  $sipcompolicies -replace "Global",""
      $sipcomPolicies =  $sipcompolicies -replace "Tag:",""
      
    }
  
    process {
      foreach ($policy in $sipcomPolicies) {
        Get-CsOnlineUser | Where-Object {($_.EnterpriseVoiceEnabled -eq $true) -and ($_.OnlineVoiceRoutingPolicy -like $policy) } | Select-Object LineUri,identity,DisplayName | ForEach-Object {
          $activeLicensedUsersWithVoiceRoutingPolicy.Add($_)
        }
      }

      foreach ($policy in $sipcomPolicies){
        $Groups += Get-CsGroupPolicyAssignment | where PolicyName -eq $policy | Select GroupId
      }
  
      foreach ($group in $Groups){
        $Members += Get-AzureADGroupMember -ObjectId $group.GroupId -All $true | Select UserPrincipalName
      }

      foreach ($member in $Members) {
        Get-CsOnlineUser -Identity $member.UserPrincipalName | Select-Object LineUri,identity,DisplayName | ForEach-Object {
          $activeLicensedUsersWithVoiceRoutingPolicy.Add($_)
        }
      }

      foreach ($user in $activeLicensedUsersWithVoiceRoutingPolicy) {
        if ([string]::IsNullOrWhiteSpace($user.LineUri)) {
          continue
        }
  
        $phoneNumber = ($user.LineUri -split ":")[-1]
        $teamZone = (Get-ZoneFromPhoneNumbers -PhoneNumbers $phoneNumber -Countries $Countries).TeamZone
  
        $user | Add-Member NoteProperty -Name TeamZone -Value $teamZone
      }
      
        #Get Call Queue and AutoAttendants
    $CQ = Get-CsCallQueue | select identity, Name
    $AA = Get-CsAutoAttendant  | select identity, Name
  
    #Count Call Queues on HALO
    $CallQueueData = 0

    foreach ($i in $CQ){
        if ($activeLicensedUsersWithVoiceRoutingPolicy.DisplayName -match $i.Name){
        $CallQueueData++
        }
    }

 #Count AutoAttendants on HALO
    $AutoAttendantData = 0

    foreach ($i in $AA){
        if ($activeLicensedUsersWithVoiceRoutingPolicy.DisplayName -match $i.Name){
        $AutoAttendantData++
        }
    }
    $TotalHALOUsers = $activeLicensedUsersWithVoiceRoutingPolicy.Count-$AutoAttendantData-$CallQueueData
    }

    
  
    end {
      $output.Add([PSCustomObject]@{
          CustomerID          = $customerId
          VerifiedDomin       = $verifiedDomain
          TenantEnabledUsers  = $activeUsers.Count
          HALODREndpoints     = $activeLicensedUsersWithVoiceRoutingPolicy.Count
          HALODRCallQueues    = $CallQueueData
          HALODRAutoAttendants= $AutoAttendantData
          HALODRUsers         = $TotalHALOUsers
          GlobalVRP           = $GlobalVRP
          GlobalPSTN          = $GlobalPSTN
          ZoneOne             = @($activeLicensedUsersWithVoiceRoutingPolicy | Where-Object { $_.TeamZone -eq 1 }).Count
          ZoneTwo             = @($activeLicensedUsersWithVoiceRoutingPolicy | Where-Object { $_.TeamZone -eq 2 }).Count
          ZoneThree           = @($activeLicensedUsersWithVoiceRoutingPolicy | Where-Object { $_.TeamZone -eq 3 }).Count
          ZoneFour            = @($activeLicensedUsersWithVoiceRoutingPolicy | Where-Object { $_.TeamZone -eq 4 }).Count
          ZoneFive            = @($activeLicensedUsersWithVoiceRoutingPolicy | Where-Object { $_.TeamZone -eq 5 }).Count
          ZoneSix             = @($activeLicensedUsersWithVoiceRoutingPolicy | Where-Object { $_.TeamZone -eq 6 }).Count
        })
  
      return $output
    }
  }
  
  function Invoke-UploadTenantDataToBlobStorage {
    [CmdletBinding()]
    param (
      [Parameter(Mandatory = $true)]
      [ValidateScript({ -not ([string]::IsNullOrWhiteSpace($_)) })]
      [string]$AccountName,
  
      [Parameter(Mandatory = $true)]
      [ValidateScript({ -not ([string]::IsNullOrWhiteSpace($_)) })]
      [string]$ContainerName,
  
      [Parameter(Mandatory = $true)]
      [ValidateScript({ $_ -ne $null })]
      [PSCustomObject]$TenantData,
  
      [Parameter(Mandatory = $true)]
      [ValidateScript({ -not ([string]::IsNullOrWhiteSpace($_)) })]
      [string]$SasToken
      
    )
  
    begin {
      $output = New-Object System.Collections.Generic.List[PSCustomObject]
      $blobName = "$($TenantData.CustomerID).csv"
  
      $csvData = $TenantData | ConvertTo-Csv -NoTypeInformation
      $csvData = $csvData -replace '"', ''
      $bytes = [System.Text.Encoding]::UTF8.GetBytes($csvData -join "`n")
  
      $url = "https://saservicemanagerdata.blob.core.windows.net/customerdata/$($blobName)$SasToken"
  
      $headers = @{
        "x-ms-blob-type"         = "BlockBlob"
        "x-ms-blob-content-type" = "application/octet-stream"
      }
    }
  
    process {
      $response = Invoke-WebRequest -Uri $url -Method Put -Headers $headers -Body $bytes -ContentType "application/octet-stream" -ErrorAction SilentlyContinue
  
      if ($response.StatusCode -ge 200 -and $response.StatusCode -lt 300) {
        $output.Add([PSCustomObject]@{
            AccountName   = $AccountName
            ContainerName = $ContainerName
            BlobName      = $blobName
            Uploaded      = $true
            StatusCode    = $response.StatusCode
            Message       = "Successfully uploaded"
          })
      }
      else {
        $output.Add([PSCustomObject]@{
            AccountName   = $AccountName
            ContainerName = $ContainerName
            BlobName      = $blobName
            Uploaded      = $false
            StatusCode    = $response.StatusCode
            Message       = "Failed to upload"
          })
      }
    }
  
    end {
      return $output
    }
  }

#Import Required Module
import-module Az.Storage
import-module AzureAD
import-module MicrosoftTeams

 #Start execution 
  Connect-MicrosoftTeams
  Connect-AzureAD

  $token = Read-Host "Please Enter The SAS Token"
  $sas = "?sv=2022-11-02&ss=bf&srt=co&sp=rwdlaciytfx&se=2025-05-01T00:23:50Z&st=2023-11-28T17:23:50Z&spr=https&sig=" + $token
  
  #Get Storage Account context
  $context = New-AzStorageContext -StorageAccountName "saservicemanagerdata" -SasToken $sas
  #Get reference for container
  $container = Get-AzStorageContainer -Name "customerdata" -Context $context
  #Get reference for file
  $client = $container.CloudBlobContainer.GetBlockBlobReference("CountryData.json")
  #Read file contents into memory
  $file = $client.DownloadText()
  
  $countryData = $file | ConvertFrom-Json
  
  $tenantData = Get-TenantData -Countries $countryData
  
  Invoke-UploadTenantDataToBlobStorage -TenantData $tenantData -AccountName "saservicemanagerdata" -ContainerName "customerdata" -SasToken $sas

  Write-Output "Tenant Results:"
  Write-Output " "
  foreach ($i in $tenantData){
  Write-Output $i
  }
  Write-Output " "
  Write-Output " "
  Write-Output "###################################"
  Write-Output "###################################"
  Write-Output " "
  Write-Output "Script Complete"

  Disconnect-AzAccount
  Disconnect-MicrosoftTeams
