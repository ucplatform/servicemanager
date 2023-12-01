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
      $activeUsers = Get-CsOnlineUser | Where-Object { $_.AccountEnabled -eq $true }
      $activeLicensedUsersWithVoiceRoutingPolicy = [System.Collections.Generic.List[PSCustomObject]]::new()
      $callQueueData = Get-CsCallQueue
      $voiceRoutes = Get-CsOnlineVoiceRoute
      $pstnUsages = [System.Collections.Generic.List[PSCustomObject]]::new()
      $sipcomPolicies = [System.Collections.Generic.List[PSCustomObject]]::new()
      $customerId = ""
      foreach ($voiceRoute in $voiceRoutes) {
          $customerId = $voiceRoute.OnlinePstnGatewayList
        if ($voiceRoute.OnlinePstnGatewayList -like "*halo.sipcom.cloud") {
          $pstnUsages.Add($voiceRoute.OnlinePstnUsages)
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
  
      $sipcomPolicies = $sipcomPolicies.substring(4)
    }
  
    process {
      foreach ($policy in $sipcomPolicies) {
        Get-CsOnlineUser | Where-Object { ($_.AccountEnabled -eq $true) -and ($_.EnterpriseVoiceEnabled -eq $true) -and ($_.OnlineVoiceRoutingPolicy -like $policy) } | ForEach-Object {
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
    }
  
    end {
      $output.Add([PSCustomObject]@{
          CustomerID          = $customerId.SubString(3,6)
          VerifiedDomin       = $verifiedDomain
          TenantActiveUsers   = $activeUsers.Count
          SipcomPlatformUsers = $activeLicensedUsersWithVoiceRoutingPolicy.Count
          TenantCallQueues    = $callQueueData.Length
          ZoneOne             = ($activeLicensedUsersWithVoiceRoutingPolicy | Where-Object { $_.TeamZone -eq 1 }).Count
          ZoneTwo             = ($activeLicensedUsersWithVoiceRoutingPolicy | Where-Object { $_.TeamZone -eq 2 }).Count
          ZoneThree           = ($activeLicensedUsersWithVoiceRoutingPolicy | Where-Object { $_.TeamZone -eq 3 }).Count
          ZoneFour            = ($activeLicensedUsersWithVoiceRoutingPolicy | Where-Object { $_.TeamZone -eq 4 }).Count
          ZoneFive            = ($activeLicensedUsersWithVoiceRoutingPolicy | Where-Object { $_.TeamZone -eq 5 }).Count
          ZoneSix             = ($activeLicensedUsersWithVoiceRoutingPolicy | Where-Object { $_.TeamZone -eq 6 }).Count
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

 #Start execution 
  Connect-MicrosoftTeams

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
