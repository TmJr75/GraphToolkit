#
# Get-AutRefreshToken.ps1
#
Function Get-AutRefreshToken {
    [cmdletbinding()]
    Param($appid,
      $appSecret, 
      $tenantId, 
      [Validateset("outlook.office.com/.default", 
        "manage.office.com", 
        "graph.microsoft.com", 
        "management.core.windows.net", 
        "vault.azure.net")]$api = "graph.microsoft.com"
      )
   
    $Authority = "https://login.microsoftonline.com/$tenantId"
    
    $tokenEndpointUri = "$authority/oauth2/token"
    $Contentbody = @{"grant_type" = "client_credentials";
      "Client_Id"                 = $AppId;
      "Client_Secret"             = $AppSecret;
      "Resource"                  = "https://$api"
    }
  
    
    # Try to collect an access token using your app.
    try {
      $response = Invoke-RestMethod -Uri $tokenEndpointUri -Body $Contentbody -Method Post -UseBasicParsing -ErrorAction Stop
      return $response.access_token
    }
    Catch {
      Write-Output "Unable to request access token."
      Write-Output $error[0].exception
    }
    
    
  }
    
  Write-Output "Performed on host: $($env:ComputerName)"
