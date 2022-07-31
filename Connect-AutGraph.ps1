# https://gcits.com/knowledge-base/export-customers-microsoft-secure-scores-to-csv-and-html-reports/
# Function is based on this Securescore report


Function Connect-AutGraph {
    [CmdletBinding()]
    Param(
        [ValidateSet("V1.0","Beta")]$GraphVersion = "Beta",
        [ValidateSet("manage.microsoft.com/api","graph.microsoft.com","manage.office.com/api")]$APIProvider ="Graph.Microsoft.Com" ,
        $EndPoint = "Users",
        [validateset ("Yes","No")]$All="NO",
        [ValidateSet ("0","3","5","7")]$Retry = "0",
        [Parameter(Mandatory = $True)]$access_Token)

        $URIForApplicationPermissionCall = https://$ApiProvider/$GraphVersion/$EndPoint

    # Try to access graph information X times
    $Stoploop = $false
        [int]$Retrycount = "1"
        do {
            try {
                $GraphResponse = $null
        
                $GraphResponse = (Invoke-RestMethod `
                    -Uri $URIForApplicationPermissionCall `
                    -Headers @{"Authorization" = "Bearer $access_token"} `
                    -Method GET -DisableKeepAlive -UseBasicParsing)  #  -ContentType "application/json" 
                        
                    $Stoploop = $true
            }
            catch {
                if ($Retrycount -gt $Retry) {
                    $Stoploop = $true
                    throw $Error[0].Exception

                }
                else {
                    Write-Host "Could not get Graph content. Retrying in 5 seconds..." -ForegroundColor DarkYellow
                    Start-Sleep -Seconds 5
                    $Retrycount ++
                    throw $Error[0].Exception
                }
            }
        }
        While ($Stoploop -eq $false)

        # Added this check because some endpoints gives you the return data in ().value and some endpoints just gives normal info, 
        # so this is to check what kind of information we receive back
        if ($graphresponse.Psobject.Properties.Name -notcontains "value") {
            $GraphOutput = $GraphResponse
        }
        else {
            $GraphOutput = $GraphResponse.value
        }
        


            # Check if there is presented a next link if the result is paged, then loop trough until all information is collected if NextLink parameter is set to Yes
            $nextGraphLink = $graphResponse.'@odata.nextLink'

            if ($All -like "Yes") {

            while ($null -ne $NextGraphLink){
            
                    Write-Verbose "Next link is: $nextGraphLink"

                    $graphResponse = (Invoke-RestMethod -Method Get -Uri $nextGraphLink -Headers @{"Authorization"="Bearer $access_Token"} -UseBasicParsing)
            
                    $nextGraphLink = $graphResponse.'@odata.nextLink'

                    if ($graphresponse.Psobject.Properties.Name -notcontains "value") {
                        $GraphOutput += $GraphResponse
                    }
                    Else {
                        $GraphOutput += $GraphResponse.Value
                    }
                    Start-Sleep -Seconds 1
            }
        }   
        
        Return $GraphOutput
        
}
