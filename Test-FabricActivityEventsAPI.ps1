# PowerShell script to test Fabric Admin Activity Events API
# Requires: Az.PowerShell module and admin permissions

# Execution examples: 
# .\Test-FabricActivityEventsAPI.ps1 -StartDateTime "2025-10-15T09:00:00.000" -EndDateTime "2025-10-15T10:00:00.000" 
# .\Test-FabricActivityEventsAPI.ps1 -StartDateTime "2025-10-15T09:00:00.000" -EndDateTime "2025-10-15T10:00:00.000" -Activity "ReadArtifact"

param(
    [Parameter(Mandatory=$false)]
    [string]$StartDateTime,
    
    [Parameter(Mandatory=$false)]
    [string]$EndDateTime,
    
    [Parameter(Mandatory=$false)]
    [string]$OperationId,
    
    [Parameter(Mandatory=$false)]
    [string]$Activity,
    
    [Parameter(Mandatory=$false)]
    [string]$UserId
)

# Note: Activity Events API requires UTC dates and the time range must be within the last 30 days
# The API returns events in 1-hour chunks maximum
if (-not $StartDateTime) {
    # Default to 1 hour ago, rounded to the hour
    $start = (Get-Date).AddHours(-1).ToUniversalTime()
    $StartDateTime = Get-Date $start -Format "yyyy-MM-dd'T'HH:00:00.000"
}

if (-not $EndDateTime) {
    # Default to current hour, rounded
    $end = (Get-Date).ToUniversalTime()
    $EndDateTime = Get-Date $end -Format "yyyy-MM-dd'T'HH:00:00.000"
}

# Function to check and install required modules
function Install-RequiredModules {
    $modules = @("Az.Accounts")
    
    foreach ($module in $modules) {
        if (!(Get-Module -ListAvailable -Name $module)) {
            Write-Host "Installing module: $module" -ForegroundColor Yellow
            Install-Module -Name $module -Force -AllowClobber -Scope CurrentUser
        }
        Import-Module $module -Force
    }
}

# Function to get access token for Power BI API
function Get-PowerBIAccessToken {
    try {
        # Interactive login to Azure AD
        Write-Host "Initiating interactive login..." -ForegroundColor Green
        Write-Host "A browser window should open for authentication. Please complete the login there." -ForegroundColor Yellow
        
        $context = Connect-AzAccount -Force
        
        if (!$context) {
            throw "Failed to authenticate with Azure AD"
        }
        
        Write-Host "Successfully authenticated as: $($context.Context.Account.Id)" -ForegroundColor Green
        Write-Host "Tenant: $($context.Context.Tenant.Id)" -ForegroundColor Gray
        
        # Get access token for Power BI API using the recommended method
        Write-Host "Requesting Power BI API access token..." -ForegroundColor Gray
        
        try {
            # Primary method using Get-AzAccessToken
            # Suppress the breaking change warning by setting WarningAction
            $tokenRequest = Get-AzAccessToken -ResourceUrl "https://analysis.windows.net/powerbi/api" -WarningAction SilentlyContinue
            $token = $tokenRequest.Token
            
            # Handle both string and SecureString token types for future compatibility
            if ($token -is [System.Security.SecureString]) {
                # Convert SecureString to plain text
                $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($token)
                $token = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
                [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
            }
            
            if ($token) {
                Write-Host "Token acquired successfully (expires: $($tokenRequest.ExpiresOn))" -ForegroundColor Gray
                return $token
            }
        }
        catch {
            Write-Warning "Primary token method failed: $($_.Exception.Message)"
        }
        
        # Fallback method using REST API directly
        Write-Host "Trying alternative token acquisition method..." -ForegroundColor Yellow
        
        $azContext = Get-AzContext
        $currentAzureContext = Get-AzContext
        $tenantId = $currentAzureContext.Tenant.Id
        $accountId = $currentAzureContext.Account.Id
        
        # Get token using Azure PowerShell profile
        $profileClient = [Microsoft.Azure.Commands.Common.Authentication.Abstractions.AzureRmProfileProvider]::Instance.Profile.DefaultContext
        $token = [Microsoft.Azure.Commands.Common.Authentication.AzureSession]::Instance.AuthenticationFactory.Authenticate(
            $profileClient.Account,
            $profileClient.Environment,
            $profileClient.Tenant.Id,
            $null,
            [Microsoft.Azure.Commands.Common.Authentication.ShowDialog]::Never,
            $null,
            "https://analysis.windows.net/powerbi/api"
        ).AccessToken
        
        if ($token) {
            Write-Host "Token acquired using fallback method" -ForegroundColor Gray
            return $token
        }
        
        throw "Unable to acquire access token using any available method"
    }
    catch {
        Write-Error "Failed to get access token: $($_.Exception.Message)"
        Write-Host "`nTroubleshooting tips:" -ForegroundColor Yellow
        Write-Host "1. Ensure you completed the browser authentication" -ForegroundColor Yellow
        Write-Host "2. Verify you have Power BI admin permissions" -ForegroundColor Yellow
        Write-Host "3. Try running: Disconnect-AzAccount, then run this script again" -ForegroundColor Yellow
        return $null
    }
}

# Function to call the Fabric Activity Events API
function Invoke-FabricActivityEventsAPI {
    param(
        [string]$AccessToken,
        [string]$StartTime,
        [string]$EndTime,
        [string]$FilterOperationId,
        [string]$FilterActivity,
        [string]$FilterUserId
    )
    
    try {
        # Construct the API URL with query parameters
        # Note: The API format should be 'yyyy-MM-ddTHH:mm:ss.fff' (no Z at the end)
        $baseUrl = "https://api.powerbi.com/v1.0/myorg/admin/activityevents"
        
        # Ensure dates are properly formatted (remove Z if present and ensure proper format)
        $startFormatted = $StartTime -replace 'Z$', ''
        $endFormatted = $EndTime -replace 'Z$', ''
        
        # Build URL with properly encoded parameters
        $fullUrl = "$baseUrl`?startDateTime='$startFormatted'&endDateTime='$endFormatted'"
        
        # Build OData filter for Activity and/or UserId (API only supports these two properties)
        $filterParts = @()
        
        if ($FilterActivity) {
            $filterParts += "Activity eq '$FilterActivity'"
            Write-Host "Filtering by Activity: $FilterActivity" -ForegroundColor Gray
        }
        
        if ($FilterUserId) {
            $filterParts += "UserId eq '$FilterUserId'"
            Write-Host "Filtering by UserId: $FilterUserId" -ForegroundColor Gray
        }
        
        if ($filterParts.Count -gt 0) {
            $filterString = $filterParts -join ' and '
            $fullUrl += "&`$filter=$filterString"
        }
        
        if ($FilterOperationId) {
            Write-Host "Note: Will filter results locally for Operation ID: $FilterOperationId (API doesn't support Id filtering)" -ForegroundColor Gray
        }
        
        Write-Host "Calling API endpoint: $fullUrl" -ForegroundColor Cyan
        Write-Host "Time range: $startFormatted to $endFormatted (UTC)" -ForegroundColor Gray
        
        # Calculate time difference
        $startDate = [DateTime]::Parse($startFormatted)
        $endDate = [DateTime]::Parse($endFormatted)
        $duration = $endDate - $startDate
        
        if ($duration.TotalDays -gt 30) {
            Write-Warning "Time range exceeds 30 days. The API only supports queries within the last 30 days."
        }
        
        if ($duration.TotalHours -gt 24) {
            Write-Warning "Time range exceeds 24 hours. Large time ranges may require pagination."
        }
        
        # Prepare headers
        $headers = @{
            'Authorization' = "Bearer $AccessToken"
            'Content-Type' = 'application/json'
        }
        
        # Make the API call
        $response = Invoke-RestMethod -Uri $fullUrl -Method GET -Headers $headers -ErrorAction Stop
        
        # Apply local filtering if OperationId is specified
        if ($FilterOperationId -and $response.activityEventEntities) {
            $originalCount = $response.activityEventEntities.Count
            $response.activityEventEntities = @($response.activityEventEntities | Where-Object { $_.Id -eq $FilterOperationId })
            Write-Host "Filtered from $originalCount events to $($response.activityEventEntities.Count) matching Operation ID" -ForegroundColor Gray
        }
        
        return $response
    }
    catch {
        Write-Error "API call failed: $($_.Exception.Message)"
        if ($_.Exception.Response) {
            $statusCode = $_.Exception.Response.StatusCode
            Write-Error "Status Code: $statusCode"
            
            # Try to read the response content for more details
            try {
                $result = $_.Exception.Response.Content.ReadAsStringAsync().Result
                Write-Error "Response Content: $result"
            }
            catch {
                Write-Warning "Could not read response content: $($_.Exception.Message)"
            }
        }
        
        Write-Host "`nCommon issues:" -ForegroundColor Yellow
        Write-Host "- Date format must be: yyyy-MM-ddTHH:mm:ss.fff" -ForegroundColor Yellow
        Write-Host "- Time range must be within the last 30 days" -ForegroundColor Yellow
        Write-Host "- Dates must be in UTC" -ForegroundColor Yellow
        Write-Host "- You must have Fabric/Power BI admin permissions" -ForegroundColor Yellow
        
        return $null
    }
}

# Function to display results
function Show-Results {
    param($Results)
    
    if ($Results) {
        Write-Host "`n=== API Response ===" -ForegroundColor Green
        
        # Check if results contain activity events
        if ($Results.activityEventEntities) {
            Write-Host "Found $($Results.activityEventEntities.Count) activity events" -ForegroundColor Green
            
            # Display first few events as examples
            $Results.activityEventEntities | Select-Object -First 5 | ForEach-Object {
                Write-Host "`nActivity: $($_.Activity)" -ForegroundColor Yellow
                Write-Host "Operation ID: $($_.Id)" -ForegroundColor Yellow
                Write-Host "User: $($_.UserId)" -ForegroundColor Yellow
                Write-Host "Creation Time: $($_.CreationTime)" -ForegroundColor Yellow
                Write-Host "Workspace: $($_.WorkSpaceName)" -ForegroundColor Yellow
                Write-Host "---"
            }
            
            if ($Results.activityEventEntities.Count -gt 5) {
                Write-Host "... and $($Results.activityEventEntities.Count - 5) more events" -ForegroundColor Gray
            }
        }
        else {
            Write-Host "No activity events found in the specified time range" -ForegroundColor Yellow
        }
        
        # Display continuation URL if present
        if ($Results.continuationUri) {
            Write-Host "`nContinuation URL available for pagination: $($Results.continuationUri)" -ForegroundColor Cyan
        }
        
        # Save full results to file
        $outputFile = "ActivityEvents_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
        $Results | ConvertTo-Json -Depth 10 | Out-File -FilePath $outputFile -Encoding UTF8
        Write-Host "`nFull results saved to: $outputFile" -ForegroundColor Green
    }
    else {
        Write-Host "No results to display" -ForegroundColor Red
    }
}

# Main execution
try {
    Write-Host "=== Fabric Admin Activity Events API Test ===" -ForegroundColor Magenta
    Write-Host "Start Time: $StartDateTime" -ForegroundColor Gray
    Write-Host "End Time: $EndDateTime" -ForegroundColor Gray
    if ($Activity) {
        Write-Host "Filter by Activity: $Activity" -ForegroundColor Gray
    }
    if ($UserId) {
        Write-Host "Filter by UserId: $UserId" -ForegroundColor Gray
    }
    if ($OperationId) {
        Write-Host "Filter by Operation ID (local): $OperationId" -ForegroundColor Gray
    }
    Write-Host ""
    
    # Install required modules
    Write-Host "Checking required PowerShell modules..." -ForegroundColor Blue
    Install-RequiredModules
    
    # Get access token
    Write-Host "`nGetting access token..." -ForegroundColor Blue
    $accessToken = Get-PowerBIAccessToken
    
    if (!$accessToken) {
        throw "Failed to obtain access token"
    }
    
    Write-Host "Access token obtained successfully" -ForegroundColor Green
    
    # Call the API
    Write-Host "`nCalling Fabric Activity Events API..." -ForegroundColor Blue
    $results = Invoke-FabricActivityEventsAPI -AccessToken $accessToken -StartTime $StartDateTime -EndTime $EndDateTime -FilterOperationId $OperationId -FilterActivity $Activity -FilterUserId $UserId
    
    # Display results
    Show-Results -Results $results
    
    Write-Host "`n=== Test Completed Successfully ===" -ForegroundColor Magenta
}
catch {
    Write-Error "Script execution failed: $($_.Exception.Message)"
    Write-Host "`n=== Test Failed ===" -ForegroundColor Red
}

# Cleanup
Write-Host "`nPress any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
