<#
Power BI Scanner (Admin) – JSON + CSV export with improved column/datatype handling
Runs with your Entra ID account (interactive login), no service principal required.
Requires: MicrosoftPowerBIMgmt (Profile/Admin modules)

References:
- Connect-PowerBIServiceAccount (interactive sign-in) https://learn.microsoft.com/powershell/module/microsoftpowerbimgmt.profile/connect-powerbiserviceaccount
- Admin – GetGroupsAsAdmin (list workspaces) https://learn.microsoft.com/rest/api/power-bi/admin/groups-get-groups-as-admin
- Admin – WorkspaceInfo getInfo / scanStatus / scanResult https://learn.microsoft.com/rest/api/power-bi/admin/workspace-info-post-workspace-info
                                                    https://learn.microsoft.com/rest/api/power-bi/admin/workspace-info-get-scan-status
                                                    https://learn.microsoft.com/rest/api/power-bi/admin/workspace-info-get-scan-result
- Metadata scanning overview https://learn.microsoft.com/fabric/governance/metadata-scanning-overview
#>

[CmdletBinding()]
param(
  [string]$ExportRoot = (Join-Path $HOME "PowerBI_Metadata"),
  [switch]$SelectWorkspaces,          # show interactive selector
  [string]$WorkspaceNameLike,         # wildcard filter for workspace names
  [switch]$IncludeLineage = $true,    # lineage edges
  [switch]$IncludeDatasourceDetails = $true,
  [switch]$IncludeDatasetSchema = $true,        # tables/columns/measures
  [switch]$IncludeDatasetExpressions = $true,   # Requires metadata scanning to be enabled in tenant settings
  [switch]$IncludeRefreshHistory = $true,      # fetch last refresh datetime for datasets
  [int]$MaxRetries = 3,               # retry attempts for failed operations
  [int]$RetryDelaySeconds = 10        # delay between retries
)

# -------- Global Variables for Tracking --------
$script:ScanStats = @{
  TotalWorkspaces = 0
  SuccessfulScans = 0
  FailedScans = 0
  TotalDatasets = 0
  DatasetsWithSchema = 0
  DatasetsWithoutSchema = 0
  TotalTables = 0
  TotalColumns = 0
  DatasetsWithRefreshHistory = 0
  ProcessingErrors = @()
}

# -------- Helpers --------
function New-Folder { 
  param([string]$Path) 
  if (-not (Test-Path $Path)) { 
    New-Item -ItemType Directory -Path $Path | Out-Null 
  } 
}

function Write-Progress-Info {
  param(
    [string]$Message,
    [string]$Color = "Yellow"
  )
  $timestamp = Get-Date -Format "HH:mm:ss"
  Write-Host "[$timestamp] $Message" -ForegroundColor $Color
}

function Invoke-PBIAdminGet {
  param(
    [Parameter(Mandatory)][string]$Url
  )
  
  $attempt = 0
  do {
    try {
      $resp = Invoke-PowerBIRestMethod -Url $Url -Method Get -ErrorAction Stop | ConvertFrom-Json
      return $resp
    } catch {
      $attempt++
      
      # Check if it's a 404 error - don't retry these as they indicate the resource doesn't exist/support the API
      if ($_.Exception.Message -like "*404*" -or $_.Exception.Message -like "*Not Found*") {
        throw $_ # Re-throw 404 errors immediately without retrying
      }
      
      if ($attempt -le $MaxRetries) {
        Write-Progress-Info "API call failed (attempt $attempt/$MaxRetries), retrying in $RetryDelaySeconds seconds: $($_.Exception.Message)" "Yellow"
        Start-Sleep -Seconds $RetryDelaySeconds
      } else {
        throw $_
      }
    }
  } while ($attempt -le $MaxRetries)
}

function Invoke-PBIAdminPostJson {
  param(
    [Parameter(Mandatory)][string]$Url,
    [Parameter(Mandatory)][object]$BodyObject
  )
  
  $json = $BodyObject | ConvertTo-Json -Depth 10
  $attempt = 0
  
  do {
    try {
      $resp = Invoke-PowerBIRestMethod -Url $Url -Method Post -ContentType "application/json" -Body $json | ConvertFrom-Json
      return $resp
    } catch {
      $attempt++
      if ($attempt -le $MaxRetries) {
        Write-Progress-Info "API call failed (attempt $attempt/$MaxRetries), retrying in $RetryDelaySeconds seconds: $($_.Exception.Message)" "Yellow"
        Start-Sleep -Seconds $RetryDelaySeconds
      } else {
        throw $_
      }
    }
  } while ($attempt -le $MaxRetries)
}

function Get-AllWorkspacesAsAdmin {
  [CmdletBinding()]
  param()
  $top  = 5000
  $skip = 0
  $all = @()
  
  Write-Progress-Info "Fetching workspaces from Power BI Admin API..."
  
  do {
    Write-Progress-Info "Fetching workspaces batch: skip=$skip, top=$top"
    $page = Invoke-PBIAdminGet -Url "/admin/groups?`$top=$top&`$skip=$skip"
    $all += $page.value
    $skip += $top
    Write-Progress-Info "Retrieved $($page.value.Count) workspaces in this batch (total: $($all.Count))"
  } while ($page.value.Count -eq $top)
  
  Write-Progress-Info "Retrieved $($all.Count) total workspaces" "Green"
  return $all
}

function Select-WorkspaceIds {
  param(
    [Parameter(Mandatory)][array]$AllWorkspaces,
    [switch]$Select,
    [string]$Like
  )
  $candidates = $AllWorkspaces | Where-Object { $_.state -eq "Active" -and $_.type -eq "Workspace" }
  if ($Like) { $candidates = $candidates | Where-Object { $_.name -like $Like } }

  Write-Progress-Info "Found $($candidates.Count) matching workspaces"

  if ($Select) {
    $supportOGV = Get-Command Out-GridView -ErrorAction SilentlyContinue
    if ($supportOGV) {
      $picked = $candidates | Select-Object name,id,isOnDedicatedCapacity,capacityId | Out-GridView -Title "Select workspaces to scan (multi-select), then click OK" -PassThru
      return @($picked.id)
    } else {
      Write-Host "`nInteractive selection:" -ForegroundColor Cyan
      $list = $candidates | ForEach-Object -Begin { $i=1 } -Process {
        "{0}. {1}  ({2})" -f $i, $_.name, $_.id; $i++
      }
      $list | ForEach-Object { Write-Host $_ }
      $inputIdx = Read-Host "Enter comma-separated numbers to select, or press Enter for ALL shown"
      if ([string]::IsNullOrWhiteSpace($inputIdx)) { 
        Write-Progress-Info "Returning all $($candidates.Count) workspace IDs"
        return @($candidates.id)
      }
      
      $indices = $inputIdx -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^\d+$' } | ForEach-Object { [int]$_ }
      $selected = @()
      foreach ($index in $indices) {
        if ($index -ge 1 -and $index -le $candidates.Count) {
          $selected += $candidates[$index - 1].id
        } else {
          Write-Warning "Invalid selection: $index (must be between 1 and $($candidates.Count))"
        }
      }
      
      Write-Progress-Info "Selected $($selected.Count) workspace(s): $($selected -join ', ')" "Green"
      return @($selected)
    }
  }

  Write-Progress-Info "Returning all $($candidates.Count) workspace IDs (no selection mode)"
  return @($candidates.id)
}

function Start-WorkspaceScan {
  param(
    [Parameter(Mandatory)][string[]]$WorkspaceIds,
    [switch]$Lineage,
    [switch]$DatasourceDetails,
    [switch]$DatasetSchema,
    [switch]$DatasetExpressions
  )
  $qs = @()
  if ($Lineage)            { $qs += "lineage=true" }
  if ($DatasourceDetails)  { $qs += "datasourceDetails=true" }
  if ($DatasetSchema)      { $qs += "datasetSchema=true" }
  if ($DatasetExpressions) { $qs += "datasetExpressions=true" }
  $qsStr = if ($qs) { "?" + ($qs -join "&") } else { "" }

  Write-Progress-Info "Scan options: $($qs -join ', ')"
  Write-Progress-Info "API URL: /admin/workspaces/getInfo$qsStr"
  Write-Progress-Info "Workspace IDs to scan: $($WorkspaceIds.Count) workspaces"

  # Validate workspace IDs
  foreach ($wsId in $WorkspaceIds) {
    if ([string]::IsNullOrWhiteSpace($wsId)) {
      throw "Empty workspace ID detected"
    }
    if ($wsId -notmatch '^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$') {
      Write-Warning "Workspace ID $wsId doesn't appear to be a valid GUID"
    }
  }

  $body = @{ workspaces = $WorkspaceIds }
  
  try {
    $resp = Invoke-PBIAdminPostJson -Url ("/admin/workspaces/getInfo{0}" -f $qsStr) -BodyObject $body
    
    if ($resp.id) {
      Write-Progress-Info "Scan ID: $($resp.id)" "Green"
      return $resp.id
    } else {
      throw "No scan ID returned from API response"
    }
  } catch {
    Write-Progress-Info "Scan request failed: $($_.Exception.Message)" "Red"
    if ($_.Exception.Message -like "*BadRequest*" -or $_.Exception.Message -like "*Invalid value*") {
      Write-Progress-Info "This often indicates:" "Yellow"
      Write-Progress-Info "  - Workspace IDs are invalid or don't exist" "Yellow"
      Write-Progress-Info "  - You don't have access to the specified workspaces" "Yellow"
      Write-Progress-Info "  - Metadata scanning is not enabled in tenant settings" "Yellow"
    }
    throw
  }
}

function Wait-ScanSucceeded {
  param([Parameter(Mandatory)][string]$ScanId)
  
  $startTime = Get-Date
  $checkCount = 0
  
  do {
    Start-Sleep -Seconds 5
    $checkCount++
    $status = Invoke-PBIAdminGet -Url "/admin/workspaces/scanStatus/$ScanId"
    $elapsed = ((Get-Date) - $startTime).TotalSeconds
    Write-Progress-Info "Scan $ScanId status: $($status.status) (check #$checkCount, ${elapsed}s elapsed)"
  } while ($status.status -in @("NotStarted","Running"))

  if ($status.status -ne "Succeeded") {
    throw "Scanner failed with status: $($status.status)"
  }
  
  $totalTime = ((Get-Date) - $startTime).TotalSeconds
  Write-Progress-Info "Scan completed successfully in ${totalTime}s" "Green"
}

function Get-ScanResult {
  param([Parameter(Mandatory)][string]$ScanId)
  $raw = Invoke-PowerBIRestMethod -Url "/admin/workspaces/scanResult/$ScanId" -Method Get
  return $raw
}

function Get-DatasetRefreshHistory {
  param(
    [Parameter(Mandatory)][string]$WorkspaceId,
    [Parameter(Mandatory)][string]$DatasetId,
    [int]$Top = 1  # Get only the most recent refresh
  )
  
  try {
    $url = "/groups/$WorkspaceId/datasets/$DatasetId/refreshes?`$top=$Top"
    
    # Use ErrorAction Stop to catch errors and prevent console output
    $refreshHistory = $null
    $refreshHistory = Invoke-PowerBIRestMethod -Url $url -Method Get -ErrorAction Stop | ConvertFrom-Json
    
    if ($refreshHistory.value -and $refreshHistory.value.Count -gt 0) {
      $lastRefresh = $refreshHistory.value[0]
      return @{
        LastRefreshStartTime = $lastRefresh.startTime
        LastRefreshEndTime = $lastRefresh.endTime
        LastRefreshStatus = $lastRefresh.status
        LastRefreshType = $lastRefresh.refreshType
        LastRefreshRequestId = $lastRefresh.requestId
        HasRefreshHistory = $true
      }
    } else {
      return @{
        LastRefreshStartTime = $null
        LastRefreshEndTime = $null
        LastRefreshStatus = "NoHistory"
        LastRefreshType = $null
        LastRefreshRequestId = $null
        HasRefreshHistory = $false
      }
    }
  } catch {
    # Check if it's a 404 error (dataset doesn't support refresh API)
    if ($_.Exception.Message -like "*404*" -or $_.Exception.Message -like "*Not Found*") {
      Write-Progress-Info "Dataset $DatasetId does not support refresh API (likely Fabric lakehouse/warehouse/dataflow)" "Cyan"
      return @{
        LastRefreshStartTime = $null
        LastRefreshEndTime = $null
        LastRefreshStatus = "NotSupported"
        LastRefreshType = $null
        LastRefreshRequestId = $null
        HasRefreshHistory = $false
        RefreshHistoryError = "Dataset type does not support refresh API"
      }
    } else {
      Write-Progress-Info "Could not get refresh history for dataset $DatasetId in workspace $WorkspaceId : $($_.Exception.Message)" "Yellow"
      return @{
        LastRefreshStartTime = $null
        LastRefreshEndTime = $null
        LastRefreshStatus = "Error"
        LastRefreshType = $null
        LastRefreshRequestId = $null
        HasRefreshHistory = $false
        RefreshHistoryError = $_.Exception.Message
      }
    }
  }
}

function Get-CapacityRefreshables {
  param([Parameter(Mandatory)][string]$CapacityId)
  
  try {
    Write-Progress-Info "Fetching refreshables for capacity: $CapacityId"
    $url = "/admin/capacities/$CapacityId/refreshables?`$top=1000"
    $refreshables = Invoke-PBIAdminGet -Url $url
    
    if ($refreshables.value) {
      Write-Progress-Info "Found $($refreshables.value.Count) refreshables in capacity $CapacityId" "Green"
      return $refreshables.value
    } else {
      Write-Progress-Info "No refreshables found for capacity $CapacityId" "Yellow"
      return @()
    }
  } catch {
    Write-Progress-Info "Could not get refreshables for capacity $CapacityId : $($_.Exception.Message)" "Yellow"
    return @()
  }
}

function Test-ScanResultQuality {
  param(
    [Parameter(Mandatory)][pscustomobject]$ScanObject,
    [Parameter(Mandatory)][string[]]$ExpectedWorkspaceIds
  )
  
  Write-Progress-Info "Validating scan result quality..."
  
  $issues = @()
  
  # Check if we got all expected workspaces
  $returnedIds = @($ScanObject.workspaces | ForEach-Object { $_.id })
  $missingIds = $ExpectedWorkspaceIds | Where-Object { $_ -notin $returnedIds }
  
  if ($missingIds) {
    $issues += "Missing $($missingIds.Count) expected workspaces: $($missingIds -join ', ')"
  }
  
  # Check for workspaces without datasets
  $workspacesWithoutDatasets = $ScanObject.workspaces | Where-Object { 
    $null -eq $_.datasets -or $_.datasets.Count -eq 0 
  }
  
  if ($workspacesWithoutDatasets) {
    $issues += "$($workspacesWithoutDatasets.Count) workspaces have no datasets"
  }
  
  # Check datasets without schema when schema was requested
  if ($IncludeDatasetSchema) {
    $datasetsWithoutSchema = foreach ($ws in $ScanObject.workspaces) {
      foreach ($ds in ($ws.datasets | Where-Object { $_ })) {
        if ($null -eq $ds.tables -or $ds.tables.Count -eq 0) {
          "$($ws.name)/$($ds.name)"
        }
      }
    }
    
    if ($datasetsWithoutSchema) {
      $issues += "$($datasetsWithoutSchema.Count) datasets missing schema: $($datasetsWithoutSchema -join ', ')"
      $script:ScanStats.DatasetsWithoutSchema += $datasetsWithoutSchema.Count
    }
  }
  
  if ($issues) {
    Write-Progress-Info "Scan quality issues detected:" "Yellow"
    $issues | ForEach-Object { Write-Progress-Info "  - $_" "Yellow" }
    $script:ScanStats.ProcessingErrors += $issues
  }
  
  return $issues.Count -eq 0
}

function Merge-ScanResults {
  param([Parameter(Mandatory)][string[]]$JsonPayloads)
  # Simple merge: concatenate .workspaces arrays and de-duplicate by id
  $merged = @{ workspaces = @() }
  $h = @{}
  foreach ($jp in $JsonPayloads) {
    $o = $jp | ConvertFrom-Json
    foreach ($w in ($o.workspaces | Where-Object { $_ })) {
      if (-not $h.ContainsKey($w.id)) {
        $h[$w.id] = $true
        $merged.workspaces += $w
      }
    }
  }
  return ($merged | ConvertTo-Json -Depth 50)
}

function Export-FlatCsvs {
  param(
    [Parameter(Mandatory)][pscustomobject]$ScanObject,
    [Parameter(Mandatory)][string]$OutDir,
    [Parameter(Mandatory)][string]$Stamp
  )
  
  Write-Progress-Info "Analyzing scan results for CSV export..."
  Write-Progress-Info "Found $($ScanObject.workspaces.Count) workspaces"
  
  # Update statistics
  $script:ScanStats.TotalWorkspaces = $ScanObject.workspaces.Count
  
  # Debug: Show what's in the first workspace
  if ($ScanObject.workspaces.Count -gt 0) {
    $firstWs = $ScanObject.workspaces[0]
    Write-Progress-Info "First workspace sample:" "Cyan"
    Write-Progress-Info "  - Reports: $($firstWs.reports.Count)" "Cyan"
    Write-Progress-Info "  - Datasets: $($firstWs.datasets.Count)" "Cyan"
    if ($firstWs.datasets.Count -gt 0) {
      $firstDs = $firstWs.datasets[0]
      Write-Progress-Info "  - First dataset tables: $($firstDs.tables.Count)" "Cyan"
      if ($firstDs.tables.Count -gt 0) {
        $firstTable = $firstDs.tables[0]
        Write-Progress-Info "  - First table columns: $($firstTable.columns.Count)" "Cyan"
        Write-Progress-Info "  - First table measures: $($firstTable.measures.Count)" "Cyan"
      }
    }
  }
  
  # Count datasets with schema
  $datasetsWithSchema = 0
  foreach ($w in $ScanObject.workspaces) {
    foreach ($d in ($w.datasets | Where-Object { $_ })) {
      $script:ScanStats.TotalDatasets++
      if ($d.tables -and $d.tables.Count -gt 0) {
        $datasetsWithSchema++
        $script:ScanStats.DatasetsWithSchema++
      }
    }
  }
  
  Write-Progress-Info "Datasets with schema information: $datasetsWithSchema of $($script:ScanStats.TotalDatasets)" "Green"
  
  # Workspaces
  Write-Progress-Info "Exporting workspaces..."
  $ScanObject.workspaces |
    Select-Object id,name,type,state,isOnDedicatedCapacity,capacityId,defaultDatasetStorageFormat |
    Export-Csv (Join-Path $OutDir "workspaces_$Stamp.csv") -NoTypeInformation -Encoding UTF8

  # Reports
  Write-Progress-Info "Exporting reports..."
  $reports = foreach ($w in $ScanObject.workspaces) {
    foreach ($r in ($w.reports | Where-Object { $_ })) {
      [pscustomobject]@{
        workspaceId = $w.id
        workspaceName = $w.name
        reportId    = $r.id
        reportName  = $r.name
        datasetId   = $r.datasetId
        reportType  = $r.reportType
        created     = $r.createdDateTime
        modified    = $r.modifiedDateTime
        webUrl      = $r.webUrl
        embedUrl    = $r.embedUrl
      }
    }
  }
  if ($reports) { 
    Write-Progress-Info "Exported $($reports.Count) reports"
    $reports | Export-Csv (Join-Path $OutDir "reports_$Stamp.csv") -NoTypeInformation -Encoding UTF8 
  }

  # Datasets with refresh history information
  Write-Progress-Info "Exporting datasets with refresh history..."
  $datasets = foreach ($w in $ScanObject.workspaces) {
    foreach ($d in ($w.datasets | Where-Object { $_ })) {
      # Get refresh history if enabled
      $refreshInfo = @{
        LastRefreshStartTime = $null
        LastRefreshEndTime = $null
        LastRefreshStatus = "NotChecked"
        LastRefreshType = $null
        HasRefreshHistory = $false
      }
      
      if ($IncludeRefreshHistory) {
        $refreshInfo = Get-DatasetRefreshHistory -WorkspaceId $w.id -DatasetId $d.id
        if ($refreshInfo.HasRefreshHistory) {
          $script:ScanStats.DatasetsWithRefreshHistory++
        }
      }
      
      [pscustomobject]@{
        workspaceId = $w.id
        workspaceName = $w.name
        datasetId   = $d.id
        datasetName = $d.name
        isOnPrem    = $d.isOnPremGatewayRequired
        configuredBy= $d.configuredBy
        isRefreshable = $d.isRefreshable
        isEffectiveIdentityRequired = $d.isEffectiveIdentityRequired
        isEffectiveIdentityRolesRequired = $d.isEffectiveIdentityRolesRequired
        targetStorageMode = $d.targetStorageMode
        actualStorage = $d.actualStorage
        createdDate = $d.createdDate
        contentProviderType = $d.contentProviderType
        hasSchemaData = ($null -ne $d.tables -and $d.tables.Count -gt 0)
        tableCount = if ($d.tables) { $d.tables.Count } else { 0 }
        # Refresh history information
        lastRefreshStartTime = $refreshInfo.LastRefreshStartTime
        lastRefreshEndTime = $refreshInfo.LastRefreshEndTime
        lastRefreshStatus = $refreshInfo.LastRefreshStatus
        lastRefreshType = $refreshInfo.LastRefreshType
        hasRefreshHistory = $refreshInfo.HasRefreshHistory
        refreshHistoryError = $refreshInfo.RefreshHistoryError
      }
    }
  }
  if ($datasets) { 
    Write-Progress-Info "Exported $($datasets.Count) datasets"
    $datasets | Export-Csv (Join-Path $OutDir "datasets_$Stamp.csv") -NoTypeInformation -Encoding UTF8 
  }

  # Tables with better error handling
  Write-Progress-Info "Exporting tables..."
  $tables = foreach ($w in $ScanObject.workspaces) {
    foreach ($d in ($w.datasets | Where-Object { $_ })) {
      if ($null -ne $d.tables) {
        foreach ($t in ($d.tables | Where-Object { $_ })) {
          $script:ScanStats.TotalTables++
          [pscustomobject]@{
            workspaceId = $w.id
            workspaceName = $w.name
            datasetId   = $d.id
            datasetName = $d.name
            tableName   = $t.name
            isHidden    = $t.isHidden
            description = $t.description
            source      = $t.source
            columnCount = if ($t.columns) { $t.columns.Count } else { 0 }
            measureCount = if ($t.measures) { $t.measures.Count } else { 0 }
          }
        }
      }
    }
  }
  if ($tables) { 
    Write-Progress-Info "Exported $($tables.Count) tables"
    $tables | Export-Csv (Join-Path $OutDir "dataset_tables_$Stamp.csv") -NoTypeInformation -Encoding UTF8 
  }

  #  Columns export with better null handling and additional metadata
  Write-Progress-Info "Exporting columns with metadata..."
  $columns = foreach ($w in $ScanObject.workspaces) {
    foreach ($d in ($w.datasets | Where-Object { $_ })) {
      if ($null -ne $d.tables) {
        foreach ($t in ($d.tables | Where-Object { $_ })) {
          if ($null -ne $t.columns) {
            foreach ($c in ($t.columns | Where-Object { $_ })) {
              $script:ScanStats.TotalColumns++
              [pscustomobject]@{
                workspaceId = $w.id
                workspaceName = $w.name
                datasetId   = $d.id
                datasetName = $d.name
                tableName   = $t.name
                columnName  = $c.name
                dataType    = $c.dataType
                isHidden    = $c.isHidden
                isCalculated= $c.isCalculated
                isKey       = $c.isKey
                isNullable  = $c.isNullable
                dataCategory= $c.dataCategory
                description = $c.description
                expression  = $c.expression
                formatString= $c.formatString
                sortByColumn= $c.sortByColumn
                summarizeBy = $c.summarizeBy
                type        = $c.type
              }
            }
          }
        }
      }
    }
  }
  if ($columns) { 
    Write-Progress-Info "Exported $($columns.Count) columns" "Green"
    $columns | Export-Csv (Join-Path $OutDir "dataset_columns_$Stamp.csv") -NoTypeInformation -Encoding UTF8 
  } else {
    Write-Progress-Info "No columns found - this may indicate missing schema permissions or datasets without schema support" "Yellow"
  }

  # Measures (includes DAX expression when datasetExpressions=true)
  Write-Progress-Info "Exporting measures..."
  $measures = foreach ($w in $ScanObject.workspaces) {
    foreach ($d in ($w.datasets | Where-Object { $_ })) {
      if ($null -ne $d.tables) {
        foreach ($t in ($d.tables | Where-Object { $_ })) {
          if ($null -ne $t.measures) {
            foreach ($m in ($t.measures | Where-Object { $_ })) {
              [pscustomobject]@{
                workspaceId = $w.id
                workspaceName = $w.name
                datasetId   = $d.id
                datasetName = $d.name
                tableName   = $t.name
                measureName = $m.name
                expression  = $m.expression
                isHidden    = $m.isHidden
                description = $m.description
                displayFolder = $m.displayFolder
                formatString = $m.formatString
              }
            }
          }
        }
      }
    }
  }
  if ($measures) { 
    Write-Progress-Info "Exported $($measures.Count) measures"
    $measures | Export-Csv (Join-Path $OutDir "dataset_measures_$Stamp.csv") -NoTypeInformation -Encoding UTF8 
  }

  # Export Datasources with connection details
  Write-Progress-Info "Exporting datasources..."
  $datasources = foreach ($w in $ScanObject.workspaces) {
    foreach ($d in ($w.datasets | Where-Object { $_ })) {
      if ($null -ne $d.datasources) {
        foreach ($s in ($d.datasources | Where-Object { $_ })) {
          [pscustomobject]@{
            workspaceId  = $w.id
            workspaceName = $w.name
            datasetId    = $d.id
            datasetName  = $d.name
            datasourceId = $s.id
            datasourceType = $s.datasourceType
            connectionDetails = if ($s.connectionDetails) { $s.connectionDetails | ConvertTo-Json -Compress } else { $null }
            gatewayId    = $s.gatewayId
            server       = $s.connectionDetails.server
            database     = $s.connectionDetails.database
          }
        }
      }
    }
  }
  if ($datasources) { 
    Write-Progress-Info "Exported $($datasources.Count) datasources"
    $datasources | Export-Csv (Join-Path $OutDir "dataset_datasources_$Stamp.csv") -NoTypeInformation -Encoding UTF8 
  }

  # Lineage edges
  Write-Progress-Info "Exporting lineage..."
  $edges = foreach ($w in $ScanObject.workspaces) {
    foreach ($d in ($w.datasets | Where-Object { $_ })) {
      if ($null -ne $d.upstreamDataflows) {
        foreach ($u in ($d.upstreamDataflows | Where-Object { $_ })) {
          [pscustomobject]@{
            workspaceId = $w.id
            workspaceName = $w.name
            datasetId   = $d.id
            datasetName = $d.name
            upstreamType= "Dataflow"
            upstreamId  = $u.id
            upstreamName = $u.name
          }
        }
      }
      if ($null -ne $d.upstreamDatasets) {
        foreach ($u in ($d.upstreamDatasets | Where-Object { $_ })) {
          [pscustomobject]@{
            workspaceId = $w.id
            workspaceName = $w.name
            datasetId   = $d.id
            datasetName = $d.name
            upstreamType= "Dataset"
            upstreamId  = $u.id
            upstreamName = $u.name
          }
        }
      }
    }
  }
  if ($edges) { 
    Write-Progress-Info "Exported $($edges.Count) lineage edges"
    $edges | Export-Csv (Join-Path $OutDir "lineage_edges_$Stamp.csv") -NoTypeInformation -Encoding UTF8 
  }
  
  # Dataset Refresh History (if enabled and available)
  if ($IncludeRefreshHistory) {
    Write-Progress-Info "Exporting detailed dataset refresh history..."
    $detailedRefreshHistory = foreach ($w in $ScanObject.workspaces) {
      foreach ($d in ($w.datasets | Where-Object { $_ })) {
        try {
          # Get more detailed refresh history (last 5 refreshes) - use direct API call
          $url = "/groups/$($w.id)/datasets/$($d.id)/refreshes?`$top=5"
          $fullHistory = $null
          $fullHistory = Invoke-PowerBIRestMethod -Url $url -Method Get -ErrorAction Stop | ConvertFrom-Json
          
          if ($fullHistory.value -and $fullHistory.value.Count -gt 0) {
            foreach ($refresh in $fullHistory.value) {
              [pscustomobject]@{
                workspaceId = $w.id
                workspaceName = $w.name
                datasetId = $d.id
                datasetName = $d.name
                refreshRequestId = $refresh.requestId
                refreshType = $refresh.refreshType
                refreshStatus = $refresh.status
                startTime = $refresh.startTime
                endTime = $refresh.endTime
                durationMinutes = if ($refresh.startTime -and $refresh.endTime) {
                  try {
                    $start = [DateTime]::Parse($refresh.startTime)
                    $end = [DateTime]::Parse($refresh.endTime)
                    [math]::Round(($end - $start).TotalMinutes, 2)
                  } catch { $null }
                } else { $null }
                serviceExceptionJson = $refresh.serviceExceptionJson
                refreshAttempts = if ($refresh.refreshAttempts) { $refresh.refreshAttempts.Count } else { 0 }
              }
            }
          }
        } catch {
          # Skip datasets that don't support refresh API (lakehouses, warehouses, etc.)
          if ($_.Exception.Message -like "*404*" -or $_.Exception.Message -like "*Not Found*") {
            Write-Progress-Info "Skipping refresh history for $($d.name) - dataset type does not support refresh API" "Cyan"
          } else {
            Write-Progress-Info "Could not get detailed refresh history for dataset $($d.name): $($_.Exception.Message)" "Yellow"
          }
        }
      }
    }
    
    if ($detailedRefreshHistory) {
      Write-Progress-Info "Exported $($detailedRefreshHistory.Count) refresh history records"
      $detailedRefreshHistory | Export-Csv (Join-Path $OutDir "dataset_refresh_history_$Stamp.csv") -NoTypeInformation -Encoding UTF8
    } else {
      Write-Progress-Info "No detailed refresh history available for export" "Yellow"
    }
  }
  
  # Export scan statistics
  Write-Progress-Info "Exporting scan statistics..."
  $statsPath = Join-Path $OutDir "scan_statistics_$Stamp.json"
  $script:ScanStats | ConvertTo-Json -Depth 5 | Out-File -FilePath $statsPath -Encoding utf8
  
  Write-Progress-Info "CSV export completed successfully" "Green"
}

# ------------------------
# -------- [Main] --------
# ------------------------
Write-Host "== Power BI Scanner – Interactive ==" -ForegroundColor Cyan
Write-Progress-Info "Starting Power BI metadata scanner with improved column/datatype handling"

New-Folder -Path $ExportRoot
$stamp = (Get-Date -Format "yyyyMMdd_HHmmss")
$runDir = Join-Path $ExportRoot $stamp
New-Folder -Path $runDir

# 1) Sign in interactively
Import-Module MicrosoftPowerBIMgmt.Profile -ErrorAction Stop
Import-Module MicrosoftPowerBIMgmt.Admin   -ErrorAction Stop
Write-Progress-Info "Sign in to Power BI..."
Connect-PowerBIServiceAccount | Out-Null   # user interactive sign-in

# 2) Verify admin API access
try {
  Invoke-PBIAdminGet -Url "/admin/groups?`$top=1"
  Write-Progress-Info "Admin API access verified" "Green"
} catch {
  Write-Error "You don't have access to Admin APIs. Ensure you're a Power BI/Fabric admin and that tenant settings allow your user. See docs for GetGroupsAsAdmin and Admin Portal."
  break
}

# 3) Get all workspaces
Write-Progress-Info "Retrieving workspaces..."
$allWs = Get-AllWorkspacesAsAdmin
$activeCount = ($allWs | Where-Object { $_.state -eq 'Active' -and $_.type -eq 'Workspace' }).Count
Write-Progress-Info "Found $activeCount active workspaces" "Green"

# 4) Decide which to scan
$idsToScan = Select-WorkspaceIds -AllWorkspaces $allWs -Select:$SelectWorkspaces -Like $WorkspaceNameLike
if (-not $idsToScan -or $idsToScan.Count -eq 0) { throw "No workspaces selected." }

Write-Progress-Info "Selected $($idsToScan.Count) workspaces to scan" "Green"

# 5) Batch into max 100 IDs per scan (API limit)
$batchSize = 100
$scanJsonPayloads = @()

# Ensure idsToScan is always treated as an array
$idsArray = @($idsToScan)

for ($i=0; $i -lt $idsArray.Count; $i+=$batchSize) {
  $j = [Math]::Min($i + $batchSize - 1, $idsArray.Count - 1)
  $batch = $idsArray[$i..$j]
  
  Write-Progress-Info "Starting scan for batch $([math]::floor($i/$batchSize)+1) of $([math]::ceiling($idsArray.Count/$batchSize)): workspaces $($i+1)-$($j+1) of $($idsArray.Count)"

  try {
    $scanId = Start-WorkspaceScan -WorkspaceIds $batch `
                                 -Lineage:$IncludeLineage `
                                 -DatasourceDetails:$IncludeDatasourceDetails `
                                 -DatasetSchema:$IncludeDatasetSchema `
                                 -DatasetExpressions:$IncludeDatasetExpressions
    
    if ([string]::IsNullOrWhiteSpace($scanId)) {
      Write-Warning "Scan ID is empty, skipping this batch"
      $script:ScanStats.FailedScans++
      continue
    }
    
    Wait-ScanSucceeded -ScanId $scanId
    $raw = Get-ScanResult -ScanId $scanId

    # Validate scan result quality
    $tempObj = $raw | ConvertFrom-Json
    Test-ScanResultQuality -ScanObject $tempObj -ExpectedWorkspaceIds $batch

    # Save per-batch raw for traceability
    $batchJsonPath = Join-Path $runDir ("scanner_batch_{0:D4}.json" -f (($i/$batchSize)+1))
    $raw | Out-File -FilePath $batchJsonPath -Encoding utf8
    $scanJsonPayloads += $raw
    
    $script:ScanStats.SuccessfulScans++
    Write-Progress-Info "Batch scan completed successfully" "Green"
  } catch {
    $script:ScanStats.FailedScans++
    $script:ScanStats.ProcessingErrors += "Batch $([math]::floor($i/$batchSize)+1) failed: $($_.Exception.Message)"
    Write-Warning "Batch scan failed: $($_.Exception.Message)"
    Write-Progress-Info "Continuing with next batch..."
    continue
  }
}

# 6) Merge & save final JSON
if ($scanJsonPayloads.Count -eq 0) {
  Write-Warning "No successful scans completed. Creating empty result file."
  $emptyResult = @{ workspaces = @() } | ConvertTo-Json
  $finalJsonPath = Join-Path $runDir "scanner_result_$stamp.json"
  $emptyResult | Out-File -FilePath $finalJsonPath -Encoding utf8
  Write-Progress-Info "Saved empty JSON result: $finalJsonPath"
  
  Write-Progress-Info "No data to export to CSV files"
  Write-Progress-Info "Check the errors above and ensure:" "Yellow"
  Write-Progress-Info "  - You have access to the selected workspaces" "Yellow"
  Write-Progress-Info "  - Metadata scanning is enabled in Power BI Admin Portal" "Yellow"
  Write-Progress-Info "  - The workspaces contain semantic models (datasets)" "Yellow"
} else {
  $mergedJson = Merge-ScanResults -JsonPayloads $scanJsonPayloads
  $finalJsonPath = Join-Path $runDir "scanner_result_$stamp.json"
  $mergedJson | Out-File -FilePath $finalJsonPath -Encoding utf8
  Write-Progress-Info "Saved merged JSON: $finalJsonPath" "Green"

  # 7) Flatten to CSVs
  $scanObj = $mergedJson | ConvertFrom-Json
  Export-FlatCsvs -ScanObject $scanObj -OutDir $runDir -Stamp $stamp
}

# 8) Display final statistics
Write-Progress-Info "=== SCAN STATISTICS ===" "Cyan"
Write-Progress-Info "Total workspaces processed: $($script:ScanStats.TotalWorkspaces)" "Green"
Write-Progress-Info "Successful scan batches: $($script:ScanStats.SuccessfulScans)" "Green"
Write-Progress-Info "Failed scan batches: $($script:ScanStats.FailedScans)" "Red"
Write-Progress-Info "Total datasets found: $($script:ScanStats.TotalDatasets)" "Green"
Write-Progress-Info "Datasets with schema: $($script:ScanStats.DatasetsWithSchema)" "Green"
Write-Progress-Info "Datasets without schema: $($script:ScanStats.DatasetsWithoutSchema)" "Yellow"
if ($IncludeRefreshHistory) {
  Write-Progress-Info "Datasets with refresh history: $($script:ScanStats.DatasetsWithRefreshHistory)" "Green"
}
Write-Progress-Info "Total tables found: $($script:ScanStats.TotalTables)" "Green"
Write-Progress-Info "Total columns found: $($script:ScanStats.TotalColumns)" "Green"

if ($script:ScanStats.ProcessingErrors.Count -gt 0) {
  Write-Progress-Info "Processing errors encountered:" "Yellow"
  $script:ScanStats.ProcessingErrors | ForEach-Object { Write-Progress-Info "  - $_" "Yellow" }
}

Write-Progress-Info "=== COMPLETION ===" "Cyan"
Write-Progress-Info "Outputs saved in: $runDir" "Green"
Write-Progress-Info "Scanner completed successfully!" "Green"