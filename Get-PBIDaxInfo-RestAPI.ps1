<#
.SYNOPSIS
    Automated PowerShell script to extract Power BI semantic model metadata using DAX queries.

.DESCRIPTION
    This script provides an automated way to:
    1. Authenticate to Power BI using device code flow
    2. Find the specified workspace and semantic model by name
    3. Execute multiple INFO.VIEW.* DAX queries to extract model metadata:
       - INFO.VIEW.COLUMNS() - Column definitions and properties
       - INFO.VIEW.TABLES() - Table information and storage modes
       - INFO.VIEW.RELATIONSHIPS() - Model relationships and cardinality
       - INFO.VIEW.MEASURES() - Measure definitions and expressions
    4. Save results to separate timestamped CSV files for each query

.PARAMETER WorkspaceName
    Required. The name of the Power BI workspace containing the dataset.
    Supports exact match or partial match (if unique).

.PARAMETER DatasetName
    Required. The name of the dataset/semantic model to query.
    Supports exact match or partial match (if unique).

.EXAMPLE
    .\Get-PBIDaxInfo-RestAPI.ps1 -WorkspaceName "ws-powerbi-test" -DatasetName "rum_test"
    Extract metadata from the rum_test dataset in ws-powerbi-test workspace

.EXAMPLE
    .\Get-PBIDaxInfo-RestAPI.ps1 -WorkspaceName "Sales" -DatasetName "Model"
    Find workspace containing "Sales" and dataset containing "Model" (if unique matches)

.OUTPUTS
    Four CSV files with timestamped names:
    - INFO_VIEW_COLUMNS_<dataset>_<timestamp>.csv
    - INFO_VIEW_TABLES_<dataset>_<timestamp>.csv  
    - INFO_VIEW_RELATIONSHIPS_<dataset>_<timestamp>.csv
    - INFO_VIEW_MEASURES_<dataset>_<timestamp>.csv

.NOTES
    Requirements:
    - PowerShell 7 or later: 
    - https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows?view=powershell-7.5#install-powershell-using-winget-recommended
    - MicrosoftPowerBIMgmt module (auto-installed if missing)
    - Premium workspace or Premium Per User license for DAX query execution
    - Read permissions on the target dataset

#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$WorkspaceName,
    
    [Parameter(Mandatory = $true)]
    [string]$DatasetName
)

#region Functions

function Write-ColorOutput {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [System.ConsoleColor]$ForegroundColor = [System.ConsoleColor]::White
    )
    
    $originalColor = $Host.UI.RawUI.ForegroundColor
    $Host.UI.RawUI.ForegroundColor = $ForegroundColor
    Write-Output $Message
    $Host.UI.RawUI.ForegroundColor = $originalColor
}

function Install-RequiredModules {
    <#
    .SYNOPSIS
        Checks for and installs required PowerShell modules.
    .DESCRIPTION
        Verifies that MicrosoftPowerBIMgmt module is installed with minimum version.
        Installs or updates the module if needed.
    #>
    [CmdletBinding()]
    param()
    
    Write-ColorOutput -Message "=== Checking Required Modules ===" -ForegroundColor Cyan
    
    $requiredModules = @(
        @{ Name = "MicrosoftPowerBIMgmt"; MinVersion = "1.2.1" }
    )
    
    foreach ($module in $requiredModules) {
        $installedModule = Get-Module -Name $module.Name -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
        
        if (-not $installedModule -or $installedModule.Version -lt [Version]$module.MinVersion) {
            Write-ColorOutput -Message "Installing/Updating module: $($module.Name)" -ForegroundColor Yellow
            try {
                Install-Module -Name $module.Name -MinimumVersion $module.MinVersion -Scope CurrentUser -Force -AllowClobber
                Write-ColorOutput -Message "Successfully installed $($module.Name)" -ForegroundColor Green
            }
            catch {
                Write-ColorOutput -Message "Failed to install $($module.Name): $_" -ForegroundColor Red
                throw
            }
        }
        else {
            Write-ColorOutput -Message "Module $($module.Name) is already installed (Version: $($installedModule.Version))" -ForegroundColor Green
        }
    }
}

function Connect-PowerBIInteractive {
    <#
    .SYNOPSIS
        Authenticates to Power BI using device code flow.
    .DESCRIPTION
        Handles Power BI authentication using interactive device code flow.
        Checks for existing authentication before attempting new login.
        Opens browser for user authentication.
    #>
    [CmdletBinding()]
    param()
    
    Write-ColorOutput -Message "`n=== Power BI Authentication ===" -ForegroundColor Cyan
    
    try {
        # Check if already connected
        $context = Get-PowerBIAccessToken -AsString -ErrorAction SilentlyContinue
        if ($context) {
            Write-ColorOutput -Message "Already authenticated to Power BI" -ForegroundColor Green
            return $true
        }
    }
    catch {
        # Not authenticated, proceed with login
    }
    
    try {
        Write-ColorOutput -Message "Starting Power BI authentication..." -ForegroundColor Yellow
        Write-ColorOutput -Message "Please complete the authentication in your browser." -ForegroundColor Yellow
        
        # Use device code flow for interactive authentication
        Connect-PowerBIServiceAccount -Environment Public
        
        # Verify connection
        $context = Get-PowerBIAccessToken -AsString
        if ($context) {
            Write-ColorOutput -Message "Successfully authenticated to Power BI!" -ForegroundColor Green
            return $true
        }
        else {
            throw "Authentication verification failed"
        }
    }
    catch {
        Write-ColorOutput -Message "Power BI authentication failed: $_" -ForegroundColor Red
        return $false
    }
}

function Select-PowerBIWorkspace {
    <#
    .SYNOPSIS
        Finds a Power BI workspace by name.
    .DESCRIPTION
        Searches for a workspace using exact match first, then partial match.
        Returns the workspace object if found uniquely, otherwise shows available options.
    .PARAMETER FilterName
        The workspace name to search for (exact or partial match).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$FilterName
    )
    
    Write-ColorOutput -Message "`n=== Finding Workspace ===" -ForegroundColor Cyan
    
    try {
        # Get workspaces
        Write-ColorOutput -Message "Retrieving Power BI workspaces..." -ForegroundColor Yellow
        $workspaces = Get-PowerBIWorkspace -All
        
        if (-not $workspaces) {
            Write-ColorOutput -Message "No workspaces found or insufficient permissions" -ForegroundColor Red
            return $null
        }
        
        # Filter by name if provided
        if ($FilterName) {
            # Try exact match first
            $exactMatch = $workspaces | Where-Object { $_.Name -eq $FilterName }
            if ($exactMatch) {
                Write-ColorOutput -Message "Found workspace (exact match): $($exactMatch.Name)" -ForegroundColor Green
                return $exactMatch
            }
            
            # Try partial match
            $partialMatch = $workspaces | Where-Object { $_.Name -like "*$FilterName*" }
            if ($partialMatch) {
                if ($partialMatch.Count -eq 1) {
                    Write-ColorOutput -Message "Found workspace (partial match): $($partialMatch.Name)" -ForegroundColor Green
                    return $partialMatch
                }
                else {
                    Write-ColorOutput -Message "Multiple workspaces found matching '$FilterName':" -ForegroundColor Yellow
                    $partialMatch | ForEach-Object { Write-ColorOutput -Message "  - $($_.Name)" -ForegroundColor White }
                    Write-ColorOutput -Message "Please use a more specific name" -ForegroundColor Red
                    return $null
                }
            }
            else {
                Write-ColorOutput -Message "No workspaces found matching '$FilterName'" -ForegroundColor Red
                Write-ColorOutput -Message "Available workspaces:" -ForegroundColor Yellow
                $workspaces | Select-Object -First 10 | ForEach-Object { Write-ColorOutput -Message "  - $($_.Name)" -ForegroundColor White }
                if ($workspaces.Count -gt 10) {
                    Write-ColorOutput -Message "  ... and $($workspaces.Count - 10) more" -ForegroundColor Gray
                }
                return $null
            }
        }
        else {
            # No filter provided, show available workspaces for reference
            Write-ColorOutput -Message "No workspace name provided. Available workspaces:" -ForegroundColor Yellow
            $workspaces | Select-Object -First 10 | ForEach-Object { Write-ColorOutput -Message "  - $($_.Name)" -ForegroundColor White }
            if ($workspaces.Count -gt 10) {
                Write-ColorOutput -Message "  ... and $($workspaces.Count - 10) more" -ForegroundColor Gray
            }
            return $null
        }
    }
    catch {
        Write-ColorOutput -Message "Error retrieving workspaces: $_" -ForegroundColor Red
        return $null
    }
}

function Select-PowerBIDataset {
    <#
    .SYNOPSIS
        Finds a Power BI dataset/semantic model by name within a workspace.
    .DESCRIPTION
        Searches for a dataset using exact match first, then partial match.
        Returns the dataset object if found uniquely, otherwise shows available options.
    .PARAMETER WorkspaceId
        The ID of the workspace to search within.
    .PARAMETER FilterName
        The dataset name to search for (exact or partial match).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$WorkspaceId,
        
        [Parameter(Mandatory = $false)]
        [string]$FilterName
    )
    
    Write-ColorOutput -Message "`n=== Finding Dataset/Semantic Model ===" -ForegroundColor Cyan
    
    try {
        # Get datasets in the workspace
        Write-ColorOutput -Message "Retrieving datasets from workspace..." -ForegroundColor Yellow
        $datasets = Get-PowerBIDataset -WorkspaceId $WorkspaceId
        
        if (-not $datasets) {
            Write-ColorOutput -Message "No datasets found in this workspace" -ForegroundColor Red
            return $null
        }
        
        # Filter by name if provided
        if ($FilterName) {
            # Try exact match first
            $exactMatch = $datasets | Where-Object { $_.Name -eq $FilterName }
            if ($exactMatch) {
                Write-ColorOutput -Message "Found dataset (exact match): $($exactMatch.Name)" -ForegroundColor Green
                return $exactMatch
            }
            
            # Try partial match
            $partialMatch = $datasets | Where-Object { $_.Name -like "*$FilterName*" }
            if ($partialMatch) {
                if ($partialMatch.Count -eq 1) {
                    Write-ColorOutput -Message "Found dataset (partial match): $($partialMatch.Name)" -ForegroundColor Green
                    return $partialMatch
                }
                else {
                    Write-ColorOutput -Message "Multiple datasets found matching '$FilterName':" -ForegroundColor Yellow
                    $partialMatch | ForEach-Object { Write-ColorOutput -Message "  - $($_.Name)" -ForegroundColor White }
                    Write-ColorOutput -Message "Please use a more specific name" -ForegroundColor Red
                    return $null
                }
            }
            else {
                Write-ColorOutput -Message "No datasets found matching '$FilterName'" -ForegroundColor Red
                Write-ColorOutput -Message "Available datasets:" -ForegroundColor Yellow
                $datasets | ForEach-Object { Write-ColorOutput -Message "  - $($_.Name)" -ForegroundColor White }
                return $null
            }
        }
        else {
            # No filter provided, show available datasets for reference
            Write-ColorOutput -Message "No dataset name provided. Available datasets:" -ForegroundColor Yellow
            $datasets | ForEach-Object { Write-ColorOutput -Message "  - $($_.Name)" -ForegroundColor White }
            return $null
        }
    }
    catch {
        Write-ColorOutput -Message "Error retrieving datasets: $_" -ForegroundColor Red
        return $null
    }
}

function Invoke-DAXQuery {
    <#
    .SYNOPSIS
        Executes a single DAX query against a Power BI semantic model.
    .DESCRIPTION
        Uses the Power BI REST API executeQueries endpoint to run DAX queries.
        Displays results in console and automatically saves to CSV file.
    .PARAMETER WorkspaceId
        The ID of the Power BI workspace.
    .PARAMETER DatasetId
        The ID of the dataset/semantic model.
    .PARAMETER DatasetName
        The name of the dataset (used for file naming).
    .PARAMETER DAXQuery
        The DAX query to execute.
    .PARAMETER QueryName
        A name for the query (used in CSV filename).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$WorkspaceId,
        
        [Parameter(Mandatory = $true)]
        [string]$DatasetId,
        
        [Parameter(Mandatory = $true)]
        [string]$DatasetName,
        
        [Parameter(Mandatory = $false)]
        [string]$DAXQuery = "EVALUATE INFO.VIEW.COLUMNS()",
        
        [Parameter(Mandatory = $false)]
        [string]$QueryName = "INFO_VIEW_COLUMNS"
    )
    
    Write-ColorOutput -Message "`n=== Executing DAX Query ===" -ForegroundColor Cyan
    
    try {
        Write-ColorOutput -Message "Using Power BI REST API for DAX query execution..." -ForegroundColor Yellow
        Write-ColorOutput -Message "Workspace: $WorkspaceId" -ForegroundColor Gray
        Write-ColorOutput -Message "Dataset: $DatasetName ($DatasetId)" -ForegroundColor Gray
        
        # Display the DAX query
        Write-ColorOutput -Message "`nExecuting DAX Query:" -ForegroundColor White
        Write-ColorOutput -Message $DAXQuery -ForegroundColor Cyan
        
        # Prepare the REST API call body
        $requestBody = @{
            queries = @(
                @{
                    query = $DAXQuery
                }
            )
            serializerSettings = @{
                includeNulls = $true
            }
        } | ConvertTo-Json -Depth 3
        
        Write-ColorOutput -Message "`nExecuting query via Power BI REST API..." -ForegroundColor Yellow
        
        # Execute the query using Invoke-PowerBIRestMethod
        $response = Invoke-PowerBIRestMethod -Url "groups/$WorkspaceId/datasets/$DatasetId/executeQueries" -Method Post -Body $requestBody
        
        if ($response) {
            $result = $response | ConvertFrom-Json
            
            Write-ColorOutput -Message "`n=== Query Results ===" -ForegroundColor Green
            
            if ($result.results -and $result.results[0].tables -and $result.results[0].tables[0]) {
                $table = $result.results[0].tables[0]
                $rowCount = $table.rows.Count
                
                Write-ColorOutput -Message "Rows returned: $rowCount" -ForegroundColor Green
                
                if ($rowCount -gt 0) {
                    # Get column names from first row
                    $columnNames = $table.rows[0].PSObject.Properties.Name
                    
                    Write-ColorOutput -Message "`nColumns: $($columnNames -join ', ')" -ForegroundColor White
                    Write-ColorOutput -Message ("-" * 120) -ForegroundColor Gray
                    
                    # Display preview (limit to first 10 rows for readability)
                    $displayRows = [Math]::Min($rowCount, 10)
                    
                    for ($i = 0; $i -lt $displayRows; $i++) {
                        $row = $table.rows[$i]
                        $rowData = @()
                        foreach ($columnName in $columnNames) {
                            $value = if ($null -eq $row.$columnName) { "NULL" } else { $row.$columnName.ToString() }
                            # Truncate long values for display
                            if ($value.Length -gt 30) {
                                $value = $value.Substring(0, 27) + "..."
                            }
                            $rowData += $value
                        }
                        Write-Output ($rowData -join " | ")
                    }
                    
                    if ($rowCount -gt 10) {
                        Write-ColorOutput -Message "`n... (showing first 10 rows of $rowCount total)" -ForegroundColor Yellow
                    }
                    
                    # Automatically save results to CSV
                    Save-ResultsToCSV -ResultTable $table -DatasetName $DatasetName -QueryName $QueryName
                }
                else {
                    Write-ColorOutput -Message "No data returned from query" -ForegroundColor Yellow
                }
            }
            else {
                Write-ColorOutput -Message "No data returned or unexpected response format" -ForegroundColor Yellow
                Write-ColorOutput -Message "Response: $response" -ForegroundColor Gray
            }
        }
        else {
            Write-ColorOutput -Message "No response received from Power BI API" -ForegroundColor Red
        }
    }
    catch {
        Write-ColorOutput -Message "Error executing DAX query: $_" -ForegroundColor Red
        
        # Provide troubleshooting information
        Write-ColorOutput -Message "`nTroubleshooting Tips:" -ForegroundColor Yellow
        Write-ColorOutput -Message "1. Ensure the dataset is in a Premium workspace or you have Premium Per User" -ForegroundColor White
        Write-ColorOutput -Message "2. Check that you have Read permissions on the dataset" -ForegroundColor White
        Write-ColorOutput -Message "3. Ensure the DAX query syntax is correct" -ForegroundColor White
        Write-ColorOutput -Message "4. Try a simpler query first, like: EVALUATE INFO.VIEW.TABLES()" -ForegroundColor White
        Write-ColorOutput -Message "5. Verify your Power BI authentication is still valid" -ForegroundColor White
    }
}

function Save-ResultsToCSV {
    <#
    .SYNOPSIS
        Saves DAX query results to a CSV file.
    .DESCRIPTION
        Converts Power BI REST API response to CSV format and saves with timestamped filename.
        Handles proper CSV escaping and UTF-8 encoding.
    .PARAMETER ResultTable
        The result table object from Power BI REST API response.
    .PARAMETER DatasetName
        The dataset name (used in filename).
    .PARAMETER QueryName
        The query name (used as filename prefix).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [PSObject]$ResultTable,
        
        [Parameter(Mandatory = $true)]
        [string]$DatasetName,
        
        [Parameter(Mandatory = $false)]
        [string]$QueryName = "DAX_Results"
    )
    
    try {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $safeDatasetName = $DatasetName -replace '[<>:"/\\|?*]', '_'
        $fileName = "$($QueryName)_$($safeDatasetName)_$timestamp.csv"
        $filePath = Join-Path -Path $PWD -ChildPath $fileName
        
        Write-ColorOutput -Message "Saving results to: $filePath" -ForegroundColor Yellow
        
        # Create CSV content
        $csvContent = @()
        
        if ($ResultTable.rows -and $ResultTable.rows.Count -gt 0) {
            # Get column names from first row
            $columnNames = $ResultTable.rows[0].PSObject.Properties.Name
            
            # Add headers
            $headers = @()
            foreach ($columnName in $columnNames) {
                $headers += "`"$columnName`""
            }
            $csvContent += $headers -join ","
            
            # Add data rows
            foreach ($row in $ResultTable.rows) {
                $rowData = @()
                foreach ($columnName in $columnNames) {
                    $value = if ($null -eq $row.$columnName) { "" } else { $row.$columnName.ToString() }
                    # Escape quotes in CSV
                    $value = $value -replace '"', '""'
                    $rowData += "`"$value`""
                }
                $csvContent += $rowData -join ","
            }
            
            # Write to file
            $csvContent | Out-File -FilePath $filePath -Encoding UTF8
            
            Write-ColorOutput -Message "Results saved successfully!" -ForegroundColor Green
            Write-ColorOutput -Message "File: $filePath" -ForegroundColor Green
        }
        else {
            Write-ColorOutput -Message "No data to save" -ForegroundColor Yellow
        }
    }
    catch {
        Write-ColorOutput -Message "Error saving to CSV: $_" -ForegroundColor Red
    }
}

function Invoke-MultipleDAXQueries {
    <#
    .SYNOPSIS
        Executes multiple INFO.VIEW.* DAX queries to extract semantic model metadata.
    .DESCRIPTION
        Runs four predefined DAX queries to extract comprehensive model information:
        - INFO.VIEW.COLUMNS() for column metadata
        - INFO.VIEW.TABLES() for table information  
        - INFO.VIEW.RELATIONSHIPS() for relationship details
        - INFO.VIEW.MEASURES() for measure definitions
        Each query result is saved to a separate CSV file.
    .PARAMETER WorkspaceId
        The ID of the Power BI workspace.
    .PARAMETER DatasetId
        The ID of the dataset/semantic model.
    .PARAMETER DatasetName
        The name of the dataset (used for file naming).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$WorkspaceId,
        
        [Parameter(Mandatory = $true)]
        [string]$DatasetId,
        
        [Parameter(Mandatory = $true)]
        [string]$DatasetName
    )
    
    Write-ColorOutput -Message "`n=== Executing Multiple DAX Queries ===" -ForegroundColor Cyan
    
    # Define the INFO.VIEW.* DAX queries for metadata extraction
    $daxQueries = @(
        @{
            Query = "EVALUATE INFO.VIEW.COLUMNS()"
            Name = "INFO_VIEW_COLUMNS"
            Description = "Column metadata and properties"
        },
        @{
            Query = "EVALUATE INFO.VIEW.TABLES()"
            Name = "INFO_VIEW_TABLES"
            Description = "Table information and storage modes"
        },
        @{
            Query = "EVALUATE INFO.VIEW.RELATIONSHIPS()"
            Name = "INFO_VIEW_RELATIONSHIPS"
            Description = "Model relationships and cardinality"
        },
        @{
            Query = "EVALUATE INFO.VIEW.MEASURES()"
            Name = "INFO_VIEW_MEASURES"
            Description = "Measure definitions and expressions"
        }
    )
    
    $successCount = 0
    $totalQueries = $daxQueries.Count
    
    Write-ColorOutput -Message "Executing $totalQueries metadata extraction queries..." -ForegroundColor Yellow
    
    foreach ($queryInfo in $daxQueries) {
        try {
            Write-ColorOutput -Message "`n--- Processing: $($queryInfo.Description) ---" -ForegroundColor White
            
            # Execute the DAX query via Power BI REST API
            Invoke-DAXQuery -WorkspaceId $WorkspaceId -DatasetId $DatasetId -DatasetName $DatasetName -DAXQuery $queryInfo.Query -QueryName $queryInfo.Name
            
            $successCount++
            Write-ColorOutput -Message "V Completed: $($queryInfo.Name)" -ForegroundColor Green
        }
        catch {
            Write-ColorOutput -Message "X Failed: $($queryInfo.Name) - $_" -ForegroundColor Red
        }
    }
    
    Write-ColorOutput -Message "`n=== Summary ===" -ForegroundColor Cyan
    Write-ColorOutput -Message "Successfully executed: $successCount/$totalQueries queries" -ForegroundColor Green
    
    if ($successCount -eq $totalQueries) {
        Write-ColorOutput -Message "All queries completed successfully!" -ForegroundColor Green
    }
    elseif ($successCount -gt 0) {
        Write-ColorOutput -Message "Some queries completed with errors. Check output above." -ForegroundColor Yellow
    }
    else {
        Write-ColorOutput -Message "All queries failed. Check authentication and permissions." -ForegroundColor Red
    }
}

#endregion Functions

#region Main Script

# Script header
Write-ColorOutput -Message "=================================================" -ForegroundColor Cyan
Write-ColorOutput -Message "    Power BI DAX Query Tool - PowerShell Edition" -ForegroundColor Cyan
Write-ColorOutput -Message "=================================================" -ForegroundColor Cyan
Write-ColorOutput -Message "This script will connect to Power BI and execute multiple DAX queries" -ForegroundColor White
Write-ColorOutput -Message "to extract model metadata and save results to CSV files.`n" -ForegroundColor White

try {
    # Step 1: Install and import required PowerShell modules
    Install-RequiredModules
    
    # Step 2: Import Power BI management module
    Write-ColorOutput -Message "`nImporting PowerShell modules..." -ForegroundColor Yellow
    Import-Module MicrosoftPowerBIMgmt -Force
    
    # Step 3: Authenticate to Power BI service using device code flow
    $authResult = Connect-PowerBIInteractive
    if (-not $authResult) {
        Write-ColorOutput -Message "Authentication failed. Exiting script." -ForegroundColor Red
        exit 1
    }
    
    # Step 4: Find the specified workspace by name
    $selectedWorkspace = Select-PowerBIWorkspace -FilterName $WorkspaceName
    if (-not $selectedWorkspace) {
        Write-ColorOutput -Message "Workspace '$WorkspaceName' not found. Exiting script." -ForegroundColor Red
        exit 1
    }
    
    # Step 5: Find the specified dataset/semantic model by name
    $selectedDataset = Select-PowerBIDataset -WorkspaceId $selectedWorkspace.Id -FilterName $DatasetName
    if (-not $selectedDataset) {
        Write-ColorOutput -Message "Dataset '$DatasetName' not found in workspace '$WorkspaceName'. Exiting script." -ForegroundColor Red
        exit 1
    }
    
    # Step 6: Execute all INFO.VIEW.* DAX queries for metadata extraction
    Invoke-MultipleDAXQueries -WorkspaceId $selectedWorkspace.Id -DatasetId $selectedDataset.Id -DatasetName $selectedDataset.Name
    
    Write-ColorOutput -Message "`n=== Script Completed Successfully ===" -ForegroundColor Green
}
catch {
    Write-ColorOutput -Message "`nScript failed with error: $_" -ForegroundColor Red
    Write-ColorOutput -Message "Stack trace: $($_.ScriptStackTrace)" -ForegroundColor Red
}
finally {
    # Script completed - automated execution, no user interaction needed
}

#endregion Main Script