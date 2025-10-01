<#
.SYNOPSIS
    Fetch Tabular Model Objects via Power BI XMLA Endpoint

.DESCRIPTION
    This script uses interactive authentication and XMLA connectivity to Power BI
    to retrieve all objects (tables, columns, measures, etc.) from a tabular semantic model.
    
    The script connects to Power BI Premium, Premium Per User, or Fabric capacity workspaces
    and uses Analysis Services cmdlets to query the tabular model metadata through XMLA endpoints.

.PARAMETER WorkspaceName
    The name of the Power BI workspace containing the dataset. The workspace must be assigned 
    to a Premium capacity, Premium Per User, or Fabric capacity with XMLA endpoint enabled.

.PARAMETER DatasetName
    The name of the Power BI dataset (semantic model) to analyze. You must have Build permissions
    on this dataset to access it via XMLA.

.PARAMETER ExportFormat
    Export format for the retrieved metadata. Valid options: 'JSON', 'CSV', or 'None'.
    - JSON: Export all metadata as a single JSON file
    - CSV: Export each object type to separate CSV files
    - None: Display results only (no export)

.PARAMETER ExportPath
    Custom directory path where the export files will be saved. If not specified, 
    the script directory will be used. The directory will be created if it doesn't exist.

.PARAMETER Locale
    Locale ID for the XMLA connection. Default is 1033 (English - United States).
    Common values: 1033 (en-US), 1031 (de-DE), 1036 (fr-FR), 1040 (it-IT), 1034 (es-ES).
    This can help resolve locale-related warnings in XMLA connections.

.EXAMPLE
    .\Get-PowerBITabularObjects.ps1 -WorkspaceName "Sales Analytics" -DatasetName "Sales Model"
    
    Basic usage to retrieve and display all tabular model objects.

.EXAMPLE
    .\Get-PowerBITabularObjects.ps1 -WorkspaceName "Sales Analytics" -DatasetName "Sales Model" -ExportFormat JSON
    
    Retrieve objects and export to JSON file in the script directory.

.EXAMPLE
    .\Get-PowerBITabularObjects.ps1 -WorkspaceName "Sales Analytics" -DatasetName "Sales Model" -ExportFormat CSV -ExportPath "C:\Exports"
    
    Retrieve objects and export each object type to separate CSV files in a custom directory.

.EXAMPLE
    .\Get-PowerBITabularObjects.ps1 -WorkspaceName "Sales Analytics" -DatasetName "Sales Model" -ExportFormat JSON -ExportPath "C:\Exports"
    
    Retrieve objects and export to JSON file in a custom directory.

.EXAMPLE
    .\Get-PowerBITabularObjects.ps1 -WorkspaceName "Sales Analytics" -DatasetName "Sales Model" -Locale 1031
    
    Retrieve objects using German locale (1031) to avoid locale warnings.

.NOTES
    Requirements:
    - PowerShell 5.1 or later
    - SqlServer PowerShell module
    - MicrosoftPowerBIMgmt PowerShell module
    - Power BI Pro or Premium Per User license
    - Premium capacity, Premium Per User, or Fabric capacity for the workspace
    - XMLA endpoint enabled for read operations
    - Build permission on the target dataset

.LINK
    https://docs.microsoft.com/en-us/power-bi/admin/service-premium-connect-tools
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = "Name of the Power BI workspace")]
    [string]$WorkspaceName,
    
    [Parameter(Mandatory = $true, HelpMessage = "Name of the Power BI dataset/semantic model")]
    [string]$DatasetName,
    
    [Parameter(Mandatory = $false, HelpMessage = "Export format: JSON, CSV, or None (default: None)")]
    [ValidateSet("JSON", "CSV", "None")]
    [string]$ExportFormat = "None",
    
    [Parameter(Mandatory = $false, HelpMessage = "Custom directory path for export files")]
    [string]$ExportPath = $PSScriptRoot,
    
    [Parameter(Mandatory = $false, HelpMessage = "Locale ID for XMLA connection (default: 1033 = English US)")]
    [int]$Locale = 1033
)

# Import required modules
Write-Host "Checking and importing required modules..." -ForegroundColor Yellow

# Check and install SqlServer module (required for XMLA connectivity)
if (-not (Get-Module -ListAvailable -Name SqlServer)) {
    Write-Host "Installing SqlServer module..." -ForegroundColor Green
    Install-Module -Name SqlServer -Force -AllowClobber -Scope CurrentUser
}

# Check and install MicrosoftPowerBIMgmt module (for interactive authentication)
if (-not (Get-Module -ListAvailable -Name MicrosoftPowerBIMgmt)) {
    Write-Host "Installing MicrosoftPowerBIMgmt module..." -ForegroundColor Green
    Install-Module -Name MicrosoftPowerBIMgmt -Force -Scope CurrentUser
}

# Import modules
Import-Module SqlServer
Import-Module MicrosoftPowerBIMgmt

Write-Host "Modules imported successfully." -ForegroundColor Green

# Function to convert Analysis Services XML result to PowerShell objects
function ConvertFrom-ASResult {
    param(
        [string]$RawResult,
        [string]$QueryType
    )
    
    try {
        # Parse the XML result
        [xml]$xmlResult = $RawResult
        
        # Check for exceptions or errors first
        $exceptions = $xmlResult.return.root.Exception
        $messages = $xmlResult.return.root.Messages
        
        if ($exceptions -or $messages) {
            if ($messages -and $messages.Error) {
                Write-Warning "XMLA Error for $QueryType`: $($messages.Error.Description)"
            }
            if ($messages -and $messages.Warning) {
                Write-Host "    XMLA Warning: $($messages.Warning.Description)" -ForegroundColor Yellow
            }
        }
        
        # Navigate to the data rows in the XML structure
        $dataRows = $xmlResult.return.root.row
        
        if (-not $dataRows -or $dataRows.Count -eq 0) {
            Write-Host "    No data rows found in result for $QueryType" -ForegroundColor Gray
            return @()
        }
        
        # Extract column mapping from schema if available
        $columnMapping = @{}
        if ($xmlResult.return.root.schema) {
            $schemaElements = $xmlResult.return.root.schema.complexType | Where-Object { $_.name -eq "row" }
            if ($schemaElements -and $schemaElements.sequence -and $schemaElements.sequence.element) {
                foreach ($element in $schemaElements.sequence.element) {
                    if ($element.name -and $element.field) {
                        $columnMapping[$element.name] = $element.field
                    }
                }
            }
        }
        
        # Convert XML rows to PowerShell objects
        $objects = @()
        foreach ($row in $dataRows) {
            $obj = New-Object PSObject
            
            # Add properties based on the row's child elements
            foreach ($property in $row.ChildNodes) {
                if ($property.NodeType -eq [System.Xml.XmlNodeType]::Element) {
                    # Use mapped field name if available, otherwise use the element name
                    $propertyName = if ($columnMapping.ContainsKey($property.LocalName)) {
                        $columnMapping[$property.LocalName]
                    } else {
                        $property.LocalName
                    }
                    
                    $obj | Add-Member -MemberType NoteProperty -Name $propertyName -Value $property.InnerText
                }
            }
            
            $objects += $obj
        }
        
        return $objects
    }
    catch {
        Write-Warning "Failed to parse XML result for $QueryType`: $($_.Exception.Message)"
        Write-Host "Raw result preview: $($RawResult.Substring(0, [Math]::Min(200, $RawResult.Length)))" -ForegroundColor Gray
        return @()
    }
}

# Function to test XMLA endpoint connectivity using Analysis Services cmdlets
function Test-XmlaConnectivity {
    param(
        [string]$ServerEndpoint,
        [string]$DatasetName
    )
    
    Write-Host "Testing XMLA endpoint connectivity using Analysis Services..." -ForegroundColor Yellow
    
    try {
        # Try using a simple discover request for testing
        $testRequest = @{
            discover = @{
                requestType = "DISCOVER_DATASOURCES"
                restrictions = @{}
                properties = @{}
            }
        } | ConvertTo-Json -Depth 3
        
        $testResult = Invoke-ASCmd -Server $ServerEndpoint -Query $testRequest
        
        if ($testResult) {
            Write-Host "✓ XMLA endpoint connectivity test successful" -ForegroundColor Green
            return $true
        } else {
            Write-Host "✗ XMLA endpoint test failed - no response" -ForegroundColor Red
            return $false
        }
    }
    catch {
        Write-Host "✗ XMLA endpoint connectivity test failed: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Function to authenticate and get workspace information
function Get-WorkspaceInfo {
    param(
        [string]$WorkspaceName
    )
    
    Write-Host "Authenticating to Power BI Service..." -ForegroundColor Yellow
    try {
        Connect-PowerBIServiceAccount | Out-Null
        Write-Host "Successfully authenticated to Power BI Service." -ForegroundColor Green
        
        # Get workspace information
        $workspace = Get-PowerBIWorkspace -Name $WorkspaceName
        if (-not $workspace) {
            throw "Workspace '$WorkspaceName' not found or you don't have access to it."
        }
        
        # Check workspace capacity status
        Write-Host "Workspace found: $($workspace.Name)" -ForegroundColor Green
        Write-Host "  Type: $($workspace.Type)" -ForegroundColor Gray
        Write-Host "  Is on dedicated capacity: $($workspace.IsOnDedicatedCapacity)" -ForegroundColor Gray
        
        # Verify XMLA capability
        if (-not $workspace.IsOnDedicatedCapacity -and $workspace.Type -ne "PersonalGroup") {
            Write-Warning "This workspace may not support XMLA endpoint connectivity."
            Write-Warning "XMLA endpoints require Premium capacity, Premium Per User, or Fabric capacity."
        }
        
        return $workspace
    }
    catch {
        Write-Error "Failed to authenticate or access workspace: $($_.Exception.Message)"
        return $null
    }
}

# Function to build XMLA connection string
function Get-XmlaConnectionString {
    param(
        [string]$WorkspaceName,
        [string]$DatasetName,
        [int]$Locale = 1033
    )
    
    $serverEndpoint = "powerbi://api.powerbi.com/v1.0/myorg/$WorkspaceName"
    
    # For interactive authentication, we don't need to specify credentials in connection string
    # The current user context will be used
    # Include Locale to avoid locale-related warnings
    return "Data Source=$serverEndpoint;Initial Catalog=$DatasetName;Locale Identifier=$Locale;"
}

# Function to execute XMLA queries and retrieve tabular model objects using DMV queries
function Get-TabularModelObjects {
    param(
        [string]$ConnectionString
    )
    
    Write-Host "Connecting to XMLA endpoint using Analysis Services..." -ForegroundColor Yellow
    
    # Extract server endpoint
    $serverEndpoint = ($ConnectionString -split ';' | Where-Object { $_ -like 'Data Source=*' }) -replace 'Data Source=', ''
    $datasetName = ($ConnectionString -split ';' | Where-Object { $_ -like 'Initial Catalog=*' }) -replace 'Initial Catalog=', ''
    
    # Test connectivity first
    if (-not (Test-XmlaConnectivity -ServerEndpoint $serverEndpoint -DatasetName $datasetName)) {
        Write-Host "`nTroubleshooting steps:" -ForegroundColor Yellow
        Write-Host "1. Verify the workspace is assigned to a Premium capacity, Premium Per User, or Fabric capacity" -ForegroundColor White
        Write-Host "2. Check that XMLA endpoint is enabled in the capacity settings:" -ForegroundColor White
        Write-Host "   - Power BI Admin Portal > Capacity settings > [Your Capacity] > Workloads > XMLA Endpoint: Read or Read Write" -ForegroundColor White
        Write-Host "3. Ensure you have Build permission on the dataset '$datasetName'" -ForegroundColor White
        Write-Host "4. Verify the dataset has Enhanced metadata format enabled" -ForegroundColor White
        Write-Host "5. Try connecting with SQL Server Management Studio (SSMS) to the same endpoint: $serverEndpoint" -ForegroundColor White
        return $null
    }
    
    try {
        # Define DMV queries for each object type - these work reliably with Power BI XMLA
        $dmvQueries = @{
            "Tables" = "SELECT * FROM `$SYSTEM.TMSCHEMA_TABLES"
            "Columns" = "SELECT * FROM `$SYSTEM.TMSCHEMA_COLUMNS"  
            "Measures" = "SELECT * FROM `$SYSTEM.TMSCHEMA_MEASURES"
            "Hierarchies" = "SELECT * FROM `$SYSTEM.TMSCHEMA_HIERARCHIES"
            "Partitions" = "SELECT * FROM `$SYSTEM.TMSCHEMA_PARTITIONS"
            "Relationships" = "SELECT * FROM `$SYSTEM.TMSCHEMA_RELATIONSHIPS"
        }
        
        $results = @{}
        
        foreach ($queryType in $dmvQueries.Keys) {
            Write-Host "Fetching $queryType..." -ForegroundColor Cyan
            
            try {
                Write-Host "    Using DMV query..." -ForegroundColor Gray
                
                $dmvQuery = $dmvQueries[$queryType]
                Write-Host "    Query: $dmvQuery" -ForegroundColor Gray
                
                # Execute the DMV query
                $rawResult = Invoke-ASCmd -Server $serverEndpoint -Database $datasetName -Query $dmvQuery
                
                # Parse the XML result to extract data
                $parsedResult = ConvertFrom-ASResult -RawResult $rawResult -QueryType $queryType
                
                $results[$queryType] = $parsedResult
                Write-Host "  Found $($parsedResult.Count) $queryType" -ForegroundColor Green
            }
            catch {
                Write-Warning "Could not retrieve $queryType`: $($_.Exception.Message)"
                Write-Host "    Error details: $($_.Exception.InnerException.Message)" -ForegroundColor Gray
                $results[$queryType] = @()
            }
        }
        
        return $results
    }
    catch {
        Write-Error "Failed to connect to XMLA endpoint or execute queries: $($_.Exception.Message)"
        return $null
    }
}

# Function to display results in a formatted way
function Show-TabularModelResults {
    param(
        [hashtable]$Results
    )
    
    # Define the display order and specific properties to show for each object type
    $displayOrder = @("Tables", "Columns", "Measures", "Hierarchies", "Relationships", "Partitions")
    
    foreach ($objectType in $displayOrder) {
        if (-not $Results.ContainsKey($objectType)) { continue }
        
        Write-Host "`n=== $objectType ===" -ForegroundColor Magenta
        
        if ($Results[$objectType].Count -eq 0) {
            Write-Host "No $objectType found." -ForegroundColor Gray
            continue
        }
        
        # Display customized information based on object type
        $displayCount = [Math]::Min(10, $Results[$objectType].Count)
        
        switch ($objectType) {
            "Tables" {
                $Results[$objectType] | Select-Object -First $displayCount | 
                    Select-Object Name, Description, IsHidden | 
                    Format-Table -AutoSize
            }
            "Columns" {
                # Show available properties to identify the correct field names
                if ($Results[$objectType].Count -gt 0) {
                    $sampleObj = $Results[$objectType][0]
                    $availableProps = $sampleObj | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
                    
                    # Try to find the best matching properties
                    $nameField = $availableProps | Where-Object { $_ -like "*Name*" } | Select-Object -First 1
                    $tableField = $availableProps | Where-Object { $_ -like "*Table*" } | Select-Object -First 1
                    $typeField = $availableProps | Where-Object { $_ -like "*Type*" -or $_ -like "*DataType*" } | Select-Object -First 1
                    $hiddenField = $availableProps | Where-Object { $_ -like "*Hidden*" } | Select-Object -First 1
                    
                    if ($nameField) {
                        $selectProps = @($nameField)
                        if ($tableField) { $selectProps += $tableField }
                        if ($typeField) { $selectProps += $typeField }
                        if ($hiddenField) { $selectProps += $hiddenField }
                        
                        $Results[$objectType] | Select-Object -First $displayCount | 
                            Select-Object $selectProps | Format-Table -AutoSize
                    } else {
                        $Results[$objectType] | Select-Object -First $displayCount | Format-Table -AutoSize
                    }
                }
            }
            "Measures" {
                if ($Results[$objectType].Count -gt 0) {
                    $sampleObj = $Results[$objectType][0]
                    $availableProps = $sampleObj | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
                    
                    $nameField = $availableProps | Where-Object { $_ -like "*Name*" } | Select-Object -First 1
                    $tableField = $availableProps | Where-Object { $_ -like "*Table*" } | Select-Object -First 1
                    $descField = $availableProps | Where-Object { $_ -like "*Description*" } | Select-Object -First 1
                    $hiddenField = $availableProps | Where-Object { $_ -like "*Hidden*" } | Select-Object -First 1
                    
                    if ($nameField) {
                        $selectProps = @($nameField)
                        if ($tableField) { $selectProps += $tableField }
                        if ($descField) { $selectProps += $descField }
                        if ($hiddenField) { $selectProps += $hiddenField }
                        
                        $Results[$objectType] | Select-Object -First $displayCount | 
                            Select-Object $selectProps | Format-Table -AutoSize
                    } else {
                        $Results[$objectType] | Select-Object -First $displayCount | Format-Table -AutoSize
                    }
                }
            }
            "Hierarchies" {
                if ($Results[$objectType].Count -gt 0) {
                    $sampleObj = $Results[$objectType][0]
                    $availableProps = $sampleObj | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
                    
                    $nameField = $availableProps | Where-Object { $_ -like "*Name*" } | Select-Object -First 1
                    $tableField = $availableProps | Where-Object { $_ -like "*Table*" } | Select-Object -First 1
                    $descField = $availableProps | Where-Object { $_ -like "*Description*" } | Select-Object -First 1
                    
                    if ($nameField) {
                        $selectProps = @($nameField)
                        if ($tableField) { $selectProps += $tableField }
                        if ($descField) { $selectProps += $descField }
                        
                        $Results[$objectType] | Select-Object -First $displayCount | 
                            Select-Object $selectProps | Format-Table -AutoSize
                    } else {
                        $Results[$objectType] | Select-Object -First $displayCount | Format-Table -AutoSize
                    }
                }
            }
            "Relationships" {
                if ($Results[$objectType].Count -gt 0) {
                    $sampleObj = $Results[$objectType][0]
                    $availableProps = $sampleObj | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
                    
                    $nameField = $availableProps | Where-Object { $_ -like "*Name*" } | Select-Object -First 1
                    $fromTableField = $availableProps | Where-Object { $_ -like "*FromTable*" } | Select-Object -First 1
                    $fromColumnField = $availableProps | Where-Object { $_ -like "*FromColumn*" } | Select-Object -First 1
                    $toTableField = $availableProps | Where-Object { $_ -like "*ToTable*" } | Select-Object -First 1
                    $toColumnField = $availableProps | Where-Object { $_ -like "*ToColumn*" } | Select-Object -First 1
                    
                    if ($nameField) {
                        $selectProps = @($nameField)
                        if ($fromTableField) { $selectProps += $fromTableField }
                        if ($fromColumnField) { $selectProps += $fromColumnField }
                        if ($toTableField) { $selectProps += $toTableField }
                        if ($toColumnField) { $selectProps += $toColumnField }
                        
                        $Results[$objectType] | Select-Object -First $displayCount | 
                            Select-Object $selectProps | Format-Table -AutoSize
                    } else {
                        $Results[$objectType] | Select-Object -First $displayCount | Format-Table -AutoSize
                    }
                }
            }
            "Partitions" {
                if ($Results[$objectType].Count -gt 0) {
                    $sampleObj = $Results[$objectType][0]
                    $availableProps = $sampleObj | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
                    
                    $nameField = $availableProps | Where-Object { $_ -like "*Name*" } | Select-Object -First 1
                    $tableField = $availableProps | Where-Object { $_ -like "*Table*" } | Select-Object -First 1
                    $modeField = $availableProps | Where-Object { $_ -like "*Mode*" } | Select-Object -First 1
                    $sourceField = $availableProps | Where-Object { $_ -like "*Source*" } | Select-Object -First 1
                    
                    if ($nameField) {
                        $selectProps = @($nameField)
                        if ($tableField) { $selectProps += $tableField }
                        if ($modeField) { $selectProps += $modeField }
                        if ($sourceField) { $selectProps += $sourceField }
                        
                        $Results[$objectType] | Select-Object -First $displayCount | 
                            Select-Object $selectProps | Format-Table -AutoSize
                    } else {
                        $Results[$objectType] | Select-Object -First $displayCount | Format-Table -AutoSize
                    }
                }
            }
            default {
                $Results[$objectType] | Select-Object -First $displayCount | Format-Table -AutoSize
            }
        }
        
        if ($Results[$objectType].Count -gt $displayCount) {
            Write-Host "... and $($Results[$objectType].Count - $displayCount) more $objectType" -ForegroundColor Gray
        }
    }
}

# Function to export results to JSON
function Export-ResultsToJson {
    param(
        [hashtable]$Results,
        [string]$WorkspaceName,
        [string]$DatasetName,
        [string]$ExportPath = $PSScriptRoot
    )
    
    # Ensure the export directory exists
    if (-not (Test-Path -Path $ExportPath)) {
        try {
            New-Item -Path $ExportPath -ItemType Directory -Force | Out-Null
            Write-Host "Created export directory: $ExportPath" -ForegroundColor Yellow
        }
        catch {
            Write-Warning "Could not create export directory '$ExportPath': $($_.Exception.Message)"
            Write-Warning "Using script directory instead."
            $ExportPath = $PSScriptRoot
        }
    }
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $filename = "TabularModel_${WorkspaceName}_${DatasetName}_${timestamp}.json"
    $filepath = Join-Path $ExportPath $filename
    
    $exportData = @{
        Workspace = $WorkspaceName
        Dataset = $DatasetName
        ExportDateTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        ExportPath = $filepath
        Results = $Results
    }
    
    try {
        $exportData | ConvertTo-Json -Depth 10 | Out-File -FilePath $filepath -Encoding UTF8
        Write-Host "`nResults exported to: $filepath" -ForegroundColor Green
        return $filepath
    }
    catch {
        Write-Error "Failed to export results to '$filepath': $($_.Exception.Message)"
        return $null
    }
}

# Function to export results to CSV files (one file per object type)
function Export-ResultsToCsv {
    param(
        [hashtable]$Results,
        [string]$WorkspaceName,
        [string]$DatasetName,
        [string]$ExportPath = $PSScriptRoot
    )
    
    # Ensure the export directory exists
    if (-not (Test-Path -Path $ExportPath)) {
        try {
            New-Item -Path $ExportPath -ItemType Directory -Force | Out-Null
            Write-Host "Created export directory: $ExportPath" -ForegroundColor Yellow
        }
        catch {
            Write-Warning "Could not create export directory '$ExportPath': $($_.Exception.Message)"
            Write-Warning "Using script directory instead."
            $ExportPath = $PSScriptRoot
        }
    }
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $exportedFiles = @()
    
    Write-Host "`nExporting to CSV files..." -ForegroundColor Green
    
    foreach ($objectType in $Results.Keys) {
        if ($Results[$objectType] -and $Results[$objectType].Count -gt 0) {
            $filename = "TabularModel_${WorkspaceName}_${DatasetName}_${objectType}_${timestamp}.csv"
            $filepath = Join-Path $ExportPath $filename
            
            try {
                # Export to CSV with UTF8 encoding
                $Results[$objectType] | Export-Csv -Path $filepath -NoTypeInformation -Encoding UTF8
                Write-Host "  $objectType exported to: $filepath" -ForegroundColor Green
                $exportedFiles += $filepath
            }
            catch {
                Write-Warning "Failed to export $objectType to '$filepath': $($_.Exception.Message)"
            }
        }
        else {
            Write-Host "  No data found for $objectType, skipping CSV export" -ForegroundColor Gray
        }
    }
    
    # Create a summary file with metadata
    $summaryFilename = "TabularModel_${WorkspaceName}_${DatasetName}_Summary_${timestamp}.csv"
    $summaryFilepath = Join-Path $ExportPath $summaryFilename
    
    try {
        $summaryData = @()
        foreach ($objectType in $Results.Keys) {
            $summaryData += [PSCustomObject]@{
                ObjectType = $objectType
                Count = if ($Results[$objectType]) { $Results[$objectType].Count } else { 0 }
                ExportDateTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                Workspace = $WorkspaceName
                Dataset = $DatasetName
            }
        }
        
        $summaryData | Export-Csv -Path $summaryFilepath -NoTypeInformation -Encoding UTF8
        Write-Host "  Summary exported to: $summaryFilepath" -ForegroundColor Green
        $exportedFiles += $summaryFilepath
    }
    catch {
        Write-Warning "Failed to create summary file '$summaryFilepath': $($_.Exception.Message)"
    }
    
    return $exportedFiles
}

##################
##################
#
# Main execution
#
##################
##################
try {
    Write-Host "=== Power BI Tabular Model Object Retrieval ===" -ForegroundColor Cyan
    Write-Host "Workspace: $WorkspaceName" -ForegroundColor White
    Write-Host "Dataset: $DatasetName" -ForegroundColor White
    Write-Host "Authentication Mode: Interactive" -ForegroundColor White
    Write-Host ""
    
    # Step 1: Authenticate and validate workspace access
    $workspaceInfo = Get-WorkspaceInfo -WorkspaceName $WorkspaceName
    if (-not $workspaceInfo) {
        throw "Failed to access workspace information."
    }
    
    # Step 1.5: Verify dataset exists and is accessible
    Write-Host "Checking dataset availability..." -ForegroundColor Yellow
    try {
        $datasets = Get-PowerBIDataset -WorkspaceId $workspaceInfo.Id
        $targetDataset = $datasets | Where-Object { $_.Name -eq $DatasetName }
        
        if (-not $targetDataset) {
            Write-Host "Available datasets in workspace:" -ForegroundColor Yellow
            foreach ($ds in $datasets | Select-Object -First 10) {
                Write-Host "  - $($ds.Name)" -ForegroundColor Gray
            }
            throw "Dataset '$DatasetName' not found in workspace '$WorkspaceName'."
        }
        
        Write-Host "✓ Dataset '$DatasetName' found" -ForegroundColor Green
        Write-Host "  Dataset ID: $($targetDataset.Id)" -ForegroundColor Gray
        Write-Host "  Is refreshable: $($targetDataset.IsRefreshable)" -ForegroundColor Gray
    }
    catch {
        throw "Failed to verify dataset: $($_.Exception.Message)"
    }
    
    # Step 2: Build connection string
    $connectionString = Get-XmlaConnectionString -WorkspaceName $WorkspaceName -DatasetName $DatasetName -Locale $Locale
    
    Write-Host "Connection string prepared (credentials hidden for security)." -ForegroundColor Green
    Write-Host "Using Locale: $Locale" -ForegroundColor Gray
    
    # Step 3: Retrieve tabular model objects
    $modelObjects = Get-TabularModelObjects -ConnectionString $connectionString
    
    if ($modelObjects) {
        # Step 4: Display results
        Show-TabularModelResults -Results $modelObjects
        
        # Step 5: Export results if requested
        if ($ExportFormat -eq "JSON") {
            $exportFilePath = Export-ResultsToJson -Results $modelObjects -WorkspaceName $WorkspaceName -DatasetName $DatasetName -ExportPath $ExportPath
        }
        elseif ($ExportFormat -eq "CSV") {
            $exportedFiles = Export-ResultsToCsv -Results $modelObjects -WorkspaceName $WorkspaceName -DatasetName $DatasetName -ExportPath $ExportPath
            Write-Host "`nAll CSV files exported successfully!" -ForegroundColor Green
        }
        
        # Summary
        Write-Host "`n=== Summary ===" -ForegroundColor Cyan
        foreach ($objectType in $modelObjects.Keys) {
            Write-Host "$objectType`: $($modelObjects[$objectType].Count)" -ForegroundColor White
        }
        
        Write-Host "`nScript completed successfully!" -ForegroundColor Green
    }
    else {
        Write-Error "Failed to retrieve model objects."
    }
}
catch {
    Write-Error "Script execution failed: $($_.Exception.Message)"
    Write-Host "`nTroubleshooting tips:" -ForegroundColor Yellow
    Write-Host "  1. Verify your permissions on the workspace and dataset" -ForegroundColor Yellow
    Write-Host "  2. Ensure the workspace is on Premium/Fabric capacity with XMLA enabled" -ForegroundColor Yellow
    Write-Host "  3. Check that the workspace and dataset names are spelled correctly" -ForegroundColor Yellow
    exit 1
}
