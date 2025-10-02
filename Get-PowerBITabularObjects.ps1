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

.PARAMETER TestConnectivity
    Test XMLA endpoint connectivity before retrieving objects. Default is false for faster execution.
    Use this parameter to diagnose connectivity issues with detailed troubleshooting information.

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

.EXAMPLE
    .\Get-PowerBITabularObjects.ps1 -WorkspaceName "Sales Analytics" -DatasetName "Sales Model" -TestConnectivity
    
    Test XMLA endpoint connectivity first, then retrieve objects. Useful for troubleshooting connection issues.

.NOTES
    Requirements:
    - PowerShell 7 or later
    - SqlServer PowerShell module
    - MicrosoftPowerBIMgmt PowerShell module
    - Power BI Pro or Premium Per User license
    - Premium capacity, Premium Per User, or Fabric capacity for the workspace
    - XMLA endpoint enabled for read operations
    - Build permission on the target dataset

.LINK
    https://docs.microsoft.com/en-us/power-bi/admin/service-premium-connect-tools
#>

#Requires -Version 5.1

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
    [int]$Locale = 1033,
    
    [Parameter(Mandatory = $false, HelpMessage = "Test XMLA endpoint connectivity before retrieving objects (default: false)")]
    [switch]$TestConnectivity = $false
)

# Check PowerShell version and provide guidance
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Warning "You are running PowerShell $($PSVersionTable.PSVersion)"
    Write-Warning "This script works best with PowerShell 7+ due to module compatibility."
    Write-Host ""
    Write-Host "Current environment:" -ForegroundColor Yellow
    Write-Host "  PowerShell: $($PSVersionTable.PSVersion)" -ForegroundColor Gray
    Write-Host "  .NET: $([System.Runtime.InteropServices.RuntimeInformation]::FrameworkDescription)" -ForegroundColor Gray
    Write-Host ""
    Write-Host "For best results, run this script with PowerShell 7:" -ForegroundColor Green
    Write-Host "  pwsh.exe -File `"$($MyInvocation.MyCommand.Path)`" -WorkspaceName `"$WorkspaceName`" -DatasetName `"$DatasetName`"" -ForegroundColor Green
    Write-Host ""
    $response = Read-Host "Continue with current PowerShell version? (y/n)"
    if ($response -notmatch '^y|yes$') {
        Write-Host "Script cancelled. Please run with PowerShell 7+ for best compatibility." -ForegroundColor Yellow
        exit 0
    }
}

# Import required modules
Write-Host "Checking and importing required modules..." -ForegroundColor Yellow
Write-Host "PowerShell Version: $($PSVersionTable.PSVersion)" -ForegroundColor Gray
Write-Host ".NET Version: $([System.Runtime.InteropServices.RuntimeInformation]::FrameworkDescription)" -ForegroundColor Gray

#region Module and Environment Management

# Function to check module compatibility
function Test-ModuleCompatibility {
    param([string]$ModuleName)
    
    try {
        $module = Get-Module -ListAvailable -Name $ModuleName | Sort-Object Version -Descending | Select-Object -First 1
        if ($module) {
            Write-Host "  V $ModuleName version $($module.Version) found" -ForegroundColor Green
            return $true
        } else {
            Write-Host "  X $ModuleName not found" -ForegroundColor Red
            return $false
        }
    }
    catch {
        Write-Host "  X Error checking $ModuleName`: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Check and install SqlServer module (required for XMLA connectivity)
if (-not (Test-ModuleCompatibility -ModuleName "SqlServer")) {
    Write-Host "Installing SqlServer module..." -ForegroundColor Green
    try {
        Install-Module -Name SqlServer -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
        Write-Host "  V SqlServer module installed successfully" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to install SqlServer module: $($_.Exception.Message)"
        Write-Host "Manual installation: Install-Module -Name SqlServer -Scope CurrentUser" -ForegroundColor Yellow
        exit 1
    }
}

# Check and install PowerBI module (required for authentication and workspace access)
if (-not (Test-ModuleCompatibility -ModuleName "MicrosoftPowerBIMgmt")) {
    Write-Host "Installing MicrosoftPowerBIMgmt module..." -ForegroundColor Green
    try {
        Install-Module -Name MicrosoftPowerBIMgmt -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
        Write-Host "  V MicrosoftPowerBIMgmt module installed successfully" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to install MicrosoftPowerBIMgmt module: $($_.Exception.Message)"
        Write-Host "Manual installation: Install-Module -Name MicrosoftPowerBIMgmt -Scope CurrentUser" -ForegroundColor Yellow
        exit 1
    }
}

# Import modules with better error handling
Write-Host "Importing modules..." -ForegroundColor Yellow
try {
    Import-Module SqlServer -Force -ErrorAction Stop
    Write-Host "  V SqlServer module imported" -ForegroundColor Green
    
    Import-Module MicrosoftPowerBIMgmt -Force -ErrorAction Stop
    Write-Host "  V MicrosoftPowerBIMgmt module imported" -ForegroundColor Green
    
    Write-Host "Modules imported successfully." -ForegroundColor Green
}
catch {
    Write-Error "Failed to import required modules: $($_.Exception.Message)"
    Write-Host ""
    Write-Host "Troubleshooting:" -ForegroundColor Yellow
    Write-Host "1. Try running with PowerShell 7+: pwsh.exe" -ForegroundColor White
    Write-Host "2. Or use the wrapper: .\Run-With-PowerShell7.ps1" -ForegroundColor White
    Write-Host "3. Manual module installation:" -ForegroundColor White
    Write-Host "   Install-Module -Name SqlServer -Scope CurrentUser" -ForegroundColor Gray
    Write-Host "   Install-Module -Name MicrosoftPowerBIMgmt -Scope CurrentUser" -ForegroundColor Gray
    exit 1
}

#endregion

#region XMLA and Data Handling

# Function to extract column mappings from XML schema dynamically
function Get-ColumnMappingFromSchema {
    param([System.Xml.XmlDocument]$XmlDocument)
    
    $columnMapping = @{}
    
    try {
        # Create namespace manager for XSD parsing
        $xsdNsManager = New-Object System.Xml.XmlNamespaceManager($xmlDocument.NameTable)
        $xsdNsManager.AddNamespace("xsd", "http://www.w3.org/2001/XMLSchema")
        $xsdNsManager.AddNamespace("sql", "urn:schemas-microsoft-com:xml-sql")
        
        # Navigate to the XSD schema definition within the XML
        $schemaNode = $xmlDocument.SelectSingleNode("//xsd:schema", $xsdNsManager)
        if (-not $schemaNode) {
            Write-Host "    Warning: No XSD schema found in XML response" -ForegroundColor Yellow
            return $columnMapping
        }
        
        # Find the row complex type definition
        $rowComplexType = $xmlDocument.SelectSingleNode("//xsd:complexType[@name='row']", $xsdNsManager)
        if (-not $rowComplexType) {
            Write-Host "    Warning: No 'row' complex type found in schema" -ForegroundColor Yellow
            return $columnMapping
        }
        
        # Extract all element definitions from the row type
        $elements = $rowComplexType.SelectNodes(".//xsd:element", $xsdNsManager)
        
        foreach ($element in $elements) {
            $elementName = $element.GetAttribute("name")      # This is like "C00", "C01", etc.
            $fieldName = $element.GetAttribute("sql:field")    # This is the actual column name
            $elementType = $element.GetAttribute("type")       # This is the XSD type (e.g., "xsd:long")
            
            if (-not [string]::IsNullOrEmpty($elementName) -and -not [string]::IsNullOrEmpty($fieldName)) {
                $columnMapping[$elementName] = @{
                    Name = $fieldName
                    Type = $elementType
                }
            }
        }
        
        Write-Host "    Extracted $($columnMapping.Count) column mappings from schema" -ForegroundColor Gray
        return $columnMapping
    }
    catch {
        Write-Host "    Warning: Failed to parse schema for column mapping: $($_.Exception.Message)" -ForegroundColor Yellow
        return @{} # Return an empty hashtable on failure
    }
}

# Function to safely invoke XMLA commands with error handling
function Invoke-XmlaQuery {
    param(
        [string]$ServerEndpoint,
        [string]$Query,
        [string]$QueryType = "Generic",
        [string]$DatabaseName = ""
    )
    
    Write-Host "  Using DMV query..." -ForegroundColor Gray
    Write-Host "  Query: $Query" -ForegroundColor Gray
    
    try {
        # For DMV queries, we need to specify the database/catalog in the connection
        if ([string]::IsNullOrWhiteSpace($DatabaseName)) {
            Write-Host "    Warning: No database name specified, this may cause 'CurrentCatalog not specified' errors" -ForegroundColor Yellow
            $rawResult = Invoke-ASCmd -Server $ServerEndpoint -Query $Query
        } else {
            Write-Host "    Using database: $DatabaseName" -ForegroundColor Gray
            $rawResult = Invoke-ASCmd -Server $ServerEndpoint -Database $DatabaseName -Query $Query
        }
        
        # Debug: Show raw result length (but not full content to reduce noise)
        Write-Host "    Raw result length: $($rawResult.Length) characters" -ForegroundColor Gray
        
        # Check for XMLA warnings and display them
        if ($rawResult -match '<Warning>') {
            $warningPattern = '<Warning><Description>(.*?)</Description></Warning>'
            $warningMatches = [regex]::Matches($rawResult, $warningPattern)
            foreach ($match in $warningMatches) {
                Write-Host "    XMLA Warning: $($match.Groups[1].Value)" -ForegroundColor Yellow
            }
        }
        
        # Parse XML result
        $xml = [xml]$rawResult
        $objects = @()
        
        # Create namespace manager for proper XML parsing
        $nsManager = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
        $nsManager.AddNamespace("rowset", "urn:schemas-microsoft-com:xml-analysis:rowset")
        
        # Navigate through XML to extract data rows with proper namespace
        $rowElements = $xml.SelectNodes("//rowset:row", $nsManager)
        Write-Host "    Found $($rowElements.Count) row elements in XML" -ForegroundColor Gray
        
        # Extract column mapping dynamically from XML schema
        $columnMapping = Get-ColumnMappingFromSchema -XmlDocument $xml
        
        if ($rowElements.Count -gt 0) {
            foreach ($row in $rowElements) {
                $obj = @{}
                
                # Parse child elements (C00, C01, C02, etc.) and map to descriptive names
                foreach ($childElement in $row.ChildNodes) {
                    if ($childElement.NodeType -eq [System.Xml.XmlNodeType]::Element) {
                        $value = $childElement.InnerText
                        $columnName = $childElement.Name
                        
                        $propertyName = $columnName
                        $convertedValue = $value

                        # Check if a mapping exists for the column
                        if ($columnMapping -and $columnMapping.ContainsKey($columnName)) {
                            $columnInfo = $columnMapping[$columnName]
                            $propertyName = $columnInfo.Name
                            $xsdType = $columnInfo.Type
                            
                            # Handle null/empty values properly
                            if ([string]::IsNullOrWhiteSpace($value) -or $value -eq "null") {
                                $convertedValue = $null
                            }
                            else {
                                # Convert value based on the XSD type from the schema
                                try {
                                    switch ($xsdType) {
                                        'xsd:long'      { $convertedValue = [long]$value; break }
                                        'xsd:int'       { $convertedValue = [int]$value; break }
                                        'xsd:short'     { $convertedValue = [int16]$value; break }
                                        'xsd:double'    { $convertedValue = [double]$value; break }
                                        'xsd:float'     { $convertedValue = [float]$value; break }
                                        'xsd:decimal'   { $convertedValue = [decimal]$value; break }
                                        'xsd:boolean'   { $convertedValue = [bool]::Parse($value); break }
                                        'xsd:dateTime'  { $convertedValue = [datetime]$value; break }
                                        'xsd:string'    { $convertedValue = $value; break }
                                        default         { $convertedValue = $value } # Default to string for unknown types
                                    }
                                }
                                catch {
                                    # If conversion fails, keep the original string value
                                    Write-Host "    Could not convert value `"$value`" to type `"$xsdType`". Keeping as string." -ForegroundColor DarkGray
                                    $convertedValue = $value
                                }
                            }
                        }
                        
                        $obj[$propertyName] = $convertedValue
                    }
                }
                $objects += [PSCustomObject]$obj
            }
        }
        
        return $objects
    }
    catch {
        Write-Warning "Failed to parse XML result for $QueryType`: $($_.Exception.Message)"
        Write-Host "Raw result preview: $($rawResult.Substring(0, [Math]::Min(200, $rawResult.Length)))" -ForegroundColor Gray
        return @()
    }
}

#endregion

#region Power BI and Connectivity

# Function to test XMLA endpoint connectivity using Analysis Services cmdlets
function Test-XmlaConnectivity {
    param(
        [string]$ServerEndpoint,
        [string]$DatasetName,
        [switch]$ShowDebugging = $false
    )
    
    Write-Host "Testing XMLA endpoint connectivity using Analysis Services..." -ForegroundColor Yellow
    
    if ($ShowDebugging) {
        Write-Host "`nDebugging Information:" -ForegroundColor Cyan
        Write-Host "  Server Endpoint: $ServerEndpoint" -ForegroundColor Gray
        Write-Host "  Dataset Name: $DatasetName" -ForegroundColor Gray
        Write-Host "  PowerShell Version: $($PSVersionTable.PSVersion)" -ForegroundColor Gray
        Write-Host "  .NET Version: $([System.Runtime.InteropServices.RuntimeInformation]::FrameworkDescription)" -ForegroundColor Gray
        
        # Check if SqlServer module is properly loaded
        $sqlServerModule = Get-Module SqlServer
        if ($sqlServerModule) {
            Write-Host "  SqlServer Module: $($sqlServerModule.Version) (Loaded)" -ForegroundColor Gray
        } else {
            Write-Host "  SqlServer Module: Not loaded" -ForegroundColor Red
        }
        
        Write-Host "`nAttempting connection test..." -ForegroundColor Cyan
    }
    
    try {
        # Try using a simple discover request for testing
        $testRequest = @{
            discover = @{
                requestType = "DISCOVER_DATASOURCES"
                restrictions = @{}
                properties = @{}
            }
        } | ConvertTo-Json -Depth 3
        
        if ($ShowDebugging) {
            Write-Host "  Test Request: $testRequest" -ForegroundColor Gray
        }
        
        $testResult = Invoke-ASCmd -Server $ServerEndpoint -Query $testRequest
        
        if ($testResult) {
            Write-Host "V XMLA endpoint connectivity test successful" -ForegroundColor Green
            if ($ShowDebugging) {
                Write-Host "  Response received: $($testResult.Length) characters" -ForegroundColor Gray
            }
            return $true
        }
        else {
            Write-Host "X XMLA endpoint test failed - no response" -ForegroundColor Red
            if ($ShowDebugging) {
                Write-Host "  No response received from server" -ForegroundColor Red
            }
            return $false
        }
    }
    catch {
        Write-Host "X XMLA endpoint connectivity test failed: $($_.Exception.Message)" -ForegroundColor Red
        
        if ($ShowDebugging) {
            Write-Host "`nDetailed Error Information:" -ForegroundColor Red
            Write-Host "  Exception Type: $($_.Exception.GetType().Name)" -ForegroundColor Gray
            Write-Host "  Error Message: $($_.Exception.Message)" -ForegroundColor Gray
            if ($_.Exception.InnerException) {
                Write-Host "  Inner Exception: $($_.Exception.InnerException.Message)" -ForegroundColor Gray
            }
            Write-Host "  Stack Trace: $($_.Exception.StackTrace)" -ForegroundColor Gray
        }
        
        # Show troubleshooting steps when debugging is enabled
        if ($ShowDebugging) {
            Write-Host "`nTroubleshooting Steps:" -ForegroundColor Yellow
            Write-Host "1. Verify the workspace is assigned to a Premium capacity, Premium Per User, or Fabric capacity" -ForegroundColor White
            Write-Host "2. Check that XMLA endpoint is enabled in the capacity settings:" -ForegroundColor White
            Write-Host "   - Power BI Admin Portal > Capacity settings > [Your Capacity] > Workloads > XMLA Endpoint: Read or Read Write" -ForegroundColor White
            Write-Host "3. Ensure you have Build permission on the dataset '$DatasetName'" -ForegroundColor White
            Write-Host "4. Verify the dataset has Enhanced metadata format enabled" -ForegroundColor White
            Write-Host "5. Try connecting with SQL Server Management Studio (SSMS) to the same endpoint: $ServerEndpoint" -ForegroundColor White
            Write-Host "6. Check network connectivity and firewall settings" -ForegroundColor White
            Write-Host "7. Verify your Power BI authentication is still valid" -ForegroundColor White
            Write-Host "8. Try running the script with PowerShell 7+ for better module compatibility" -ForegroundColor White
        }
        
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
    
    # Extract server endpoint and dataset name for reference
    $serverEndpoint = ($ConnectionString -split ';' | Where-Object { $_ -like 'Data Source=*' }) -replace 'Data Source=', ''
    $datasetName = ($ConnectionString -split ';' | Where-Object { $_ -like 'Initial Catalog=*' }) -replace 'Initial Catalog=', ''
    
    # Define DMV queries for different object types
    $dmvQueries = @{
        "Tables" = "SELECT * FROM `$SYSTEM.TMSCHEMA_TABLES"
        "Columns" = "SELECT * FROM `$SYSTEM.TMSCHEMA_COLUMNS"
        "Measures" = "SELECT * FROM `$SYSTEM.TMSCHEMA_MEASURES"
        "Hierarchies" = "SELECT * FROM `$SYSTEM.TMSCHEMA_HIERARCHIES"
        "Relationships" = "SELECT * FROM `$SYSTEM.TMSCHEMA_RELATIONSHIPS"
        "Partitions" = "SELECT * FROM `$SYSTEM.TMSCHEMA_PARTITIONS"
    }
    
    $results = @{}
    
    try {
        foreach ($objectType in $dmvQueries.Keys) {
            Write-Host "Fetching $objectType..." -ForegroundColor Green
            
            $queryResult = Invoke-XmlaQuery -ServerEndpoint $serverEndpoint -Query $dmvQueries[$objectType] -QueryType $objectType -DatabaseName $datasetName
            
            if ($queryResult -and $queryResult.Count -gt 0) {
                $results[$objectType] = $queryResult
                Write-Host "  Found $($queryResult.Count) $objectType" -ForegroundColor Green
            }
            else {
                Write-Host "  No $objectType found or query failed" -ForegroundColor Yellow
                $results[$objectType] = @()
            }
        }
        
        return $results
    }
    catch {
        Write-Error "Failed to retrieve tabular model objects: $($_.Exception.Message)"
        Write-Host "Connection string (sanitized): Data Source=...; Initial Catalog=$datasetName; ..." -ForegroundColor Gray
        return $null
    }
}

#endregion

#region Output and Export

# Function to display results in a readable format
function Show-TabularModelResults {
    param(
        [hashtable]$Results
    )
    
    foreach ($objectType in $Results.Keys) {
        $objects = $Results[$objectType]
        
        if ($objects -and $objects.Count -gt 0) {
            Write-Host "`n=== $objectType ===" -ForegroundColor Cyan
            
            # Display first 10 items with key properties
            $displayObjects = $objects | Select-Object -First 10
            
            switch ($objectType) {
                'Tables' {
                    $displayObjects | Select-Object Name, Description, IsHidden, ID | Format-Table -AutoSize
                    if ($objects.Count -gt 10) {
                        Write-Host "... and $($objects.Count - 10) more $objectType" -ForegroundColor Gray
                    }
                }
                'Columns' {
                    $displayObjects | Select-Object ExplicitName, TableID, ExplicitDataType, IsHidden | Format-Table -AutoSize
                    if ($objects.Count -gt 10) {
                        Write-Host "... and $($objects.Count - 10) more $objectType" -ForegroundColor Gray
                    }
                }
                'Measures' {
                    $displayObjects | Select-Object Name, TableID, IsHidden, Expression | Format-Table -AutoSize -Wrap
                    if ($objects.Count -gt 10) {
                        Write-Host "... and $($objects.Count - 10) more $objectType" -ForegroundColor Gray
                    }
                }
                'Hierarchies' {
                    $displayObjects | Select-Object Name, TableID, Description, IsHidden | Format-Table -AutoSize
                    if ($objects.Count -gt 10) {
                        Write-Host "... and $($objects.Count - 10) more $objectType" -ForegroundColor Gray
                    }
                }
                'Relationships' {
                    $displayObjects | Select-Object Name, FromTableID, FromColumnID, ToTableID, ToColumnID, IsActive | Format-Table -AutoSize
                    if ($objects.Count -gt 10) {
                        Write-Host "... and $($objects.Count - 10) more $objectType" -ForegroundColor Gray
                    }
                }
                'Partitions' {
                    $displayObjects | Select-Object Name, TableID, Mode, Type | Format-Table -AutoSize
                    if ($objects.Count -gt 10) {
                        Write-Host "... and $($objects.Count - 10) more $objectType" -ForegroundColor Gray
                    }
                }
                default {
                    $displayObjects | Format-Table -AutoSize
                    if ($objects.Count -gt 10) {
                        Write-Host "... and $($objects.Count - 10) more $objectType" -ForegroundColor Gray
                    }
                }
            }
        }
        else {
            Write-Host "`n=== $objectType ===" -ForegroundColor Cyan
            Write-Host "No $objectType found." -ForegroundColor Yellow
        }
    }
}

# Function to export results to JSON
function Export-ResultsToJson {
    param(
        [hashtable]$Results,
        [string]$WorkspaceName,
        [string]$DatasetName,
        [string]$ExportPath
    )
    
    # Ensure export directory exists
    if (-not (Test-Path $ExportPath)) {
        try {
            New-Item -ItemType Directory -Path $ExportPath -Force | Out-Null
            Write-Host "Created export directory: $ExportPath" -ForegroundColor Yellow
        }
        catch {
            Write-Error "Failed to create export directory: $($_.Exception.Message)"
            return $null
        }
    }
    
    # Create export object with metadata
    $exportObject = @{
        ExportTimestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        WorkspaceName = $WorkspaceName
        DatasetName = $DatasetName
        TotalObjects = 0
        ObjectCounts = @{}
        Objects = $Results
    }
    
    # Calculate totals
    foreach ($objectType in $Results.Keys) {
        $count = if ($Results[$objectType]) { $Results[$objectType].Count } else { 0 }
        $exportObject.ObjectCounts[$objectType] = $count
        $exportObject.TotalObjects += $count
    }
    
    # Generate filename with timestamp
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $filename = "PowerBI_TabularObjects_${WorkspaceName}_${DatasetName}_${timestamp}.json"
    $filename = $filename -replace '[<>:"/\\|?*]', '_'  # Replace invalid filename characters
    $exportFilePath = Join-Path $ExportPath $filename
    
    try {
        $exportObject | ConvertTo-Json -Depth 10 | Out-File -FilePath $exportFilePath -Encoding UTF8 -Force
        Write-Host "`nJSON export completed successfully!" -ForegroundColor Green
        Write-Host "File saved as: $exportFilePath" -ForegroundColor Green
        Write-Host "Total objects exported: $($exportObject.TotalObjects)" -ForegroundColor Green
        return $exportFilePath
    }
    catch {
        Write-Error "Failed to export to JSON: $($_.Exception.Message)"
        return $null
    }
}

# Function to export results to CSV files
function Export-ResultsToCsv {
    param(
        [hashtable]$Results,
        [string]$WorkspaceName,
        [string]$DatasetName,
        [string]$ExportPath
    )
    
    # Ensure export directory exists
    if (-not (Test-Path $ExportPath)) {
        try {
            New-Item -ItemType Directory -Path $ExportPath -Force | Out-Null
            Write-Host "Created export directory: $ExportPath" -ForegroundColor Yellow
        }
        catch {
            Write-Error "Failed to create export directory: $($_.Exception.Message)"
            return @()
        }
    }
    
    $exportedFiles = @()
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    
    foreach ($objectType in $Results.Keys) {
        $objects = $Results[$objectType]
        
        if ($objects -and $objects.Count -gt 0) {
            try {
                $filename = "PowerBI_${objectType}_${WorkspaceName}_${DatasetName}_${timestamp}.csv"
                $filename = $filename -replace '[<>:"/\\|?*]', '_'  # Replace invalid filename characters
                $csvFilePath = Join-Path $ExportPath $filename
                
                $objects | Export-Csv -Path $csvFilePath -NoTypeInformation -Encoding UTF8 -Force
                
                Write-Host "V Exported $($objects.Count) $objectType to: $csvFilePath" -ForegroundColor Green
                $exportedFiles += $csvFilePath
            }
            catch {
                Write-Warning "Failed to export $objectType to CSV: $($_.Exception.Message)"
            }
        }
        else {
            Write-Host "! Skipped $objectType (no data to export)" -ForegroundColor Yellow
        }
    }
    
    return $exportedFiles
}

#endregion

#region Utility

# Function to validate export path
function Test-ExportPath {
    param(
        [string]$Path
    )
    
    try {
        if ([string]::IsNullOrWhiteSpace($Path)) {
            return $false
        }
        
        # Check if path is valid
        $null = [System.IO.Path]::GetFullPath($Path)
        return $true
    }
    catch {
        return $false
    }
}

#endregion

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
        
        Write-Host "V Dataset '$DatasetName' found" -ForegroundColor Green
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
    
    # Step 2.5: Test connectivity if requested
    if ($TestConnectivity) {
        Write-Host ""
        $serverEndpoint = "powerbi://api.powerbi.com/v1.0/myorg/$WorkspaceName"
        $connectivityResult = Test-XmlaConnectivity -ServerEndpoint $serverEndpoint -DatasetName $DatasetName -ShowDebugging
        
        if (-not $connectivityResult) {
            Write-Host "`nXMLA connectivity test failed. Use -TestConnectivity for detailed troubleshooting information." -ForegroundColor Red
            throw "XMLA endpoint connectivity test failed. Please check the troubleshooting steps above."
        }
        Write-Host ""
    }
    
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
        
        # Return the model objects for pipeline usage
        #return $modelObjects
    }
    else {
        Write-Error "Failed to retrieve model objects."
        return $null
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
