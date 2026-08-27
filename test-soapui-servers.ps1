# Automated testing for SoapUI for .NET Web Services

param(
    [string]$ResultsDir = "C:\results",
    [switch]$CleanupHosts = $false
)

# Requires elevation for hosts file modification
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "This script requires administrator privileges. Restarting..." -ForegroundColor Red
    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# Check if Excel is installed
function Test-ExcelInstalled {
    try {
        $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        return $true
    }
    catch {
        return $false
    }
}

# Define all available services with their SoapUI project paths
$availableServices = @(
    @{
        Name = "AccountList"
        ProjectPath = "C:\SoapUIProjects\AccountList-soapui-project.xml"
        Description = "Account List Service - 5 SOAP requests"
    },
    @{
        Name = "PaymentAPI"
        ProjectPath = "C:\SoapUIProjects\PaymentAPI-soapui-project.xml"
        Description = "Payment API Service - 5 SOAP requests"
    },
    @{
        Name = "OrderProcessing"
        ProjectPath = "C:\SoapUIProjects\OrderProcessing-soapui-project.xml"
        Description = "Order Processing Service - 5 SOAP requests"
    },
    @{
        Name = "ReportingEngine"
        ProjectPath = "C:\SoapUIProjects\ReportingEngine-soapui-project.xml"
        Description = "Reporting Engine Service - 5 SOAP requests"
    }
)

# All server configurations
$allServers = @(
    # Windows Server 2012 - AccountList
    @{ Name = "WS2012-01"; ServiceName = "AccountList"; ServerNumber = 1; OSVersionCode = "0"; IP = "10.0.1.10"; OSVersion = "2012" },
    @{ Name = "WS2012-02"; ServiceName = "AccountList"; ServerNumber = 2; OSVersionCode = "0"; IP = "10.0.1.11"; OSVersion = "2012" },
    @{ Name = "WS2012-03"; ServiceName = "AccountList"; ServerNumber = 3; OSVersionCode = "0"; IP = "10.0.1.12"; OSVersion = "2012" },
    @{ Name = "WS2012-04"; ServiceName = "AccountList"; ServerNumber = 4; OSVersionCode = "0"; IP = "10.0.1.13"; OSVersion = "2012" },
    
    # Windows Server 2012 - PaymentAPI
    @{ Name = "WS2012-05"; ServiceName = "PaymentAPI"; ServerNumber = 1; OSVersionCode = "0"; IP = "10.0.1.14"; OSVersion = "2012" },
    @{ Name = "WS2012-06"; ServiceName = "PaymentAPI"; ServerNumber = 2; OSVersionCode = "0"; IP = "10.0.1.15"; OSVersion = "2012" },
    @{ Name = "WS2012-07"; ServiceName = "PaymentAPI"; ServerNumber = 3; OSVersionCode = "0"; IP = "10.0.1.16"; OSVersion = "2012" },
    @{ Name = "WS2012-08"; ServiceName = "PaymentAPI"; ServerNumber = 4; OSVersionCode = "0"; IP = "10.0.1.17"; OSVersion = "2012" },

    # Windows Server 2022 - AccountList
    @{ Name = "WS2022-01"; ServiceName = "AccountList"; ServerNumber = 1; OSVersionCode = "1"; IP = "10.0.2.10"; OSVersion = "2022" },
    @{ Name = "WS2022-02"; ServiceName = "AccountList"; ServerNumber = 2; OSVersionCode = "1"; IP = "10.0.2.11"; OSVersion = "2022" },
    @{ Name = "WS2022-03"; ServiceName = "AccountList"; ServerNumber = 3; OSVersionCode = "1"; IP = "10.0.2.12"; OSVersion = "2022" },
    @{ Name = "WS2022-04"; ServiceName = "AccountList"; ServerNumber = 4; OSVersionCode = "1"; IP = "10.0.2.13"; OSVersion = "2022" },
    
    # Windows Server 2022 - PaymentAPI
    @{ Name = "WS2022-05"; ServiceName = "PaymentAPI"; ServerNumber = 1; OSVersionCode = "1"; IP = "10.0.2.14"; OSVersion = "2022" },
    @{ Name = "WS2022-06"; ServiceName = "PaymentAPI"; ServerNumber = 2; OSVersionCode = "1"; IP = "10.0.2.15"; OSVersion = "2022" },
    @{ Name = "WS2022-07"; ServiceName = "PaymentAPI"; ServerNumber = 3; OSVersionCode = "1"; IP = "10.0.2.16"; OSVersion = "2022" },
    @{ Name = "WS2022-08"; ServiceName = "PaymentAPI"; ServerNumber = 4; OSVersionCode = "1"; IP = "10.0.2.17"; OSVersion = "2022" }
)

$hostsFile = "C:\Windows\System32\drivers\etc\hosts"

# Initialize results storage
$allResults = @()
$detailedResults = @()
$serviceProjects = @{}

function Show-ServiceMenu {
    <#
    .SYNOPSIS
    Displays an interactive menu to select services for testing
    #>
    Write-Host "`n" + "="*70
    Write-Host "SERVICE SELECTION" -ForegroundColor Cyan
    Write-Host "="*70
    Write-Host "Available services:" -ForegroundColor Cyan
    
    for ($i = 0; $i -lt $availableServices.Count; $i++) {
        $service = $availableServices[$i]
        Write-Host "  $($i + 1). $($service.Name)" -ForegroundColor Green
        Write-Host "     └─ $($service.Description)"
        Write-Host "     └─ Project: $(Split-Path $service.ProjectPath -Leaf)"
        Write-Host ""
    }
    Write-Host "  $($availableServices.Count + 1). All Services" -ForegroundColor Green
    Write-Host ""
    
    $selectedServices = @()
    
    while ($true) {
        $input = Read-Host "Enter service number(s) to test (comma-separated, e.g., 1,2 or just $($availableServices.Count + 1) for all)"
        
        if ($input -eq "") {
            Write-Host "Please enter at least one service." -ForegroundColor Yellow
            continue
        }
        
        # Check if "All Services" option selected
        $choices = $input -split ',' | ForEach-Object { $_.Trim() }
        
        if ($choices -contains ([string]($availableServices.Count + 1))) {
            $selectedServices = $availableServices
            Write-Host "Selected: All Services" -ForegroundColor Green
            break
        }
        
        # Validate selections
        $valid = $true
        foreach ($choice in $choices) {
            if (-not [int]::TryParse($choice, [ref]$null) -or $choice -lt 1 -or $choice -gt $availableServices.Count) {
                Write-Host "Invalid selection: $choice" -ForegroundColor Red
                $valid = $false
                break
            }
            $selectedServices += $availableServices[([int]$choice - 1)]
        }
        
        if ($valid) {
            # Remove duplicates by service name
            $selectedServices = $selectedServices | Sort-Object -Property Name -Unique
            Write-Host "`nSelected services:" -ForegroundColor Green
            $selectedServices | ForEach-Object { 
                Write-Host "  ✓ $($_.Name)" -ForegroundColor Green
                Write-Host "    Project: $(Split-Path $_.ProjectPath -Leaf)" -ForegroundColor Gray
            }
            break
        }
    }
    
    return $selectedServices
}

function Validate-ServiceProjects {
    param(
        [array]$Services
    )
    
    Write-Host "`nValidating SoapUI project files..." -ForegroundColor Cyan
    
    $allValid = $true
    foreach ($service in $Services) {
        if (Test-Path $service.ProjectPath) {
            Write-Host "  ✓ $($service.Name): $($service.ProjectPath)" -ForegroundColor Green
        } else {
            Write-Host "  ✗ $($service.Name): File not found - $($service.ProjectPath)" -ForegroundColor Red
            $allValid = $false
        }
    }
    
    if (-not $allValid) {
        Write-Host "`nError: Some SoapUI project files are missing. Please check the paths above." -ForegroundColor Red
        exit
    }
    
    Write-Host "`nAll project files validated successfully!" -ForegroundColor Green
}

function Add-HostEntry {
    param(
        [string]$IP,
        [string]$Hostname
    )
    
    $hosts = Get-Content $hostsFile
    if ($hosts -notcontains "$IP`t$Hostname") {
        Add-Content -Path $hostsFile -Value "`n$IP`t$Hostname"
        Write-Host "Added host entry: $IP $Hostname" -ForegroundColor Green
        return $true
    }
    return $false
}

function Remove-HostEntry {
    param(
        [string]$Hostname
    )
    
    try {
        $hosts = Get-Content $hostsFile
        $updatedHosts = $hosts | Where-Object { $_ -notmatch "\s+$Hostname\s*$" }
        Set-Content -Path $hostsFile -Value $updatedHosts
        Write-Host "Removed host entry: $Hostname" -ForegroundColor Yellow
        return $true
    }
    catch {
        Write-Host "Warning: Could not remove host entry for $Hostname : $_" -ForegroundColor Yellow
        return $false
    }
}

function Run-SoapUITest {
    param(
        [string]$ServerName,
        [string]$ServiceName,
        [int]$ServerNumber,
        [string]$OSVersionCode,
        [string]$IP,
        [string]$OSVersion,
        [string]$SoapUIProject
    )
    
    # Format hostname as: ServiceName-ServerNumber-OSVersionCode
    $hostname = "$ServiceName-$ServerNumber-$OSVersionCode"
    
    $testName = "$ServerName-$OSVersion"
    $outputDir = Join-Path $ResultsDir $testName
    
    # Create output directory
    if (-not (Test-Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir | Out-Null
    }
    
    Write-Host "`n" + "="*70
    Write-Host "Testing: $ServerName | $IP $hostname | Windows Server $OSVersion" -ForegroundColor Cyan
    Write-Host "Service: $ServiceName" -ForegroundColor Cyan
    Write-Host "Project: $(Split-Path $SoapUIProject -Leaf)" -ForegroundColor Gray
    Write-Host "="*70
    
    try {
        # Add host entry
        Add-HostEntry -IP $IP -Hostname $hostname
        
        # Allow DNS/hosts cache to refresh
        Start-Sleep -Milliseconds 500
        
        # Run SoapUI with response time capturing
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        
        $soapUIArgs = @(
            "-e", "default",
            "-r",
            "-a",
            "-f", $outputDir,
            "-Dcom.eviware.soapui.analytics.enabled=false",
            "-Dproperty_name=$hostname",
            $SoapUIProject
        )
        
        # Execute SoapUI testrunner
        $process = Start-Process -FilePath "C:\Program Files\SmartBear\SoapUI-5.7.0\bin\testrunner.bat" `
            -ArgumentList $soapUIArgs `
            -NoNewWindow `
            -RedirectStandardOutput (Join-Path $outputDir "soapui-output.txt") `
            -PassThru
        
        $process.WaitForExit()
        $stopwatch.Stop()
        
        # Parse results
        $resultsFile = Get-ChildItem -Path $outputDir -Filter "*.txt" -Recurse | Sort-Object LastWriteTime | Select-Object -Last 1
        
        $testPassed = $process.ExitCode -eq 0
        
        # Extract response times from SoapUI output
        $responseTimes = @()
        if ($resultsFile) {
            $content = Get-Content $resultsFile.FullName -Raw
            # Parse SoapUI response times (adjust regex based on your format)
            $matches = [regex]::Matches($content, 'took\s+(\d+)\s*ms')
            $responseTimes = $matches | ForEach-Object { [int]$_.Groups[1].Value }
        }
        
        $result = @{
            ServerName = $ServerName
            Hostname = $hostname
            ServiceName = $ServiceName
            ServerNumber = $ServerNumber
            IP = $IP
            OSVersion = $OSVersion
            Status = if ($testPassed) { "PASSED" } else { "FAILED" }
            TotalTime = $stopwatch.ElapsedMilliseconds
            ExitCode = $process.ExitCode
            ResponseTimes = $responseTimes
            AvgResponseTime = if ($responseTimes.Count -gt 0) { [int]($responseTimes | Measure-Object -Average).Average } else { 0 }
            MaxResponseTime = if ($responseTimes.Count -gt 0) { ($responseTimes | Measure-Object -Maximum).Maximum } else { 0 }
            MinResponseTime = if ($responseTimes.Count -gt 0) { ($responseTimes | Measure-Object -Minimum).Minimum } else { 0 }
            Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }
        
        # Store detailed response times for Excel
        if ($responseTimes.Count -gt 0) {
            for ($i = 0; $i -lt $responseTimes.Count; $i++) {
                $detailedResults += @{
                    Timestamp = $result.Timestamp
                    ServerName = $ServerName
                    Hostname = $hostname
                    ServiceName = $ServiceName
                    ServerNumber = $ServerNumber
                    IP = $IP
                    OSVersion = $OSVersion
                    RequestNumber = $i + 1
                    ResponseTime = $responseTimes[$i]
                }
            }
        }
        
        Write-Host "Status: $($result.Status)" -ForegroundColor $(if ($testPassed) { "Green" } else { "Red" })
        Write-Host "Total Execution Time: $($result.TotalTime)ms"
        Write-Host "Avg Response Time: $($result.AvgResponseTime)ms | Min: $($result.MinResponseTime)ms | Max: $($result.MaxResponseTime)ms"
        
    }
    catch {
        Write-Host "Error during test: $_" -ForegroundColor Red
        $result = @{
            ServerName = $ServerName
            Hostname = $hostname
            ServiceName = $ServiceName
            ServerNumber = $ServerNumber
            IP = $IP
            OSVersion = $OSVersion
            Status = "ERROR"
            Error = $_.Exception.Message
            Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }
    }
    finally {
        # ALWAYS remove the host entry after testing this server
        Write-Host "Cleaning up host entry for this test..." -ForegroundColor Gray
        Remove-HostEntry -Hostname $hostname
        Start-Sleep -Milliseconds 300
    }
    
    return $result
}

function Create-ExcelReport {
    param(
        [array]$SummaryResults,
        [array]$DetailedResults,
        [string]$OutputPath
    )
    
    if (-not (Test-ExcelInstalled)) {
        Write-Host "Excel not found. Skipping Excel report generation." -ForegroundColor Yellow
        return
    }
    
    Write-Host "`nGenerating Excel report..." -ForegroundColor Cyan
    
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $workbook = $excel.Workbooks.Add()
        
        # Remove default sheets
        while ($workbook.Sheets.Count -gt 1) {
            $workbook.Sheets.Item(1).Delete()
        }
        
        # ===== SUMMARY DATA SHEET =====
        $summarySheet = $workbook.Sheets.Item(1)
        $summarySheet.Name = "Summary"
        
        # Add headers
        $summarySheet.Cells.Item(1, 1) = "Timestamp"
        $summarySheet.Cells.Item(1, 2) = "Server Name"
        $summarySheet.Cells.Item(1, 3) = "Service Name"
        $summarySheet.Cells.Item(1, 4) = "Server Number"
        $summarySheet.Cells.Item(1, 5) = "Hostname"
        $summarySheet.Cells.Item(1, 6) = "IP"
        $summarySheet.Cells.Item(1, 7) = "OS Version"
        $summarySheet.Cells.Item(1, 8) = "Status"
        $summarySheet.Cells.Item(1, 9) = "Total Time (ms)"
        $summarySheet.Cells.Item(1, 10) = "Avg Response (ms)"
        $summarySheet.Cells.Item(1, 11) = "Min Response (ms)"
        $summarySheet.Cells.Item(1, 12) = "Max Response (ms)"
        
        # Format header row
        $headerRange = $summarySheet.Range("A1:L1")
        $headerRange.Font.Bold = $true
        $headerRange.Interior.ColorIndex = 15
        $headerRange.HorizontalAlignment = -4108  # xlCenter
        
        # Add data
        $row = 2
        foreach ($result in $SummaryResults) {
            $summarySheet.Cells.Item($row, 1) = $result.Timestamp
            $summarySheet.Cells.Item($row, 2) = $result.ServerName
            $summarySheet.Cells.Item($row, 3) = $result.ServiceName
            $summarySheet.Cells.Item($row, 4) = $result.ServerNumber
            $summarySheet.Cells.Item($row, 5) = $result.Hostname
            $summarySheet.Cells.Item($row, 6) = $result.IP
            $summarySheet.Cells.Item($row, 7) = $result.OSVersion
            $summarySheet.Cells.Item($row, 8) = $result.Status
            $summarySheet.Cells.Item($row, 9) = $result.TotalTime
            $summarySheet.Cells.Item($row, 10) = $result.AvgResponseTime
            $summarySheet.Cells.Item($row, 11) = $result.MinResponseTime
            $summarySheet.Cells.Item($row, 12) = $result.MaxResponseTime
            
            # Color code status
            if ($result.Status -eq "PASSED") {
                $summarySheet.Cells.Item($row, 8).Interior.Color = 0x00B050  # Green
            } else {
                $summarySheet.Cells.Item($row, 8).Interior.Color = 0xFF0000  # Red
            }
            
            $row++
        }
        
        # Autofit columns
        $summarySheet.UsedRange.Columns.AutoFit() | Out-Null
        
        # ===== DETAILED DATA SHEET =====
        $detailSheet = $workbook.Sheets.Add()
        $detailSheet.Name = "Detailed Results"
        
        # Add headers
        $detailSheet.Cells.Item(1, 1) = "Timestamp"
        $detailSheet.Cells.Item(1, 2) = "Server Name"
        $detailSheet.Cells.Item(1, 3) = "Service Name"
        $detailSheet.Cells.Item(1, 4) = "Server Number"
        $detailSheet.Cells.Item(1, 5) = "Hostname"
        $detailSheet.Cells.Item(1, 6) = "IP"
        $detailSheet.Cells.Item(1, 7) = "OS Version"
        $detailSheet.Cells.Item(1, 8) = "Request Number"
        $detailSheet.Cells.Item(1, 9) = "Response Time (ms)"
        
        # Format header row
        $detailHeaderRange = $detailSheet.Range("A1:I1")
        $detailHeaderRange.Font.Bold = $true
        $detailHeaderRange.Interior.ColorIndex = 15
        $detailHeaderRange.HorizontalAlignment = -4108  # xlCenter
        
        # Add data
        $row = 2
        foreach ($detail in $DetailedResults) {
            $detailSheet.Cells.Item($row, 1) = $detail.Timestamp
            $detailSheet.Cells.Item($row, 2) = $detail.ServerName
            $detailSheet.Cells.Item($row, 3) = $detail.ServiceName
            $detailSheet.Cells.Item($row, 4) = $detail.ServerNumber
            $detailSheet.Cells.Item($row, 5) = $detail.Hostname
            $detailSheet.Cells.Item($row, 6) = $detail.IP
            $detailSheet.Cells.Item($row, 7) = $detail.OSVersion
            $detailSheet.Cells.Item($row, 8) = $detail.RequestNumber
            $detailSheet.Cells.Item($row, 9) = $detail.ResponseTime
            $row++
        }
        
        # Autofit columns
        $detailSheet.UsedRange.Columns.AutoFit() | Out-Null
        
        # ===== PIVOT TABLE SHEET =====
        $pivotSheet = $workbook.Sheets.Add()
        $pivotSheet.Name = "Pivot Analysis"
        
        # Create pivot table data source (from Summary sheet)
        $dataRange = $summarySheet.UsedRange
        $pivotCache = $workbook.PivotCaches().Create([Microsoft.Office.Interop.Excel.XlPivotTableSourceType]::xlDatabase, $dataRange)
        
        # Create pivot table on new sheet
        $pivotTable = $pivotCache.CreatePivotTable($pivotSheet.Range("A1"), "ResponseTimePivot")
        
        # Configure pivot table
        # Row: Service Name
        $pivotTable.PivotFields("Service Name").Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlRowField
        
        # Row: OS Version
        $pivotTable.PivotFields("OS Version").Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlRowField
        
        # Column: Status
        $pivotTable.PivotFields("Status").Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlColumnField
        
        # Data: Avg Response Time
        $avgField = $pivotTable.PivotFields("Avg Response (ms)")
        $avgField.Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlDataField
        $avgField.Function = [Microsoft.Office.Interop.Excel.XlConsolidationFunction]::xlAverage
        $avgField.Name = "Avg Response Time (ms)"
        
        # Add another data field: Count
        $countField = $pivotTable.PivotFields("Server Name")
        $countField.Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlDataField
        $countField.Function = [Microsoft.Office.Interop.Excel.XlConsolidationFunction]::xlCount
        
        # ===== COMPARISON SHEET =====
        $comparisonSheet = $workbook.Sheets.Add()
        $comparisonSheet.Name = "Service Comparison"
        
        # Calculate summary statistics by service and OS
        $services = $SummaryResults | Select-Object -ExpandProperty ServiceName | Sort-Object -Unique
        $osVersions = $SummaryResults | Select-Object -ExpandProperty OSVersion | Sort-Object -Unique
        
        $row = 1
        $col = 1
        
        # Header
        $comparisonSheet.Cells.Item($row, $col) = "Service / OS"
        $col++
        
        foreach ($os in $osVersions) {
            $comparisonSheet.Cells.Item($row, $col) = "Server $os - Avg Response (ms)"
            $col++
            $comparisonSheet.Cells.Item($row, $col) = "Server $os - Count"
            $col++
        }
        
        # Format header row
        $headerRange = $comparisonSheet.Range("A1:Z1")
        $headerRange.Font.Bold = $true
        $headerRange.Interior.ColorIndex = 15
        
        # Data rows
        $row = 2
        foreach ($service in $services) {
            $comparisonSheet.Cells.Item($row, 1) = $service
            $col = 2
            
            foreach ($os in $osVersions) {
                $serviceOSData = $SummaryResults | Where-Object { $_.ServiceName -eq $service -and $_.OSVersion -eq $os }
                $avgResponse = if ($serviceOSData.Count -gt 0) { [int]($serviceOSData.AvgResponseTime | Measure-Object -Average).Average } else { 0 }
                $count = if ($serviceOSData.Count -gt 0) { $serviceOSData.Count } else { 0 }
                
                $comparisonSheet.Cells.Item($row, $col) = $avgResponse
                $col++
                $comparisonSheet.Cells.Item($row, $col) = $count
                $col++
            }
            
            $row++
        }
        
        $comparisonSheet.UsedRange.Columns.AutoFit() | Out-Null
        
        # Save workbook
        $workbook.SaveAs($OutputPath)
        $workbook.Close()
        $excel.Quit()
        
        # Release COM objects
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pivotTable) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pivotCache) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        
        Write-Host "Excel report saved to: $OutputPath" -ForegroundColor Green
        
        # Open the file
        Invoke-Item $OutputPath
    }
    catch {
        Write-Host "Error creating Excel report: $_" -ForegroundColor Red
    }
}

# Main execution
Write-Host "`n╔════════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║   SOAP UI Response Time Testing - Multi-Service Suite              ║" -ForegroundColor Cyan
Write-Host "╚════════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan

Write-Host "`nConfiguration:" -ForegroundColor Cyan
Write-Host "Results Directory: $ResultsDir"

# Show service menu and get selections
$selectedServices = Show-ServiceMenu

# Validate that all project files exist
Validate-ServiceProjects -Services $selectedServices

# Build service project mapping
foreach ($service in $selectedServices) {
    $serviceProjects[$service.Name] = $service.ProjectPath
}

# Filter servers based on selected services
$servers = $allServers | Where-Object { $_.ServiceName -in $selectedServices.Name }

Write-Host "`nServers to test: $($servers.Count)" -ForegroundColor Cyan

# Display server configuration
Write-Host "`nServer Configuration:" -ForegroundColor Cyan
$servers | Format-Table -Property Name, ServiceName, ServerNumber, OSVersionCode, IP, OSVersion -AutoSize

# Create results directory
if (-not (Test-Path $ResultsDir)) {
    New-Item -ItemType Directory -Path $ResultsDir | Out-Null
}

# Backup hosts file
$backupPath = "$hostsFile.backup.$(Get-Date -Format 'yyyyMMdd-HHmmss')"
Copy-Item -Path $hostsFile -Destination $backupPath -Force
Write-Host "Backed up hosts file to: $backupPath" -ForegroundColor Yellow

# Confirm start
Write-Host "`n" + "="*70
$confirm = Read-Host "Ready to start testing? (Y/N)"
if ($confirm -ne "Y" -and $confirm -ne "y") {
    Write-Host "Test cancelled." -ForegroundColor Yellow
    exit
}

# Run tests
$startTime = Get-Date
Write-Host "`nStarting tests at $(Get-Date -Format 'HH:mm:ss')..." -ForegroundColor Cyan

foreach ($server in $servers) {
    # Get the correct SoapUI project for this service
    $soapUIProject = $serviceProjects[$server.ServiceName]
    
    $result = Run-SoapUITest -ServerName $server.Name `
                              -ServiceName $server.ServiceName `
                              -ServerNumber $server.ServerNumber `
                              -OSVersionCode $server.OSVersionCode `
                              -IP $server.IP `
                              -OSVersion $server.OSVersion `
                              -SoapUIProject $soapUIProject
    $allResults += $result
    
    # Small delay between tests to avoid overwhelming the network
    Start-Sleep -Seconds 2
}

# Generate summary report
Write-Host "`n" + "="*70
Write-Host "TEST SUMMARY" -ForegroundColor Cyan
Write-Host "="*70

$passed = @($allResults | Where-Object { $_.Status -eq "PASSED" }).Count
$failed = @($allResults | Where-Object { $_.Status -eq "FAILED" }).Count
$errors = @($allResults | Where-Object { $_.Status -eq "ERROR" }).Count

Write-Host "Passed: $passed | Failed: $failed | Errors: $errors" -ForegroundColor $(if ($failed -eq 0 -and $errors -eq 0) { "Green" } else { "Red" })

# Summary by Service and OS Version
Write-Host "`nSummary by Service and Windows Server Version:" -ForegroundColor Cyan
$allResults | Group-Object -Property @{ Expression = { "$($_.ServiceName) - Server $($_.OSVersion)" } } | ForEach-Object {
    $groupPassed = @($_.Group | Where-Object { $_.Status -eq "PASSED" }).Count
    $groupFailed = @($_.Group | Where-Object { $_.Status -eq "FAILED" }).Count
    $groupAvgResponse = if ($_.Group.AvgResponseTime) { [int]($_.Group.AvgResponseTime | Measure-Object -Average).Average } else { 0 }
    Write-Host "  $($_.Name): $groupPassed/$($_.Count) Passed | Avg Response: $groupAvgResponse ms"
}

# Export detailed CSV report
$csvPath = Join-Path $ResultsDir "test-results-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
$allResults | Select-Object Timestamp, ServerName, ServiceName, ServerNumber, Hostname, IP, OSVersion, Status, TotalTime, AvgResponseTime, MinResponseTime, MaxResponseTime |
    Export-Csv -Path $csvPath -NoTypeInformation
Write-Host "`nCSV report saved to: $csvPath" -ForegroundColor Green

# Create Excel report with pivot tables
$excelPath = Join-Path $ResultsDir "test-results-$(Get-Date -Format 'yyyyMMdd-HHmmss').xlsx"
Create-ExcelReport -SummaryResults $allResults -DetailedResults $detailedResults -OutputPath $excelPath

# Final hosts file status check
Write-Host "`nVerifying hosts file cleanup..." -ForegroundColor Gray
$currentHosts = Get-Content $hostsFile
$testEntries = $servers | Where-Object { $currentHosts -match "$($_.ServiceName)-$($_.ServerNumber)-$($_.OSVersionCode)" }
if ($testEntries.Count -eq 0) {
    Write-Host "✓ All test entries successfully removed from hosts file" -ForegroundColor Green
} else {
    Write-Host "⚠ Warning: Some test entries remain in hosts file:" -ForegroundColor Yellow
    $testEntries | ForEach-Object { Write-Host "  - $($_.ServiceName)-$($_.ServerNumber)-$($_.OSVersionCode)" }
}

Write-Host "`nTest run completed at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
Write-Host "Total execution time: $([math]::Round((Get-Date - $startTime).TotalMinutes, 2)) minutes" -ForegroundColor Cyan
