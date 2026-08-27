# SoapUI Testing for .NET Web Services against multiple servers

param(
    [string]$SoapUIProject = "C:\path\to\your\project.xml",
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

# Server configurations
$servers = @(
    # Windows Server 2012
    @{ Name = "WS2012-01"; Hostname = "ws2012-01.local"; IP = "10.0.1.10"; OSVersion = "2012" },
    @{ Name = "WS2012-02"; Hostname = "ws2012-02.local"; IP = "10.0.1.11"; OSVersion = "2012" },
    @{ Name = "WS2012-03"; Hostname = "ws2012-03.local"; IP = "10.0.1.12"; OSVersion = "2012" },
    @{ Name = "WS2012-04"; Hostname = "ws2012-04.local"; IP = "10.0.1.13"; OSVersion = "2012" },
    @{ Name = "WS2012-05"; Hostname = "ws2012-05.local"; IP = "10.0.1.14"; OSVersion = "2012" },
    @{ Name = "WS2012-06"; Hostname = "ws2012-06.local"; IP = "10.0.1.15"; OSVersion = "2012" },
    @{ Name = "WS2012-07"; Hostname = "ws2012-07.local"; IP = "10.0.1.16"; OSVersion = "2012" },
    @{ Name = "WS2012-08"; Hostname = "ws2012-08.local"; IP = "10.0.1.17"; OSVersion = "2012" },

    # Windows Server 2022
    @{ Name = "WS2022-01"; Hostname = "ws2022-01.local"; IP = "10.0.2.10"; OSVersion = "2022" },
    @{ Name = "WS2022-02"; Hostname = "ws2022-02.local"; IP = "10.0.2.11"; OSVersion = "2022" },
    @{ Name = "WS2022-03"; Hostname = "ws2022-03.local"; IP = "10.0.2.12"; OSVersion = "2022" },
    @{ Name = "WS2022-04"; Hostname = "ws2022-04.local"; IP = "10.0.2.13"; OSVersion = "2022" },
    @{ Name = "WS2022-05"; Hostname = "ws2022-05.local"; IP = "10.0.2.14"; OSVersion = "2022" },
    @{ Name = "WS2022-06"; Hostname = "ws2022-06.local"; IP = "10.0.2.15"; OSVersion = "2022" },
    @{ Name = "WS2022-07"; Hostname = "ws2022-07.local"; IP = "10.0.2.16"; OSVersion = "2022" },
    @{ Name = "WS2022-08"; Hostname = "ws2022-08.local"; IP = "10.0.2.17"; OSVersion = "2022" }
)

$hostsFile = "C:\Windows\System32\drivers\etc\hosts"

# Initialize results storage
$allResults = @()
$detailedResults = @()

function Add-HostEntry {
    param(
        [string]$IP,
        [string]$Hostname
    )
    
    $hosts = Get-Content $hostsFile
    if ($hosts -notcontains "$IP`t$Hostname") {
        Add-Content -Path $hostsFile -Value "`n$IP`t$Hostname"
        Write-Host "Added host entry: $IP -> $Hostname" -ForegroundColor Green
        return $true
    }
    return $false
}

function Remove-HostEntry {
    param(
        [string]$Hostname
    )
    
    $hosts = Get-Content $hostsFile
    $updatedHosts = $hosts | Where-Object { $_ -notmatch "\s+$Hostname\s*$" }
    Set-Content -Path $hostsFile -Value $updatedHosts
    Write-Host "Removed host entry: $Hostname" -ForegroundColor Yellow
}

function Run-SoapUITest {
    param(
        [string]$ServerName,
        [string]$Hostname,
        [string]$IP,
        [string]$OSVersion
    )
    
    $testName = "$ServerName-$OSVersion"
    $outputDir = Join-Path $ResultsDir $testName
    
    # Create output directory
    if (-not (Test-Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir | Out-Null
    }
    
    Write-Host "`n" + "="*60
    Write-Host "Testing: $ServerName ($Hostname - $IP) | Windows Server $OSVersion" -ForegroundColor Cyan
    Write-Host "="*60
    
    try {
        # Add host entry
        Add-HostEntry -IP $IP -Hostname $Hostname
        
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
            "-Dproperty_name=$Hostname",
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
            Hostname = $Hostname
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
                    Hostname = $Hostname
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
            Hostname = $Hostname
            IP = $IP
            OSVersion = $OSVersion
            Status = "ERROR"
            Error = $_.Exception.Message
            Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }
    }
    finally {
        if ($CleanupHosts) {
            Remove-HostEntry -Hostname $Hostname
        }
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
        $summarySheet.Cells.Item(1, 3) = "Hostname"
        $summarySheet.Cells.Item(1, 4) = "IP"
        $summarySheet.Cells.Item(1, 5) = "OS Version"
        $summarySheet.Cells.Item(1, 6) = "Status"
        $summarySheet.Cells.Item(1, 7) = "Total Time (ms)"
        $summarySheet.Cells.Item(1, 8) = "Avg Response (ms)"
        $summarySheet.Cells.Item(1, 9) = "Min Response (ms)"
        $summarySheet.Cells.Item(1, 10) = "Max Response (ms)"
        
        # Format header row
        $headerRange = $summarySheet.Range("A1:J1")
        $headerRange.Font.Bold = $true
        $headerRange.Interior.ColorIndex = 15
        $headerRange.HorizontalAlignment = -4108  # xlCenter
        
        # Add data
        $row = 2
        foreach ($result in $SummaryResults) {
            $summarySheet.Cells.Item($row, 1) = $result.Timestamp
            $summarySheet.Cells.Item($row, 2) = $result.ServerName
            $summarySheet.Cells.Item($row, 3) = $result.Hostname
            $summarySheet.Cells.Item($row, 4) = $result.IP
            $summarySheet.Cells.Item($row, 5) = $result.OSVersion
            $summarySheet.Cells.Item($row, 6) = $result.Status
            $summarySheet.Cells.Item($row, 7) = $result.TotalTime
            $summarySheet.Cells.Item($row, 8) = $result.AvgResponseTime
            $summarySheet.Cells.Item($row, 9) = $result.MinResponseTime
            $summarySheet.Cells.Item($row, 10) = $result.MaxResponseTime
            
            # Color code status
            if ($result.Status -eq "PASSED") {
                $summarySheet.Cells.Item($row, 6).Interior.Color = 0x00B050  # Green
            } else {
                $summarySheet.Cells.Item($row, 6).Interior.Color = 0xFF0000  # Red
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
        $detailSheet.Cells.Item(1, 3) = "Hostname"
        $detailSheet.Cells.Item(1, 4) = "IP"
        $detailSheet.Cells.Item(1, 5) = "OS Version"
        $detailSheet.Cells.Item(1, 6) = "Request Number"
        $detailSheet.Cells.Item(1, 7) = "Response Time (ms)"
        
        # Format header row
        $detailHeaderRange = $detailSheet.Range("A1:G1")
        $detailHeaderRange.Font.Bold = $true
        $detailHeaderRange.Interior.ColorIndex = 15
        $detailHeaderRange.HorizontalAlignment = -4108  # xlCenter
        
        # Add data
        $row = 2
        foreach ($detail in $DetailedResults) {
            $detailSheet.Cells.Item($row, 1) = $detail.Timestamp
            $detailSheet.Cells.Item($row, 2) = $detail.ServerName
            $detailSheet.Cells.Item($row, 3) = $detail.Hostname
            $detailSheet.Cells.Item($row, 4) = $detail.IP
            $detailSheet.Cells.Item($row, 5) = $detail.OSVersion
            $detailSheet.Cells.Item($row, 6) = $detail.RequestNumber
            $detailSheet.Cells.Item($row, 7) = $detail.ResponseTime
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
        $comparisonSheet.Name = "OS Comparison"
        
        # Calculate summary statistics
        $ws2012Data = $SummaryResults | Where-Object { $_.OSVersion -eq "2012" }
        $ws2022Data = $SummaryResults | Where-Object { $_.OSVersion -eq "2022" }
        
        $comparisonSheet.Cells.Item(1, 1) = "Metric"
        $comparisonSheet.Cells.Item(1, 2) = "Windows Server 2012"
        $comparisonSheet.Cells.Item(1, 3) = "Windows Server 2022"
        $comparisonSheet.Cells.Item(1, 4) = "Difference"
        $comparisonSheet.Cells.Item(1, 5) = "% Improvement"
        
        $headerRange = $comparisonSheet.Range("A1:E1")
        $headerRange.Font.Bold = $true
        $headerRange.Interior.ColorIndex = 15
        
        $row = 2
        
        # Average Response Time
        $comparisonSheet.Cells.Item($row, 1) = "Avg Response Time (ms)"
        $avg2012 = if ($ws2012Data.Count -gt 0) { [int]($ws2012Data.AvgResponseTime | Measure-Object -Average).Average } else { 0 }
        $avg2022 = if ($ws2022Data.Count -gt 0) { [int]($ws2022Data.AvgResponseTime | Measure-Object -Average).Average } else { 0 }
        $comparisonSheet.Cells.Item($row, 2) = $avg2012
        $comparisonSheet.Cells.Item($row, 3) = $avg2022
        $comparisonSheet.Cells.Item($row, 4) = $avg2012 - $avg2022
        
        if ($avg2012 -ne 0) {
            $improvement = [math]::Round((($avg2012 - $avg2022) / $avg2012) * 100, 2)
            $comparisonSheet.Cells.Item($row, 5) = "$improvement%"
        }
        
        $row++
        
        # Max Response Time
        $comparisonSheet.Cells.Item($row, 1) = "Max Response Time (ms)"
        $max2012 = if ($ws2012Data.Count -gt 0) { ($ws2012Data.MaxResponseTime | Measure-Object -Maximum).Maximum } else { 0 }
        $max2022 = if ($ws2022Data.Count -gt 0) { ($ws2022Data.MaxResponseTime | Measure-Object -Maximum).Maximum } else { 0 }
        $comparisonSheet.Cells.Item($row, 2) = $max2012
        $comparisonSheet.Cells.Item($row, 3) = $max2022
        $comparisonSheet.Cells.Item($row, 4) = $max2012 - $max2022
        
        if ($max2012 -ne 0) {
            $improvement = [math]::Round((($max2012 - $max2022) / $max2012) * 100, 2)
            $comparisonSheet.Cells.Item($row, 5) = "$improvement%"
        }
        
        $row++
        
        # Min Response Time
        $comparisonSheet.Cells.Item($row, 1) = "Min Response Time (ms)"
        $min2012 = if ($ws2012Data.Count -gt 0) { ($ws2012Data.MinResponseTime | Measure-Object -Minimum).Minimum } else { 0 }
        $min2022 = if ($ws2022Data.Count -gt 0) { ($ws2022Data.MinResponseTime | Measure-Object -Minimum).Minimum } else { 0 }
        $comparisonSheet.Cells.Item($row, 2) = $min2012
        $comparisonSheet.Cells.Item($row, 3) = $min2022
        $comparisonSheet.Cells.Item($row, 4) = $min2012 - $min2022
        
        if ($min2012 -ne 0) {
            $improvement = [math]::Round((($min2012 - $min2022) / $min2012) * 100, 2)
            $comparisonSheet.Cells.Item($row, 5) = "$improvement%"
        }
        
        $row++
        
        # Test Success Rate
        $comparisonSheet.Cells.Item($row, 1) = "Success Rate"
        $passed2012 = @($ws2012Data | Where-Object { $_.Status -eq "PASSED" }).Count
        $passed2022 = @($ws2022Data | Where-Object { $_.Status -eq "PASSED" }).Count
        $rate2012 = if ($ws2012Data.Count -gt 0) { [math]::Round(($passed2012 / $ws2012Data.Count) * 100, 0) } else { 0 }
        $rate2022 = if ($ws2022Data.Count -gt 0) { [math]::Round(($passed2022 / $ws2022Data.Count) * 100, 0) } else { 0 }
        $comparisonSheet.Cells.Item($row, 2) = "$rate2012%"
        $comparisonSheet.Cells.Item($row, 3) = "$rate2022%"
        $comparisonSheet.Cells.Item($row, 4) = "$($rate2022 - $rate2012)%"
        
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
Write-Host "Starting SOAP UI Response Time Tests" -ForegroundColor Cyan
Write-Host "Target Project: $SoapUIProject" -ForegroundColor Cyan
Write-Host "Results Directory: $ResultsDir" -ForegroundColor Cyan
Write-Host "Servers: $($servers.Count) total (8x Server 2012, 8x Server 2022)" -ForegroundColor Cyan

# Create results directory
if (-not (Test-Path $ResultsDir)) {
    New-Item -ItemType Directory -Path $ResultsDir | Out-Null
}

# Backup hosts file
$backupPath = "$hostsFile.backup.$(Get-Date -Format 'yyyyMMdd-HHmmss')"
Copy-Item -Path $hostsFile -Destination $backupPath -Force
Write-Host "Backed up hosts file to: $backupPath" -ForegroundColor Yellow

# Run tests
$startTime = Get-Date
foreach ($server in $servers) {
    $result = Run-SoapUITest -ServerName $server.Name `
                              -Hostname $server.Hostname `
                              -IP $server.IP `
                              -OSVersion $server.OSVersion
    $allResults += $result
    
    # Small delay between tests to avoid overwhelming the network
    Start-Sleep -Seconds 2
}

# Generate summary report
Write-Host "`n" + "="*60
Write-Host "TEST SUMMARY" -ForegroundColor Cyan
Write-Host "="*60

$passed = @($allResults | Where-Object { $_.Status -eq "PASSED" }).Count
$failed = @($allResults | Where-Object { $_.Status -eq "FAILED" }).Count
$errors = @($allResults | Where-Object { $_.Status -eq "ERROR" }).Count

Write-Host "Passed: $passed | Failed: $failed | Errors: $errors" -ForegroundColor $(if ($failed -eq 0 -and $errors -eq 0) { "Green" } else { "Red" })

# Summary by OS Version
Write-Host "`nSummary by Windows Server Version:" -ForegroundColor Cyan
$allResults | Group-Object -Property OSVersion | ForEach-Object {
    $groupPassed = @($_.Group | Where-Object { $_.Status -eq "PASSED" }).Count
    $groupFailed = @($_.Group | Where-Object { $_.Status -eq "FAILED" }).Count
    $groupAvgResponse = if ($_.Group.AvgResponseTime) { [int]($_.Group.AvgResponseTime | Measure-Object -Average).Average } else { 0 }
    Write-Host "  Server $($_.Name): $groupPassed/$($_.Count) Passed | Avg Response: $groupAvgResponse ms"
}

# Export detailed CSV report
$csvPath = Join-Path $ResultsDir "test-results-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
$allResults | Select-Object Timestamp, ServerName, Hostname, IP, OSVersion, Status, TotalTime, AvgResponseTime, MinResponseTime, MaxResponseTime |
    Export-Csv -Path $csvPath -NoTypeInformation
Write-Host "`nCSV report saved to: $csvPath" -ForegroundColor Green

# Create Excel report with pivot tables
$excelPath = Join-Path $ResultsDir "test-results-$(Get-Date -Format 'yyyyMMdd-HHmmss').xlsx"
Create-ExcelReport -SummaryResults $allResults -DetailedResults $detailedResults -OutputPath $excelPath

Write-Host "`nTest run completed at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
Write-Host "Total execution time: $([math]::Round((Get-Date - $startTime).TotalMinutes, 2)) minutes" -ForegroundColor Cyan
