# test-soapui-servers.ps1

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
Write-Host "`nDetailed report saved to: $csvPath" -ForegroundColor Green

# Optional: Create Excel pivot table summary (if Excel is installed)
# ... you can add additional reporting here

Write-Host "`nTest run completed at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
Write-Host "Total execution time: $((Get-Date) - $startTime).TotalMinutes minutes" -ForegroundColor Cyan
