<#
.SYNOPSIS
    Generates an HTML report for Windows LAPS deployment status across domain machines.

.DESCRIPTION
    This script queries Active Directory for all computer objects and checks their LAPS status.
    It supports both Legacy LAPS and Windows LAPS (native).
    It differentiates between Windows Clients and Windows Servers and generates a detailed HTML report.

.NOTES
    Requires: Active Directory PowerShell module
    Legacy LAPS Attribute: ms-Mcs-AdmPwdExpirationTime
    Windows LAPS Attribute: msLAPS-PasswordExpirationTime
#>

# Import required module
Import-Module ActiveDirectory -ErrorAction Stop

# Output file path
$reportPath = "$PSScriptRoot\LAPS-Deployment-Report.html"

# Ensure output directory exists
$outputDir = Split-Path $reportPath -Parent
if (!(Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

Write-Host "Gathering computer information from Active Directory..." -ForegroundColor Cyan

# Query all computer objects with relevant properties (including both Legacy and Windows LAPS attributes)
$computers = Get-ADComputer -Filter * -Properties Name, OperatingSystem, OperatingSystemVersion, `
    'msLAPS-PasswordExpirationTime', 'ms-Mcs-AdmPwdExpirationTime', Enabled, LastLogonDate, DistinguishedName | 
    Where-Object { $_.OperatingSystem -like "*Windows*" }

Write-Host "Processing $($computers.Count) computers..." -ForegroundColor Cyan

# Process computers and determine LAPS status
$results = foreach ($comp in $computers) {
    
    # Determine OS Type
    if ($comp.OperatingSystem -like "*Server*") {
        $osType = "Windows Server"
    } else {
        $osType = "Windows Client"
    }
    
    # Check LAPS status (both Legacy and Windows LAPS)
    $legacyLAPS = $comp.'ms-Mcs-AdmPwdExpirationTime'
    $windowsLAPS = $comp.'msLAPS-PasswordExpirationTime'
    
    if ($windowsLAPS) {
        $lapsEnabled = "Enabled"
        $lapsType = "Windows LAPS"
    } elseif ($legacyLAPS) {
        $lapsEnabled = "Enabled"
        $lapsType = "Legacy LAPS"
    } else {
        $lapsEnabled = "Not Enabled"
        $lapsType = "N/A"
    }
    
    # Format last logon date
    $lastLogon = if ($comp.LastLogonDate) {
        $comp.LastLogonDate.ToString("yyyy-MM-dd HH:mm:ss")
    } else {
        "Never"
    }
    
    [PSCustomObject]@{
        ComputerName = $comp.Name
        OSType = $osType
        OperatingSystem = $comp.OperatingSystem
        OSVersion = $comp.OperatingSystemVersion
        LAPSStatus = $lapsEnabled
        LAPSType = $lapsType
        AccountEnabled = $comp.Enabled
        LastLogon = $lastLogon
        OU = ($comp.DistinguishedName -split ',', 2)[1]
    }
}

# Calculate statistics
$totalComputers = $results.Count
$serversTotal = ($results | Where-Object { $_.OSType -eq "Windows Server" }).Count
$clientsTotal = ($results | Where-Object { $_.OSType -eq "Windows Client" }).Count
$lapsEnabledTotal = ($results | Where-Object { $_.LAPSStatus -eq "Enabled" }).Count
$lapsDisabledTotal = ($results | Where-Object { $_.LAPSStatus -eq "Not Enabled" }).Count
$serversLAPSEnabled = ($results | Where-Object { $_.OSType -eq "Windows Server" -and $_.LAPSStatus -eq "Enabled" }).Count
$clientsLAPSEnabled = ($results | Where-Object { $_.OSType -eq "Windows Client" -and $_.LAPSStatus -eq "Enabled" }).Count

# Calculate LAPS type statistics
$legacyLAPSCount = ($results | Where-Object { $_.LAPSType -eq "Legacy LAPS" }).Count
$windowsLAPSCount = ($results | Where-Object { $_.LAPSType -eq "Windows LAPS" }).Count

# Calculate percentages
$lapsPercentage = if ($totalComputers -gt 0) { [math]::Round(($lapsEnabledTotal / $totalComputers) * 100, 2) } else { 0 }
$serversPercentage = if ($serversTotal -gt 0) { [math]::Round(($serversLAPSEnabled / $serversTotal) * 100, 2) } else { 0 }
$clientsPercentage = if ($clientsTotal -gt 0) { [math]::Round(($clientsLAPSEnabled / $clientsTotal) * 100, 2) } else { 0 }

# Generate HTML Report
$html = @"
<!DOCTYPE html>
<html>
<head>
    <title>Windows LAPS Deployment Report</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 20px;
            background-color: #f5f5f5;
        }
        h1 {
            color: #2c3e50;
            border-bottom: 3px solid #3498db;
            padding-bottom: 10px;
        }
        h2 {
            color: #34495e;
            margin-top: 30px;
        }
        .summary {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 15px;
            margin: 20px 0;
        }
        .stat-card {
            background: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            border-left: 4px solid #3498db;
        }
        .stat-card.server {
            border-left-color: #9b59b6;
        }
        .stat-card.client {
            border-left-color: #1abc9c;
        }
        .stat-card.enabled {
            border-left-color: #27ae60;
        }
        .stat-card.disabled {
            border-left-color: #e74c3c;
        }
        .stat-card.legacy {
            border-left-color: #f39c12;
        }
        .stat-card.native {
            border-left-color: #16a085;
        }
        .stat-label {
            font-size: 14px;
            color: #7f8c8d;
            margin-bottom: 5px;
        }
        .stat-value {
            font-size: 28px;
            font-weight: bold;
            color: #2c3e50;
        }
        .stat-percentage {
            font-size: 14px;
            color: #27ae60;
            margin-top: 5px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            background: white;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            margin: 20px 0;
        }
        th {
            background-color: #3498db;
            color: white;
            padding: 12px;
            text-align: left;
            font-weight: 600;
            cursor: pointer;
            user-select: none;
        }
        th:hover {
            background-color: #2980b9;
        }
        td {
            padding: 10px 12px;
            border-bottom: 1px solid #ecf0f1;
        }
        tr:hover {
            background-color: #f8f9fa;
        }
        .enabled-status {
            color: white;
            padding: 4px 8px;
            border-radius: 4px;
            font-size: 12px;
            font-weight: bold;
            display: inline-block;
        }
        .status-enabled {
            background-color: #27ae60;
        }
        .status-disabled {
            background-color: #e74c3c;
        }
        .laps-type {
            font-size: 11px;
            padding: 2px 6px;
            border-radius: 3px;
            margin-left: 5px;
            display: inline-block;
        }
        .type-legacy {
            background-color: #f39c12;
            color: white;
        }
        .type-native {
            background-color: #16a085;
            color: white;
        }
        .os-server {
            color: #9b59b6;
            font-weight: 600;
        }
        .os-client {
            color: #1abc9c;
            font-weight: 600;
        }
        .footer {
            margin-top: 30px;
            padding: 20px;
            background: white;
            border-radius: 8px;
            text-align: center;
            color: #7f8c8d;
            font-size: 14px;
        }
        .filter-container {
            margin: 20px 0;
            padding: 15px;
            background: white;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .filter-container label {
            margin-right: 15px;
            font-weight: 600;
        }
        .filter-container select, .filter-container input {
            padding: 8px;
            margin-right: 15px;
            border: 1px solid #bdc3c7;
            border-radius: 4px;
        }
    </style>
</head>
<body>
    <h1>🔐 Windows LAPS Deployment Report</h1>
    
    <div class="summary">
        <div class="stat-card">
            <div class="stat-label">Total Computers</div>
            <div class="stat-value">$totalComputers</div>
        </div>
        <div class="stat-card enabled">
            <div class="stat-label">LAPS Enabled</div>
            <div class="stat-value">$lapsEnabledTotal</div>
            <div class="stat-percentage">$lapsPercentage%</div>
        </div>
        <div class="stat-card disabled">
            <div class="stat-label">LAPS Not Enabled</div>
            <div class="stat-value">$lapsDisabledTotal</div>
        </div>
        <div class="stat-card legacy">
            <div class="stat-label">Legacy LAPS</div>
            <div class="stat-value">$legacyLAPSCount</div>
        </div>
        <div class="stat-card native">
            <div class="stat-label">Windows LAPS</div>
            <div class="stat-value">$windowsLAPSCount</div>
        </div>
        <div class="stat-card server">
            <div class="stat-label">Windows Servers</div>
            <div class="stat-value">$serversTotal</div>
            <div class="stat-percentage">$serversLAPSEnabled enabled ($serversPercentage%)</div>
        </div>
        <div class="stat-card client">
            <div class="stat-label">Windows Clients</div>
            <div class="stat-value">$clientsTotal</div>
            <div class="stat-percentage">$clientsLAPSEnabled enabled ($clientsPercentage%)</div>
        </div>
    </div>

    <div class="filter-container">
        <label>Filter:</label>
        <select id="osTypeFilter" onchange="filterTable()">
            <option value="all">All OS Types</option>
            <option value="Windows Server">Windows Server</option>
            <option value="Windows Client">Windows Client</option>
        </select>
        <select id="lapsFilter" onchange="filterTable()">
            <option value="all">All LAPS Status</option>
            <option value="Enabled">LAPS Enabled</option>
            <option value="Not Enabled">LAPS Not Enabled</option>
        </select>
        <select id="lapsTypeFilter" onchange="filterTable()">
            <option value="all">All LAPS Types</option>
            <option value="Legacy LAPS">Legacy LAPS</option>
            <option value="Windows LAPS">Windows LAPS</option>
        </select>
        <select id="accountFilter" onchange="filterTable()">
            <option value="all">All Account Status</option>
            <option value="Yes">Account Enabled</option>
            <option value="No">Account Disabled</option>
        </select>
        <input type="text" id="searchBox" placeholder="Search computer name..." onkeyup="filterTable()">
    </div>

    <h2>Detailed Computer Report</h2>
    <table id="computerTable">
        <thead>
            <tr>
                <th onclick="sortTable(0)">Computer Name ▼</th>
                <th onclick="sortTable(1)">OS Type ▼</th>
                <th onclick="sortTable(2)">Operating System ▼</th>
                <th onclick="sortTable(3)">LAPS Status ▼</th>
                <th onclick="sortTable(4)">LAPS Type ▼</th>
                <th onclick="sortTable(5)">Account Enabled ▼</th>
                <th onclick="sortTable(6)">Last Logon ▼</th>
                <th onclick="sortTable(7)">Organizational Unit ▼</th>
            </tr>
        </thead>
        <tbody>
"@

foreach ($comp in $results | Sort-Object OSType, ComputerName) {
    $lapsClass = if ($comp.LAPSStatus -eq "Enabled") { "status-enabled" } else { "status-disabled" }
    $osClass = if ($comp.OSType -eq "Windows Server") { "os-server" } else { "os-client" }
    $accountStatus = if ($comp.AccountEnabled) { "Yes" } else { "No" }
    
    # LAPS Type badge
    $lapsTypeBadge = if ($comp.LAPSType -eq "Legacy LAPS") {
        '<span class="laps-type type-legacy">Legacy</span>'
    } elseif ($comp.LAPSType -eq "Windows LAPS") {
        '<span class="laps-type type-native">Native</span>'
    } else {
        ''
    }
    
    $html += @"
            <tr>
                <td><strong>$($comp.ComputerName)</strong></td>
                <td class="$osClass">$($comp.OSType)</td>
                <td>$($comp.OperatingSystem)</td>
                <td><span class="enabled-status $lapsClass">$($comp.LAPSStatus)</span></td>
                <td>$($comp.LAPSType) $lapsTypeBadge</td>
                <td>$accountStatus</td>
                <td>$($comp.LastLogon)</td>
                <td><small>$($comp.OU)</small></td>
            </tr>
"@
}

$html += @"
        </tbody>
    </table>

    <div class="footer">
        Report generated on $(Get-Date -Format "yyyy-MM-dd HH:mm:ss") | Domain: $env:USERDNSDOMAIN
    </div>

    <script>
        function sortTable(n) {
            var table, rows, switching, i, x, y, shouldSwitch, dir, switchcount = 0;
            table = document.getElementById("computerTable");
            switching = true;
            dir = "asc";
            
            while (switching) {
                switching = false;
                rows = table.rows;
                
                for (i = 1; i < (rows.length - 1); i++) {
                    shouldSwitch = false;
                    x = rows[i].getElementsByTagName("TD")[n];
                    y = rows[i + 1].getElementsByTagName("TD")[n];
                    
                    if (dir == "asc") {
                        if (x.innerHTML.toLowerCase() > y.innerHTML.toLowerCase()) {
                            shouldSwitch = true;
                            break;
                        }
                    } else if (dir == "desc") {
                        if (x.innerHTML.toLowerCase() < y.innerHTML.toLowerCase()) {
                            shouldSwitch = true;
                            break;
                        }
                    }
                }
                
                if (shouldSwitch) {
                    rows[i].parentNode.insertBefore(rows[i + 1], rows[i]);
                    switching = true;
                    switchcount++;
                } else {
                    if (switchcount == 0 && dir == "asc") {
                        dir = "desc";
                        switching = true;
                    }
                }
            }
        }

        function filterTable() {
            var osFilter = document.getElementById("osTypeFilter").value;
            var lapsFilter = document.getElementById("lapsFilter").value;
            var lapsTypeFilter = document.getElementById("lapsTypeFilter").value;
            var accountFilter = document.getElementById("accountFilter").value;
            var searchText = document.getElementById("searchBox").value.toLowerCase();
            var table = document.getElementById("computerTable");
            var tr = table.getElementsByTagName("tr");

            for (var i = 1; i < tr.length; i++) {
                var tdName = tr[i].getElementsByTagName("td")[0];
                var tdOS = tr[i].getElementsByTagName("td")[1];
                var tdLAPS = tr[i].getElementsByTagName("td")[3];
                var tdLAPSType = tr[i].getElementsByTagName("td")[4];
                var tdAccount = tr[i].getElementsByTagName("td")[5];
                
                if (tdName && tdOS && tdLAPS && tdLAPSType && tdAccount) {
                    var nameMatch = tdName.textContent.toLowerCase().indexOf(searchText) > -1;
                    var osMatch = (osFilter === "all" || tdOS.textContent === osFilter);
                    var lapsText = tdLAPS.textContent.trim();
                    var lapsMatch = (lapsFilter === "all" || lapsText === lapsFilter);
                    var lapsTypeText = tdLAPSType.textContent.trim().split(' ')[0] + ' ' + tdLAPSType.textContent.trim().split(' ')[1];
                    var lapsTypeMatch = (lapsTypeFilter === "all" || lapsTypeText === lapsTypeFilter);
                    var accountText = tdAccount.textContent.trim();
                    var accountMatch = (accountFilter === "all" || accountText === accountFilter);
                    
                    if (nameMatch && osMatch && lapsMatch && lapsTypeMatch && accountMatch) {
                        tr[i].style.display = "";
                    } else {
                        tr[i].style.display = "none";
                    }
                }
            }
        }
    </script>
</body>
</html>
"@

# Save the report
$html | Out-File -FilePath $reportPath -Encoding UTF8

Write-Host "`nReport generated successfully!" -ForegroundColor Green
Write-Host "Location: $reportPath" -ForegroundColor Yellow
Write-Host "`nSummary:" -ForegroundColor Cyan
Write-Host "  Total Computers: $totalComputers"
Write-Host "  LAPS Enabled: $lapsEnabledTotal ($lapsPercentage%)" -ForegroundColor Green
Write-Host "    - Legacy LAPS: $legacyLAPSCount" -ForegroundColor Yellow
Write-Host "    - Windows LAPS: $windowsLAPSCount" -ForegroundColor Cyan
Write-Host "  LAPS Not Enabled: $lapsDisabledTotal" -ForegroundColor Red
Write-Host "  Windows Servers: $serversTotal ($serversLAPSEnabled enabled, $serversPercentage%)"
Write-Host "  Windows Clients: $clientsTotal ($clientsLAPSEnabled enabled, $clientsPercentage%)"

# Open the report in default browser
Start-Process $reportPath
