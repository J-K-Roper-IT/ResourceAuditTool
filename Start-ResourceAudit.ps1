[CmdletBinding()]
param(
    [string]$OutputRoot = "$env:USERPROFILE\Documents\ResourceAudit",
    [string[]]$AuditPaths = @(),
    [int]$KeepDays = 14,
    [switch]$ExportExcel,
    [string]$ExcelPath
)

$ErrorActionPreference = "Stop"

function Get-MachineTypeInfo {
    $cs = Get-CimInstance Win32_ComputerSystem
    $model = [string]$cs.Model
    $type = switch -Regex ($model) {
        'Virtual|VMware|VirtualBox|HVM domU|KVM' { 'VM'; break }
        default { 'Physical' }
    }

    [pscustomobject]@{
        ComputerName = $env:COMPUTERNAME
        Manufacturer = $cs.Manufacturer
        Model = $model
        MachineType = $type
        TotalMemoryGB = [math]::Round($cs.TotalPhysicalMemory / 1GB, 2)
    }
}

function Get-FixedDriveAudit {
    Get-CimInstance Win32_LogicalDisk -Filter "DriveType = 3" | ForEach-Object {
        [pscustomobject]@{
            ComputerName = $env:COMPUTERNAME
            Drive = $_.DeviceID
            VolumeName = $_.VolumeName
            SizeGB = [math]::Round($_.Size / 1GB, 2)
            FreeGB = [math]::Round($_.FreeSpace / 1GB, 2)
            UsedGB = [math]::Round(($_.Size - $_.FreeSpace) / 1GB, 2)
            FreePercent = if ($_.Size) { [math]::Round(($_.FreeSpace / $_.Size) * 100, 2) } else { $null }
        }
    }
}

function Get-DirectorySize {
    param([string]$Path)

    $itemCount = 0L
    $totalBytes = 0L

    Get-ChildItem -LiteralPath $Path -File -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object {
        $itemCount++
        $totalBytes += $_.Length
    }

    [pscustomobject]@{
        ComputerName = $env:COMPUTERNAME
        Path = $Path
        FileCount = $itemCount
        SizeGB = [math]::Round($totalBytes / 1GB, 2)
        Error = $null
    }
}

function Get-IisAudit {
    $sites = @()
    if (Get-Command Get-Website -ErrorAction SilentlyContinue) {
        $sites = Get-Website | ForEach-Object {
            [pscustomobject]@{
                ComputerName = $env:COMPUTERNAME
                SiteName = $_.Name
                State = $_.State
                PhysicalPath = $_.PhysicalPath
                ApplicationPool = $_.ApplicationPool
            }
        }
    }
    $sites
}

function Get-DatabaseFootprint {
    $results = @()

    if (Test-Path 'HKLM:\SOFTWARE\Oracle') {
        $results += [pscustomobject]@{
            ComputerName = $env:COMPUTERNAME
            Engine = 'Oracle'
            Detected = $true
            Notes = 'Oracle registry keys detected'
        }
    }

    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server' -or Test-Path 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Microsoft SQL Server') {
        $results += [pscustomobject]@{
            ComputerName = $env:COMPUTERNAME
            Engine = 'SQL Server'
            Detected = $true
            Notes = 'SQL Server registry keys detected'
        }
    }

    if (-not $results) {
        $results += [pscustomobject]@{
            ComputerName = $env:COMPUTERNAME
            Engine = 'Unknown'
            Detected = $false
            Notes = 'No common database engine registry keys detected'
        }
    }

    $results
}

function Initialize-OutputRoot {
    param(
        [string]$Root,
        [int]$RetentionDays
    )

    if (-not (Test-Path -LiteralPath $Root)) {
        New-Item -ItemType Directory -Path $Root -Force | Out-Null
    }

    $cutoff = (Get-Date).AddDays(-1 * $RetentionDays)
    Get-ChildItem -LiteralPath $Root -File -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -lt $cutoff } |
        Remove-Item -Force -ErrorAction SilentlyContinue
}

function Export-AuditWorkbook {
    param(
        [string]$WorkbookPath,
        [object]$Summary,
        [object[]]$Drives,
        [object[]]$Paths,
        [object[]]$Sites,
        [object[]]$Databases
    )

    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Warning 'ImportExcel module not found. Skipping Excel export.'
        return
    }

    Import-Module ImportExcel -ErrorAction Stop

    $sheetSummary = @($Summary.Machine)
    $sheetSummary | Export-Excel -Path $WorkbookPath -WorksheetName 'Machine' -AutoSize -ClearSheet
    $Drives | Export-Excel -Path $WorkbookPath -WorksheetName 'Drives' -AutoSize -ClearSheet
    $Paths | Export-Excel -Path $WorkbookPath -WorksheetName 'Paths' -AutoSize -ClearSheet
    $Sites | Export-Excel -Path $WorkbookPath -WorksheetName 'IIS Sites' -AutoSize -ClearSheet
    $Databases | Export-Excel -Path $WorkbookPath -WorksheetName 'Databases' -AutoSize -ClearSheet
}

Initialize-OutputRoot -Root $OutputRoot -RetentionDays $KeepDays

$stamp = Get-Date -Format 'yyyy-MM-dd_HHmmss'
$summaryPath = Join-Path $OutputRoot "resource_audit_summary_$stamp.json"
$drivesPath = Join-Path $OutputRoot "resource_audit_drives_$stamp.csv"
$pathsPath = Join-Path $OutputRoot "resource_audit_paths_$stamp.csv"
$sitesPath = Join-Path $OutputRoot "resource_audit_sites_$stamp.csv"
$dbPath = Join-Path $OutputRoot "resource_audit_databases_$stamp.csv"

if (-not $ExcelPath) {
    $ExcelPath = Join-Path $OutputRoot "resource_audit_$stamp.xlsx"
}

$machine = Get-MachineTypeInfo
$drives = @(Get-FixedDriveAudit)
$sites = @(Get-IisAudit)
$databases = @(Get-DatabaseFootprint)

$pathAudit = foreach ($path in $AuditPaths) {
    if (Test-Path -LiteralPath $path) {
        Get-DirectorySize -Path $path
    }
    else {
        [pscustomobject]@{
            ComputerName = $env:COMPUTERNAME
            Path = $path
            FileCount = $null
            SizeGB = $null
            Error = 'Path not found'
        }
    }
}

$summary = [pscustomobject]@{
    GeneratedAt = (Get-Date).ToString('s')
    ComputerName = $env:COMPUTERNAME
    Machine = $machine
    DriveCount = @($drives).Count
    AuditedPathCount = @($pathAudit).Count
    IisSiteCount = @($sites).Count
    DatabaseSignals = @($databases).Count
    ExcelExportRequested = [bool]$ExportExcel
}

$summary | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $summaryPath -Encoding UTF8
$drives | Export-Csv -LiteralPath $drivesPath -NoTypeInformation -Encoding UTF8
$pathAudit | Export-Csv -LiteralPath $pathsPath -NoTypeInformation -Encoding UTF8
$sites | Export-Csv -LiteralPath $sitesPath -NoTypeInformation -Encoding UTF8
$databases | Export-Csv -LiteralPath $dbPath -NoTypeInformation -Encoding UTF8

if ($ExportExcel) {
    Export-AuditWorkbook -WorkbookPath $ExcelPath -Summary $summary -Drives $drives -Paths $pathAudit -Sites $sites -Databases $databases
}

Write-Host "Resource audit complete."
Write-Host "Summary:   $summaryPath"
Write-Host "Drives:    $drivesPath"
Write-Host "Paths:     $pathsPath"
Write-Host "IIS Sites: $sitesPath"
Write-Host "Databases: $dbPath"
if ($ExportExcel) {
    Write-Host "Workbook:  $ExcelPath"
}
