<#
Opens the Vesta dashboard with the CSVs that sit next to it, already refreshed.

  1. Picks the CSV folder: data_private\ if it holds both CSVs, otherwise data\.
  2. Points the DataFolder parameter at that folder (absolute paths differ per computer).
  3. Drops the stale data cache, opens VestaDashboard.pbip in Power BI Desktop.
  4. Waits for the model to load and refreshes it through Power BI's local engine.

Run via "Open Dashboard.bat" (double-click). Close any open copy of the dashboard first.
#>
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$names = "vesta_Document.csv", "vesta_DocumentVersion.csv"

# ---------------------------------------------------------------- 1. which CSVs
$folder = Join-Path $root "data_private"
$hasPrivate = ($names | Where-Object { Test-Path (Join-Path $folder $_) }).Count -eq 2
if (-not $hasPrivate) { $folder = Join-Path $root "data" }
foreach ($n in $names) {
    if (-not (Test-Path (Join-Path $folder $n))) { throw "Missing $n in $folder" }
}
Write-Host "Using CSVs from: $folder"

# ---------------------------------------------------------------- 2. DataFolder
$tmdl = Join-Path $root "VestaDashboard.SemanticModel\definition\expressions.tmdl"
$text = [IO.File]::ReadAllText($tmdl)
$value = $folder.TrimEnd('\') + '\'
$text = [regex]::Replace($text, '(expression DataFolder = )"[^"]*"', { param($m) $m.Groups[1].Value + '"' + $value + '"' }, 1)
[IO.File]::WriteAllText($tmdl, $text, (New-Object Text.UTF8Encoding $false))

# ---------------------------------------------------------------- 3. open
$cache = Join-Path $root "VestaDashboard.SemanticModel\.pbi\cache.abf"
if (Test-Path $cache) { Remove-Item $cache -Force }
$pkg = Get-AppxPackage *MicrosoftPowerBIDesktop* | Select-Object -First 1
$bin = if ($pkg) { Join-Path $pkg.InstallLocation "bin" } else { "C:\Program Files\Microsoft Power BI Desktop\bin" }
$started = Get-Date
Write-Host "Opening Power BI Desktop..."
Start-Process (Join-Path $root "VestaDashboard.pbip")

# ---------------------------------------------------------------- 4. refresh
Add-Type -Path (Join-Path $bin "Microsoft.PowerBI.Tabular.dll")
$deadline = (Get-Date).AddMinutes(5)
$server = $null
while ((Get-Date) -lt $deadline -and -not $server) {
    Start-Sleep -Seconds 3
    $procs = Get-Process msmdsrv -ErrorAction SilentlyContinue | Where-Object { $_.StartTime -gt $started.AddSeconds(-5) }
    foreach ($p in $procs) {
        $port = (Get-NetTCPConnection -OwningProcess $p.Id -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1).LocalPort
        if (-not $port) { continue }
        try {
            $s = New-Object Microsoft.AnalysisServices.Tabular.Server
            $s.Connect("Data Source=localhost:$port")
            if ($s.Databases.Count -gt 0 -and $s.Databases[0].Model.Tables.Find("vesta_DocumentVersion")) { $server = $s }
            else { $s.Disconnect() }
        } catch { }
    }
}
if (-not $server) { throw "Power BI did not finish loading the model. Click Refresh in Power BI instead." }

Start-Sleep -Seconds 5  # let Desktop finish wiring up the report
Write-Host "Refreshing data..."
$model = $server.Databases[0].Model
$model.RequestRefresh([Microsoft.AnalysisServices.Tabular.RefreshType]::Full)
$null = $model.SaveChanges()
$server.Disconnect()
Write-Host "Done - the dashboard now shows the CSVs from $folder"
