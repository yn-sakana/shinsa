Import-Module (Join-Path $PSScriptRoot 'Common.psm1') -Force -DisableNameChecking

$ErrorActionPreference = 'Stop'
$config = Get-ShinsaConfig -ScriptPath $MyInvocation.MyCommand.Path
$paths = Get-ShinsaDataPaths -Config $config

Ensure-ShinsaState -Paths $paths

$sourceNames = Get-ShinsaSourceNames -Config $config
$cache = Read-ShinsaCache -Paths $paths
$results = @()
$allConflicts = @()

foreach ($name in $sourceNames) {
    $src = $config.sources[$name]
    $sourcePath = Get-ShinsaSourcePath -Config $config -SourceName $name

    if ([string]::IsNullOrWhiteSpace($sourcePath) -or -not (Test-Path $sourcePath)) {
        Write-Host ("  skip {0}: source not found" -f $name) -ForegroundColor Yellow
        continue
    }

    try {
        $remoteRecords = @(Import-ShinsaSourceRecords -Config $config -SourceName $name)
    } catch {
        Write-Host ("  skip {0}: {1}" -f $name, $_.Exception.Message) -ForegroundColor Yellow
        continue
    }

    $jsonPath = Join-Path $paths.JsonRoot $src.file
    $snapshotPath = ($jsonPath -replace '\.json$', '.snapshot.json')

    $srcMap = ConvertTo-ShinsaMap -InputObject $src
    $view = if ($srcMap.Contains('view')) { [string]$srcMap['view'] } else { 'fields' }
    $keyColumn = if ($null -ne $src.key_column) { [string]$src.key_column } else { '' }
    $editableCols = @(Get-ShinsaEditableColumnNames -Config $config -Cache $cache -SourceName $name)

    $doMerge = ($view -eq 'fields' -and $editableCols.Count -gt 0 -and
                -not [string]::IsNullOrWhiteSpace($keyColumn) -and
                (Test-Path $jsonPath) -and (Test-Path $snapshotPath))

    if ($doMerge) {
        # Smart sync: 3-way merge
        $localRecords = @(Read-ShinsaJson -Path $jsonPath)
        $snapshotRecords = @(Read-ShinsaJson -Path $snapshotPath)

        $mergeResult = Merge-ShinsaSourceRecords `
            -KeyColumn $keyColumn `
            -EditableColumns $editableCols `
            -RemoteRecords $remoteRecords `
            -LocalRecords $localRecords `
            -SnapshotRecords $snapshotRecords

        Write-ShinsaJson -Path $jsonPath -Data $mergeResult.merged

        # Write changelog
        $logPath = Join-Path $paths.JsonRoot 'changelog.jsonl'
        foreach ($entry in $mergeResult.log_entries) {
            Write-ShinsaChangeLog -LogPath $logPath -SourceName $name `
                -KeyValue $entry.key -FieldName $entry.field `
                -OldValue $entry.old -NewValue $entry.new -Origin $entry.origin
        }

        $allConflicts += @($mergeResult.conflicts | ForEach-Object {
            $_ | Add-Member -NotePropertyName 'source_name' -NotePropertyValue $name -PassThru
        })
    } else {
        # Simple overwrite (first sync, or non-editable source)
        Write-ShinsaJson -Path $jsonPath -Data $remoteRecords
    }

    # Always update snapshot to latest remote
    Write-ShinsaJson -Path $snapshotPath -Data $remoteRecords

    $results += [pscustomobject]@{ name = $name; count = $remoteRecords.Count }
}

$uiState = ConvertTo-ShinsaMap -InputObject $cache.ui_state
$uiState['last_sync_at'] = (Get-Date).ToString('o')
$cache.ui_state = [pscustomobject]$uiState
Save-ShinsaCache -Paths $paths -Cache $cache

Write-Host ''
Write-Host 'sync completed' -ForegroundColor Cyan
foreach ($r in $results) {
    Write-Host ("  {0}: {1} records" -f $r.name, $r.count)
}
Write-Host ("  json: {0}" -f $paths.JsonRoot)

# Return conflicts for GUI caller
if ($allConflicts.Count -gt 0) {
    $allConflicts
}
