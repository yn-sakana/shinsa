Set-StrictMode -Version Latest

# Load EPPlus for Excel I/O (no COM, no locks)
$epplusPath = Join-Path $PSScriptRoot '..\lib\EPPlus.dll'
if (Test-Path $epplusPath) {
    Add-Type -Path $epplusPath -ErrorAction SilentlyContinue
}

# =============================================================================
# App Root
# =============================================================================

function Get-ShinsaAppRoot {
    param([Parameter(Mandatory = $true)][string]$ScriptPath)

    $scriptDirectory = Split-Path -Parent $ScriptPath
    if (Test-Path (Join-Path $scriptDirectory 'config\config.base.json')) {
        return $scriptDirectory
    }

    Split-Path -Parent $scriptDirectory
}

# =============================================================================
# Config utilities
# =============================================================================

function ConvertTo-ShinsaConfigValue {
    param($Value)

    if ($null -eq $Value) { return $null }

    if ($Value -is [System.Management.Automation.PSCustomObject] -or $Value -is [System.Collections.IDictionary]) {
        return ConvertTo-ShinsaMap -InputObject $Value
    }

    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        return @($Value | ForEach-Object { ConvertTo-ShinsaConfigValue -Value $_ })
    }

    $Value
}

function ConvertTo-ShinsaMap {
    param($InputObject)

    $map = [ordered]@{}
    if ($null -eq $InputObject) { return $map }

    if ($InputObject -is [System.Collections.IDictionary]) {
        foreach ($key in $InputObject.Keys) {
            $map[[string]$key] = ConvertTo-ShinsaConfigValue -Value $InputObject[$key]
        }
        return $map
    }

    foreach ($property in $InputObject.PSObject.Properties) {
        $map[$property.Name] = ConvertTo-ShinsaConfigValue -Value $property.Value
    }

    $map
}

function Merge-ShinsaMap {
    param(
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$Base,
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$Overlay
    )

    foreach ($key in $Overlay.Keys) {
        if (
            $Base.Contains($key) -and
            $Base[$key] -is [System.Collections.IDictionary] -and
            $Overlay[$key] -is [System.Collections.IDictionary]
        ) {
            Merge-ShinsaMap -Base $Base[$key] -Overlay $Overlay[$key]
            continue
        }

        $Base[$key] = ConvertTo-ShinsaConfigValue -Value $Overlay[$key]
    }
}

# =============================================================================
# JSON I/O
# =============================================================================

function Read-ShinsaJson {
    param([Parameter(Mandatory = $true)][string]$Path)

    Get-Content -Path $Path -Raw -Encoding UTF8 | ConvertFrom-Json
}

function Write-ShinsaJson {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)]$Data
    )

    $directory = Split-Path -Parent $Path
    if (-not [string]::IsNullOrWhiteSpace($directory) -and -not (Test-Path $directory)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }

    if (
        $Data -is [System.Collections.IEnumerable] -and
        -not ($Data -is [string]) -and
        -not ($Data -is [System.Collections.IDictionary]) -and
        -not ($Data -is [System.Management.Automation.PSCustomObject])
    ) {
        $items = @($Data)
        if ($items.Count -eq 0) {
            $json = '[]'
        }
        else {
            $json = "[`r`n" + (($items | ForEach-Object { $_ | ConvertTo-Json -Depth 16 }) -join ",`r`n") + "`r`n]"
        }
    }
    else {
        $json = ConvertTo-Json -InputObject $Data -Depth 16
    }

    Set-Content -Path $Path -Value $json -Encoding UTF8
}

# =============================================================================
# Path utilities
# =============================================================================

function Resolve-ShinsaPath {
    param(
        [Parameter(Mandatory = $true)][string]$AppRoot,
        [string]$PathValue
    )

    if ([string]::IsNullOrWhiteSpace($PathValue)) { return '' }

    if ([System.IO.Path]::IsPathRooted($PathValue)) {
        return [System.IO.Path]::GetFullPath($PathValue)
    }

    [System.IO.Path]::GetFullPath((Join-Path $AppRoot $PathValue))
}

function Ensure-ShinsaDirectory {
    param([string[]]$Paths)

    foreach ($path in @($Paths)) {
        if ([string]::IsNullOrWhiteSpace($path)) { continue }
        if (-not (Test-Path $path)) {
            New-Item -ItemType Directory -Path $path -Force | Out-Null
        }
    }
}

# =============================================================================
# Config
# =============================================================================

function Get-ShinsaConfig {
    param([Parameter(Mandatory = $true)][string]$ScriptPath)

    $appRoot = Get-ShinsaAppRoot -ScriptPath $ScriptPath
    $basePath = Join-Path $appRoot 'config\config.base.json'
    $localPath = Join-Path $appRoot 'config\config.local.json'
    $samplePath = Join-Path $appRoot 'config\config.local.sample.json'

    if (-not (Test-Path $localPath)) {
        Copy-Item -Path $samplePath -Destination $localPath -Force
    }

    $baseConfig = ConvertTo-ShinsaMap -InputObject (Read-ShinsaJson -Path $basePath)
    $localConfig = ConvertTo-ShinsaMap -InputObject (Read-ShinsaJson -Path $localPath)
    Merge-ShinsaMap -Base $baseConfig -Overlay $localConfig
    $baseConfig['app_root'] = $appRoot

    [pscustomobject]$baseConfig
}

function Get-ShinsaDataPaths {
    param([Parameter(Mandatory = $true)]$Config)

    $appRoot = [string]$Config.app_root
    $jsonRoot = Resolve-ShinsaPath -AppRoot $appRoot -PathValue $Config.paths.json_root

    [pscustomobject]@{
        AppRoot = $appRoot
        JsonRoot = $jsonRoot
        CacheJsonPath = Join-Path $jsonRoot 'cache.json'
    }
}

function Get-ShinsaSourcePath {
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)][string]$SourceName
    )

    $src = $Config.sources[$SourceName]
    if ($null -eq $src) { return '' }
    Resolve-ShinsaPath -AppRoot $Config.app_root -PathValue $src.source_path
}

function Get-ShinsaSourceNames {
    param([Parameter(Mandatory = $true)]$Config)

    if ($Config.sources -is [System.Collections.IDictionary]) {
        return @($Config.sources.Keys)
    }

    @($Config.sources.PSObject.Properties | ForEach-Object { $_.Name })
}

function Test-ShinsaSourceHasJoin {
    param($SourceConfig)

    if ($null -eq $SourceConfig) { return $false }
    if ($SourceConfig -is [System.Collections.IDictionary]) {
        return $SourceConfig.Contains('join') -and $null -ne $SourceConfig['join']
    }
    return ($SourceConfig.PSObject.Properties.Name -contains 'join') -and $null -ne $SourceConfig.join
}

function Get-ShinsaPrimarySourceNames {
    param([Parameter(Mandatory = $true)]$Config)

    $names = Get-ShinsaSourceNames -Config $Config
    @($names | Where-Object {
        $src = $Config.sources[$_]
        -not (Test-ShinsaSourceHasJoin -SourceConfig $src)
    })
}

function Get-ShinsaJoinedSourceNames {
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)][string]$PrimaryName
    )

    $names = Get-ShinsaSourceNames -Config $Config
    @($names | Where-Object {
        $src = $Config.sources[$_]
        (Test-ShinsaSourceHasJoin -SourceConfig $src) -and [string]$src.join.source -eq $PrimaryName
    })
}

# =============================================================================
# Cache
# =============================================================================

function Get-DefaultCacheState {
    [ordered]@{
        ui_state = [ordered]@{}
        field_settings = [ordered]@{}
    }
}

function Ensure-ShinsaCacheShape {
    param($Cache)

    $normalized = ConvertTo-ShinsaMap -InputObject $Cache

    if (-not $normalized.Contains('ui_state') -or -not ($normalized['ui_state'] -is [System.Collections.IDictionary])) {
        $normalized['ui_state'] = [ordered]@{}
    }
    if (-not $normalized.Contains('field_settings') -or -not ($normalized['field_settings'] -is [System.Collections.IDictionary])) {
        $normalized['field_settings'] = [ordered]@{}
    }

    [pscustomobject]$normalized
}

function Get-ShinsaFieldSettings {
    param(
        [Parameter(Mandatory)]$Cache,
        [Parameter(Mandatory)][string]$SourceName
    )
    $fs = ConvertTo-ShinsaMap -InputObject $Cache.field_settings
    if ($fs.Contains($SourceName)) {
        return ConvertTo-ShinsaMap -InputObject $fs[$SourceName]
    }
    return [ordered]@{}
}

function Set-ShinsaFieldSettings {
    param(
        [Parameter(Mandatory)]$Cache,
        [Parameter(Mandatory)][string]$SourceName,
        [Parameter(Mandatory)]$Settings
    )
    $fs = ConvertTo-ShinsaMap -InputObject $Cache.field_settings
    $fs[$SourceName] = [pscustomobject]$Settings
    $Cache | Add-Member -MemberType NoteProperty -Name 'field_settings' -Value ([pscustomobject]$fs) -Force
}

function Get-ShinsaFieldTypeGuess {
    param([string]$FieldName, [string[]]$SampleValues)

    # Name-based heuristics
    if ($FieldName -match '(?i)(date|日|期限|年月)') { return 'date' }
    if ($FieldName -match '(?i)(amount|fee|cost|price|金額|費|額|率)') { return 'number' }

    # Value-based heuristics
    $nonEmpty = @($SampleValues | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($nonEmpty.Count -eq 0) { return 'text' }

    $dateCount = 0; $numCount = 0
    foreach ($v in $nonEmpty) {
        $d = [datetime]::MinValue
        if ([datetime]::TryParse($v, [ref]$d)) { $dateCount++ }
        $n = 0.0
        if ([double]::TryParse($v, [ref]$n)) { $numCount++ }
    }
    if ($dateCount -gt ($nonEmpty.Count * 0.5)) { return 'date' }
    if ($numCount -gt ($nonEmpty.Count * 0.7)) { return 'number' }
    return 'text'
}

function Initialize-ShinsaFieldSettings {
    param(
        [Parameter(Mandatory)]$Config,
        [Parameter(Mandatory)]$Cache,
        [Parameter(Mandatory)][string]$SourceName,
        [Parameter(Mandatory)]$Records
    )
    $srcMap = ConvertTo-ShinsaMap -InputObject $Config.sources[$SourceName]
    $existing = Get-ShinsaFieldSettings -Cache $Cache -SourceName $SourceName
    if ($existing.Count -gt 0) { return $existing }

    $settings = [ordered]@{}

    # Get field names from config or data
    $fields = @()
    if ($srcMap.Contains('columns')) {
        $fields = @((ConvertTo-ShinsaMap -InputObject $srcMap['columns']).Keys)
    }
    if ($fields.Count -eq 0 -and $Records.Count -gt 0) {
        $fields = @($Records[0].PSObject.Properties | Where-Object { $_.Name -notlike '_*' } | ForEach-Object { $_.Name })
    }

    # Config-based defaults for display/multiline
    $configDisplay = @()
    $configMultiline = @()
    if ($srcMap.Contains('display_columns')) { $configDisplay = @($srcMap['display_columns']) }
    if ($srcMap.Contains('multiline_columns')) { $configMultiline = @($srcMap['multiline_columns']) }

    foreach ($fn in $fields) {
        $samples = @($Records | ForEach-Object { ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $_ -Name $fn) } | Select-Object -First 10)
        $guessedType = Get-ShinsaFieldTypeGuess -FieldName $fn -SampleValues $samples
        $inList = if ($configDisplay.Count -gt 0) { $fn -in $configDisplay } else { $false }
        $multiline = $fn -in $configMultiline

        $settings[$fn] = [ordered]@{
            type = $guessedType
            in_list = $inList
            editable = $true
            multiline = $multiline
        }
    }

    # If no display columns in config, default to first 3-4 fields
    if ($configDisplay.Count -eq 0 -and $fields.Count -gt 0) {
        $defaultShow = [Math]::Min(4, $fields.Count)
        for ($i = 0; $i -lt $defaultShow; $i++) {
            $settings[$fields[$i]].in_list = $true
        }
    }

    Set-ShinsaFieldSettings -Cache $Cache -SourceName $SourceName -Settings $settings
    return $settings
}

function Ensure-ShinsaState {
    param([Parameter(Mandatory = $true)]$Paths)

    Ensure-ShinsaDirectory -Paths @($Paths.JsonRoot)
    if (-not (Test-Path $Paths.CacheJsonPath)) {
        Write-ShinsaJson -Path $Paths.CacheJsonPath -Data (Get-DefaultCacheState)
    }
}

function Read-ShinsaCache {
    param([Parameter(Mandatory = $true)]$Paths)

    Ensure-ShinsaState -Paths $Paths
    Ensure-ShinsaCacheShape -Cache (Read-ShinsaJson -Path $Paths.CacheJsonPath)
}

function Save-ShinsaCache {
    param(
        [Parameter(Mandatory = $true)]$Paths,
        [Parameter(Mandatory = $true)]$Cache
    )

    Write-ShinsaJson -Path $Paths.CacheJsonPath -Data (Ensure-ShinsaCacheShape -Cache $Cache)
}

# =============================================================================
# Record utilities
# =============================================================================

function ConvertTo-ShinsaString {
    param($Value)

    if ($null -eq $Value) { return '' }

    if ($Value -is [datetime]) {
        return $Value.ToString('o')
    }

    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        return (@($Value | ForEach-Object { ConvertTo-ShinsaString -Value $_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join ', ')
    }

    [string]$Value
}

function Test-ShinsaRecordProperty {
    param(
        [Parameter(Mandatory = $true)]$Record,
        [Parameter(Mandatory = $true)][string]$Name
    )

    if ($Record -is [System.Collections.IDictionary]) {
        return $Record.Contains($Name)
    }

    return $Record.PSObject.Properties.Name -contains $Name
}

function Get-ShinsaRecordValue {
    param(
        [Parameter(Mandatory = $true)]$Record,
        [Parameter(Mandatory = $true)][string]$Name
    )

    if ($Record -is [System.Collections.IDictionary]) {
        if ($Record.Contains($Name)) { return $Record[$Name] }
        return $null
    }

    $property = $Record.PSObject.Properties[$Name]
    if ($null -eq $property) { return $null }

    $property.Value
}

function Set-ShinsaRecordValue {
    param(
        [Parameter(Mandatory = $true)]$Record,
        [Parameter(Mandatory = $true)][string]$Name,
        $Value
    )

    if ($Record -is [System.Collections.IDictionary]) {
        $Record[$Name] = $Value
        return
    }

    if ($Record.PSObject.Properties.Name -contains $Name) {
        $Record.$Name = $Value
    }
    else {
        $Record | Add-Member -NotePropertyName $Name -NotePropertyValue $Value
    }
}

function Copy-ShinsaRecord {
    param([Parameter(Mandatory = $true)]$Record)

    [pscustomobject](ConvertTo-ShinsaMap -InputObject $Record)
}

# =============================================================================
# Table / Fields source import
# =============================================================================

function ConvertFrom-ExcelCellValue {
    param(
        $Value,
        [string]$NumberFormat = '',
        [string]$FieldType = ''
    )

    if ($null -eq $Value) { return '' }

    # Detect date by Excel NumberFormat or explicit field type
    $isDate = $false
    if ($Value -is [double]) {
        if ($FieldType -eq 'date') { $isDate = $true }
        elseif (-not [string]::IsNullOrWhiteSpace($NumberFormat)) {
            $nf = $NumberFormat.ToLowerInvariant()
            if ($nf -ne 'general' -and $nf -ne '0' -and $nf -ne '@' -and
                ($nf -match '[ymd]' -or $nf -match 'date' -or $nf -match '[ghse]')) {
                $isDate = $true
            }
        }
    }

    if ($isDate) {
        try { return ([datetime]::FromOADate($Value)).ToString('yyyy-MM-dd') } catch { }
    }

    if ($Value -is [double] -and [math]::Floor($Value) -eq $Value) {
        return [string][int64]$Value
    }

    [string]$Value
}

function Convert-ShinsaSourceRow {
    param(
        [Parameter(Mandatory = $true)]$Row,
        [Parameter(Mandatory = $true)]$SourceConfig,
        [Parameter(Mandatory = $true)][string]$SourcePath,
        [string]$SheetName = '',
        [int]$RowId = 0
    )

    $columnsMap = if ($null -ne $SourceConfig.columns) { ConvertTo-ShinsaMap -InputObject $SourceConfig.columns } else { $null }
    $keyColumn = if ($null -ne $SourceConfig.key_column) { [string]$SourceConfig.key_column } else { '' }

    $record = [ordered]@{}

    foreach ($propertyName in (ConvertTo-ShinsaMap -InputObject $Row).Keys) {
        $record[$propertyName] = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $Row -Name $propertyName)
    }

    if ($null -ne $columnsMap) {
        foreach ($logicalName in $columnsMap.Keys) {
            $sourceName = [string]$columnsMap[$logicalName]
            if (Test-ShinsaRecordProperty -Record $Row -Name $sourceName) {
                $record[$logicalName] = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $Row -Name $sourceName)
            }
            elseif (Test-ShinsaRecordProperty -Record $Row -Name $logicalName) {
                $record[$logicalName] = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $Row -Name $logicalName)
            }
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($keyColumn)) {
        $keyValue = ''
        if ($record.Contains($keyColumn)) {
            $keyValue = ConvertTo-ShinsaString -Value $record[$keyColumn]
        }
        if ([string]::IsNullOrWhiteSpace($keyValue)) {
            return $null
        }
    }

    $record['_source_path'] = $SourcePath
    $record['_source_sheet'] = $SheetName
    $record['_source_row_id'] = $RowId

    [pscustomobject]$record
}

function Import-ShinsaFieldsFromJson {
    param(
        [Parameter(Mandatory = $true)]$SourceConfig,
        [Parameter(Mandatory = $true)][string]$SourcePath
    )

    $source = Read-ShinsaJson -Path $SourcePath
    if ($source -is [System.Collections.IEnumerable] -and -not ($source -is [string]) -and -not ($source -is [System.Management.Automation.PSCustomObject])) {
        $rows = @($source)
    }
    elseif (Test-ShinsaRecordProperty -Record $source -Name 'rows') {
        $rows = @($source.rows)
    }
    else {
        $rows = @($source)
    }

    $records = @()
    $rowIndex = 2
    foreach ($row in $rows) {
        $record = Convert-ShinsaSourceRow -Row $row -SourceConfig $SourceConfig -SourcePath $SourcePath -RowId $rowIndex
        if ($null -ne $record) { $records += $record }
        $rowIndex += 1
    }

    $records
}

function Import-ShinsaFieldsFromExcel {
    param(
        [Parameter(Mandatory = $true)]$SourceConfig,
        [Parameter(Mandatory = $true)][string]$SourcePath
    )

    $sourceTable = if ($null -ne $SourceConfig.source_table) { [string]$SourceConfig.source_table } else { '' }

    # Open with EPPlus (read-only, no locks)
    $stream = [System.IO.File]::Open($SourcePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    $package = $null
    try {
        $package = New-Object OfficeOpenXml.ExcelPackage($stream)

        if (-not [string]::IsNullOrWhiteSpace($sourceTable)) {
            # Find structured table by name
            $tbl = $null
            $ws = $null
            foreach ($worksheet in $package.Workbook.Worksheets) {
                foreach ($t in $worksheet.Tables) {
                    if ($t.Name -eq $sourceTable) { $tbl = $t; $ws = $worksheet; break }
                }
                if ($null -ne $tbl) { break }
            }
            if ($null -eq $tbl) { throw "Structured table '$sourceTable' not found in '$SourcePath'." }

            $sheetName = $ws.Name
            $addr = $tbl.Address
            $startRow = $addr.Start.Row
            $startCol = $addr.Start.Column
            $endRow = $addr.End.Row
            $endCol = $addr.End.Column

            # Header row
            $headers = @{}
            for ($c = $startCol; $c -le $endCol; $c++) {
                $hv = $ws.Cells[$startRow, $c].Text
                if (-not [string]::IsNullOrWhiteSpace($hv)) { $headers[$c] = $hv.Trim() }
            }

            $records = @()
            for ($r = ($startRow + 1); $r -le $endRow; $r++) {
                $row = [ordered]@{}
                $hasValue = $false
                foreach ($c in $headers.Keys) {
                    $cell = $ws.Cells[$r, $c]
                    $cellValue = ConvertFrom-ExcelCellValue -Value $cell.Value -NumberFormat ([string]$cell.Style.Numberformat.Format)
                    if (-not [string]::IsNullOrWhiteSpace($cellValue)) { $hasValue = $true }
                    $row[$headers[$c]] = $cellValue
                }
                if (-not $hasValue) { continue }

                $record = Convert-ShinsaSourceRow -Row ([pscustomobject]$row) -SourceConfig $SourceConfig -SourcePath $SourcePath -SheetName $sheetName -RowId $r
                if ($null -ne $record) { $records += $record }
            }
            return $records
        }
        else {
            # Fallback: first sheet, first row as header
            $ws = $package.Workbook.Worksheets[1]
            if ($null -eq $ws.Dimension) { return @() }
            $startRow = $ws.Dimension.Start.Row
            $startCol = $ws.Dimension.Start.Column
            $endRow = $ws.Dimension.End.Row
            $endCol = $ws.Dimension.End.Column

            $headers = @{}
            for ($c = $startCol; $c -le $endCol; $c++) {
                $hv = $ws.Cells[$startRow, $c].Text
                if (-not [string]::IsNullOrWhiteSpace($hv)) { $headers[$c] = $hv.Trim() }
            }

            $records = @()
            for ($r = ($startRow + 1); $r -le $endRow; $r++) {
                $row = [ordered]@{}
                $hasValue = $false
                foreach ($c in $headers.Keys) {
                    $cell = $ws.Cells[$r, $c]
                    $cellValue = ConvertFrom-ExcelCellValue -Value $cell.Value -NumberFormat ([string]$cell.Style.Numberformat.Format)
                    if (-not [string]::IsNullOrWhiteSpace($cellValue)) { $hasValue = $true }
                    $row[$headers[$c]] = $cellValue
                }
                if (-not $hasValue) { continue }

                $record = Convert-ShinsaSourceRow -Row ([pscustomobject]$row) -SourceConfig $SourceConfig -SourcePath $SourcePath -SheetName $ws.Name -RowId $r
                if ($null -ne $record) { $records += $record }
            }
            return $records
        }
    }
    finally {
        if ($null -ne $package) { $package.Dispose() }
        if ($null -ne $stream) { $stream.Dispose() }
    }
}

function Import-ShinsaFieldsRecords {
    param(
        [Parameter(Mandatory = $true)]$SourceConfig,
        [Parameter(Mandatory = $true)][string]$SourcePath
    )

    $extension = [System.IO.Path]::GetExtension($SourcePath).ToLowerInvariant()
    $keyColumn = if ($null -ne $SourceConfig.key_column) { [string]$SourceConfig.key_column } else { '' }

    $records = switch ($extension) {
        '.json' { Import-ShinsaFieldsFromJson -SourceConfig $SourceConfig -SourcePath $SourcePath }
        '.xlsx'  { Import-ShinsaFieldsFromExcel -SourceConfig $SourceConfig -SourcePath $SourcePath }
        '.xlsm'  { Import-ShinsaFieldsFromExcel -SourceConfig $SourceConfig -SourcePath $SourcePath }
        '.xlsb'  { Import-ShinsaFieldsFromExcel -SourceConfig $SourceConfig -SourcePath $SourcePath }
        '.xls'   { Import-ShinsaFieldsFromExcel -SourceConfig $SourceConfig -SourcePath $SourcePath }
        default  { throw "Unsupported source format: $extension" }
    }

    if (-not [string]::IsNullOrWhiteSpace($keyColumn)) {
        @($records | Sort-Object $keyColumn)
    }
    else {
        @($records)
    }
}

# =============================================================================
# Mail source import
# =============================================================================

function Resolve-ShinsaArchiveItemPath {
    param(
        [Parameter(Mandatory = $true)][string]$MailDirectory,
        [string]$RelativePath
    )

    if ([string]::IsNullOrWhiteSpace($RelativePath)) { return '' }

    if ([System.IO.Path]::IsPathRooted($RelativePath)) {
        return [System.IO.Path]::GetFullPath($RelativePath)
    }

    $candidate = Join-Path $MailDirectory $RelativePath
    if (Test-Path $candidate) {
        return [System.IO.Path]::GetFullPath($candidate)
    }

    $attachmentsCandidate = Join-Path (Join-Path $MailDirectory 'attachments') $RelativePath
    if (Test-Path $attachmentsCandidate) {
        return [System.IO.Path]::GetFullPath($attachmentsCandidate)
    }

    [System.IO.Path]::GetFullPath($candidate)
}

function Import-ShinsaMailRecords {
    param([Parameter(Mandatory = $true)][string]$SourcePath)

    if (-not (Test-Path $SourcePath)) { return @() }

    $manifestFiles = @(Get-ChildItem -Path $SourcePath -Recurse -File -Include '*.json' | Where-Object {
        $_.Name -eq 'meta.json' -or $_.Name -like 'mail_*.json'
    })

    $records = @()
    foreach ($manifestFile in $manifestFiles) {
        $manifest = Read-ShinsaJson -Path $manifestFile.FullName
        if (-not (Test-ShinsaRecordProperty -Record $manifest -Name 'mail_id')) { continue }

        $mailDirectory = Split-Path -Parent $manifestFile.FullName
        $attachmentSource = @()
        if (Test-ShinsaRecordProperty -Record $manifest -Name 'attachment_paths') {
            $attachmentSource = @($manifest.attachment_paths)
        }
        elseif (Test-ShinsaRecordProperty -Record $manifest -Name 'attachments') {
            $attachmentSource = @($manifest.attachments)
        }

        $attachmentPaths = @()
        foreach ($attachment in $attachmentSource) {
            if ($attachment -is [string]) {
                $attachmentPaths += Resolve-ShinsaArchiveItemPath -MailDirectory $mailDirectory -RelativePath $attachment
                continue
            }
            if (Test-ShinsaRecordProperty -Record $attachment -Name 'path') {
                $attachmentPaths += Resolve-ShinsaArchiveItemPath -MailDirectory $mailDirectory -RelativePath ([string]$attachment.path)
            }
        }

        if (Test-ShinsaRecordProperty -Record $manifest -Name 'folder_path') {
            $folderPath = ConvertTo-ShinsaString -Value $manifest.folder_path
        }
        else {
            $folderPath = $mailDirectory.Substring($SourcePath.Length).TrimStart('\')
        }

        $records += [pscustomobject]@{
            mail_id = ConvertTo-ShinsaString -Value $manifest.mail_id
            entry_id = if (Test-ShinsaRecordProperty -Record $manifest -Name 'entry_id') { ConvertTo-ShinsaString -Value $manifest.entry_id } else { ConvertTo-ShinsaString -Value $manifest.mail_id }
            mailbox_address = if (Test-ShinsaRecordProperty -Record $manifest -Name 'mailbox_address') { ConvertTo-ShinsaString -Value $manifest.mailbox_address } else { '' }
            folder_path = $folderPath
            received_at = if (Test-ShinsaRecordProperty -Record $manifest -Name 'received_at') { ConvertTo-ShinsaString -Value $manifest.received_at } else { '' }
            sender_name = if (Test-ShinsaRecordProperty -Record $manifest -Name 'sender_name') { ConvertTo-ShinsaString -Value $manifest.sender_name } else { '' }
            sender_email = if (Test-ShinsaRecordProperty -Record $manifest -Name 'sender_email') { ConvertTo-ShinsaString -Value $manifest.sender_email } else { '' }
            subject = if (Test-ShinsaRecordProperty -Record $manifest -Name 'subject') { ConvertTo-ShinsaString -Value $manifest.subject } else { '' }
            body_path = if (Test-ShinsaRecordProperty -Record $manifest -Name 'body_path') { Resolve-ShinsaArchiveItemPath -MailDirectory $mailDirectory -RelativePath ([string]$manifest.body_path) } else { '' }
            msg_path = if (Test-ShinsaRecordProperty -Record $manifest -Name 'msg_path') { Resolve-ShinsaArchiveItemPath -MailDirectory $mailDirectory -RelativePath ([string]$manifest.msg_path) } else { '' }
            attachment_paths = @($attachmentPaths)
        }
    }

    @($records | Sort-Object received_at, mail_id -Descending)
}

# =============================================================================
# Folder source import
# =============================================================================

function Import-ShinsaFolderRecords {
    param([Parameter(Mandatory = $true)][string]$SourcePath)

    if (-not (Test-Path $SourcePath)) { return @() }

    $root = [System.IO.Path]::GetFullPath($SourcePath)
    $records = @()
    foreach ($file in Get-ChildItem -Path $root -Recurse -File) {
        $fullPath = [System.IO.Path]::GetFullPath($file.FullName)
        $relativePath = $fullPath.Substring($root.Length).TrimStart('\')
        if ([string]::IsNullOrWhiteSpace($relativePath)) { continue }

        $segments = $relativePath -split '\\'
        $caseId = if ($segments.Count -gt 0) { $segments[0] } else { '' }
        if ([string]::IsNullOrWhiteSpace($caseId)) { continue }

        $records += [pscustomobject]@{
            case_id = $caseId
            folder_path = Join-Path $root $caseId
            file_path = $fullPath
            relative_path = if ($segments.Count -gt 1) { ($segments[1..($segments.Count - 1)] -join '\') } else { $file.Name }
            file_name = $file.Name
            extension = $file.Extension
            modified_at = $file.LastWriteTime.ToString('o')
            size = [int64]$file.Length
        }
    }

    @($records | Sort-Object case_id, relative_path)
}

# =============================================================================
# Source dispatcher
# =============================================================================

function Import-ShinsaSourceRecords {
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)][string]$SourceName
    )

    $src = $Config.sources[$SourceName]
    $view = if ($null -ne $src -and $null -ne $src.view) { [string]$src.view } else { 'fields' }
    $sourcePath = Get-ShinsaSourcePath -Config $Config -SourceName $SourceName

    if ([string]::IsNullOrWhiteSpace($sourcePath)) {
        return @()
    }

    if (-not (Test-Path $sourcePath)) {
        throw "Source path not found for '$SourceName': $sourcePath"
    }

    switch ($view) {
        'fields' { return @(Import-ShinsaFieldsRecords -SourceConfig $src -SourcePath $sourcePath) }
        'mail'   { return @(Import-ShinsaMailRecords -SourcePath $sourcePath) }
        'tree'   { return @(Import-ShinsaFolderRecords -SourcePath $sourcePath) }
        default  { throw "Unknown view type '$view' for source '$SourceName'." }
    }
}

# =============================================================================
# Join
# =============================================================================

function Get-ShinsaJoinedRecords {
    param(
        [Parameter(Mandatory = $true)]$JoinConfig,
        [Parameter(Mandatory = $true)]$SourceRecord,
        [Parameter(Mandatory = $true)][array]$TargetRecords
    )

    $joinMap = ConvertTo-ShinsaMap -InputObject $JoinConfig
    $localKey = [string]$joinMap['local_key']
    $foreignKey = [string]$joinMap['foreign_key']
    $matchMode = if ($joinMap.Contains('match_mode')) { [string]$joinMap['match_mode'] } else { 'exact' }

    $foreignRaw = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $SourceRecord -Name $foreignKey)

    # Support multi-value fields (semicolon separated)
    $foreignValues = @($foreignRaw.Split(@(';', [char]0xFF1B), [System.StringSplitOptions]::RemoveEmptyEntries) | ForEach-Object { $_.Trim().ToLowerInvariant() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

    if ($foreignValues.Count -eq 0) {
        return @()
    }

    @($TargetRecords | Where-Object {
        $localValue = (ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $_ -Name $localKey)).ToLowerInvariant()
        if ([string]::IsNullOrWhiteSpace($localValue)) { return $false }

        switch ($matchMode) {
            'domain' {
                $localDomain = if ($localValue -match '@(.+)$') { $Matches[1] } else { $localValue }
                foreach ($fv in $foreignValues) {
                    $foreignDomain = if ($fv -match '@(.+)$') { $Matches[1] } else { $fv }
                    if ($localDomain -eq $foreignDomain) { return $true }
                }
                return $false
            }
            default {
                # exact match
                foreach ($fv in $foreignValues) {
                    if ($localValue -eq $fv) { return $true }
                }
                return $false
            }
        }
    })
}

# =============================================================================
# Writeback
# =============================================================================

function Get-ShinsaWritebackPlan {
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)]$Paths,
        [Parameter(Mandatory = $true)][string]$SourceName
    )

    $src = $Config.sources[$SourceName]
    $cache = Read-ShinsaCache -Paths $Paths
    $editableColumns = @(Get-ShinsaEditableColumnNames -Config $Config -Cache $cache -SourceName $SourceName)
    if ($editableColumns.Count -eq 0) { throw "Source '$SourceName' has no editable columns." }

    $jsonPath = Join-Path $Paths.JsonRoot $src.file
    if (-not (Test-Path $jsonPath)) {
        throw "Local JSON not found for '$SourceName'. Run sync first."
    }

    $keyColumn = if ($null -ne $src.key_column) { [string]$src.key_column } else { '' }
    if ([string]::IsNullOrWhiteSpace($keyColumn)) {
        throw "Source '$SourceName' has no key_column for writeback."
    }

    $currentRecords = @(Read-ShinsaJson -Path $jsonPath)
    $sourcePath = Get-ShinsaSourcePath -Config $Config -SourceName $SourceName
    $sourceRecords = @(Import-ShinsaFieldsRecords -SourceConfig $src -SourcePath $sourcePath)

    $sourceByKey = @{}
    foreach ($record in $sourceRecords) {
        $keyValue = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $record -Name $keyColumn)
        if (-not [string]::IsNullOrWhiteSpace($keyValue)) {
            $sourceByKey[$keyValue] = $record
        }
    }

    $changes = @()
    foreach ($record in $currentRecords) {
        $keyValue = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $record -Name $keyColumn)
        if ([string]::IsNullOrWhiteSpace($keyValue) -or -not $sourceByKey.ContainsKey($keyValue)) {
            continue
        }

        $sourceRecord = $sourceByKey[$keyValue]
        $fieldChanges = [ordered]@{}
        foreach ($columnName in $editableColumns) {
            $targetValue = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $record -Name $columnName)
            $sourceValue = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $sourceRecord -Name $columnName)
            if ($targetValue -ne $sourceValue) {
                $fieldChanges[$columnName] = [ordered]@{
                    from = $sourceValue
                    to = $targetValue
                }
            }
        }

        if ($fieldChanges.Count -eq 0) { continue }

        $changes += [pscustomobject]@{
            key_value = $keyValue
            row_id = [int](Get-ShinsaRecordValue -Record $record -Name '_source_row_id')
            changes = [pscustomobject]$fieldChanges
        }
    }

    $changeCount = (@($changes | ForEach-Object { @($_.changes.PSObject.Properties).Count } | Measure-Object -Sum).Sum)
    if ($null -eq $changeCount) { $changeCount = 0 }

    [pscustomobject]@{
        source_name = $SourceName
        source_kind = [System.IO.Path]::GetExtension($sourcePath).ToLowerInvariant()
        key_column = $keyColumn
        changes = $changes
        change_count = [int]$changeCount
        case_count = @($changes).Count
    }
}

function Invoke-ShinsaJsonWriteback {
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)][string]$SourceName,
        [Parameter(Mandatory = $true)]$Plan
    )

    $src = $Config.sources[$SourceName]
    $sourcePath = Get-ShinsaSourcePath -Config $Config -SourceName $SourceName
    $columnsMap = if ($null -ne $src.columns) { ConvertTo-ShinsaMap -InputObject $src.columns } else { $null }

    $source = Read-ShinsaJson -Path $sourcePath
    if ($source -is [System.Collections.IEnumerable] -and -not ($source -is [string]) -and -not ($source -is [System.Management.Automation.PSCustomObject])) {
        $rows = @($source)
    }
    elseif (Test-ShinsaRecordProperty -Record $source -Name 'rows') {
        $rows = @($source.rows)
    }
    else {
        $rows = @($source)
    }

    $keyColumn = [string]$Plan.key_column
    $sourceKeyName = $keyColumn
    if ($null -ne $columnsMap -and $columnsMap.Contains($keyColumn)) {
        $sourceKeyName = [string]$columnsMap[$keyColumn]
    }

    $rowByKey = @{}
    foreach ($row in $rows) {
        $keyValue = ''
        if (Test-ShinsaRecordProperty -Record $row -Name $sourceKeyName) {
            $keyValue = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $row -Name $sourceKeyName)
        }
        elseif (Test-ShinsaRecordProperty -Record $row -Name $keyColumn) {
            $keyValue = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $row -Name $keyColumn)
        }
        if (-not [string]::IsNullOrWhiteSpace($keyValue)) {
            $rowByKey[$keyValue] = $row
        }
    }

    foreach ($change in @($Plan.changes)) {
        if (-not $rowByKey.ContainsKey($change.key_value)) { continue }
        $row = $rowByKey[$change.key_value]
        foreach ($property in $change.changes.PSObject.Properties) {
            $logicalName = $property.Name
            $sourceName = if ($null -ne $columnsMap -and $columnsMap.Contains($logicalName)) { [string]$columnsMap[$logicalName] } else { $logicalName }
            Set-ShinsaRecordValue -Record $row -Name $sourceName -Value $property.Value.to
            if ($sourceName -ne $logicalName -and (Test-ShinsaRecordProperty -Record $row -Name $logicalName)) {
                Set-ShinsaRecordValue -Record $row -Name $logicalName -Value $property.Value.to
            }
        }
    }

    Write-ShinsaJson -Path $sourcePath -Data $source
}

function Invoke-ShinsaExcelWriteback {
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)][string]$SourceName,
        [Parameter(Mandatory = $true)]$Plan
    )

    $src = $Config.sources[$SourceName]
    $sourcePath = Get-ShinsaSourcePath -Config $Config -SourceName $SourceName
    $sourceTable = if ($null -ne $src.source_table) { [string]$src.source_table } else { '' }
    $columnsMap = if ($null -ne $src.columns) { ConvertTo-ShinsaMap -InputObject $src.columns } else { $null }

    $fileInfo = New-Object System.IO.FileInfo($sourcePath)
    $package = New-Object OfficeOpenXml.ExcelPackage($fileInfo)
    try {
        if (-not [string]::IsNullOrWhiteSpace($sourceTable)) {
            $tbl = $null
            $ws = $null
            foreach ($worksheet in $package.Workbook.Worksheets) {
                foreach ($t in $worksheet.Tables) {
                    if ($t.Name -eq $sourceTable) { $tbl = $t; $ws = $worksheet; break }
                }
                if ($null -ne $tbl) { break }
            }
            if ($null -eq $tbl) { throw "Structured table '$sourceTable' not found." }

            $addr = $tbl.Address
            $headerRow = $addr.Start.Row
            $startCol = $addr.Start.Column
            $endCol = $addr.End.Column

            $headerMap = @{}
            for ($c = $startCol; $c -le $endCol; $c++) {
                $hv = $ws.Cells[$headerRow, $c].Text
                if (-not [string]::IsNullOrWhiteSpace($hv)) { $headerMap[$hv.Trim()] = $c }
            }

            foreach ($change in @($Plan.changes)) {
                $rowId = [int]$change.row_id
                foreach ($property in $change.changes.PSObject.Properties) {
                    $logicalName = $property.Name
                    $headerName = if ($null -ne $columnsMap -and $columnsMap.Contains($logicalName)) { [string]$columnsMap[$logicalName] } else { $logicalName }
                    if (-not $headerMap.ContainsKey($headerName)) { throw "Column '$headerName' not found in table '$sourceTable'." }
                    $ws.Cells[$rowId, $headerMap[$headerName]].Value = $property.Value.to
                }
            }
        }
        else {
            $ws = $package.Workbook.Worksheets[1]
            $startCol = $ws.Dimension.Start.Column
            $endCol = $ws.Dimension.End.Column

            $headerMap = @{}
            for ($c = $startCol; $c -le $endCol; $c++) {
                $hv = $ws.Cells[1, $c].Text
                if (-not [string]::IsNullOrWhiteSpace($hv)) { $headerMap[$hv.Trim()] = $c }
            }

            foreach ($change in @($Plan.changes)) {
                $rowId = [int]$change.row_id
                foreach ($property in $change.changes.PSObject.Properties) {
                    $logicalName = $property.Name
                    $headerName = if ($null -ne $columnsMap -and $columnsMap.Contains($logicalName)) { [string]$columnsMap[$logicalName] } else { $logicalName }
                    if (-not $headerMap.ContainsKey($headerName)) { throw "Column '$headerName' not found in worksheet." }
                    $ws.Cells[$rowId, $headerMap[$headerName]].Value = $property.Value.to
                }
            }
        }

        $package.Save()
    }
    finally {
        if ($null -ne $package) { $package.Dispose() }
    }
}

function Invoke-ShinsaWriteback {
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)][string]$SourceName,
        [Parameter(Mandatory = $true)]$Plan
    )

    if ($Plan.case_count -eq 0) { return }

    switch ($Plan.source_kind) {
        '.json' { Invoke-ShinsaJsonWriteback -Config $Config -SourceName $SourceName -Plan $Plan }
        '.xlsx' { Invoke-ShinsaExcelWriteback -Config $Config -SourceName $SourceName -Plan $Plan }
        '.xlsm' { Invoke-ShinsaExcelWriteback -Config $Config -SourceName $SourceName -Plan $Plan }
        '.xlsb' { Invoke-ShinsaExcelWriteback -Config $Config -SourceName $SourceName -Plan $Plan }
        '.xls'  { Invoke-ShinsaExcelWriteback -Config $Config -SourceName $SourceName -Plan $Plan }
        default { throw "Writeback not supported for '$($Plan.source_kind)'." }
    }
}

# =============================================================================
# Misc
# =============================================================================

function Start-ShinsaItem {
    param([Parameter(Mandatory = $true)][string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { throw 'Path is empty.' }
    if (-not (Test-Path $Path)) { throw "Path not found: $Path" }

    Start-Process -FilePath $Path | Out-Null
}

# =============================================================================
# Change Log
# =============================================================================

function Write-ShinsaChangeLog {
    param(
        [Parameter(Mandatory = $true)][string]$LogPath,
        [Parameter(Mandatory = $true)][string]$SourceName,
        [Parameter(Mandatory = $true)][string]$KeyValue,
        [Parameter(Mandatory = $true)][string]$FieldName,
        [string]$OldValue = '',
        [string]$NewValue = '',
        [Parameter(Mandatory = $true)][string]$Origin
    )
    $entry = [ordered]@{
        ts     = (Get-Date).ToString('o')
        src    = $SourceName
        key    = $KeyValue
        field  = $FieldName
        old    = $OldValue
        new    = $NewValue
        origin = $Origin
    }
    $line = ($entry | ConvertTo-Json -Compress)
    $dir = Split-Path -Parent $LogPath
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    Add-Content -Path $LogPath -Value $line -Encoding UTF8
}

# =============================================================================
# Editable Columns
# =============================================================================

function Get-ShinsaEditableColumnNames {
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)]$Cache,
        [Parameter(Mandatory = $true)][string]$SourceName
    )
    $srcMap = ConvertTo-ShinsaMap -InputObject $Config.sources[$SourceName]
    $keyColumn = if ($srcMap.Contains('key_column')) { [string]$srcMap['key_column'] } else { '' }

    $cols = @()
    # From config (if property exists)
    if ($srcMap.Contains('editable_columns')) { $cols = @($srcMap['editable_columns']) }

    # Also from cache field_settings
    $fs = Get-ShinsaFieldSettings -Cache $Cache -SourceName $SourceName
    foreach ($fn in $fs.Keys) {
        $setting = ConvertTo-ShinsaMap -InputObject $fs[$fn]
        if ($setting.Contains('editable') -and $setting['editable'] -eq $true) {
            if ($cols -notcontains $fn) { $cols += $fn }
        }
    }

    # Exclude key_column
    @($cols | Where-Object { $_ -ne $keyColumn })
}

# =============================================================================
# 3-Way Merge
# =============================================================================

function Merge-ShinsaSourceRecords {
    param(
        [Parameter(Mandatory = $true)][string]$KeyColumn,
        [Parameter(Mandatory = $true)][string[]]$EditableColumns,
        [Parameter(Mandatory = $true)][array]$RemoteRecords,
        [array]$LocalRecords = @(),
        [array]$SnapshotRecords = @()
    )

    $merged = @()
    $conflicts = @()
    $logEntries = @()

    # Build lookup tables
    $localByKey = @{}
    foreach ($r in $LocalRecords) {
        $k = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $r -Name $KeyColumn)
        if (-not [string]::IsNullOrWhiteSpace($k)) { $localByKey[$k] = $r }
    }

    $snapshotByKey = @{}
    foreach ($r in $SnapshotRecords) {
        $k = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $r -Name $KeyColumn)
        if (-not [string]::IsNullOrWhiteSpace($k)) { $snapshotByKey[$k] = $r }
    }

    # Process remote records (authoritative for existence)
    foreach ($remote in $RemoteRecords) {
        $key = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $remote -Name $KeyColumn)
        if ([string]::IsNullOrWhiteSpace($key)) { $merged += $remote; continue }

        if (-not $localByKey.ContainsKey($key)) {
            # New record from remote
            $merged += $remote
            $logEntries += @{ key = $key; field = '*'; old = ''; new = 'added'; origin = 'remote' }
            continue
        }

        $local = $localByKey[$key]
        $snapshot = if ($snapshotByKey.ContainsKey($key)) { $snapshotByKey[$key] } else { $null }

        # If no snapshot, treat as first sync - take remote
        if ($null -eq $snapshot) {
            $merged += $remote
            continue
        }

        # 3-way merge per field
        $mergedRecord = [ordered]@{}
        $recordConflicts = @()

        # Collect all field names from remote
        $allFields = @()
        foreach ($prop in $remote.PSObject.Properties) { if ($prop.Name -notlike '_*') { $allFields += $prop.Name } }

        foreach ($fn in $allFields) {
            $remoteVal = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $remote -Name $fn)
            $localVal = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $local -Name $fn)
            $snapVal = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $snapshot -Name $fn)

            if ($EditableColumns -contains $fn) {
                # Editable field: 3-way merge
                $localChanged = ($localVal -ne $snapVal)
                $remoteChanged = ($remoteVal -ne $snapVal)

                if (-not $localChanged -and -not $remoteChanged) {
                    $mergedRecord[$fn] = $remoteVal
                }
                elseif (-not $localChanged -and $remoteChanged) {
                    $mergedRecord[$fn] = $remoteVal
                    $logEntries += @{ key = $key; field = $fn; old = $snapVal; new = $remoteVal; origin = 'remote' }
                }
                elseif ($localChanged -and -not $remoteChanged) {
                    $mergedRecord[$fn] = $localVal
                }
                elseif ($localVal -eq $remoteVal) {
                    $mergedRecord[$fn] = $localVal
                }
                else {
                    # CONFLICT
                    $mergedRecord[$fn] = $localVal  # keep local as tentative
                    $recordConflicts += [pscustomobject]@{
                        field    = $fn
                        original = $snapVal
                        local    = $localVal
                        remote   = $remoteVal
                    }
                }
            }
            else {
                # Non-editable: always take remote
                $mergedRecord[$fn] = $remoteVal
                if ($remoteVal -ne $localVal) {
                    $logEntries += @{ key = $key; field = $fn; old = $localVal; new = $remoteVal; origin = 'remote' }
                }
            }
        }

        # Copy internal metadata from remote
        foreach ($prop in $remote.PSObject.Properties) {
            if ($prop.Name -like '_*') {
                $mergedRecord[$prop.Name] = $prop.Value
            }
        }

        $merged += [pscustomobject]$mergedRecord

        if ($recordConflicts.Count -gt 0) {
            $conflicts += [pscustomobject]@{
                key_value = $key
                fields    = $recordConflicts
            }
        }
    }

    # Check for deleted records (in local but not in remote)
    foreach ($key in $localByKey.Keys) {
        $inRemote = $false
        foreach ($r in $RemoteRecords) {
            $rk = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $r -Name $KeyColumn)
            if ($rk -eq $key) { $inRemote = $true; break }
        }
        if (-not $inRemote) {
            $logEntries += @{ key = $key; field = '*'; old = 'existed'; new = 'deleted'; origin = 'remote' }
        }
    }

    [pscustomobject]@{
        merged      = @($merged)
        conflicts   = @($conflicts)
        log_entries = @($logEntries)
    }
}

Export-ModuleMember -Function *-Shinsa*
