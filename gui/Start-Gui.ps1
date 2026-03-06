Import-Module (Join-Path $PSScriptRoot '..\scripts\Common.psm1') -Force -DisableNameChecking
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Add-Type -Name NativeMethods -Namespace Win32 -MemberDefinition @'
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
'@

$ErrorActionPreference = 'Stop'
$script:AppRoot = Split-Path -Parent $PSScriptRoot
$script:Config = Get-ShinsaConfig -ScriptPath $MyInvocation.MyCommand.Path
$script:Paths = Get-ShinsaDataPaths -Config $script:Config

Ensure-ShinsaState -Paths $script:Paths

# Auto-sync on startup if any source JSON missing
$needsSync = $false
foreach ($name in (Get-ShinsaSourceNames -Config $script:Config)) {
    $src = $script:Config.sources[$name]
    $jp = Join-Path $script:Paths.JsonRoot $src.file
    if (-not (Test-Path $jp)) { $needsSync = $true; break }
}
if ($needsSync) {
    try { & (Join-Path $script:AppRoot 'scripts\Sync-Data.ps1') }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Sync failed: $($_.Exception.Message)", 'shinsa') | Out-Null
        exit 1
    }
}

# --- Style ---
$script:M = 6
$script:BgColor = [System.Drawing.SystemColors]::Control
$script:Font = New-Object System.Drawing.Font($script:Config.gui.font_name, [single]$script:Config.gui.font_size)
$script:MonoFont = New-Object System.Drawing.Font('Consolas', 10)

# --- Data ---
$script:allData = [ordered]@{}
$script:fieldEditors = @{}
$script:filteredIndices = @()
$script:cacheState = $null
$script:joinedTabs = @{}
$script:fieldGroupTabs = @()
$script:dirty = $false
$script:undoStack = [System.Collections.ArrayList]::new()
$script:undoMaxSize = 50
$script:lastSourceTimestamps = @{}

# --- Helpers ---

function Get-FieldLabelText {
    param([string]$FieldName)
    $words = ($FieldName -replace '_', ' ') -split ' '
    ($words | ForEach-Object { if ($_.Length -gt 0) { $_.Substring(0,1).ToUpper() + $_.Substring(1) } else { $_ } }) -join ' '
}

function Get-FieldGroupAndName {
    param([string]$FieldName)
    $pos = $FieldName.IndexOf('_')
    if ($pos -gt 0 -and $pos -lt ($FieldName.Length - 1)) {
        $prefix = $FieldName.Substring(0, $pos)
        $rest = $FieldName.Substring($pos + 1)
        return @{ Group = $prefix; Name = $rest }
    }
    return @{ Group = ''; Name = $FieldName }
}

function Get-FieldGroups {
    param([string[]]$Fields)
    $groups = [ordered]@{}
    foreach ($fn in $Fields) {
        $gn = Get-FieldGroupAndName $fn
        $g = $gn.Group
        if (-not $groups.Contains($g)) { $groups[$g] = @() }
        $groups[$g] += $fn
    }
    $groups
}

function Test-HasGroups {
    param([string[]]$Fields)
    $groups = Get-FieldGroups $Fields
    if ($groups.Count -le 1) { return $false }
    $named = @($groups.Keys | Where-Object { $_ -ne '' })
    return ($named.Count -ge 2)
}

function Format-FieldValue {
    param($Value)
    if ($Value -is [System.Array] -or $Value -is [System.Collections.IList]) {
        return ($Value | ForEach-Object { [string]$_ }) -join '; '
    }
    return ConvertTo-ShinsaString -Value $Value
}

function Get-SourceConfig {
    param([string]$Name)
    $script:Config.sources[$Name]
}

function Get-SourceMap {
    param([string]$Name)
    ConvertTo-ShinsaMap -InputObject (Get-SourceConfig $Name)
}

function Ensure-FieldSettings {
    param([string]$Name)
    $fs = Get-ShinsaFieldSettings -Cache $script:cacheState -SourceName $Name
    if ($fs.Count -eq 0) {
        $records = if ($script:allData.Contains($Name)) { @($script:allData[$Name]) } else { @() }
        $fs = Initialize-ShinsaFieldSettings -Config $script:Config -Cache $script:cacheState -SourceName $Name -Records $records
        Save-ShinsaCache -Paths $script:Paths -Cache $script:cacheState
    }
    return $fs
}

function Get-DisplayColumns {
    param([string]$Name)
    $fs = Ensure-FieldSettings $Name
    $cols = @($fs.Keys | Where-Object { $fsm = ConvertTo-ShinsaMap -InputObject $fs[$_]; $fsm.Contains('in_list') -and $fsm['in_list'] -eq $true })
    if ($cols.Count -gt 0) { return $cols }
    # Fallback
    $recs = $script:allData[$Name]
    if ($recs -and $recs.Count -gt 0) { return @($recs[0].PSObject.Properties | Where-Object { $_.Name -notlike '_*' } | Select-Object -First 3 | ForEach-Object { $_.Name }) }
    return @()
}

function Get-DetailColumns {
    param([string]$Name)
    $fs = Ensure-FieldSettings $Name
    if ($fs.Count -gt 0) { return @($fs.Keys) }
    $recs = $script:allData[$Name]
    if ($recs -and $recs.Count -gt 0) {
        return @($recs[0].PSObject.Properties | Where-Object { $_.Name -notlike '_*' } | ForEach-Object { $_.Name })
    }
    return @()
}

function Get-EditableColumns {
    param([string]$Name)
    $fs = Ensure-FieldSettings $Name
    $cols = @($fs.Keys | Where-Object { $fsm = ConvertTo-ShinsaMap -InputObject $fs[$_]; -not $fsm.Contains('editable') -or $fsm['editable'] -eq $true })
    if ($cols.Count -gt 0) { return $cols }
    return @(Get-DetailColumns $Name)
}

function Get-MultilineColumns {
    param([string]$Name)
    $fs = Ensure-FieldSettings $Name
    return @($fs.Keys | Where-Object { $fsm = ConvertTo-ShinsaMap -InputObject $fs[$_]; $fsm.Contains('multiline') -and $fsm['multiline'] -eq $true })
}

function Get-UiState { ConvertTo-ShinsaMap -InputObject $script:cacheState.ui_state }

function Save-UiState {
    $uiState = Get-UiState
    $bounds = if ($form.WindowState -eq 'Normal') { $form.Bounds } else { $form.RestoreBounds }
    $uiState['window_left']   = $bounds.Left
    $uiState['window_top']    = $bounds.Top
    $uiState['window_width']  = $bounds.Width
    $uiState['window_height'] = $bounds.Height
    $uiState['window_state']  = switch ($form.WindowState) { 'Maximized' { 'Maximized' }; default { 'Normal' } }
    $uiState['main_splitter_distance'] = $mainSplit.SplitterDistance
    $uiState['outer_splitter_distance'] = $outerSplit.SplitterDistance
    $uiState['selected_source'] = [string]$cmbSource.SelectedItem
    $uiState['search_text']     = $txtFilter.Text
    $script:cacheState.ui_state = [pscustomobject]$uiState
    Save-ShinsaCache -Paths $script:Paths -Cache $script:cacheState
}

function Set-SafeSplitterLayout {
    param([Parameter(Mandatory)][System.Windows.Forms.SplitContainer]$Ctl, [int]$Min1, [int]$Min2, [int]$Preferred)
    $Ctl.Panel1MinSize = [Math]::Max(0, $Min1)
    $avail = if ($Ctl.Orientation -eq [System.Windows.Forms.Orientation]::Vertical) { $Ctl.ClientSize.Width } else { $Ctl.ClientSize.Height }
    if ($avail -le 0) { return }
    $Ctl.Panel2MinSize = [Math]::Max(0, [Math]::Min($Min2, [Math]::Max(0, $avail - $Ctl.Panel1MinSize - 4)))
    $max = [Math]::Max($Ctl.Panel1MinSize, $avail - $Ctl.Panel2MinSize - 4)
    $Ctl.SplitterDistance = [Math]::Max($Ctl.Panel1MinSize, [Math]::Min($Preferred, $max))
}

function Load-AllData {
    $script:allData = [ordered]@{}
    foreach ($name in (Get-ShinsaSourceNames -Config $script:Config)) {
        $src = $script:Config.sources[$name]
        $jp = Join-Path $script:Paths.JsonRoot $src.file
        if (Test-Path $jp) {
            $script:allData[$name] = @((Read-ShinsaJson -Path $jp))
        } else {
            $script:allData[$name] = @()
        }
    }
    $script:cacheState = Read-ShinsaCache -Paths $script:Paths
}

function Save-SourceData {
    param([string]$Name, [switch]$Quiet)
    $editableCols = Get-EditableColumns $Name
    if ($editableCols.Count -eq 0) { return }

    $idx = $listRecords.SelectedIndex
    if ($idx -lt 0 -or $idx -ge $script:filteredIndices.Count) { return }

    $recIdx = $script:filteredIndices[$idx]
    $rec = $script:allData[$Name][$recIdx]
    $src = Get-SourceConfig $Name
    $keyColumn = if ($null -ne $src.key_column) { [string]$src.key_column } else { '' }
    $keyValue = if (-not [string]::IsNullOrWhiteSpace($keyColumn)) { ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $rec -Name $keyColumn) } else { '' }
    $logPath = Join-Path $script:Paths.JsonRoot 'changelog.jsonl'

    $changed = $false
    foreach ($fn in $editableCols) {
        if (-not $script:fieldEditors.ContainsKey($fn)) { continue }
        $newVal = $script:fieldEditors[$fn].Text.Trim()
        $oldVal = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $rec -Name $fn)
        if ($newVal -ne $oldVal) {
            # Push to undo stack
            if ($script:undoStack.Count -ge $script:undoMaxSize) { $script:undoStack.RemoveAt(0) }
            [void]$script:undoStack.Add(@{ source = $Name; key = $keyValue; field = $fn; oldValue = $oldVal; newValue = $newVal })
            # Log
            Write-ShinsaChangeLog -LogPath $logPath -SourceName $Name -KeyValue $keyValue -FieldName $fn -OldValue $oldVal -NewValue $newVal -Origin 'local'
            Add-LogEntry ([pscustomobject]@{ ts = (Get-Date).ToString('o'); src = $Name; key = $keyValue; field = $fn; old = $oldVal; new = $newVal; origin = 'local' })
            Set-ShinsaRecordValue -Record $rec -Name $fn -Value $newVal
            $changed = $true
        }
    }
    if ($changed) {
        $jp = Join-Path $script:Paths.JsonRoot $src.file
        Write-ShinsaJson -Path $jp -Data $script:allData[$Name]
    }
    $script:dirty = $false
    if (-not $Quiet -and $changed) { $statusBar.Text = 'Saved.' }
}

function Save-CurrentRecordEdits {
    if (-not $script:dirty) { return }
    $name = [string]$cmbSource.SelectedItem
    if (-not $name) { return }
    Save-SourceData $name -Quiet
}

function Invoke-GuiSync {
    Save-CurrentRecordEdits
    $statusBar.Text = 'Syncing...'
    $form.Refresh()
    $conflicts = @(& (Join-Path $script:AppRoot 'scripts\Sync-Data.ps1'))
    Load-AllData
    Load-ChangeLog
    if ($conflicts.Count -gt 0) {
        $resolved = Show-ConflictDialog -Conflicts $conflicts
        if ($resolved) { Apply-ConflictResolutions -Resolutions $resolved }
    }
    Initialize-SourceTimestamps
    Update-RecordList
    $statusBar.Text = "Synced at $(Get-Date -Format 'HH:mm:ss')"
}

function Invoke-GuiWriteback {
    $name = [string]$cmbSource.SelectedItem
    if (-not $name) { return }
    Save-CurrentRecordEdits

    # Smart sync first to get latest remote
    Invoke-GuiSync

    $editCols = Get-EditableColumns $name
    if ($editCols.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Source '$name' has no editable columns.", 'shinsa') | Out-Null
        return
    }

    $plan = Get-ShinsaWritebackPlan -Config $script:Config -Paths $script:Paths -SourceName $name
    if ($plan.case_count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('No changes to reflect.', 'shinsa') | Out-Null; return
    }
    $lines = @()
    foreach ($c in @($plan.changes)) { $lines += ('{0}: {1}' -f $c.key_value, (@($c.changes.PSObject.Properties.Name) -join ', ')) }
    $msg = "Reflect {0} record(s) / {1} field(s) to source?`n`n{2}" -f $plan.case_count, $plan.change_count, ($lines -join "`n")
    $r = [System.Windows.Forms.MessageBox]::Show($msg, 'shinsa', 'OKCancel', 'Question')
    if ($r -ne 'OK') { return }
    Invoke-ShinsaWriteback -Config $script:Config -SourceName $name -Plan $plan

    $src = Get-SourceConfig $name
    $sourcePath = Get-ShinsaSourcePath -Config $script:Config -SourceName $name
    $refreshed = @(Import-ShinsaFieldsRecords -SourceConfig $src -SourcePath $sourcePath)
    $jp = Join-Path $script:Paths.JsonRoot $src.file
    Write-ShinsaJson -Path $jp -Data $refreshed
    $script:allData[$name] = $refreshed

    # Update snapshot after reflect
    $snapshotPath = ($jp -replace '\.json$', '.snapshot.json')
    Write-ShinsaJson -Path $snapshotPath -Data $refreshed

    Initialize-SourceTimestamps
    Update-RecordList
    $statusBar.Text = 'Reflected to source.'
}

# =============================================================================
# Settings Dialog
# =============================================================================

function Show-SettingsDialog {
    $localPath = Join-Path $script:AppRoot 'config\config.local.json'
    $localConfig = ConvertTo-ShinsaMap -InputObject (Read-ShinsaJson -Path $localPath)
    if (-not $localConfig.Contains('sources')) { $localConfig['sources'] = [ordered]@{} }
    $localSources = $localConfig['sources']
    if (-not ($localSources -is [System.Collections.IDictionary])) { $localSources = ConvertTo-ShinsaMap -InputObject $localSources; $localConfig['sources'] = $localSources }
    if (-not $localConfig.Contains('mail')) { $localConfig['mail'] = [ordered]@{} }
    $localMail = $localConfig['mail']
    if (-not ($localMail -is [System.Collections.IDictionary])) { $localMail = ConvertTo-ShinsaMap -InputObject $localMail; $localConfig['mail'] = $localMail }

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = 'Settings'
    $dlg.Size = New-Object System.Drawing.Size(820, 640)
    $dlg.MinimumSize = New-Object System.Drawing.Size(700, 520)
    $dlg.StartPosition = 'CenterParent'
    $dlg.FormBorderStyle = 'Sizable'
    $dlg.MaximizeBox = $true
    $dlg.MinimizeBox = $false
    $dlg.Font = $script:Font
    $dlg.BackColor = $script:BgColor

    $tabCtl = New-Object System.Windows.Forms.TabControl
    $tabCtl.Dock = 'Fill'

    # ===================== Sources tab =====================
    $tabSources = New-Object System.Windows.Forms.TabPage
    $tabSources.Text = 'Sources'
    $tabSources.Padding = New-Object System.Windows.Forms.Padding(8)

    # Top: source selector + path + table
    $srcTopPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $srcTopPanel.Dock = 'Top'
    $srcTopPanel.Height = 100
    $srcTopPanel.ColumnCount = 3
    $srcTopPanel.RowCount = 3
    [void]$srcTopPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('AutoSize')))
    [void]$srcTopPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 100)))
    [void]$srcTopPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('AutoSize')))

    $lblSrc = New-Object System.Windows.Forms.Label; $lblSrc.Text = 'Source:'; $lblSrc.AutoSize = $true; $lblSrc.Anchor = 'Left'; $lblSrc.Margin = New-Object System.Windows.Forms.Padding(0, 6, 8, 0)
    $cmbSrc = New-Object System.Windows.Forms.ComboBox; $cmbSrc.DropDownStyle = 'DropDownList'; $cmbSrc.Dock = 'Fill'; $cmbSrc.Margin = New-Object System.Windows.Forms.Padding(0, 2, 0, 4)
    $lblPath = New-Object System.Windows.Forms.Label; $lblPath.Text = 'Path:'; $lblPath.AutoSize = $true; $lblPath.Anchor = 'Left'; $lblPath.Margin = New-Object System.Windows.Forms.Padding(0, 6, 8, 0)
    $txtPath = New-Object System.Windows.Forms.TextBox; $txtPath.Dock = 'Fill'; $txtPath.Margin = New-Object System.Windows.Forms.Padding(0, 2, 4, 4)
    $btnBrowse = New-Object System.Windows.Forms.Button; $btnBrowse.Text = '...'; $btnBrowse.Size = New-Object System.Drawing.Size(36, 26); $btnBrowse.FlatStyle = 'Standard'
    $lblTable = New-Object System.Windows.Forms.Label; $lblTable.Text = 'Table:'; $lblTable.AutoSize = $true; $lblTable.Anchor = 'Left'; $lblTable.Margin = New-Object System.Windows.Forms.Padding(0, 6, 8, 0)
    $txtTable = New-Object System.Windows.Forms.TextBox; $txtTable.Dock = 'Fill'; $txtTable.Margin = New-Object System.Windows.Forms.Padding(0, 2, 0, 4)

    $srcTopPanel.Controls.Add($lblSrc, 0, 0); $srcTopPanel.Controls.Add($cmbSrc, 1, 0); $srcTopPanel.SetColumnSpan($cmbSrc, 2)
    $srcTopPanel.Controls.Add($lblPath, 0, 1); $srcTopPanel.Controls.Add($txtPath, 1, 1); $srcTopPanel.Controls.Add($btnBrowse, 2, 1)
    $srcTopPanel.Controls.Add($lblTable, 0, 2); $srcTopPanel.Controls.Add($txtTable, 1, 2)

    # Bottom: field settings DataGridView
    $dgvFields = New-Object System.Windows.Forms.DataGridView
    $dgvFields.Dock = 'Fill'
    $dgvFields.ColumnHeadersHeightSizeMode = 'AutoSize'
    $dgvFields.AllowUserToAddRows = $false
    $dgvFields.AllowUserToDeleteRows = $false
    $dgvFields.RowHeadersVisible = $false
    $dgvFields.BackgroundColor = [System.Drawing.SystemColors]::Window
    $dgvFields.BorderStyle = 'Fixed3D'
    $dgvFields.AutoSizeColumnsMode = 'Fill'
    $dgvFields.SelectionMode = 'FullRowSelect'

    $colName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colName.Name = 'Field'; $colName.HeaderText = 'Field'; $colName.ReadOnly = $true; $colName.FillWeight = 40
    $colType = New-Object System.Windows.Forms.DataGridViewComboBoxColumn; $colType.Name = 'Type'; $colType.HeaderText = 'Type'; $colType.FillWeight = 15
    [void]$colType.Items.AddRange(@('text', 'date', 'number'))
    $colType.FlatStyle = 'Flat'
    $colInList = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn; $colInList.Name = 'InList'; $colInList.HeaderText = 'List'; $colInList.FillWeight = 10
    $colEdit = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn; $colEdit.Name = 'Editable'; $colEdit.HeaderText = 'Edit'; $colEdit.FillWeight = 10
    $colMulti = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn; $colMulti.Name = 'Multiline'; $colMulti.HeaderText = 'Multi'; $colMulti.FillWeight = 10
    [void]$dgvFields.Columns.AddRange(@($colName, $colType, $colInList, $colEdit, $colMulti))

    $tabSources.Controls.Add($dgvFields)
    $tabSources.Controls.Add($srcTopPanel)

    # ===================== General tab =====================
    $tabGeneral = New-Object System.Windows.Forms.TabPage
    $tabGeneral.Text = 'General'
    $tabGeneral.Padding = New-Object System.Windows.Forms.Padding(12)

    $genLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $genLayout.Dock = 'Top'
    $genLayout.Height = 80
    $genLayout.ColumnCount = 2
    $genLayout.RowCount = 2
    [void]$genLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('AutoSize')))
    [void]$genLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 100)))

    $lblSelfAddr = New-Object System.Windows.Forms.Label; $lblSelfAddr.Text = 'Self address:'; $lblSelfAddr.AutoSize = $true; $lblSelfAddr.Anchor = 'Left'; $lblSelfAddr.Margin = New-Object System.Windows.Forms.Padding(0, 6, 8, 0)
    $txtSelfAddr = New-Object System.Windows.Forms.TextBox; $txtSelfAddr.Dock = 'Fill'; $txtSelfAddr.Margin = New-Object System.Windows.Forms.Padding(0, 2, 0, 4)
    $selfAddr = if ($localMail.Contains('self_address')) { [string]$localMail['self_address'] } else { '' }
    $txtSelfAddr.Text = $selfAddr

    $lblJsonRoot = New-Object System.Windows.Forms.Label; $lblJsonRoot.Text = 'JSON root:'; $lblJsonRoot.AutoSize = $true; $lblJsonRoot.Anchor = 'Left'; $lblJsonRoot.Margin = New-Object System.Windows.Forms.Padding(0, 6, 8, 0)
    $txtJsonRoot = New-Object System.Windows.Forms.TextBox; $txtJsonRoot.Dock = 'Fill'; $txtJsonRoot.Margin = New-Object System.Windows.Forms.Padding(0, 2, 0, 4)
    $jsonRoot = ''
    if ($localConfig.Contains('paths') -and $localConfig['paths'] -is [System.Collections.IDictionary] -and $localConfig['paths'].Contains('json_root')) {
        $jsonRoot = [string]$localConfig['paths']['json_root']
    }
    $txtJsonRoot.Text = $jsonRoot

    $genLayout.Controls.Add($lblSelfAddr, 0, 0); $genLayout.Controls.Add($txtSelfAddr, 1, 0)
    $genLayout.Controls.Add($lblJsonRoot, 0, 1);  $genLayout.Controls.Add($txtJsonRoot, 1, 1)
    $tabGeneral.Controls.Add($genLayout)

    [void]$tabCtl.TabPages.Add($tabSources)
    [void]$tabCtl.TabPages.Add($tabGeneral)

    # ===================== Buttons =====================
    $btnPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $btnPanel.Dock = 'Bottom'
    $btnPanel.Height = 40
    $btnPanel.FlowDirection = 'RightToLeft'
    $btnPanel.Padding = New-Object System.Windows.Forms.Padding(4)

    $btnCancel = New-Object System.Windows.Forms.Button; $btnCancel.Text = 'Cancel'; $btnCancel.Size = New-Object System.Drawing.Size(90, 28); $btnCancel.FlatStyle = 'Standard'
    $btnSave = New-Object System.Windows.Forms.Button; $btnSave.Text = 'Save'; $btnSave.Size = New-Object System.Drawing.Size(90, 28); $btnSave.FlatStyle = 'Standard'
    $btnOpenJson = New-Object System.Windows.Forms.Button; $btnOpenJson.Text = 'Open JSON...'; $btnOpenJson.Size = New-Object System.Drawing.Size(120, 28); $btnOpenJson.FlatStyle = 'Standard'
    $btnPanel.Controls.AddRange(@($btnCancel, $btnSave, $btnOpenJson))

    $dlg.Controls.Add($tabCtl)
    $dlg.Controls.Add($btnPanel)
    $dlg.AcceptButton = $btnSave
    $dlg.CancelButton = $btnCancel

    # ===================== State =====================
    $script:settingsEditors = @{}
    $allSourceNames = Get-ShinsaSourceNames -Config $script:Config

    foreach ($sn in $allSourceNames) {
        $srcMap = Get-SourceMap $sn
        $localSrcMap = if ($localSources.Contains($sn)) { ConvertTo-ShinsaMap -InputObject $localSources[$sn] } else { [ordered]@{} }
        $sp = if ($localSrcMap.Contains('source_path')) { [string]$localSrcMap['source_path'] } elseif ($srcMap.Contains('source_path')) { [string]$srcMap['source_path'] } else { '' }
        $st = if ($localSrcMap.Contains('source_table')) { [string]$localSrcMap['source_table'] } elseif ($srcMap.Contains('source_table')) { [string]$srcMap['source_table'] } else { '' }

        $fs = Ensure-FieldSettings $sn

        $script:settingsEditors[$sn] = @{ path = $sp; table = $st; fieldSettings = $fs }
        [void]$cmbSrc.Items.Add($sn)
    }

    $script:currentSettingsSrc = ''

    function Save-CurrentSourceSettings {
        if ([string]::IsNullOrWhiteSpace($script:currentSettingsSrc)) { return }
        $ed = $script:settingsEditors[$script:currentSettingsSrc]
        $ed.path = $txtPath.Text
        $ed.table = $txtTable.Text

        $newFs = [ordered]@{}
        for ($r = 0; $r -lt $dgvFields.Rows.Count; $r++) {
            $fn = [string]$dgvFields.Rows[$r].Cells['Field'].Value
            $newFs[$fn] = [ordered]@{
                type = [string]$dgvFields.Rows[$r].Cells['Type'].Value
                in_list = [bool]$dgvFields.Rows[$r].Cells['InList'].Value
                editable = [bool]$dgvFields.Rows[$r].Cells['Editable'].Value
                multiline = [bool]$dgvFields.Rows[$r].Cells['Multiline'].Value
            }
        }
        $ed.fieldSettings = $newFs
    }

    function Load-SourceSettings {
        param([string]$SrcName)
        $ed = $script:settingsEditors[$SrcName]
        $txtPath.Text = $ed.path
        $txtTable.Text = $ed.table
        $srcMap = Get-SourceMap $SrcName
        $txtTable.Enabled = $srcMap.Contains('source_table')

        $dgvFields.Rows.Clear()
        foreach ($fn in $ed.fieldSettings.Keys) {
            $fsm = ConvertTo-ShinsaMap -InputObject $ed.fieldSettings[$fn]
            $tp = if ($fsm.Contains('type')) { [string]$fsm['type'] } else { 'text' }
            $il = if ($fsm.Contains('in_list')) { [bool]$fsm['in_list'] } else { $false }
            $ed2 = if ($fsm.Contains('editable')) { [bool]$fsm['editable'] } else { $true }
            $ml = if ($fsm.Contains('multiline')) { [bool]$fsm['multiline'] } else { $false }
            [void]$dgvFields.Rows.Add($fn, $tp, $il, $ed2, $ml)
        }

        $script:currentSettingsSrc = $SrcName
    }

    $cmbSrc.Add_SelectedIndexChanged({
        Save-CurrentSourceSettings
        Load-SourceSettings ([string]$cmbSrc.SelectedItem)
    })

    $btnBrowse.Add_Click({
        $srcMap = Get-SourceMap ([string]$cmbSrc.SelectedItem)
        $view = if ($srcMap.Contains('view')) { [string]$srcMap['view'] } else { 'fields' }
        if ($view -eq 'fields') {
            $ofd = New-Object System.Windows.Forms.OpenFileDialog
            $ofd.Filter = 'Excel / JSON|*.xlsx;*.xlsm;*.xls;*.json|All Files|*.*'
            if ($ofd.ShowDialog() -eq 'OK') { $txtPath.Text = $ofd.FileName }
        } else {
            $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
            $fbd.Description = 'Select source folder'
            if ($fbd.ShowDialog() -eq 'OK') { $txtPath.Text = $fbd.SelectedPath }
        }
    })

    # ===================== Save =====================
    $btnSave.Add_Click({
        Save-CurrentSourceSettings

        # Save source paths to config.local.json
        foreach ($sn in $script:settingsEditors.Keys) {
            $ed = $script:settingsEditors[$sn]
            if (-not $localSources.Contains($sn)) { $localSources[$sn] = [ordered]@{} }
            $ls = $localSources[$sn]
            if (-not ($ls -is [System.Collections.IDictionary])) { $ls = ConvertTo-ShinsaMap -InputObject $ls; $localSources[$sn] = $ls }
            $ls['source_path'] = $ed.path
            if (-not [string]::IsNullOrWhiteSpace($ed.table)) { $ls['source_table'] = $ed.table }

            # Save field settings to cache
            Set-ShinsaFieldSettings -Cache $script:cacheState -SourceName $sn -Settings $ed.fieldSettings
        }

        $localMail['self_address'] = $txtSelfAddr.Text.Trim()
        if (-not [string]::IsNullOrWhiteSpace($txtJsonRoot.Text.Trim())) {
            if (-not $localConfig.Contains('paths')) { $localConfig['paths'] = [ordered]@{} }
            $localConfig['paths']['json_root'] = $txtJsonRoot.Text.Trim()
        }

        Write-ShinsaJson -Path $localPath -Data ([pscustomobject]$localConfig)
        Save-ShinsaCache -Paths $script:Paths -Cache $script:cacheState
        $dlg.DialogResult = 'OK'
        $dlg.Close()
    })

    $btnCancel.Add_Click({ $dlg.DialogResult = 'Cancel'; $dlg.Close() })
    $btnOpenJson.Add_Click({
        try { Start-Process $localPath } catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null }
    })

    if ($cmbSrc.Items.Count -gt 0) { $cmbSrc.SelectedIndex = 0 }

    $result = $dlg.ShowDialog($form)
    $dlg.Dispose()

    if ($result -eq 'OK') {
        $script:Config = Get-ShinsaConfig -ScriptPath $MyInvocation.MyCommand.Path
        $script:Paths = Get-ShinsaDataPaths -Config $script:Config
        Ensure-ShinsaState -Paths $script:Paths
        & (Join-Path $script:AppRoot 'scripts\Sync-Data.ps1')
        Load-AllData

        $prevSource = [string]$cmbSource.SelectedItem
        $cmbSource.Items.Clear()
        foreach ($name in (Get-ShinsaPrimarySourceNames -Config $script:Config)) {
            [void]$cmbSource.Items.Add($name)
        }
        $idx = $cmbSource.Items.IndexOf($prevSource)
        if ($idx -ge 0) { $cmbSource.SelectedIndex = $idx }
        elseif ($cmbSource.Items.Count -gt 0) { $cmbSource.SelectedIndex = 0 }

        $statusBar.Text = 'Settings saved. Data reloaded.'
    }
}

# =============================================================================
# UI
# =============================================================================

$form = New-Object System.Windows.Forms.Form
$form.Text = $script:Config.gui.title
$form.MinimumSize = New-Object System.Drawing.Size(900, 500)
$form.Font = $script:Font
$form.BackColor = $script:BgColor
$form.Padding = New-Object System.Windows.Forms.Padding($script:M)
$form.AutoScaleMode = 'Dpi'
$form.KeyPreview = $true

# --- Menu ---
$menuBar = New-Object System.Windows.Forms.MenuStrip
$menuBar.BackColor = $script:BgColor

$menuFile = New-Object System.Windows.Forms.ToolStripMenuItem('File(&F)')
$menuSync = New-Object System.Windows.Forms.ToolStripMenuItem('Reload from sources')
$menuSync.ShortcutKeys = [System.Windows.Forms.Keys]::F5
$menuReflect = New-Object System.Windows.Forms.ToolStripMenuItem('Reflect to source...')
$menuReflect.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::W
$menuSep1 = New-Object System.Windows.Forms.ToolStripSeparator
$menuSettings = New-Object System.Windows.Forms.ToolStripMenuItem('Settings...')
$menuSep2 = New-Object System.Windows.Forms.ToolStripSeparator
$menuQuit = New-Object System.Windows.Forms.ToolStripMenuItem('Quit')
$menuQuit.ShortcutKeys = [System.Windows.Forms.Keys]::Alt -bor [System.Windows.Forms.Keys]::F4
[void]$menuFile.DropDownItems.AddRange(@($menuSync, $menuReflect, $menuSep1, $menuSettings, $menuSep2, $menuQuit))
[void]$menuBar.Items.Add($menuFile)
$form.MainMenuStrip = $menuBar

# --- Toolbar ---
$toolbar = New-Object System.Windows.Forms.ToolStrip
$toolbar.GripStyle = 'Hidden'
$toolbar.BackColor = $script:BgColor
$toolbar.Renderer = New-Object System.Windows.Forms.ToolStripSystemRenderer

$tsbSync = New-Object System.Windows.Forms.ToolStripButton('Sync')
$tsbSync.DisplayStyle = 'Text'
$tsbReflect = New-Object System.Windows.Forms.ToolStripButton('Reflect')
$tsbReflect.DisplayStyle = 'Text'
$tsbSettings = New-Object System.Windows.Forms.ToolStripButton('Settings')
$tsbSettings.DisplayStyle = 'Text'
[void]$toolbar.Items.AddRange(@($tsbSync, (New-Object System.Windows.Forms.ToolStripSeparator), $tsbReflect, (New-Object System.Windows.Forms.ToolStripSeparator), $tsbSettings))

# --- Outer split (left+center | right log) ---
$outerSplit = New-Object System.Windows.Forms.SplitContainer
$outerSplit.Dock = 'Fill'
$outerSplit.Orientation = 'Vertical'
$outerSplit.BackColor = $script:BgColor
$outerSplit.BorderStyle = 'None'
$outerSplit.FixedPanel = 'Panel2'
$outerSplit.SplitterDistance = 600

# --- Main split (left | center) ---
$mainSplit = New-Object System.Windows.Forms.SplitContainer
$mainSplit.Dock = 'Fill'
$mainSplit.Orientation = 'Vertical'
$mainSplit.BackColor = $script:BgColor
$mainSplit.BorderStyle = 'Fixed3D'

# --- Left pane ---
$leftPanel = New-Object System.Windows.Forms.TableLayoutPanel
$leftPanel.Dock = 'Fill'
$leftPanel.RowCount = 3
$leftPanel.ColumnCount = 1
$leftPanel.BackColor = $script:BgColor
[void]$leftPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$leftPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$leftPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))

$cmbSource = New-Object System.Windows.Forms.ComboBox
$cmbSource.DropDownStyle = 'DropDownList'
$cmbSource.Dock = 'Fill'
$cmbSource.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, $script:M)

$txtFilter = New-Object System.Windows.Forms.TextBox
$txtFilter.Dock = 'Fill'
$txtFilter.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, $script:M)
$txtFilter.BorderStyle = 'Fixed3D'

$listRecords = New-Object System.Windows.Forms.ListBox
$listRecords.Dock = 'Fill'
$listRecords.Font = $script:MonoFont
$listRecords.IntegralHeight = $false
$listRecords.BorderStyle = 'Fixed3D'

$leftPanel.Controls.Add($cmbSource, 0, 0)
$leftPanel.Controls.Add($txtFilter, 0, 1)
$leftPanel.Controls.Add($listRecords, 0, 2)
$mainSplit.Panel1.Controls.Add($leftPanel)

# --- Right pane: TabControl ---
$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Dock = 'Fill'

# Detail tab (always present)
$tabDetail = New-Object System.Windows.Forms.TabPage
$tabDetail.Text = 'Detail'
$tabDetail.Padding = New-Object System.Windows.Forms.Padding(0)
$tabDetail.BackColor = $script:BgColor

$pnlButtons = New-Object System.Windows.Forms.FlowLayoutPanel
$pnlButtons.Dock = 'Top'
$pnlButtons.Height = 32
$pnlButtons.FlowDirection = 'LeftToRight'
$pnlButtons.Padding = New-Object System.Windows.Forms.Padding(0, 1, 0, 1)
$pnlButtons.BackColor = $script:BgColor

$pnlFields = New-Object System.Windows.Forms.Panel
$pnlFields.Dock = 'Fill'
$pnlFields.AutoScroll = $true
$pnlFields.Padding = New-Object System.Windows.Forms.Padding(0, $script:M, 0, 0)
$pnlFields.BackColor = $script:BgColor

$tabDetail.Controls.Add($pnlFields)

[void]$tabs.TabPages.Add($tabDetail)

$mainSplit.Panel2.Controls.Add($tabs)
$mainSplit.Panel2.Controls.Add($pnlButtons)

# --- Right pane: Change Log ---
$logPanel = New-Object System.Windows.Forms.Panel
$logPanel.Dock = 'Fill'
$logPanel.BackColor = $script:BgColor

$logLabel = New-Object System.Windows.Forms.Label
$logLabel.Text = 'Change Log'
$logLabel.Dock = 'Top'
$logLabel.Height = 22
$logLabel.TextAlign = 'MiddleLeft'
$logLabel.Font = New-Object System.Drawing.Font($script:Font, [System.Drawing.FontStyle]::Bold)
$logLabel.BackColor = $script:BgColor
$logLabel.Padding = New-Object System.Windows.Forms.Padding(2, 0, 0, 0)

$logList = New-Object System.Windows.Forms.ListBox
$logList.Dock = 'Fill'
$logList.Font = $script:MonoFont
$logList.IntegralHeight = $false
$logList.BorderStyle = 'Fixed3D'

$logPanel.Controls.Add($logList)
$logPanel.Controls.Add($logLabel)

$outerSplit.Panel1.Controls.Add($mainSplit)
$outerSplit.Panel2.Controls.Add($logPanel)

# --- Status bar ---
$statusBar = New-Object System.Windows.Forms.StatusBar
$statusBar.SizingGrip = $true
$statusBar.ShowPanels = $true
$statusBarPanel = New-Object System.Windows.Forms.StatusBarPanel
$statusBarPanel.AutoSize = 'Spring'
$statusBarPanel.BorderStyle = 'Sunken'
[void]$statusBar.Panels.Add($statusBarPanel)

$lblCount = New-Object System.Windows.Forms.Label
$lblCount.Dock = 'Bottom'
$lblCount.Height = 20
$lblCount.TextAlign = 'BottomRight'
$lblCount.ForeColor = [System.Drawing.Color]::DimGray
$lblCount.BackColor = $script:BgColor

$form.Controls.Add($outerSplit)
$form.Controls.Add($toolbar)
$form.Controls.Add($lblCount)
$form.Controls.Add($statusBar)
$form.Controls.Add($menuBar)

# =============================================================================
# Dynamic joined tabs
# =============================================================================

function Build-JoinedTabs {
    param([string]$PrimaryName)

    foreach ($jt in $script:joinedTabs.Values) {
        $tabs.TabPages.Remove($jt.tab)
        $jt.tab.Dispose()
    }
    $script:joinedTabs = @{}

    $joinedNames = Get-ShinsaJoinedSourceNames -Config $script:Config -PrimaryName $PrimaryName
    foreach ($jName in $joinedNames) {
        $jSrcMap = Get-SourceMap $jName
        $jView = if ($jSrcMap.Contains('view')) { [string]$jSrcMap['view'] } else { 'fields' }

        $tab = New-Object System.Windows.Forms.TabPage
        $tab.Text = (Get-FieldLabelText $jName)
        $tab.Tag = $jName
        $tab.BackColor = $script:BgColor

        $info = @{ tab = $tab; name = $jName; view = $jView; srcMap = $jSrcMap }

        switch ($jView) {
            'mail' {
                $mailSplit = New-Object System.Windows.Forms.SplitContainer
                $mailSplit.Dock = 'Fill'
                $mailSplit.Orientation = 'Horizontal'
                $mailSplit.SplitterDistance = 80
                $mailSplit.FixedPanel = 'Panel1'
                $mailSplit.BackColor = $script:BgColor

                $mailList = New-Object System.Windows.Forms.ListBox
                $mailList.Dock = 'Fill'
                $mailList.Font = $script:MonoFont
                $mailList.IntegralHeight = $false
                $mailList.BorderStyle = 'Fixed3D'

                $mailDetailPanel = New-Object System.Windows.Forms.Panel
                $mailDetailPanel.Dock = 'Fill'

                $mailHeaderPanel = New-Object System.Windows.Forms.TableLayoutPanel
                $mailHeaderPanel.Dock = 'Top'
                $mailHeaderPanel.Height = 76
                $mailHeaderPanel.ColumnCount = 2
                $mailHeaderPanel.RowCount = 3
                $mailHeaderPanel.BackColor = $script:BgColor
                [void]$mailHeaderPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('AutoSize')))
                [void]$mailHeaderPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 100)))

                $subjectLbl = New-Object System.Windows.Forms.Label; $subjectLbl.Text = 'Subject:'; $subjectLbl.AutoSize = $true; $subjectLbl.Anchor = 'Left'; $subjectLbl.Margin = New-Object System.Windows.Forms.Padding(0, 4, $script:M, 0); $subjectLbl.Font = New-Object System.Drawing.Font($script:Font, [System.Drawing.FontStyle]::Bold)
                $subjectTxt = New-Object System.Windows.Forms.Label; $subjectTxt.AutoSize = $true; $subjectTxt.Anchor = 'Left'; $subjectTxt.Margin = New-Object System.Windows.Forms.Padding(0, 4, 0, 0)
                $fromLbl    = New-Object System.Windows.Forms.Label; $fromLbl.Text = 'From:'; $fromLbl.AutoSize = $true; $fromLbl.Anchor = 'Left'; $fromLbl.Margin = New-Object System.Windows.Forms.Padding(0, 2, $script:M, 0)
                $fromTxt    = New-Object System.Windows.Forms.Label; $fromTxt.AutoSize = $true; $fromTxt.Anchor = 'Left'; $fromTxt.Margin = New-Object System.Windows.Forms.Padding(0, 2, 0, 0)
                $dateLbl    = New-Object System.Windows.Forms.Label; $dateLbl.Text = 'Date:'; $dateLbl.AutoSize = $true; $dateLbl.Anchor = 'Left'; $dateLbl.Margin = New-Object System.Windows.Forms.Padding(0, 2, $script:M, 0)
                $dateTxt    = New-Object System.Windows.Forms.Label; $dateTxt.AutoSize = $true; $dateTxt.Anchor = 'Left'; $dateTxt.Margin = New-Object System.Windows.Forms.Padding(0, 2, 0, 0)
                $mailHeaderPanel.Controls.Add($subjectLbl, 0, 0); $mailHeaderPanel.Controls.Add($subjectTxt, 1, 0)
                $mailHeaderPanel.Controls.Add($fromLbl, 0, 1);    $mailHeaderPanel.Controls.Add($fromTxt, 1, 1)
                $mailHeaderPanel.Controls.Add($dateLbl, 0, 2);    $mailHeaderPanel.Controls.Add($dateTxt, 1, 2)

                $mailBody = New-Object System.Windows.Forms.TextBox
                $mailBody.Dock = 'Fill'
                $mailBody.Multiline = $true
                $mailBody.ScrollBars = 'Vertical'
                $mailBody.ReadOnly = $true
                $mailBody.BackColor = [System.Drawing.SystemColors]::Window
                $mailBody.BorderStyle = 'Fixed3D'

                $attachLabel = New-Object System.Windows.Forms.Label
                $attachLabel.Dock = 'Bottom'
                $attachLabel.Height = 18
                $attachLabel.Text = 'Attachments:'
                $attachLabel.ForeColor = [System.Drawing.Color]::DimGray

                $attachList = New-Object System.Windows.Forms.ListBox
                $attachList.Dock = 'Bottom'
                $attachList.Height = 56
                $attachList.Font = $script:MonoFont
                $attachList.BorderStyle = 'Fixed3D'

                $mailDetailPanel.Controls.Add($mailBody)
                $mailDetailPanel.Controls.Add($attachLabel)
                $mailDetailPanel.Controls.Add($attachList)
                $mailDetailPanel.Controls.Add($mailHeaderPanel)

                $mailSplit.Panel1.Controls.Add($mailList)
                $mailSplit.Panel2.Controls.Add($mailDetailPanel)
                $tab.Controls.Add($mailSplit)

                $info['mailList'] = $mailList
                $info['subjectTxt'] = $subjectTxt
                $info['fromTxt'] = $fromTxt
                $info['dateTxt'] = $dateTxt
                $info['mailBody'] = $mailBody
                $info['attachList'] = $attachList
                $info['mailRecords'] = @()

                $mailList.Add_SelectedIndexChanged({
                    $jTabName = $this.Parent.Parent.Parent.Tag
                    if (-not $jTabName -or -not $script:joinedTabs.ContainsKey($jTabName)) { return }
                    $jInfo = $script:joinedTabs[$jTabName]
                    $mi = $jInfo.mailList.SelectedIndex
                    if ($mi -lt 0 -or $mi -ge $jInfo.mailRecords.Count) { return }
                    $mailRec = $jInfo.mailRecords[$mi]

                    $jInfo.subjectTxt.Text = Format-FieldValue (Get-ShinsaRecordValue -Record $mailRec -Name 'subject')
                    $jInfo.fromTxt.Text    = Format-FieldValue (Get-ShinsaRecordValue -Record $mailRec -Name 'sender_email')
                    $jInfo.dateTxt.Text    = Format-FieldValue (Get-ShinsaRecordValue -Record $mailRec -Name 'received_at')

                    $bp = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $mailRec -Name 'body_path')
                    if (-not [string]::IsNullOrWhiteSpace($bp) -and (Test-Path $bp)) {
                        $jInfo.mailBody.Text = Get-Content -Path $bp -Raw -Encoding UTF8
                    } else { $jInfo.mailBody.Text = '' }

                    $jInfo.attachList.Items.Clear()
                    $aps = @((Get-ShinsaRecordValue -Record $mailRec -Name 'attachment_paths'))
                    foreach ($ap in $aps) {
                        $apStr = [string]$ap
                        if ([string]::IsNullOrWhiteSpace($apStr)) { continue }
                        [void]$jInfo.attachList.Items.Add([System.IO.Path]::GetFileName($apStr))
                    }
                })

                $attachList.Add_DoubleClick({
                    $jTabName = $this.Parent.Parent.Parent.Parent.Tag
                    if (-not $jTabName -or -not $script:joinedTabs.ContainsKey($jTabName)) { return }
                    $jInfo = $script:joinedTabs[$jTabName]
                    $mi = $jInfo.mailList.SelectedIndex
                    $ai = $jInfo.attachList.SelectedIndex
                    if ($mi -lt 0 -or $ai -lt 0) { return }
                    $mailRec = $jInfo.mailRecords[$mi]
                    $aps = @((Get-ShinsaRecordValue -Record $mailRec -Name 'attachment_paths'))
                    $ci = 0
                    foreach ($ap in $aps) {
                        $apStr = [string]$ap
                        if ([string]::IsNullOrWhiteSpace($apStr)) { continue }
                        if ($ci -eq $ai) {
                            try { Start-Process $apStr } catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null }
                            return
                        }
                        $ci++
                    }
                })

                $mailList.Add_DoubleClick({
                    $jTabName = $this.Parent.Parent.Parent.Tag
                    if (-not $jTabName -or -not $script:joinedTabs.ContainsKey($jTabName)) { return }
                    $jInfo = $script:joinedTabs[$jTabName]
                    $mi = $jInfo.mailList.SelectedIndex
                    if ($mi -lt 0 -or $mi -ge $jInfo.mailRecords.Count) { return }
                    $mailRec = $jInfo.mailRecords[$mi]
                    $msg = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $mailRec -Name 'msg_path')
                    if (-not [string]::IsNullOrWhiteSpace($msg) -and (Test-Path $msg)) { Start-Process $msg; return }
                    $bp = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $mailRec -Name 'body_path')
                    if (-not [string]::IsNullOrWhiteSpace($bp) -and (Test-Path $bp)) { Start-Process $bp }
                })
            }
            'tree' {
                $tree = New-Object System.Windows.Forms.TreeView
                $tree.Dock = 'Fill'
                $tree.HideSelection = $false
                $tree.ShowLines = $true
                $tree.ShowPlusMinus = $false
                $tree.ShowRootLines = $true
                $tree.PathSeparator = '\'
                $tree.BorderStyle = 'Fixed3D'
                $tab.Controls.Add($tree)
                $info['tree'] = $tree

                $tree.Add_NodeMouseDoubleClick({
                    param($s, $e)
                    try {
                        $path = [string]$e.Node.Tag
                        if (-not [string]::IsNullOrWhiteSpace($path) -and (Test-Path $path)) { Start-Process $path }
                    } catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null }
                })
                $tree.Add_BeforeCollapse({ $_.Cancel = $true })
            }
        }

        [void]$tabs.TabPages.Add($tab)
        $script:joinedTabs[$jName] = $info
    }
}

# =============================================================================
# Fields view (detail tab)
# =============================================================================

function Add-FieldEditorsToPanel {
    param([System.Windows.Forms.Panel]$Panel, [string[]]$Fields, [string[]]$EditableCols, [string[]]$MultilineCols, [bool]$StripGroup)
    $tip = New-Object System.Windows.Forms.ToolTip

    $tbl = New-Object System.Windows.Forms.TableLayoutPanel
    $tbl.Dock = 'Top'
    $tbl.AutoSize = $true
    $tbl.ColumnCount = 2
    $tbl.RowCount = $Fields.Count
    $tbl.Padding = New-Object System.Windows.Forms.Padding($script:M)
    [void]$tbl.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('AutoSize')))
    [void]$tbl.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 100)))

    for ($i = 0; $i -lt $Fields.Count; $i++) {
        $fieldName = $Fields[$i]
        $isMultiline = $fieldName -in $MultilineCols
        $rowH = if ($isMultiline) { 60 } else { 24 }
        [void]$tbl.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', $rowH)))

        $rawLabel = if ($StripGroup) { (Get-FieldGroupAndName $fieldName).Name } else { $fieldName }
        $displayName = Get-FieldLabelText -FieldName $rawLabel
        $fullLabel = Get-FieldLabelText -FieldName $fieldName

        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = $displayName
        $lbl.AutoSize = $false
        $lbl.Dock = 'Fill'
        $lbl.TextAlign = 'MiddleRight'
        $lbl.Margin = New-Object System.Windows.Forms.Padding(0, 0, 6, 2)
        if ($displayName -ne $fullLabel) { $tip.SetToolTip($lbl, $fullLabel) }
        $tbl.Controls.Add($lbl, 0, $i)

        $txt = New-Object System.Windows.Forms.TextBox
        $txt.Dock = 'Fill'
        $txt.BorderStyle = 'Fixed3D'
        $txt.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 2)
        # key_column is always read-only even if in editable_columns
        $srcCfg = Get-SourceConfig ([string]$cmbSource.SelectedItem)
        $keyCol = if ($null -ne $srcCfg -and $null -ne $srcCfg.key_column) { [string]$srcCfg.key_column } else { '' }
        $isEditable = ($fieldName -in $EditableCols) -and ($fieldName -ne $keyCol)
        $txt.ReadOnly = (-not $isEditable)
        if ($txt.ReadOnly) { $txt.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240) }
        if ($isMultiline) { $txt.Multiline = $true; $txt.ScrollBars = 'Vertical' }
        if ($isEditable) { $txt.Add_TextChanged({ $script:dirty = $true }) }
        $tbl.Controls.Add($txt, 1, $i)

        $script:fieldEditors[$fieldName] = $txt
    }

    $Panel.Controls.Add($tbl)
}

function Build-FieldEditors {
    param([string]$SourceName)
    $pnlFields.SuspendLayout()
    $pnlFields.Controls.Clear()
    $script:fieldEditors = @{}

    # Remove old group tabs if any
    if ($script:fieldGroupTabs) {
        foreach ($ft in $script:fieldGroupTabs) { $tabs.TabPages.Remove($ft); $ft.Dispose() }
    }
    $script:fieldGroupTabs = @()

    $fields = Get-DetailColumns $SourceName
    if (-not $fields -or $fields.Count -eq 0) { $pnlFields.ResumeLayout(); return }

    $editableCols = Get-EditableColumns $SourceName
    $multilineCols = Get-MultilineColumns $SourceName
    $hasGroups = Test-HasGroups $fields

    if (-not $hasGroups) {
        # No grouping: render all fields in the Detail tab
        if (-not $tabs.TabPages.Contains($tabDetail)) { $tabs.TabPages.Insert(0, $tabDetail) }
        $tabDetail.Text = 'Detail'
        Add-FieldEditorsToPanel -Panel $pnlFields -Fields $fields -EditableCols $editableCols -MultilineCols $multilineCols -StripGroup $false
    } else {
        # Grouped: one tab per group, remove the Detail tab
        $tabs.TabPages.Remove($tabDetail)
        $groups = Get-FieldGroups $fields
        $tabIndex = 0

        foreach ($gName in $groups.Keys) {
            $gFields = @($groups[$gName])
            $isUngrouped = ($gName -eq '')

            $gTab = New-Object System.Windows.Forms.TabPage
            $gTab.Text = if ($isUngrouped) { 'Other' } else { $gName }
            $gTab.BackColor = $script:BgColor
            $gTab.Padding = New-Object System.Windows.Forms.Padding(0)

            $gPanel = New-Object System.Windows.Forms.Panel
            $gPanel.Dock = 'Fill'
            $gPanel.AutoScroll = $true
            $gPanel.BackColor = $script:BgColor
            $gPanel.Padding = New-Object System.Windows.Forms.Padding(0, $script:M, 0, 0)
            $gTab.Controls.Add($gPanel)

            Add-FieldEditorsToPanel -Panel $gPanel -Fields $gFields -EditableCols $editableCols -MultilineCols $multilineCols -StripGroup (-not $isUngrouped)

            $tabs.TabPages.Insert($tabIndex, $gTab)
            $script:fieldGroupTabs += $gTab
            $tabIndex++
        }
    }
    $pnlFields.ResumeLayout()
}

function Fill-FieldEditors {
    param($Record)
    if (-not $Record) {
        foreach ($txt in $script:fieldEditors.Values) { $txt.Text = '' }
        return
    }
    foreach ($fn in $script:fieldEditors.Keys) {
        $val = Get-ShinsaRecordValue -Record $Record -Name $fn
        $script:fieldEditors[$fn].Text = Format-FieldValue $val
    }
}

function Update-DetailTab {
    param([string]$SourceName)
    $pnlButtons.Controls.Clear()
    $idx = $listRecords.SelectedIndex
    if ($idx -lt 0 -or $idx -ge $script:filteredIndices.Count) { Fill-FieldEditors $null; return }

    $recIdx = $script:filteredIndices[$idx]
    $rec = $script:allData[$SourceName][$recIdx]
    Fill-FieldEditors $rec
    $script:dirty = $false
}

# =============================================================================
# Joined tab updates
# =============================================================================

function Update-JoinedTabs {
    param([string]$PrimaryName)

    $idx = $listRecords.SelectedIndex
    $primaryRec = $null
    if ($idx -ge 0 -and $idx -lt $script:filteredIndices.Count) {
        $recIdx = $script:filteredIndices[$idx]
        $primaryRec = $script:allData[$PrimaryName][$recIdx]
    }

    foreach ($jName in @($script:joinedTabs.Keys)) {
        $jInfo = $script:joinedTabs[$jName]
        $jSrcMap = $jInfo.srcMap
        $joinConfig = if ($jSrcMap.Contains('join')) { $jSrcMap['join'] } else { $null }
        $jRecords = $script:allData[$jName]

        $matched = @()
        if ($null -ne $primaryRec -and $null -ne $joinConfig -and $null -ne $jRecords -and $jRecords.Count -gt 0) {
            $matched = @(Get-ShinsaJoinedRecords -JoinConfig $joinConfig -SourceRecord $primaryRec -TargetRecords $jRecords)
        }

        switch ($jInfo.view) {
            'mail' {
                $jInfo.mailRecords = $matched
                $jInfo.mailList.BeginUpdate()
                $jInfo.mailList.Items.Clear()
                $displayCols = Get-DisplayColumns $jName
                foreach ($rec in $matched) {
                    $label = ($displayCols | ForEach-Object { Format-FieldValue (Get-ShinsaRecordValue -Record $rec -Name $_) }) -join ' | '
                    [void]$jInfo.mailList.Items.Add($label)
                }
                $jInfo.mailList.EndUpdate()
                $jInfo.subjectTxt.Text = ''
                $jInfo.fromTxt.Text = ''
                $jInfo.dateTxt.Text = ''
                $jInfo.mailBody.Text = ''
                $jInfo.attachList.Items.Clear()
                if ($matched.Count -gt 0) { $jInfo.mailList.SelectedIndex = 0 }

                $jInfo.tab.Text = "{0} ({1})" -f (Get-FieldLabelText $jName), $matched.Count
            }
            'tree' {
                $tree = $jInfo.tree
                $tree.Nodes.Clear()

                $groupKey = 'folder_path'
                $relKey   = 'relative_path'
                $fullKey  = 'file_path'
                if ($jSrcMap.Contains('tree_group_key'))     { $groupKey = [string]$jSrcMap['tree_group_key'] }
                if ($jSrcMap.Contains('tree_relative_path')) { $relKey   = [string]$jSrcMap['tree_relative_path'] }
                if ($jSrcMap.Contains('tree_full_path'))     { $fullKey  = [string]$jSrcMap['tree_full_path'] }

                $groups = [ordered]@{}
                foreach ($rec in $matched) {
                    $gv = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $rec -Name $groupKey)
                    if (-not $groups.Contains($gv)) { $groups[$gv] = @() }
                    $groups[$gv] += $rec
                }

                foreach ($gv in $groups.Keys) {
                    $groupRecs = $groups[$gv]
                    $rootLabel = if ($gv) { [System.IO.Path]::GetFileName($gv) } else { '(root)' }
                    $rootNode = New-Object System.Windows.Forms.TreeNode($rootLabel)
                    $rootNode.Tag = $gv

                    foreach ($rec in $groupRecs) {
                        $relPath  = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $rec -Name $relKey)
                        $fullPath = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $rec -Name $fullKey)
                        $segments = $relPath -split '\\'
                        $parent = $rootNode
                        for ($i = 0; $i -lt $segments.Count; $i++) {
                            $seg = $segments[$i]
                            $existing = $null
                            foreach ($child in $parent.Nodes) { if ($child.Text -eq $seg) { $existing = $child; break } }
                            if ($i -eq $segments.Count - 1) {
                                $node = New-Object System.Windows.Forms.TreeNode($seg)
                                $node.Tag = $fullPath
                                [void]$parent.Nodes.Add($node)
                            } else {
                                if ($null -eq $existing) {
                                    $dirNode = New-Object System.Windows.Forms.TreeNode($seg)
                                    $partialPath = Join-Path $gv ($segments[0..$i] -join '\')
                                    $dirNode.Tag = $partialPath
                                    [void]$parent.Nodes.Add($dirNode)
                                    $parent = $dirNode
                                } else {
                                    $parent = $existing
                                }
                            }
                        }
                    }
                    [void]$tree.Nodes.Add($rootNode)
                }
                $tree.ExpandAll()

                $jInfo.tab.Text = "{0} ({1})" -f (Get-FieldLabelText $jName), $matched.Count
            }
        }
    }
}

# =============================================================================
# Record list
# =============================================================================

function Update-RecordList {
    $name = [string]$cmbSource.SelectedItem
    if (-not $name) {
        $listRecords.Items.Clear()
        $lblCount.Text = ''
        return
    }

    $records = $script:allData[$name]
    $displayCols = Get-DisplayColumns $name

    $listRecords.BeginUpdate()
    $listRecords.Items.Clear()
    $filter = $txtFilter.Text.Trim()
    $script:filteredIndices = @()

    for ($i = 0; $i -lt $records.Count; $i++) {
        $rec = $records[$i]
        if ($filter) {
            $allText = ($rec.PSObject.Properties | Where-Object { $_.Name -notlike '_*' } | ForEach-Object { Format-FieldValue $_.Value }) -join ' '
            if ($allText -notmatch [regex]::Escape($filter)) { continue }
        }
        $label = ($displayCols | ForEach-Object { Format-FieldValue (Get-ShinsaRecordValue -Record $rec -Name $_) }) -join ' | '
        $script:filteredIndices += $i
        [void]$listRecords.Items.Add($label)
    }
    $listRecords.EndUpdate()

    $total = $records.Count
    $shown = $script:filteredIndices.Count
    $lblCount.Text = if ($filter) { "$shown / $total" } else { "$total" }

    if ($listRecords.Items.Count -gt 0) { $listRecords.SelectedIndex = 0 }
}

function Update-Detail {
    $name = [string]$cmbSource.SelectedItem
    if (-not $name) { return }
    Update-DetailTab $name
    Update-JoinedTabs $name
}

# =============================================================================
# Conflict resolution dialog
# =============================================================================

function Show-ConflictDialog {
    param([array]$Conflicts)
    if ($Conflicts.Count -eq 0) { return $null }

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = 'Conflict Resolution'
    $dlg.Size = New-Object System.Drawing.Size(700, 400)
    $dlg.StartPosition = 'CenterParent'
    $dlg.Font = $script:Font
    $dlg.MinimumSize = New-Object System.Drawing.Size(500, 300)

    $currentIdx = 0
    $results = @{}

    $lblNav = New-Object System.Windows.Forms.Label
    $lblNav.Dock = 'Top'
    $lblNav.Height = 28
    $lblNav.TextAlign = 'MiddleCenter'

    $dgv = New-Object System.Windows.Forms.DataGridView
    $dgv.Dock = 'Fill'
    $dgv.AllowUserToAddRows = $false
    $dgv.AllowUserToDeleteRows = $false
    $dgv.SelectionMode = 'FullRowSelect'
    $dgv.AutoSizeColumnsMode = 'Fill'
    $dgv.RowHeadersVisible = $false
    [void]$dgv.Columns.Add('Field', 'Field')
    [void]$dgv.Columns.Add('Original', 'Original')
    [void]$dgv.Columns.Add('Local', 'Local (You)')
    [void]$dgv.Columns.Add('Remote', 'Remote (Source)')
    $colKeep = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
    $colKeep.Name = 'Keep'
    $colKeep.HeaderText = 'Keep'
    $colKeep.Items.AddRange(@('Local', 'Remote'))
    $colKeep.FlatStyle = 'Standard'
    [void]$dgv.Columns.Add($colKeep)
    $dgv.Columns['Field'].ReadOnly = $true
    $dgv.Columns['Original'].ReadOnly = $true
    $dgv.Columns['Local'].ReadOnly = $true
    $dgv.Columns['Remote'].ReadOnly = $true

    $btnPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $btnPanel.Dock = 'Bottom'
    $btnPanel.Height = 36
    $btnPanel.FlowDirection = 'RightToLeft'
    $btnPanel.Padding = New-Object System.Windows.Forms.Padding(4)

    $btnApply = New-Object System.Windows.Forms.Button; $btnApply.Text = 'Apply All'; $btnApply.Width = 90
    $btnCancel = New-Object System.Windows.Forms.Button; $btnCancel.Text = 'Cancel'; $btnCancel.Width = 90
    $btnPrev = New-Object System.Windows.Forms.Button; $btnPrev.Text = '<< Prev'; $btnPrev.Width = 70
    $btnNext = New-Object System.Windows.Forms.Button; $btnNext.Text = 'Next >>'; $btnNext.Width = 70
    $btnPanel.Controls.AddRange(@($btnApply, $btnCancel, $btnNext, $btnPrev))

    function Load-ConflictRecord {
        param([int]$Idx)
        $c = $Conflicts[$Idx]
        $lblNav.Text = "Record: $($c.key_value)  ($($Idx + 1) / $($Conflicts.Count))"
        $dgv.Rows.Clear()
        foreach ($f in $c.fields) {
            $rowIdx = $dgv.Rows.Add($f.field, $f.original, $f.local, $f.remote)
            $key = "$($c.source_name)|$($c.key_value)|$($f.field)"
            $dgv.Rows[$rowIdx].Cells['Keep'].Value = if ($results.ContainsKey($key)) { $results[$key] } else { 'Local' }
        }
    }

    function Save-CurrentChoices {
        $c = $Conflicts[$currentIdx]
        for ($r = 0; $r -lt $dgv.Rows.Count; $r++) {
            $fn = [string]$dgv.Rows[$r].Cells['Field'].Value
            $choice = [string]$dgv.Rows[$r].Cells['Keep'].Value
            if ([string]::IsNullOrWhiteSpace($choice)) { $choice = 'Local' }
            $key = "$($c.source_name)|$($c.key_value)|$fn"
            $results[$key] = $choice
        }
    }

    $btnPrev.Add_Click({ Save-CurrentChoices; if ($currentIdx -gt 0) { $currentIdx--; Load-ConflictRecord $currentIdx } })
    $btnNext.Add_Click({ Save-CurrentChoices; if ($currentIdx -lt $Conflicts.Count - 1) { $currentIdx++; Load-ConflictRecord $currentIdx } })
    $btnApply.Add_Click({ Save-CurrentChoices; $dlg.DialogResult = 'OK'; $dlg.Close() })
    $btnCancel.Add_Click({ $dlg.DialogResult = 'Cancel'; $dlg.Close() })

    $dlg.Controls.Add($dgv)
    $dlg.Controls.Add($lblNav)
    $dlg.Controls.Add($btnPanel)
    $dlg.AcceptButton = $btnApply
    $dlg.CancelButton = $btnCancel

    Load-ConflictRecord 0

    if ($dlg.ShowDialog() -eq 'OK') {
        # Return resolved results
        $resolved = @()
        foreach ($c in $Conflicts) {
            foreach ($f in $c.fields) {
                $key = "$($c.source_name)|$($c.key_value)|$($f.field)"
                $choice = if ($results.ContainsKey($key)) { $results[$key] } else { 'Local' }
                $resolved += [pscustomobject]@{
                    source_name = $c.source_name
                    key_value   = $c.key_value
                    field       = $f.field
                    keep        = $choice
                    local       = $f.local
                    remote      = $f.remote
                    original    = $f.original
                }
            }
        }
        return $resolved
    }
    return $null
}

function Apply-ConflictResolutions {
    param([array]$Resolutions)
    $logPath = Join-Path $script:Paths.JsonRoot 'changelog.jsonl'
    foreach ($res in $Resolutions) {
        $name = $res.source_name
        $src = Get-SourceConfig $name
        $keyColumn = [string]$src.key_column
        $records = $script:allData[$name]
        foreach ($rec in $records) {
            $kv = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $rec -Name $keyColumn)
            if ($kv -ne $res.key_value) { continue }
            $val = if ($res.keep -eq 'Remote') { $res.remote } else { $res.local }
            Set-ShinsaRecordValue -Record $rec -Name $res.field -Value $val
            $origin = if ($res.keep -eq 'Remote') { 'conflict_remote' } else { 'conflict_local' }
            Write-ShinsaChangeLog -LogPath $logPath -SourceName $name -KeyValue $kv -FieldName $res.field `
                -OldValue $res.original -NewValue $val -Origin $origin
            Add-LogEntry ([pscustomobject]@{ ts = (Get-Date).ToString('o'); src = $name; key = $kv; field = $res.field; old = $res.original; new = $val; origin = $origin })
            break
        }
        $jp = Join-Path $script:Paths.JsonRoot $src.file
        Write-ShinsaJson -Path $jp -Data $records
    }
}

# =============================================================================
# Undo
# =============================================================================

function Invoke-Undo {
    if ($script:undoStack.Count -eq 0) { $statusBar.Text = 'Nothing to undo.'; return }
    $entry = $script:undoStack[$script:undoStack.Count - 1]
    $script:undoStack.RemoveAt($script:undoStack.Count - 1)

    $name = $entry.source
    $src = Get-SourceConfig $name
    $keyColumn = [string]$src.key_column
    $records = $script:allData[$name]

    foreach ($rec in $records) {
        $kv = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $rec -Name $keyColumn)
        if ($kv -ne $entry.key) { continue }
        Set-ShinsaRecordValue -Record $rec -Name $entry.field -Value $entry.oldValue
        break
    }

    $jp = Join-Path $script:Paths.JsonRoot $src.file
    Write-ShinsaJson -Path $jp -Data $records

    $logPath = Join-Path $script:Paths.JsonRoot 'changelog.jsonl'
    Write-ShinsaChangeLog -LogPath $logPath -SourceName $name -KeyValue $entry.key -FieldName $entry.field `
        -OldValue $entry.newValue -NewValue $entry.oldValue -Origin 'undo'
    Add-LogEntry ([pscustomobject]@{ ts = (Get-Date).ToString('o'); src = $name; key = $entry.key; field = $entry.field; old = $entry.newValue; new = $entry.oldValue; origin = 'undo' })

    # Refresh display if we're on the right source
    $currentName = [string]$cmbSource.SelectedItem
    if ($currentName -eq $name) {
        $script:dirty = $false
        Update-Detail
    }
    $statusBar.Text = "Undone: $($entry.field)"
}

# =============================================================================
# Polling timer (real-time sync)
# =============================================================================

function Initialize-SourceTimestamps {
    $script:lastSourceTimestamps = @{}
    foreach ($name in (Get-ShinsaSourceNames -Config $script:Config)) {
        $path = Get-ShinsaSourcePath -Config $script:Config -SourceName $name
        if ($path -and (Test-Path $path)) {
            $item = Get-Item $path -ErrorAction SilentlyContinue
            if ($item) {
                $ts = if ($item.PSIsContainer) {
                    try { (Get-ChildItem $path -Recurse -File -ErrorAction SilentlyContinue | Measure-Object -Property LastWriteTime -Maximum).Maximum } catch { $item.LastWriteTime }
                } else { $item.LastWriteTime }
                $script:lastSourceTimestamps[$name] = $ts
            }
        }
    }
}

$syncTimer = New-Object System.Windows.Forms.Timer
$syncCfg = if ($null -ne $script:Config.sync) { ConvertTo-ShinsaMap -InputObject $script:Config.sync } else { @{} }
$syncTimer.Interval = if ($syncCfg.Contains('interval_ms')) { [int]$syncCfg['interval_ms'] } else { 5000 }
$syncAutoEnabled = if ($syncCfg.Contains('auto_sync')) { [bool]$syncCfg['auto_sync'] } else { $true }

$syncTimer.Add_Tick({
    try {
        $changed = $false
        foreach ($name in (Get-ShinsaSourceNames -Config $script:Config)) {
            $path = Get-ShinsaSourcePath -Config $script:Config -SourceName $name
            if (-not $path -or -not (Test-Path $path)) { continue }
            $item = Get-Item $path -ErrorAction SilentlyContinue
            if (-not $item) { continue }
            $ts = if ($item.PSIsContainer) {
                try { (Get-ChildItem $path -Recurse -File -ErrorAction SilentlyContinue | Measure-Object -Property LastWriteTime -Maximum).Maximum } catch { $item.LastWriteTime }
            } else { $item.LastWriteTime }
            $lastKnown = $script:lastSourceTimestamps[$name]
            if ($null -eq $lastKnown -or $ts -gt $lastKnown) {
                $changed = $true
                break
            }
        }
        if ($changed) {
            Invoke-GuiSync
        }
    } catch {
        $statusBar.Text = "Auto-sync error: $($_.Exception.Message)"
    }
})

# =============================================================================
# Change Log panel
# =============================================================================

$script:logColorMap = @{
    'local'           = [System.Drawing.Color]::Black
    'remote'          = [System.Drawing.Color]::DarkBlue
    'conflict_local'  = [System.Drawing.Color]::DarkGoldenrod
    'conflict_remote' = [System.Drawing.Color]::DarkGoldenrod
    'undo'            = [System.Drawing.Color]::Gray
}

function Format-LogLine {
    param([pscustomobject]$Entry)
    $ts = ''
    if ($Entry.ts) {
        try { $ts = ([datetime]$Entry.ts).ToString('HH:mm:ss') } catch { $ts = [string]$Entry.ts }
    }
    $change = ''
    if ($Entry.old -or $Entry.new) { $change = "$($Entry.old)->$($Entry.new)" }
    @($ts, $Entry.src, $Entry.key, $Entry.field, $change, $Entry.origin) -join ' | '
}

function Add-LogEntry {
    param([pscustomobject]$Entry)
    $line = Format-LogLine $Entry
    $logList.Items.Insert(0, $line)
}

function Load-ChangeLog {
    $logPath = Join-Path $script:Paths.JsonRoot 'changelog.jsonl'
    $logList.Items.Clear()
    if (-not (Test-Path $logPath)) { return }
    $lines = @(Get-Content -Path $logPath -Encoding UTF8 -ErrorAction SilentlyContinue)
    # Load last 200 entries (newest last in file, we insert at top)
    $start = if ($lines.Count -gt 200) { $lines.Count - 200 } else { 0 }
    for ($i = $start; $i -lt $lines.Count; $i++) {
        $line = $lines[$i].Trim()
        if (-not $line) { continue }
        try {
            $entry = $line | ConvertFrom-Json
            Add-LogEntry $entry
        } catch { }
    }
}

# =============================================================================
# Events
# =============================================================================

$cmbSource.Add_SelectedIndexChanged({
    Save-CurrentRecordEdits
    $name = [string]$cmbSource.SelectedItem
    $txtFilter.Clear()
    Build-FieldEditors $name
    Build-JoinedTabs $name
    Update-RecordList
})

$listRecords.Add_SelectedIndexChanged({
    Save-CurrentRecordEdits
    Update-Detail
})

$listRecords.Add_DoubleClick({
    $name = [string]$cmbSource.SelectedItem
    $idx = $listRecords.SelectedIndex
    if (-not $name -or $idx -lt 0 -or $idx -ge $script:filteredIndices.Count) { return }
    $recIdx = $script:filteredIndices[$idx]
    $rec = $script:allData[$name][$recIdx]
    $sourcePath = Get-ShinsaSourcePath -Config $script:Config -SourceName $name
    if (-not [string]::IsNullOrWhiteSpace($sourcePath) -and (Test-Path $sourcePath)) {
        try { Start-Process $sourcePath } catch { }
    }
})

$txtFilter.Add_TextChanged({ Update-RecordList })
$txtFilter.Add_KeyDown({
    if ($_.KeyCode -eq 'Enter' -or $_.KeyCode -eq 'Down') {
        $listRecords.Focus()
        if ($listRecords.Items.Count -gt 0 -and $listRecords.SelectedIndex -lt 0) { $listRecords.SelectedIndex = 0 }
        $_.Handled = $true
    }
})

$form.Add_KeyDown({
    param($s, $e)
    if ($e.Control -and $e.KeyCode -eq 'F') { $txtFilter.Focus(); $txtFilter.SelectAll(); $e.Handled = $true; $e.SuppressKeyPress = $true }
    if ($e.Control -and $e.KeyCode -eq 'Z') { Invoke-Undo; $e.Handled = $true; $e.SuppressKeyPress = $true }
})

$menuSync.Add_Click({ try { Invoke-GuiSync } catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null } })
$menuReflect.Add_Click({ try { Invoke-GuiWriteback } catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null } })
$menuSettings.Add_Click({ try { Show-SettingsDialog } catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null } })
$menuQuit.Add_Click({ $form.Close() })

$tsbSync.Add_Click({ try { Invoke-GuiSync } catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null } })
$tsbReflect.Add_Click({ try { Invoke-GuiWriteback } catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null } })
$tsbSettings.Add_Click({ try { Show-SettingsDialog } catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null } })

# =============================================================================
# Init
# =============================================================================

Load-AllData
Load-ChangeLog

$script:initUiState = Get-UiState
if ($script:initUiState.Contains('window_width') -and $script:initUiState.Contains('window_height')) {
    $form.StartPosition = 'Manual'
    $form.Size = New-Object System.Drawing.Size([int]$script:initUiState.window_width, [int]$script:initUiState.window_height)
    if ($script:initUiState.Contains('window_left') -and $script:initUiState.Contains('window_top')) {
        $form.Location = New-Object System.Drawing.Point([int]$script:initUiState.window_left, [int]$script:initUiState.window_top)
    }
} else {
    $form.StartPosition = 'CenterScreen'
    $form.Size = New-Object System.Drawing.Size([int]$script:Config.gui.window_width, [int]$script:Config.gui.window_height)
}
if ($script:initUiState.Contains('window_state') -and [string]$script:initUiState['window_state'] -eq 'Maximized') { $form.WindowState = 'Maximized' }

$script:mainSplitDist = if ($script:initUiState.Contains('main_splitter_distance')) { [int]$script:initUiState['main_splitter_distance'] } else { 300 }
$script:outerSplitDist = if ($script:initUiState.Contains('outer_splitter_distance')) { [int]$script:initUiState['outer_splitter_distance'] } else { -1 }

foreach ($name in (Get-ShinsaPrimarySourceNames -Config $script:Config)) {
    [void]$cmbSource.Items.Add($name)
}

$form.Add_Shown({
    Set-SafeSplitterLayout -Ctl $mainSplit -Min1 200 -Min2 300 -Preferred $script:mainSplitDist
    $outerPref = if ($script:outerSplitDist -gt 0) { $script:outerSplitDist } else { [Math]::Max(400, $outerSplit.ClientSize.Width - 280) }
    Set-SafeSplitterLayout -Ctl $outerSplit -Min1 400 -Min2 200 -Preferred $outerPref

    $selSource = ''
    if ($script:initUiState.Contains('selected_source')) { $selSource = [string]$script:initUiState['selected_source'] }
    $idx = $cmbSource.Items.IndexOf($selSource)
    if ($idx -ge 0) { $cmbSource.SelectedIndex = $idx }
    elseif ($cmbSource.Items.Count -gt 0) { $cmbSource.SelectedIndex = 0 }

    if ($script:initUiState.Contains('search_text') -and -not [string]::IsNullOrWhiteSpace([string]$script:initUiState['search_text'])) {
        $txtFilter.Text = [string]$script:initUiState['search_text']
    }

    $form.TopMost = $true
    [void][Win32.NativeMethods]::SetForegroundWindow($form.Handle)
    [void][Win32.NativeMethods]::ShowWindow($form.Handle, 9)
    $form.TopMost = $false

    # Start polling timer
    Initialize-SourceTimestamps
    if ($syncAutoEnabled) { $syncTimer.Start() }
})

$form.Add_FormClosing({
    $syncTimer.Stop()
    Save-CurrentRecordEdits
    Save-UiState
})

[void]$form.ShowDialog()
