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

# Auto-sync on startup if JSON missing
$needsSync = $false
foreach ($src in @($script:Config.sources)) {
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

# --- Data ---
$script:sourceConfigs = @{}   # name -> source config object
$script:allData = [ordered]@{} # name -> records array
$script:fieldEditors = @{}
$script:filteredIndices = @()
$script:cacheState = $null

# --- Helpers ---

function Get-FieldLabelText {
    param([string]$FieldName)
    $words = ($FieldName -replace '_', ' ') -split ' '
    ($words | ForEach-Object { if ($_.Length -gt 0) { $_.Substring(0,1).ToUpper() + $_.Substring(1) } else { $_ } }) -join ' '
}

function Format-FieldValue {
    param($Value)
    if ($Value -is [System.Array] -or $Value -is [System.Collections.IList]) {
        return ($Value | ForEach-Object { [string]$_ }) -join '; '
    }
    return ConvertTo-ShinsaString -Value $Value
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
    $script:sourceConfigs = @{}
    $script:allData = [ordered]@{}
    foreach ($src in @($script:Config.sources)) {
        $name = [string]$src.name
        $script:sourceConfigs[$name] = $src
        $jp = Join-Path $script:Paths.JsonRoot $src.file
        if (Test-Path $jp) {
            $script:allData[$name] = @((Read-ShinsaJson -Path $jp))
        } else {
            $script:allData[$name] = @()
        }
    }
    $script:cacheState = Read-ShinsaCache -Paths $script:Paths
}

function Get-SourceConfig {
    param([string]$Name)
    if ($script:sourceConfigs.ContainsKey($Name)) { return $script:sourceConfigs[$Name] }
    return $null
}

function Get-SourceView {
    param([string]$Name)
    $src = Get-SourceConfig $Name
    if ($src -and $null -ne $src.view) { return [string]$src.view }
    return 'fields'
}

function Get-DisplayColumns {
    param([string]$Name)
    $src = Get-SourceConfig $Name
    if ($src -and $null -ne $src.display_columns) {
        $dc = @($src.display_columns)
        if ($dc.Count -gt 0) { return $dc }
    }
    $recs = $script:allData[$Name]
    if ($recs -and $recs.Count -gt 0) { return @($recs[0].PSObject.Properties | Select-Object -First 3 | ForEach-Object { $_.Name }) }
    return @()
}

function Get-DetailColumns {
    param([string]$Name)
    $src = Get-SourceConfig $Name
    if ($src -and $null -ne $src.detail_columns) {
        $dc = @($src.detail_columns)
        if ($dc.Count -gt 0) { return $dc }
    }
    $recs = $script:allData[$Name]
    if ($recs -and $recs.Count -gt 0) { return @($recs[0].PSObject.Properties | ForEach-Object { $_.Name }) }
    return @()
}

function Get-EditableColumns {
    param([string]$Name)
    $src = Get-SourceConfig $Name
    if ($src -and $null -ne $src.editable_columns) { return @($src.editable_columns) }
    return @()
}

function Get-MultilineColumns {
    param([string]$Name)
    $src = Get-SourceConfig $Name
    if ($src -and $null -ne $src.multiline_columns) { return @($src.multiline_columns) }
    return @()
}

function Save-SourceData {
    param([string]$Name, [switch]$Quiet)
    $editableCols = Get-EditableColumns $Name
    if ($editableCols.Count -eq 0) { return }

    $idx = $listRecords.SelectedIndex
    if ($idx -lt 0 -or $idx -ge $script:filteredIndices.Count) { return }

    $recIdx = $script:filteredIndices[$idx]
    $rec = $script:allData[$Name][$recIdx]
    foreach ($fn in $editableCols) {
        if (-not $script:fieldEditors.ContainsKey($fn)) { continue }
        Set-ShinsaRecordValue -Record $rec -Name $fn -Value $script:fieldEditors[$fn].Text.Trim()
    }
    $src = Get-SourceConfig $Name
    $jp = Join-Path $script:Paths.JsonRoot $src.file
    Write-ShinsaJson -Path $jp -Data $script:allData[$Name]
    if (-not $Quiet) { $lblStatus.Text = 'Saved.' }
}

function Invoke-GuiSync {
    $name = [string]$cmbSource.SelectedItem
    if ($name) { Save-SourceData $name -Quiet }
    & (Join-Path $script:AppRoot 'scripts\Sync-Data.ps1')
    Load-AllData
    Update-RecordList
    $lblStatus.Text = 'Sync completed.'
}

function Invoke-GuiWriteback {
    $name = [string]$cmbSource.SelectedItem
    if ($name) { Save-SourceData $name -Quiet }
    $plan = Get-ShinsaTableWritebackPlan -Config $script:Config -Paths $script:Paths
    if ($plan.case_count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('No changes to reflect.', 'shinsa') | Out-Null; return
    }
    $lines = @()
    foreach ($c in @($plan.changes)) { $lines += ('{0}: {1}' -f $c.case_id, (@($c.changes.PSObject.Properties.Name) -join ', ')) }
    $msg = "Reflect {0} case(s) / {1} field(s) to the source table?`n`n{2}" -f $plan.case_count, $plan.change_count, ($lines -join "`n")
    $r = [System.Windows.Forms.MessageBox]::Show($msg, 'shinsa', 'OKCancel', 'Question')
    if ($r -ne 'OK') { return }
    Invoke-ShinsaTableWriteback -Config $script:Config -Paths $script:Paths -Plan $plan
    Write-ShinsaJson -Path $script:Paths.TableJsonPath -Data @(Import-ShinsaTableRecords -Config $script:Config -Paths $script:Paths | Sort-Object case_id)
    Load-AllData
    Update-RecordList
    $lblStatus.Text = 'Reflected to source.'
}

# =============================================================================
# UI
# =============================================================================
$script:M = 8

$form = New-Object System.Windows.Forms.Form
$form.Text = $script:Config.gui.title
$form.MinimumSize = New-Object System.Drawing.Size(700, 500)
$form.Font = New-Object System.Drawing.Font($script:Config.gui.font_name, [single]$script:Config.gui.font_size)
$form.Padding = New-Object System.Windows.Forms.Padding($script:M)
$form.AutoScaleMode = 'Dpi'
$form.KeyPreview = $true

# --- Menu ---
$menuBar = New-Object System.Windows.Forms.MenuStrip
$menuFile = New-Object System.Windows.Forms.ToolStripMenuItem('File(&F)')
$menuSync = New-Object System.Windows.Forms.ToolStripMenuItem('Reload from sources')
$menuSync.ShortcutKeys = [System.Windows.Forms.Keys]::F5
$menuReflect = New-Object System.Windows.Forms.ToolStripMenuItem('Reflect to source table...')
$menuReflect.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::W
$menuQuit = New-Object System.Windows.Forms.ToolStripMenuItem('Quit')
$menuQuit.ShortcutKeys = [System.Windows.Forms.Keys]::Alt -bor [System.Windows.Forms.Keys]::F4
[void]$menuFile.DropDownItems.AddRange(@($menuSync, (New-Object System.Windows.Forms.ToolStripSeparator), $menuReflect, (New-Object System.Windows.Forms.ToolStripSeparator), $menuQuit))
[void]$menuBar.Items.Add($menuFile)
$form.MainMenuStrip = $menuBar

# --- Main split ---
$mainSplit = New-Object System.Windows.Forms.SplitContainer
$mainSplit.Dock = 'Fill'
$mainSplit.Orientation = 'Vertical'

# --- Left pane ---
$leftPanel = New-Object System.Windows.Forms.TableLayoutPanel
$leftPanel.Dock = 'Fill'
$leftPanel.RowCount = 3
$leftPanel.ColumnCount = 1
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

$listRecords = New-Object System.Windows.Forms.ListBox
$listRecords.Dock = 'Fill'
$listRecords.Font = New-Object System.Drawing.Font('Consolas', 10)
$listRecords.IntegralHeight = $false

$leftPanel.Controls.Add($cmbSource, 0, 0)
$leftPanel.Controls.Add($txtFilter, 0, 1)
$leftPanel.Controls.Add($listRecords, 0, 2)
$mainSplit.Panel1.Controls.Add($leftPanel)

# --- Right pane: TabControl ---
$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Dock = 'Fill'
$tabs.Appearance = 'FlatButtons'
$tabs.ItemSize = New-Object System.Drawing.Size(0, 1)
$tabs.SizeMode = 'Fixed'

# Tab: fields
$tabFields = New-Object System.Windows.Forms.TabPage
$tabFields.Text = 'Fields'
$tabFields.Padding = New-Object System.Windows.Forms.Padding(0)

$pnlButtons = New-Object System.Windows.Forms.FlowLayoutPanel
$pnlButtons.Dock = 'Top'
$pnlButtons.Height = 34
$pnlButtons.FlowDirection = 'LeftToRight'
$pnlButtons.Padding = New-Object System.Windows.Forms.Padding(0, 1, 0, 1)

$pnlFields = New-Object System.Windows.Forms.Panel
$pnlFields.Dock = 'Fill'
$pnlFields.AutoScroll = $true
$pnlFields.Padding = New-Object System.Windows.Forms.Padding(0, $script:M, 0, 0)

$tabFields.Controls.Add($pnlFields)
$tabFields.Controls.Add($pnlButtons)

# Tab: mail
$tabMail = New-Object System.Windows.Forms.TabPage
$tabMail.Text = 'Mail'

$mailHeaderPanel = New-Object System.Windows.Forms.TableLayoutPanel
$mailHeaderPanel.Dock = 'Top'
$mailHeaderPanel.Height = 80
$mailHeaderPanel.ColumnCount = 2
$mailHeaderPanel.RowCount = 3
[void]$mailHeaderPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('AutoSize')))
[void]$mailHeaderPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 100)))

$mailSubjectLabel = New-Object System.Windows.Forms.Label; $mailSubjectLabel.Text = 'Subject:'; $mailSubjectLabel.AutoSize = $true; $mailSubjectLabel.Anchor = 'Left'; $mailSubjectLabel.Margin = New-Object System.Windows.Forms.Padding(0, 4, $script:M, 0)
$mailSubjectText  = New-Object System.Windows.Forms.Label; $mailSubjectText.AutoSize = $true; $mailSubjectText.Anchor = 'Left'; $mailSubjectText.Margin = New-Object System.Windows.Forms.Padding(0, 4, 0, 0)
$mailFromLabel    = New-Object System.Windows.Forms.Label; $mailFromLabel.Text = 'From:'; $mailFromLabel.AutoSize = $true; $mailFromLabel.Anchor = 'Left'; $mailFromLabel.Margin = New-Object System.Windows.Forms.Padding(0, 2, $script:M, 0)
$mailFromText     = New-Object System.Windows.Forms.Label; $mailFromText.AutoSize = $true; $mailFromText.Anchor = 'Left'; $mailFromText.Margin = New-Object System.Windows.Forms.Padding(0, 2, 0, 0)
$mailDateLabel    = New-Object System.Windows.Forms.Label; $mailDateLabel.Text = 'Date:'; $mailDateLabel.AutoSize = $true; $mailDateLabel.Anchor = 'Left'; $mailDateLabel.Margin = New-Object System.Windows.Forms.Padding(0, 2, $script:M, 0)
$mailDateText     = New-Object System.Windows.Forms.Label; $mailDateText.AutoSize = $true; $mailDateText.Anchor = 'Left'; $mailDateText.Margin = New-Object System.Windows.Forms.Padding(0, 2, 0, 0)
$mailHeaderPanel.Controls.Add($mailSubjectLabel, 0, 0); $mailHeaderPanel.Controls.Add($mailSubjectText, 1, 0)
$mailHeaderPanel.Controls.Add($mailFromLabel, 0, 1);    $mailHeaderPanel.Controls.Add($mailFromText, 1, 1)
$mailHeaderPanel.Controls.Add($mailDateLabel, 0, 2);    $mailHeaderPanel.Controls.Add($mailDateText, 1, 2)

$mailPreviewBox = New-Object System.Windows.Forms.TextBox
$mailPreviewBox.Dock = 'Fill'
$mailPreviewBox.Multiline = $true
$mailPreviewBox.ScrollBars = 'Vertical'
$mailPreviewBox.ReadOnly = $true
$mailPreviewBox.BackColor = [System.Drawing.SystemColors]::Window

$mailAttachLabel = New-Object System.Windows.Forms.Label
$mailAttachLabel.Dock = 'Bottom'
$mailAttachLabel.Height = 20
$mailAttachLabel.Text = 'Attachments:'
$mailAttachLabel.ForeColor = [System.Drawing.Color]::DimGray

$mailAttachList = New-Object System.Windows.Forms.ListBox
$mailAttachList.Dock = 'Bottom'
$mailAttachList.Height = 60
$mailAttachList.Font = New-Object System.Drawing.Font('Consolas', 10)

$tabMail.Controls.Add($mailPreviewBox)
$tabMail.Controls.Add($mailAttachLabel)
$tabMail.Controls.Add($mailAttachList)
$tabMail.Controls.Add($mailHeaderPanel)

# Tab: tree
$tabTree = New-Object System.Windows.Forms.TabPage
$tabTree.Text = 'Tree'

$fileTree = New-Object System.Windows.Forms.TreeView
$fileTree.Dock = 'Fill'
$fileTree.HideSelection = $false
$fileTree.ShowLines = $true
$fileTree.ShowPlusMinus = $false
$fileTree.ShowRootLines = $true
$fileTree.PathSeparator = '\'

$tabTree.Controls.Add($fileTree)

[void]$tabs.TabPages.Add($tabFields)
[void]$tabs.TabPages.Add($tabMail)
[void]$tabs.TabPages.Add($tabTree)

$mainSplit.Panel2.Controls.Add($tabs)

# --- Bottom ---
$lblCount = New-Object System.Windows.Forms.Label
$lblCount.Dock = 'Bottom'
$lblCount.Height = 24
$lblCount.TextAlign = 'BottomRight'
$lblCount.ForeColor = [System.Drawing.Color]::DimGray

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Dock = 'Bottom'
$lblStatus.Height = 20
$lblStatus.TextAlign = 'BottomLeft'
$lblStatus.ForeColor = [System.Drawing.Color]::DimGray

$form.Controls.Add($mainSplit)
$form.Controls.Add($lblCount)
$form.Controls.Add($lblStatus)
$form.Controls.Add($menuBar)

# =============================================================================
# View switching
# =============================================================================

function Switch-View {
    param([string]$ViewType)
    switch ($ViewType) {
        'mail'  { $tabs.SelectedTab = $tabMail }
        'tree'  { $tabs.SelectedTab = $tabTree }
        default { $tabs.SelectedTab = $tabFields }
    }
}

# =============================================================================
# Fields view
# =============================================================================

function Build-FieldEditors {
    param([string]$SourceName)
    $pnlFields.SuspendLayout()
    $pnlFields.Controls.Clear()
    $script:fieldEditors = @{}

    $fields = Get-DetailColumns $SourceName
    if (-not $fields -or $fields.Count -eq 0) { $pnlFields.ResumeLayout(); return }

    $editableCols = Get-EditableColumns $SourceName
    $multilineCols = Get-MultilineColumns $SourceName
    $m = $script:M

    $g = $pnlFields.CreateGraphics()
    $maxLblChars = 20
    $lblW = 80
    foreach ($fn in $fields) {
        $label = Get-FieldLabelText -FieldName $fn
        $w = [int][math]::Ceiling($g.MeasureString($label, $pnlFields.Font).Width) + 8
        if ($w -gt $lblW) { $lblW = $w }
    }
    if ($lblW -gt 180) { $lblW = 180 }
    $g.Dispose()

    $txtX = $lblW + $m
    $y = $m
    $tip = New-Object System.Windows.Forms.ToolTip

    foreach ($fieldName in $fields) {
        $displayName = Get-FieldLabelText -FieldName $fieldName
        if ($displayName.Length -gt $maxLblChars) {
            $head = $displayName.Substring(0, [math]::Floor($maxLblChars / 2))
            $tail = $displayName.Substring($displayName.Length - [math]::Floor($maxLblChars / 2))
            $displayName = "$head...$tail"
        }

        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = $displayName
        $lbl.Location = New-Object System.Drawing.Point($m, ($y + 3))
        $lbl.Size = New-Object System.Drawing.Size($lblW, 22)
        $lbl.TextAlign = 'MiddleRight'
        if ($displayName -ne (Get-FieldLabelText -FieldName $fieldName)) { $tip.SetToolTip($lbl, (Get-FieldLabelText -FieldName $fieldName)) }
        $pnlFields.Controls.Add($lbl)

        $isMultiline = $fieldName -in $multilineCols
        $rowH = if ($isMultiline) { 48 } else { 22 }

        $txt = New-Object System.Windows.Forms.TextBox
        $txt.Location = New-Object System.Drawing.Point($txtX, $y)
        $txt.Size = New-Object System.Drawing.Size(($pnlFields.ClientSize.Width - $txtX - $m), $rowH)
        $txt.Anchor = 'Top,Left,Right'
        $txt.ReadOnly = ($fieldName -notin $editableCols)
        if ($txt.ReadOnly) { $txt.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245) }
        if ($isMultiline) { $txt.Multiline = $true; $txt.ScrollBars = 'Vertical' }
        $pnlFields.Controls.Add($txt)

        $script:fieldEditors[$fieldName] = $txt
        $y += $rowH + 6
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

function Update-FieldsDetail {
    param([string]$SourceName)
    $pnlButtons.Controls.Clear()
    $idx = $listRecords.SelectedIndex
    if ($idx -lt 0 -or $idx -ge $script:filteredIndices.Count) { Fill-FieldEditors $null; return }

    $recIdx = $script:filteredIndices[$idx]
    $rec = $script:allData[$SourceName][$recIdx]
    Fill-FieldEditors $rec

    $editableCols = Get-EditableColumns $SourceName
    if ($editableCols.Count -gt 0) {
        $btnSave = New-Object System.Windows.Forms.Button
        $btnSave.Text = 'Save'
        $btnSave.AutoSize = $true
        $btnSave.Height = 28
        $btnSave.TabStop = $false
        $btnSave.Add_Click({ Save-SourceData ([string]$cmbSource.SelectedItem) })
        $pnlButtons.Controls.Add($btnSave)
    }
}

# =============================================================================
# Mail view
# =============================================================================

function Update-MailDetail {
    param([string]$SourceName)
    $mailSubjectText.Text = ''
    $mailFromText.Text = ''
    $mailDateText.Text = ''
    $mailPreviewBox.Text = ''
    $mailAttachList.Items.Clear()

    $idx = $listRecords.SelectedIndex
    if ($idx -lt 0 -or $idx -ge $script:filteredIndices.Count) { return }
    $recIdx = $script:filteredIndices[$idx]
    $rec = $script:allData[$SourceName][$recIdx]

    $mailSubjectText.Text = Format-FieldValue (Get-ShinsaRecordValue -Record $rec -Name 'subject')
    $mailFromText.Text    = Format-FieldValue (Get-ShinsaRecordValue -Record $rec -Name 'sender_email')
    $mailDateText.Text    = Format-FieldValue (Get-ShinsaRecordValue -Record $rec -Name 'received_at')

    $bp = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $rec -Name 'body_path')
    if (-not [string]::IsNullOrWhiteSpace($bp) -and (Test-Path $bp)) {
        $mailPreviewBox.Text = Get-Content -Path $bp -Raw -Encoding UTF8
    }

    $aps = @((Get-ShinsaRecordValue -Record $rec -Name 'attachment_paths'))
    foreach ($ap in $aps) {
        $apStr = [string]$ap
        if ([string]::IsNullOrWhiteSpace($apStr)) { continue }
        [void]$mailAttachList.Items.Add([System.IO.Path]::GetFileName($apStr))
    }
}

# =============================================================================
# Tree view
# =============================================================================

function Update-TreeDetail {
    param([string]$SourceName)
    $fileTree.Nodes.Clear()
    $src = Get-SourceConfig $SourceName
    $records = $script:allData[$SourceName]

    $groupKey = 'folder_path'
    $relKey   = 'relative_path'
    $fullKey  = 'file_path'
    if ($null -ne $src.tree_group_key)     { $groupKey = [string]$src.tree_group_key }
    if ($null -ne $src.tree_relative_path) { $relKey   = [string]$src.tree_relative_path }
    if ($null -ne $src.tree_full_path)     { $fullKey  = [string]$src.tree_full_path }

    $groups = [ordered]@{}
    foreach ($idx in $script:filteredIndices) {
        $rec = $records[$idx]
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
        [void]$fileTree.Nodes.Add($rootNode)
    }
    $fileTree.ExpandAll()
}

# =============================================================================
# Logic
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
            $allText = ($rec.PSObject.Properties | ForEach-Object { Format-FieldValue $_.Value }) -join ' '
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

    $view = Get-SourceView $name
    if ($view -eq 'tree') { Update-TreeDetail $name }
}

function Update-Detail {
    $name = [string]$cmbSource.SelectedItem
    if (-not $name) { return }

    $view = Get-SourceView $name
    switch ($view) {
        'mail'   { Update-MailDetail $name }
        'tree'   { Update-TreeDetail $name }
        default  { Update-FieldsDetail $name }
    }
}

# =============================================================================
# Events
# =============================================================================

$cmbSource.Add_SelectedIndexChanged({
    $name = [string]$cmbSource.SelectedItem
    $txtFilter.Clear()
    $view = Get-SourceView $name
    Switch-View $view
    if ($view -eq 'fields') { Build-FieldEditors $name }
    Update-RecordList
})

$listRecords.Add_SelectedIndexChanged({ Update-Detail })

$listRecords.Add_KeyDown({
    if ($_.KeyCode -eq 'Enter') { & $script:openAction; $_.Handled = $true }
})
$listRecords.Add_DoubleClick({ & $script:openAction })

$script:openAction = {
    $name = [string]$cmbSource.SelectedItem
    $idx = $listRecords.SelectedIndex
    if (-not $name -or $idx -lt 0 -or $idx -ge $script:filteredIndices.Count) { return }
    $recIdx = $script:filteredIndices[$idx]
    $rec = $script:allData[$name][$recIdx]
    $view = Get-SourceView $name

    if ($view -eq 'mail') {
        $msg = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $rec -Name 'msg_path')
        if (-not [string]::IsNullOrWhiteSpace($msg) -and (Test-Path $msg)) { Start-Process $msg; return }
        $bp = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $rec -Name 'body_path')
        if (-not [string]::IsNullOrWhiteSpace($bp) -and (Test-Path $bp)) { Start-Process $bp }
        return
    }

    if ($view -eq 'tree') {
        $src = Get-SourceConfig $name
        $fullKey = 'file_path'
        if ($null -ne $src.tree_full_path) { $fullKey = [string]$src.tree_full_path }
        $fp = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $rec -Name $fullKey)
        if (-not [string]::IsNullOrWhiteSpace($fp) -and (Test-Path $fp)) { Start-Process $fp }
        return
    }
}

$mailAttachList.Add_DoubleClick({
    $aiIdx = $mailAttachList.SelectedIndex
    if ($aiIdx -lt 0) { return }
    $name = [string]$cmbSource.SelectedItem
    $idx = $listRecords.SelectedIndex
    if ($idx -lt 0 -or $idx -ge $script:filteredIndices.Count) { return }
    $recIdx = $script:filteredIndices[$idx]
    $rec = $script:allData[$name][$recIdx]
    $aps = @((Get-ShinsaRecordValue -Record $rec -Name 'attachment_paths'))
    $ci = 0
    foreach ($ap in $aps) {
        $apStr = [string]$ap
        if ([string]::IsNullOrWhiteSpace($apStr)) { continue }
        if ($ci -eq $aiIdx) {
            try { Start-Process $apStr } catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null }
            return
        }
        $ci++
    }
})

$fileTree.Add_NodeMouseDoubleClick({
    param($s, $e)
    try {
        $path = [string]$e.Node.Tag
        if (-not [string]::IsNullOrWhiteSpace($path) -and (Test-Path $path)) { Start-Process $path }
    } catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null }
})
$fileTree.Add_BeforeCollapse({ $_.Cancel = $true })

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
    if ($e.Control -and $e.KeyCode -eq 'S') { Save-SourceData ([string]$cmbSource.SelectedItem); $e.Handled = $true; $e.SuppressKeyPress = $true }
})

$menuSync.Add_Click({ try { Invoke-GuiSync } catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null } })
$menuReflect.Add_Click({ try { Invoke-GuiWriteback } catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null } })
$menuQuit.Add_Click({ $form.Close() })

# =============================================================================
# Init
# =============================================================================

Load-AllData

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

foreach ($name in $script:allData.Keys) { [void]$cmbSource.Items.Add($name) }

$form.Add_Shown({
    Set-SafeSplitterLayout -Ctl $mainSplit -Min1 200 -Min2 300 -Preferred $script:mainSplitDist

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
})

$form.Add_FormClosing({
    $name = [string]$cmbSource.SelectedItem
    if ($name) { Save-SourceData $name -Quiet }
    Save-UiState
})

[void]$form.ShowDialog()
