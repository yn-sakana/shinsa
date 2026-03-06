Import-Module (Join-Path $PSScriptRoot '..\scripts\Common.psm1') -Force -DisableNameChecking
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$appRoot = Split-Path -Parent $PSScriptRoot
$config = Get-ShinsaConfig -ScriptPath $MyInvocation.MyCommand.Path
$paths = Get-ShinsaPaths -Config $config

if (-not (Test-Path $paths.CaseIndexPath)) {
    [System.Windows.Forms.MessageBox]::Show('Index is missing. Start shinsa first.', 'shinsa') | Out-Null
    exit 1
}

$script:caseRecords = [System.Collections.Generic.List[object]]::new()
$script:mailIndex = @{}
$script:currentRecord = $null

function New-FieldLabel {
    param([string]$Text)
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.AutoSize = $true
    $label.Margin = New-Object System.Windows.Forms.Padding(3, 8, 3, 3)
    $label
}

function Read-CaseRecords {
    $records = Read-ShinsaJson -Path $paths.CaseIndexPath
    $list = [System.Collections.Generic.List[object]]::new()
    foreach ($record in @($records)) {
        $list.Add($record)
    }
    $list
}

function Read-MailIndex {
    $map = @{}
    foreach ($mail in @(Read-ShinsaJson -Path $paths.MailIndexPath)) {
        if (-not $map.ContainsKey($mail.case_id)) {
            $map[$mail.case_id] = @()
        }
        $map[$mail.case_id] += $mail
    }
    $map
}

function Clear-Detail {
    $caseIdBox.Text = ''
    $orgBox.Text = ''
    $statusBox.Text = ''
    $assignedBox.Text = ''
    $emailBox.Text = ''
    $missingBox.Text = ''
    $noteBox.Text = ''
    $mailBox.Text = ''
    $script:currentRecord = $null
}

function Update-Grid {
    param(
        [string]$FilterText,
        [string]$SelectedCaseId = ''
    )

    $table = New-Object System.Data.DataTable
    foreach ($columnName in $config.gui.visibleColumns) {
        [void]$table.Columns.Add($columnName)
    }

    foreach ($record in $script:caseRecords) {
        $searchText = @(
            $record.case_id
            $record.organization_name
            $record.contact_name
            $record.contact_email
            $record.status
            $record.assigned_to
        ) -join ' '

        if (-not [string]::IsNullOrWhiteSpace($FilterText) -and $searchText -notmatch [regex]::Escape($FilterText)) {
            continue
        }

        $row = $table.NewRow()
        foreach ($columnName in $config.gui.visibleColumns) {
            $row[$columnName] = [string]$record.$columnName
        }
        [void]$table.Rows.Add($row)
    }

    $grid.DataSource = $table

    if (-not [string]::IsNullOrWhiteSpace($SelectedCaseId)) {
        foreach ($row in $grid.Rows) {
            if ($row.Cells['case_id'].Value -eq $SelectedCaseId) {
                $row.Selected = $true
                $grid.CurrentCell = $row.Cells['case_id']
                break
            }
        }
    }
}

function Load-Detail {
    param([string]$CaseId)

    $record = $script:caseRecords | Where-Object { $_.case_id -eq $CaseId } | Select-Object -First 1
    if (-not $record) {
        Clear-Detail
        return
    }

    $script:currentRecord = $record
    $caseIdBox.Text = [string]$record.case_id
    $orgBox.Text = [string]$record.organization_name
    $statusBox.Text = [string]$record.status
    $assignedBox.Text = [string]$record.assigned_to
    $emailBox.Text = [string]$record.contact_email
    $missingBox.Text = [string](@($record.missing_documents) -join ', ')
    $noteBox.Text = [string]$record.review_note

    $latestMail = @($script:mailIndex[$record.case_id] | Sort-Object received_at -Descending | Select-Object -First 1)
    if ($latestMail.Count -gt 0) {
        $body = if (Test-Path $latestMail[0].body_path) {
            Get-Content $latestMail[0].body_path -Raw -Encoding UTF8
        } else {
            ''
        }

        $mailBox.Text = @(
            "Subject: $($latestMail[0].subject)"
            "From   : $($latestMail[0].sender_email)"
            "Date   : $($latestMail[0].received_at)"
            ''
            $body
        ) -join [Environment]::NewLine
    }
    else {
        $mailBox.Text = ''
    }
}

function Refresh-Indexes {
    param([string]$SelectedCaseId = '')
    $script:caseRecords = Read-CaseRecords
    $script:mailIndex = Read-MailIndex
    Update-Grid -FilterText $searchBox.Text -SelectedCaseId $SelectedCaseId
    if (-not [string]::IsNullOrWhiteSpace($SelectedCaseId)) {
        Load-Detail -CaseId $SelectedCaseId
    }
}

function Save-CurrentReview {
    if (-not $script:currentRecord) { return }

    $state = Get-ReviewState -Paths $paths
    $existing = $state.reviews | Where-Object { $_.case_id -eq $script:currentRecord.case_id } | Select-Object -First 1
    if (-not $existing) {
        $existing = [pscustomobject]@{
            case_id = $script:currentRecord.case_id
            status = ''
            assigned_to = ''
            review_note = ''
            missing_documents = @()
        }
        $state.reviews += $existing
    }

    $existing.status = $statusBox.Text
    $existing.assigned_to = $assignedBox.Text
    $existing.review_note = $noteBox.Text
    $existing.missing_documents = @($missingBox.Text -split '\s*,\s*' | Where-Object { $_ })
    Save-ReviewState -Paths $paths -State $state

    $script:currentRecord.status = $existing.status
    $script:currentRecord.assigned_to = $existing.assigned_to
    $script:currentRecord.review_note = $existing.review_note
    $script:currentRecord.missing_documents = $existing.missing_documents
    Update-Grid -FilterText $searchBox.Text -SelectedCaseId $script:currentRecord.case_id
}

$form = New-Object System.Windows.Forms.Form
$form.Text = $config.gui.title
$form.StartPosition = 'CenterScreen'
$form.WindowState = 'Maximized'
$form.MinimumSize = New-Object System.Drawing.Size(960, 640)
$form.AutoScaleMode = 'Dpi'

$root = New-Object System.Windows.Forms.TableLayoutPanel
$root.Dock = 'Fill'
$root.RowCount = 2
$root.ColumnCount = 1
$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))

$toolbar = New-Object System.Windows.Forms.FlowLayoutPanel
$toolbar.Dock = 'Fill'
$toolbar.AutoSize = $true
$toolbar.WrapContents = $false
$toolbar.AutoScroll = $true
$toolbar.Padding = New-Object System.Windows.Forms.Padding(8)

$searchLabel = New-Object System.Windows.Forms.Label
$searchLabel.Text = 'Search'
$searchLabel.AutoSize = $true
$searchLabel.Margin = New-Object System.Windows.Forms.Padding(3, 10, 6, 3)

$searchBox = New-Object System.Windows.Forms.TextBox
$searchBox.Width = 280
$searchBox.Margin = New-Object System.Windows.Forms.Padding(0, 6, 12, 0)

$refreshButton = New-Object System.Windows.Forms.Button
$refreshButton.Text = 'Refresh'
$refreshButton.AutoSize = $true

$syncButton = New-Object System.Windows.Forms.Button
$syncButton.Text = 'Sync'
$syncButton.AutoSize = $true

$mailButton = New-Object System.Windows.Forms.Button
$mailButton.Text = 'Mail'
$mailButton.AutoSize = $true

$saveButton = New-Object System.Windows.Forms.Button
$saveButton.Text = 'Save Review'
$saveButton.AutoSize = $true

$writebackButton = New-Object System.Windows.Forms.Button
$writebackButton.Text = 'Writeback'
$writebackButton.AutoSize = $true

$toolbar.Controls.AddRange(@(
    $searchLabel,
    $searchBox,
    $refreshButton,
    $syncButton,
    $mailButton,
    $saveButton,
    $writebackButton
))

$split = New-Object System.Windows.Forms.SplitContainer
$split.Dock = 'Fill'
$split.Orientation = 'Vertical'
$split.Panel1MinSize = 420
$split.Panel2MinSize = 320
$split.SplitterDistance = 760

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Dock = 'Fill'
$grid.ReadOnly = $true
$grid.AutoSizeColumnsMode = 'Fill'
$grid.SelectionMode = 'FullRowSelect'
$grid.MultiSelect = $false
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.RowHeadersVisible = $false

$detail = New-Object System.Windows.Forms.TableLayoutPanel
$detail.Dock = 'Fill'
$detail.Padding = New-Object System.Windows.Forms.Padding(12)
$detail.ColumnCount = 2
$detail.RowCount = 8
$detail.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 90)))
$detail.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$detail.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$detail.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$detail.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$detail.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$detail.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$detail.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 45)))
$detail.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$detail.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 55)))

$caseIdBox = New-Object System.Windows.Forms.TextBox
$caseIdBox.ReadOnly = $true
$caseIdBox.Dock = 'Fill'

$orgBox = New-Object System.Windows.Forms.TextBox
$orgBox.ReadOnly = $true
$orgBox.Dock = 'Fill'

$statusBox = New-Object System.Windows.Forms.TextBox
$statusBox.Dock = 'Fill'

$assignedBox = New-Object System.Windows.Forms.TextBox
$assignedBox.Dock = 'Fill'

$emailBox = New-Object System.Windows.Forms.TextBox
$emailBox.ReadOnly = $true
$emailBox.Dock = 'Fill'

$missingBox = New-Object System.Windows.Forms.TextBox
$missingBox.Dock = 'Fill'

$noteBox = New-Object System.Windows.Forms.TextBox
$noteBox.Multiline = $true
$noteBox.ScrollBars = 'Vertical'
$noteBox.Dock = 'Fill'

$mailBox = New-Object System.Windows.Forms.TextBox
$mailBox.Multiline = $true
$mailBox.ScrollBars = 'Vertical'
$mailBox.ReadOnly = $true
$mailBox.Dock = 'Fill'

$detail.Controls.Add((New-FieldLabel 'Case ID'), 0, 0)
$detail.Controls.Add($caseIdBox, 1, 0)
$detail.Controls.Add((New-FieldLabel 'Org'), 0, 1)
$detail.Controls.Add($orgBox, 1, 1)
$detail.Controls.Add((New-FieldLabel 'Status'), 0, 2)
$detail.Controls.Add($statusBox, 1, 2)
$detail.Controls.Add((New-FieldLabel 'Assigned'), 0, 3)
$detail.Controls.Add($assignedBox, 1, 3)
$detail.Controls.Add((New-FieldLabel 'Email'), 0, 4)
$detail.Controls.Add($emailBox, 1, 4)
$detail.Controls.Add((New-FieldLabel 'Missing'), 0, 5)
$detail.Controls.Add($missingBox, 1, 5)
$detail.Controls.Add((New-FieldLabel 'Review Note'), 0, 6)
$detail.Controls.Add($noteBox, 0, 7)
$detail.SetColumnSpan($noteBox, 2)

$mailLabel = New-FieldLabel 'Latest Mail'
$detail.Controls.Add($mailLabel, 0, 8)
$detail.Controls.Add($mailBox, 0, 9)
$detail.SetColumnSpan($mailBox, 2)
$detail.RowCount = 10
$detail.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$detail.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 55)))

$split.Panel1.Controls.Add($grid)
$split.Panel2.Controls.Add($detail)

$root.Controls.Add($toolbar, 0, 0)
$root.Controls.Add($split, 0, 1)
$form.Controls.Add($root)

$searchBox.Add_TextChanged({
    Update-Grid -FilterText $searchBox.Text -SelectedCaseId $caseIdBox.Text
})

$refreshButton.Add_Click({
    Refresh-Indexes -SelectedCaseId $caseIdBox.Text
})

$syncButton.Add_Click({
    & (Join-Path $appRoot 'scripts\Sync-Clone.ps1')
    & (Join-Path $appRoot 'scripts\Build-Index.ps1')
    Refresh-Indexes -SelectedCaseId $caseIdBox.Text
    [System.Windows.Forms.MessageBox]::Show('Sync completed.', 'shinsa') | Out-Null
})

$mailButton.Add_Click({
    & (Join-Path $appRoot 'scripts\Export-Outlook.ps1')
    & (Join-Path $appRoot 'scripts\Sync-Clone.ps1')
    & (Join-Path $appRoot 'scripts\Build-Index.ps1')
    Refresh-Indexes -SelectedCaseId $caseIdBox.Text
    [System.Windows.Forms.MessageBox]::Show('Mail import completed.', 'shinsa') | Out-Null
})

$saveButton.Add_Click({
    Save-CurrentReview
    [System.Windows.Forms.MessageBox]::Show('Review saved.', 'shinsa') | Out-Null
})

$writebackButton.Add_Click({
    Save-CurrentReview
    & (Join-Path $appRoot 'scripts\Writeback-Review.ps1')
    [System.Windows.Forms.MessageBox]::Show('Writeback completed.', 'shinsa') | Out-Null
})

$grid.Add_SelectionChanged({
    if ($grid.SelectedRows.Count -gt 0) {
        Load-Detail -CaseId $grid.SelectedRows[0].Cells['case_id'].Value
    }
})

Refresh-Indexes
[void]$form.ShowDialog()
