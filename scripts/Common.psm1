function Get-ShinsaAppRoot {
    param([string]$ScriptPath)
    Split-Path -Parent $ScriptPath | Split-Path -Parent
}

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
    if (-not (Test-Path $directory)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }

    $Data | ConvertTo-Json -Depth 8 | Set-Content -Path $Path -Encoding UTF8
}

function Merge-Hashtable {
    param(
        [hashtable]$Base,
        [hashtable]$Overlay
    )

    foreach ($key in $Overlay.Keys) {
        $Base[$key] = $Overlay[$key]
    }
}

function Get-ShinsaConfig {
    param([string]$ScriptPath)

    $appRoot = Get-ShinsaAppRoot -ScriptPath $ScriptPath
    $base = Read-ShinsaJson -Path (Join-Path $appRoot 'config\config.base.json')
    $localPath = Join-Path $appRoot 'config\config.local.json'
    if (-not (Test-Path $localPath)) {
        Copy-Item (Join-Path $appRoot 'config\config.local.sample.json') $localPath -Force
    }
    $local = Read-ShinsaJson -Path $localPath

    $config = [ordered]@{
        app = @{}
        paths = @{}
        outlook = @{}
        ledger = @{}
        gui = @{}
        review = @{}
    }

    foreach ($name in $base.app.PSObject.Properties.Name) { $config.app[$name] = $base.app.$name }
    foreach ($name in $base.paths.PSObject.Properties.Name) { $config.paths[$name] = $base.paths.$name }
    foreach ($name in $base.outlook.PSObject.Properties.Name) { $config.outlook[$name] = $base.outlook.$name }
    foreach ($name in $base.ledger.PSObject.Properties.Name) { $config.ledger[$name] = $base.ledger.$name }
    foreach ($name in $base.gui.PSObject.Properties.Name) { $config.gui[$name] = $base.gui.$name }
    foreach ($name in $base.review.PSObject.Properties.Name) { $config.review[$name] = $base.review.$name }

    if ($local.paths) { foreach ($name in $local.paths.PSObject.Properties.Name) { $config.paths[$name] = $local.paths.$name } }
    if ($local.outlook) { foreach ($name in $local.outlook.PSObject.Properties.Name) { $config.outlook[$name] = $local.outlook.$name } }
    if ($local.gui) { foreach ($name in $local.gui.PSObject.Properties.Name) { $config.gui[$name] = $local.gui.$name } }

    [pscustomobject]$config
}

function Ensure-ShinsaDirectory {
    param([string[]]$Paths)
    foreach ($path in $Paths) {
        if (-not (Test-Path $path)) {
            New-Item -ItemType Directory -Path $path -Force | Out-Null
        }
    }
}

function Get-ShinsaPaths {
    param($Config)

    [pscustomobject]@{
        CloneRoot = $Config.paths.cloneRoot
        CloneLedgerRoot = Join-Path $Config.paths.cloneRoot 'ledger'
        CloneCaseRoot = Join-Path $Config.paths.cloneRoot 'cases'
        CloneMailRoot = Join-Path $Config.paths.cloneRoot 'mail'
        IndexRoot = $Config.paths.indexRoot
        WorkRoot = $Config.paths.workRoot
        ReviewStatePath = Join-Path $Config.paths.workRoot 'review_state.json'
        PendingChangesPath = Join-Path $Config.paths.workRoot 'pending_changes.json'
        CaseIndexPath = Join-Path $Config.paths.indexRoot 'case_index.json'
        MailIndexPath = Join-Path $Config.paths.indexRoot 'mail_index.json'
        ContactsIndexPath = Join-Path $Config.paths.indexRoot 'contacts_index.json'
        LedgerCloneCasesPath = Join-Path $Config.paths.cloneRoot ('ledger\' + $Config.ledger.casesFileName)
        LedgerCloneContactsPath = Join-Path $Config.paths.cloneRoot ('ledger\' + $Config.ledger.contactsFileName)
        SourceCasesPath = Join-Path $Config.paths.onedriveLedgerRoot $Config.ledger.casesFileName
        SourceContactsPath = Join-Path $Config.paths.onedriveLedgerRoot $Config.ledger.contactsFileName
    }
}

function Write-ShinsaLog {
    param(
        [string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR')][string]$Level = 'INFO',
        [string]$ScriptPath
    )

    $config = Get-ShinsaConfig -ScriptPath $ScriptPath
    Ensure-ShinsaDirectory -Paths @($config.paths.logRoot)
    $logFile = Join-Path $config.paths.logRoot ('shinsa_' + (Get-Date -Format 'yyyyMMdd') + '.log')
    $line = '[{0}] [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Add-Content -Path $logFile -Value $line -Encoding UTF8
    Write-Host $line
}

function Get-ReviewState {
    param($Paths)

    if (Test-Path $Paths.ReviewStatePath) {
        return Read-ShinsaJson -Path $Paths.ReviewStatePath
    }

    [pscustomobject]@{ reviews = @() }
}

function Save-ReviewState {
    param(
        $Paths,
        $State
    )

    Write-ShinsaJson -Path $Paths.ReviewStatePath -Data $State
    Write-ShinsaJson -Path $Paths.PendingChangesPath -Data $State
}

function Convert-ToSafeName {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return 'blank' }
    $safe = $Value -replace '[\\/:*?"<>|]', '_'
    if ($safe.Length -gt 80) { $safe = $safe.Substring(0, 80) }
    $safe
}

Export-ModuleMember -Function *
