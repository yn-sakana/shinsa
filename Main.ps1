param(
    [switch]$SkipStartupSync
)

$ErrorActionPreference = 'Stop'
$script:AppRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

function Invoke-ShinsaStep {
    param([string]$RelativePath)
    & (Join-Path $script:AppRoot $RelativePath)
}

function Start-ShinsaGui {
    $guiScript = Join-Path $script:AppRoot 'gui\Start-Gui.ps1'
    Start-Process powershell.exe -ArgumentList @(
        '-NoLogo'
        '-NoProfile'
        '-ExecutionPolicy'
        'Bypass'
        '-File'
        $guiScript
    ) | Out-Null
}

function Show-Help {
    Write-Host ''
    Write-Host '=== shinsa commands ===' -ForegroundColor Cyan
    Write-Host '  gui        open the review GUI'
    Write-Host '  sync       sync clone from OneDrive sources and rebuild index'
    Write-Host '  index      rebuild index only'
    Write-Host '  writeback  write review changes back to the source ledger'
    Write-Host '  status     show current paths and counts'
    Write-Host '  config     show config file path'
    Write-Host '  help       show this help'
    Write-Host '  quit       exit'
    Write-Host ''
}

function Show-Status {
    $configPath = Join-Path $script:AppRoot 'config\config.local.json'
    $cloneRoot = Join-Path $script:AppRoot 'data\clone'
    $indexRoot = Join-Path $script:AppRoot 'data\index'
    $workRoot = Join-Path $script:AppRoot 'data\work'

    $mailCount = @(Get-ChildItem (Join-Path $cloneRoot 'mail') -Recurse -Filter meta.json -File -ErrorAction SilentlyContinue).Count
    $caseCount = 0
    $caseIndexPath = Join-Path $indexRoot 'case_index.json'
    if (Test-Path $caseIndexPath) {
        $loaded = Get-Content $caseIndexPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $caseCount = @($loaded).Count
    }

    Write-Host ''
    Write-Host '=== shinsa status ===' -ForegroundColor Cyan
    Write-Host "  config : $configPath"
    Write-Host "  clone  : $cloneRoot"
    Write-Host "  index  : $indexRoot"
    Write-Host "  work   : $workRoot"
    Write-Host "  mails  : $mailCount"
    Write-Host "  cases  : $caseCount"
    Write-Host ''
}

function Invoke-StartupSync {
    if ($SkipStartupSync) {
        Write-Host 'startup sync skipped' -ForegroundColor Yellow
        return
    }

    Write-Host ''
    Write-Host 'shinsa starting: sync + index' -ForegroundColor Cyan
    Invoke-ShinsaStep 'scripts\Sync-Clone.ps1'
    Invoke-ShinsaStep 'scripts\Build-Index.ps1'
}

function Start-ShinsaLoop {
    Show-Help

    while ($true) {
        $inputLine = Read-Host 'shinsa'
        if ([string]::IsNullOrWhiteSpace($inputLine)) { continue }

        $command = ($inputLine.Trim() -split '\s+', 2)[0].ToLowerInvariant()

        switch ($command) {
            'gui' {
                Start-ShinsaGui
            }
            'sync' {
                Invoke-ShinsaStep 'scripts\Sync-Clone.ps1'
                Invoke-ShinsaStep 'scripts\Build-Index.ps1'
            }
            'index' {
                Invoke-ShinsaStep 'scripts\Build-Index.ps1'
            }
            'writeback' {
                Invoke-ShinsaStep 'scripts\Writeback-Review.ps1'
            }
            'status' {
                Show-Status
            }
            'config' {
                Write-Host (Join-Path $script:AppRoot 'config\config.local.json')
            }
            'help' {
                Show-Help
            }
            'quit' {
                return
            }
            default {
                Write-Host "unknown command: $command" -ForegroundColor Yellow
                Write-Host "type 'help' for commands" -ForegroundColor Yellow
            }
        }
    }
}

try {
    Invoke-StartupSync
    Start-ShinsaLoop
}
catch {
    Write-Host ''
    Write-Host "shinsa error: $($_.Exception.Message)" -ForegroundColor Red
}
