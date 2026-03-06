param(
    [switch]$SyncOnStart
)

$ErrorActionPreference = 'Stop'
$script:MainScriptPath = $MyInvocation.MyCommand.Path
$script:AppRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Import-Module (Join-Path $script:AppRoot 'scripts\Common.psm1') -Force -DisableNameChecking

function Invoke-ShinsaScript {
    param(
        [Parameter(Mandatory = $true)][string]$RelativePath,
        [string[]]$Arguments = @()
    )

    & (Join-Path $script:AppRoot $RelativePath) @Arguments
}

function Start-ShinsaGui {
    $guiScript = Join-Path $script:AppRoot 'gui\Start-Gui.ps1'
    Start-Process powershell.exe -WindowStyle Hidden -ArgumentList @(
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
    Write-Host 'shinsa commands' -ForegroundColor Cyan
    Write-Host '  gui        open the WinForms review GUI'
    Write-Host '  sync       rebuild table.json / mails.json / folders.json'
    Write-Host '  status     show configured paths and local JSON status'
    Write-Host '  writeback  reflect edited table fields back to the source table'
    Write-Host '  help       show commands'
    Write-Host '  quit       exit the shell'
}

function Get-JsonCount {
    param([string]$Path)

    if (-not (Test-Path $Path)) {
        return 0
    }

    $data = Read-ShinsaJson -Path $Path
    if ($data -is [System.Collections.IEnumerable] -and -not ($data -is [string]) -and -not ($data -is [System.Management.Automation.PSCustomObject])) {
        return @($data).Count
    }

    if ($null -eq $data) {
        return 0
    }

    return 1
}

function Show-Status {
    $config = Get-ShinsaConfig -ScriptPath $script:MainScriptPath
    $paths = Get-ShinsaDataPaths -Config $config

    Ensure-ShinsaState -Paths $paths
    $cache = Read-ShinsaCache -Paths $paths
    $lastSyncAt = ''
    if ($cache.ui_state.PSObject.Properties.Name -contains 'last_sync_at') {
        $lastSyncAt = [string]$cache.ui_state.last_sync_at
    }

    Write-Host ''
    Write-Host 'shinsa status' -ForegroundColor Cyan
    Write-Host ("  mail archive : {0}" -f $paths.MailArchiveRoot)
    Write-Host ("  table src   : {0}" -f $paths.SharePointTablePath)
    Write-Host ("  case root    : {0}" -f $paths.SharePointCaseRoot)
    Write-Host ("  json root    : {0}" -f $paths.JsonRoot)
    Write-Host ("  table.json  : {0}" -f (Get-JsonCount -Path $paths.TableJsonPath))
    Write-Host ("  mails.json   : {0}" -f (Get-JsonCount -Path $paths.MailsJsonPath))
    Write-Host ("  folders.json : {0}" -f (Get-JsonCount -Path $paths.FoldersJsonPath))
    Write-Host ("  cache links  : {0}" -f @($cache.mail_links).Count)
    Write-Host ("  last sync    : {0}" -f $(if ([string]::IsNullOrWhiteSpace($lastSyncAt)) { '(never)' } else { $lastSyncAt }))
}

function Start-ShinsaLoop {
    Show-Help

    while ($true) {
        $inputLine = Read-Host 'shinsa'
        if ([string]::IsNullOrWhiteSpace($inputLine)) {
            continue
        }

        $command = ($inputLine.Trim() -split '\s+', 2)[0].ToLowerInvariant()
        try {
            switch ($command) {
                'gui' {
                    Start-ShinsaGui
                }
                'sync' {
                    Invoke-ShinsaScript -RelativePath 'scripts\Sync-Data.ps1'
                }
                'status' {
                    Show-Status
                }
                'writeback' {
                    Invoke-ShinsaScript -RelativePath 'scripts\Writeback-Review.ps1'
                }
                'help' {
                    Show-Help
                }
                'quit' {
                    return
                }
                default {
                    Write-Host ("unknown command: {0}" -f $command) -ForegroundColor Yellow
                }
            }
        }
        catch {
            Write-Host ("shinsa error: {0}" -f $_.Exception.Message) -ForegroundColor Red
        }
    }
}

if ($SyncOnStart) {
    try {
        Invoke-ShinsaScript -RelativePath 'scripts\Sync-Data.ps1'
    }
    catch {
        Write-Host ("startup sync failed: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
    }
}

Start-ShinsaLoop
