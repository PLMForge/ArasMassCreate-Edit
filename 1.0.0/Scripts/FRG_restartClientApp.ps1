<#
    FRG_restartClientApp.ps1

    Restarts the Innovator Client web application so it picks up the new
    filesRevision written to InnovatorClient.config. Without a restart,
    browsers may keep using their cached copies of the client files.

    The restart is triggered by updating the last-write time of the Client
    app's web.config. Its content is left unchanged. IIS restarts only that
    application, not the application pool, so other applications sharing the
    pool keep running.
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$ClientDir
)

$ErrorActionPreference = 'Stop'

$webConfig = Join-Path $ClientDir 'web.config'
if (-not (Test-Path -LiteralPath $webConfig)) {
    Write-Output "[FRG][ERROR] Client web.config not found: $webConfig"
    exit 1
}

(Get-Item -LiteralPath $webConfig).LastWriteTime = Get-Date
Write-Output '[FRG] restarted the Client application so the new filesRevision takes effect'
exit 0
