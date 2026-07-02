# Launch Outlook on Windows with add-in DevTools enabled.
#
# New Outlook (olk.exe) does NOT support F12 / right-click -> Inspect inside an
# add-in taskpane. The only supported way to debug is to start the client with
# the --devtools flag, which opens Microsoft Edge DevTools next to Outlook and
# captures the taskpane's console/network from the moment it boots:
#   https://learn.microsoft.com/office/dev/add-ins/outlook/one-outlook#debug-your-add-in
#
# Usage (from a Windows PowerShell prompt):
#   .\scripts\debug-outlook.ps1            # new Outlook + DevTools
#   .\scripts\debug-outlook.ps1 -Classic   # classic Outlook, WebView2 auto-opens DevTools
#
# Usage (from WSL):
#   powershell.exe -ExecutionPolicy Bypass -File "$(wslpath -w scripts/debug-outlook.ps1)"
#
# Notes:
#   - Outlook must be fully closed first; the script stops running instances.
#   - If you close the DevTools window, close Outlook and run this again —
#     reopening DevTools on a running instance is not supported.
#   - In DevTools, the taskpane shows up once you open the add-in. Use the
#     top-left context picker if you see the wrong frame.

param(
    # Debug classic Outlook (outlook.exe) instead of new Outlook (olk.exe).
    # Classic runs add-ins in WebView2; setting
    # WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--auto-open-devtools-for-tabs makes
    # DevTools open automatically when the add-in webview starts.
    [switch]$Classic
)

$ErrorActionPreference = 'Stop'

if ($Classic) {
    $running = Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue
    if ($running) {
        Write-Host 'Stopping classic Outlook...'
        $running | Stop-Process -Force
        Start-Sleep -Seconds 2
    }

    # Process-scoped env var: inherited by the Outlook process we start below,
    # without polluting the user/machine environment permanently.
    $env:WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS = '--auto-open-devtools-for-tabs'
    Write-Host 'Starting classic Outlook with WebView2 DevTools auto-open...'
    Start-Process 'outlook.exe'
    Write-Host 'Open the Synamail add-in; DevTools opens when its webview starts.'
    return
}

$running = Get-Process -Name 'olk' -ErrorAction SilentlyContinue
if ($running) {
    Write-Host 'Stopping new Outlook (olk.exe)...'
    $running | Stop-Process -Force
    Start-Sleep -Seconds 2
}

Write-Host 'Starting new Outlook with DevTools (olk.exe --devtools)...'
# olk.exe resolves via the WindowsApps execution alias on any machine with the
# new Outlook installed.
Start-Process 'olk.exe' -ArgumentList '--devtools'
Write-Host ''
Write-Host 'DevTools opens next to Outlook. Now open a message and click "Open Synamail".'
Write-Host 'Console + Network of the taskpane are captured from boot — look for'
Write-Host 'requests to addin.synaplan.com (prod) or localhost:3000 (dev).'
