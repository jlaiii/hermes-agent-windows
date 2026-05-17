if (-not (Get-Command Write-Log -ErrorAction SilentlyContinue)) {
    . (Join-Path $PSScriptRoot 'utils.ps1')
}
if (-not (Get-Command Test-CommandExists -ErrorAction SilentlyContinue)) {
    . (Join-Path $PSScriptRoot 'checks.ps1')
}
if (-not (Get-Command Invoke-WslShell -ErrorAction SilentlyContinue)) {
    . (Join-Path $PSScriptRoot 'wsl-manager.ps1')
}
if (-not (Get-Command Test-OllamaAvailable -ErrorAction SilentlyContinue)) {
    . (Join-Path $PSScriptRoot 'ollama-manager.ps1')
}

$script:HermesWslInstallCommand = 'curl -fsSL https://raw.githubusercontent.com/NousResearch/hermes-agent/main/scripts/install.sh | bash -s -- --skip-setup'
$script:HermesWslUser = 'admin'

function Invoke-HermesWslCommand {
    param(
        [Parameter(Mandatory)]
        [string]$Command,
        [int]$TimeoutSeconds = 120,
        [switch]$AsAdminUser
    )

    $user = if ($AsAdminUser) { $script:HermesWslUser } else { '' }
    $wrappedCommand = 'export PATH="$HOME/.local/bin:$HOME/.hermes/hermes-agent/venv/bin:$PATH"; ' + $Command
    return Invoke-WslShell -Command $wrappedCommand -User $user -TimeoutSeconds $TimeoutSeconds
}

function Test-HermesInstalled {
    if (-not (Test-WslExists)) {
        return $false
    }

    $result = Invoke-HermesWslCommand -Command 'command -v hermes' -AsAdminUser -TimeoutSeconds 30
    return ($result.Status -eq 'Success' -and $result.Details)
}

function Get-HermesVersion {
    if (-not (Test-HermesInstalled)) {
        return Format-StatusResult -Name 'Hermes Agent Version' -Status 'Missing' -Message 'Hermes Agent is not installed in WSL.' -Details 'The setup installs Hermes inside WSL with the official Nous Research Linux installer.'
    }

    $result = Invoke-HermesWslCommand -Command 'hermes --version' -AsAdminUser -TimeoutSeconds 60
    if ($result.Status -eq 'Success') {
        return Format-StatusResult -Name 'Hermes Agent Version' -Status 'Installed' -Message 'Hermes Agent version retrieved from WSL.' -Details $result.Details
    }

    return Format-StatusResult -Name 'Hermes Agent Version' -Status 'Error' -Message 'Could not read Hermes Agent version from WSL.' -Details $result.Details -ExitCode $result.ExitCode
}

function Get-HermesStatus {
    if (-not (Test-WslExists)) {
        return Format-StatusResult -Name 'Hermes Agent Status' -Status 'Missing' -Message 'WSL is required before Hermes can run.' -Details 'Install WSL first.'
    }

    $running = Invoke-HermesWslCommand -Command "ps -eo pid=,comm=,args= | awk '`$2 !~ /^(bash|sh|awk|ps)$/ && `$0 ~ /hermes/ {print}'" -AsAdminUser -TimeoutSeconds 30
    if ($running.Status -eq 'Success' -and $running.Details -match 'hermes') {
        return Format-StatusResult -Name 'Hermes Agent Status' -Status 'Running' -Message 'Hermes Agent process is running inside WSL.' -Details $running.Details
    }

    if (Test-HermesInstalled) {
        return Format-StatusResult -Name 'Hermes Agent Status' -Status 'Stopped' -Message 'Hermes Agent is installed in WSL but is not currently running.' -Details 'Use Start Hermes Agent or Enable Gateway.'
    }

    return Format-StatusResult -Name 'Hermes Agent Status' -Status 'Missing' -Message 'Hermes Agent is not installed in WSL.' -Details 'Use Install Hermes Agent.'
}

function Install-HermesAgent {
    $logFile = Get-LogFilePath -Kind 'app'
    if (-not (Test-WslExists)) {
        return Format-StatusResult -Name 'Hermes Install' -Status 'Missing' -Message 'WSL is required before installing Hermes Agent.' -Details 'Install WSL first.' -ExitCode 1
    }

    $admin = Ensure-WslAdminAccount
    if ($admin.Status -eq 'Error') {
        return $admin
    }

    if (Test-HermesInstalled) {
        return Format-StatusResult -Name 'Hermes Install' -Status 'Installed' -Message 'Hermes Agent is already installed in WSL.' -Details 'No installation was needed.'
    }

    Write-Log -Message 'Installing Hermes Agent inside WSL with the official Nous Research Linux installer.' -Level 'INFO' -LogFile $logFile | Out-Null
    $result = Invoke-HermesWslCommand -Command $script:HermesWslInstallCommand -AsAdminUser -TimeoutSeconds 1200
    if ($result.Status -eq 'Success' -and (Test-HermesInstalled)) {
        return Format-StatusResult -Name 'Hermes Install' -Status 'Installed' -Message 'Hermes Agent was installed inside WSL.' -Details $result.Details
    }

    if (Test-HermesInstalled) {
        return Format-StatusResult -Name 'Hermes Install' -Status 'Installed' -Message 'Hermes Agent was found after installer returned an error.' -Details $result.Details
    }

    return Format-StatusResult -Name 'Hermes Install' -Status 'Error' -Message 'Hermes Agent installation in WSL failed.' -Details $result.Details -ExitCode $result.ExitCode
}

function Update-HermesAgent {
    if (-not (Test-HermesInstalled)) {
        return Install-HermesAgent
    }

    $result = Invoke-HermesWslCommand -Command 'hermes update' -AsAdminUser -TimeoutSeconds 900
    if ($result.Status -eq 'Success') {
        return Format-StatusResult -Name 'Hermes Update' -Status 'Installed' -Message 'Hermes Agent update command completed in WSL.' -Details $result.Details
    }

    return Format-StatusResult -Name 'Hermes Update' -Status 'Error' -Message 'Hermes Agent update command failed in WSL.' -Details $result.Details -ExitCode $result.ExitCode
}

function Start-HermesAgent {
    if (-not (Test-HermesInstalled)) {
        $install = Install-HermesAgent
        if ($install.Status -ne 'Installed') {
            return $install
        }
    }

    $command = 'mkdir -p "$HOME/.hermes"; nohup hermes > "$HOME/.hermes/hermes.log" 2>&1 & sleep 3; ps -eo pid=,comm=,args= | awk ''$2 !~ /^(bash|sh|awk|ps)$/ && $0 ~ /hermes/ {print}'''
    $result = Invoke-HermesWslCommand -Command $command -AsAdminUser -TimeoutSeconds 60
    if ($result.Status -eq 'Success') {
        return Format-StatusResult -Name 'Start Hermes Agent' -Status 'Running' -Message 'Hermes Agent was started inside WSL.' -Details 'Logs: /home/admin/.hermes/hermes.log'
    }

    return Format-StatusResult -Name 'Start Hermes Agent' -Status 'Error' -Message 'Hermes Agent did not start cleanly.' -Details $result.Details -ExitCode $result.ExitCode
}

function Stop-HermesAgent {
    $result = Invoke-HermesWslCommand -Command 'pkill -f "python.*hermes|venv/bin/hermes|hermes gateway" 2>/dev/null || true; sleep 1; ps -eo comm=,args= | awk ''$1 !~ /^(bash|sh|awk|ps)$/ && $0 ~ /hermes/ {found=1} END {exit found ? 1 : 0}''' -AsAdminUser -TimeoutSeconds 60
    if ($result.Status -eq 'Success') {
        return Format-StatusResult -Name 'Stop Hermes Agent' -Status 'Stopped' -Message 'Hermes Agent processes were stopped inside WSL.'
    }

    return Format-StatusResult -Name 'Stop Hermes Agent' -Status 'Error' -Message 'Failed to stop Hermes Agent cleanly.' -Details $result.Details -ExitCode $result.ExitCode
}

function Restart-HermesAgent {
    Stop-HermesAgent | Out-Null
    return Start-HermesAgent
}

function Invoke-HermesDoctor {
    if (-not (Test-HermesInstalled)) {
        return Format-StatusResult -Name 'Hermes Doctor' -Status 'Missing' -Message 'Hermes Agent is not installed in WSL.' -Details 'Install Hermes Agent first.' -ExitCode 1
    }

    $result = Invoke-HermesWslCommand -Command 'HERMES_ACCEPT_HOOKS=1 hermes doctor 2>&1' -AsAdminUser -TimeoutSeconds 900
    if ($result.Status -eq 'Success') {
        return Format-StatusResult -Name 'Hermes Doctor' -Status 'Installed' -Message 'Hermes Doctor completed.' -Details $result.Details
    }

    if ($result.ExitCode -eq 124) {
        return Format-StatusResult -Name 'Hermes Doctor' -Status 'Error' -Message 'Hermes Doctor did not finish within 15 minutes.' -Details 'The command may be waiting on an external dependency or provider check. Try opening the Hermes dashboard or running hermes doctor manually inside WSL for interactive prompts.' -ExitCode $result.ExitCode
    }

    return Format-StatusResult -Name 'Hermes Doctor' -Status 'Error' -Message 'Hermes Doctor found a problem or failed to run.' -Details $result.Details -ExitCode $result.ExitCode
}

function Open-HermesCli {
    if (-not (Test-HermesInstalled)) {
        return Format-StatusResult -Name 'Launch Hermes CLI' -Status 'Missing' -Message 'Hermes Agent is not installed in WSL.' -Details 'Install Hermes Agent first.' -ExitCode 1
    }

    $projectRoot = Get-ProjectRoot
    $logsPath = Join-Path $projectRoot 'logs'
    if (-not (Test-Path $logsPath)) {
        New-Item -ItemType Directory -Path $logsPath -Force | Out-Null
    }

    $launcherPath = Join-Path $logsPath 'Launch-HermesCli.cmd'
    $shellLauncherPath = Join-Path $logsPath 'launch-hermes-cli.sh'
    $projectRootForShell = $projectRoot -replace '\\', '/'
    if ($projectRootForShell -match '^([A-Za-z]):/(.*)$') {
        $driveLetter = $Matches[1].ToLowerInvariant()
        $pathPart = $Matches[2]
        $wslStartPath = "/mnt/$driveLetter/$pathPart"
    }
    else {
        $wslStartPath = '$HOME'
    }

    $shellLauncher = @'
#!/usr/bin/env bash
START_PATH="__HERMES_START_PATH__"
if [ -d "$START_PATH" ]; then
  cd "$START_PATH" || cd "$HOME" || exit 1
else
  cd "$HOME" || exit 1
fi

export PATH="$HOME/.local/bin:$HOME/.hermes/hermes-agent/venv/bin:$PATH"
test -f "$HOME/.ollama-cloud.env" && . "$HOME/.ollama-cloud.env"
export HERMES_ACCEPT_HOOKS=1

echo "Starting Hermes Agent CLI..."
echo ""
echo "You can talk to Hermes here and ask it to inspect, edit, or explain files it can access."
echo "Working folder: $(pwd)"
echo "Provider/model come from your saved Hermes config."
echo ""
echo "Tip: type /help inside Hermes for chat commands, or Ctrl+C to stop."
echo ""
exec hermes chat --accept-hooks
'@
    $shellLauncher = $shellLauncher.Replace('__HERMES_START_PATH__', $wslStartPath)
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($shellLauncherPath, ($shellLauncher -replace "`r`n", "`n"), $utf8NoBom)

    $shellLauncherForWsl = $shellLauncherPath -replace '\\', '/'
    if ($shellLauncherForWsl -match '^([A-Za-z]):/(.*)$') {
        $shellDriveLetter = $Matches[1].ToLowerInvariant()
        $shellPathPart = $Matches[2]
        $wslShellLauncherPath = "/mnt/$shellDriveLetter/$shellPathPart"
    }
    else {
        $wslShellLauncherPath = $shellLauncherPath
    }

    $launcherLines = @(
        '@echo off'
        'title Hermes Agent CLI'
        'echo Starting Hermes Agent CLI in WSL...'
        'echo.'
        'echo If Windows asks for WSL access or this fails, open hermes-agent-windows logs for details.'
        'echo.'
        "wsl.exe -u admin -e bash ""$wslShellLauncherPath"""
        'set "EXIT_CODE=%ERRORLEVEL%"'
        'echo.'
        'echo Hermes CLI session ended. Exit code: %EXIT_CODE%'
        'echo Press any key to close this window.'
        'pause >nul'
        'exit /b %EXIT_CODE%'
    )
    $launcherScript = $launcherLines -join "`r`n"
    Set-Content -Path $launcherPath -Value $launcherScript -Encoding ASCII
    $args = @('/k', "`"$launcherPath`"")

    try {
        Write-AppLog -Message "Opening interactive Hermes Agent CLI in Windows Command Prompt using $launcherPath." -Level 'INFO' | Out-Null
        Start-Process -FilePath 'cmd.exe' -ArgumentList $args -WorkingDirectory $projectRoot | Out-Null
        return Format-StatusResult -Name 'Launch Hermes CLI' -Status 'Running' -Message 'Opened the interactive Hermes Agent CLI.' -Details "Command: hermes chat --accept-hooks`nWSL user: admin`nLauncher: $launcherPath"
    }
    catch {
        return Format-StatusResult -Name 'Launch Hermes CLI' -Status 'Error' -Message 'Could not open the Hermes CLI terminal.' -Details $_.Exception.Message -ExitCode 1
    }
}

function Get-HermesGatewayStatus {
    if (-not (Test-HermesInstalled)) {
        return Format-StatusResult -Name 'Hermes Gateway' -Status 'Missing' -Message 'Hermes Agent is not installed in WSL.' -Details 'Install Hermes first.'
    }

    $status = Invoke-HermesWslCommand -Command 'hermes gateway status 2>&1 || true' -AsAdminUser -TimeoutSeconds 60
    if ($status.Status -eq 'Success' -and $status.Details -match 'Gateway is running|running') {
        return Format-StatusResult -Name 'Hermes Gateway' -Status 'Running' -Message 'Hermes Gateway status reports running.' -Details $status.Details
    }

    $process = Invoke-HermesWslCommand -Command "ps -eo pid=,comm=,args= | awk '`$2 !~ /^(bash|sh|awk|ps)$/ && `$0 ~ /gateway/ && `$0 ~ /hermes/ {print}'" -AsAdminUser -TimeoutSeconds 30
    if ($process.Status -eq 'Success' -and $process.Details -match 'gateway') {
        return Format-StatusResult -Name 'Hermes Gateway' -Status 'Running' -Message 'Hermes Gateway appears to be running inside WSL.' -Details $process.Details
    }

    $config = Invoke-HermesWslCommand -Command 'test -f "$HOME/.hermes/config.yaml" && grep -n "use_gateway" "$HOME/.hermes/config.yaml" || true' -AsAdminUser -TimeoutSeconds 30
    if ($config.Status -eq 'Success' -and $config.Details -match 'use_gateway') {
        return Format-StatusResult -Name 'Hermes Gateway' -Status 'Installed' -Message 'Hermes gateway setting was found in config.' -Details $config.Details
    }

    return Format-StatusResult -Name 'Hermes Gateway' -Status 'Stopped' -Message 'Hermes Gateway is not running.' -Details 'Use Enable Hermes Gateway to run hermes gateway setup/start.'
}

function Enable-HermesGateway {
    if (-not (Test-HermesInstalled)) {
        $install = Install-HermesAgent
        if ($install.Status -ne 'Installed') {
            return $install
        }
    }

    $current = Get-HermesGatewayStatus
    if ($current.Status -eq 'Running') {
        return Format-StatusResult -Name 'Hermes Gateway' -Status 'Running' -Message 'Hermes Gateway is already running.' -Details $current.Details
    }

    $command = @'
mkdir -p "$HOME/.hermes"
if hermes gateway --help >/dev/null 2>&1; then
  mkdir -p "$HOME/.hermes/logs"
  nohup hermes gateway run --replace --accept-hooks > "$HOME/.hermes/logs/gateway.log" 2>&1 &
  sleep 4
  hermes gateway status 2>&1 || true
else
  echo "Hermes gateway command is not available in this installed version."
  exit 1
fi
'@
    $result = Invoke-HermesWslCommand -Command $command -AsAdminUser -TimeoutSeconds 90
    if ($result.Status -eq 'Success') {
        return Format-StatusResult -Name 'Hermes Gateway' -Status 'Running' -Message 'Hermes Gateway run command was launched inside WSL.' -Details "Logs: /home/admin/.hermes/logs/gateway.log`n$($result.Details)"
    }

    return Format-StatusResult -Name 'Hermes Gateway' -Status 'Error' -Message 'Hermes Gateway did not start cleanly.' -Details $result.Details -ExitCode $result.ExitCode
}

function Open-HermesGateway {
    if (-not (Test-HermesInstalled)) {
        $install = Install-HermesAgent
        if ($install.Status -ne 'Installed') {
            return $install
        }
    }

    $command = @'
mkdir -p "$HOME/.hermes/logs"
if ! curl -fsS --max-time 5 http://127.0.0.1:9119 >/dev/null 2>&1; then
  hermes dashboard --stop >/dev/null 2>&1 || true
  nohup hermes dashboard --host 127.0.0.1 --port 9119 --no-open --skip-build > "$HOME/.hermes/logs/dashboard.log" 2>&1 &
  sleep 8
fi
curl -fsS --max-time 5 http://127.0.0.1:9119 >/dev/null 2>&1
echo "Hermes dashboard is reachable inside WSL at http://127.0.0.1:9119"
'@
    $dashboard = Invoke-HermesWslCommand -Command $command -AsAdminUser -TimeoutSeconds 120
    if ($dashboard.Status -ne 'Success') {
        return Format-StatusResult -Name 'Open Hermes Dashboard' -Status 'Error' -Message 'Hermes dashboard did not become reachable inside WSL.' -Details "Log: /home/admin/.hermes/logs/dashboard.log`n$($dashboard.Details)" -ExitCode $dashboard.ExitCode
    }

    $probe = Invoke-CommandSafe -FilePath 'powershell.exe' -Arguments @('-NoProfile', '-Command', 'try { $r = Invoke-WebRequest -UseBasicParsing -TimeoutSec 5 http://localhost:9119; "OK $($r.StatusCode)" } catch { "ERR $($_.Exception.Message)"; exit 1 }') -LogFile (Get-LogFilePath -Kind 'app') -AllowFailure -TimeoutSeconds 15
    $url = 'http://localhost:9119'
    if ($probe.Status -ne 'Success') {
        $ipResult = Invoke-WslShell -Command "hostname -I | awk '{print `$1}'" -TimeoutSeconds 15
        if ($ipResult.Status -eq 'Success' -and $ipResult.Details) {
            $candidateUrl = "http://$($ipResult.Details.Trim()):9119"
            $probe = Invoke-CommandSafe -FilePath 'powershell.exe' -Arguments @('-NoProfile', '-Command', "try { `$r = Invoke-WebRequest -UseBasicParsing -TimeoutSec 5 $candidateUrl; `"OK `$(`$r.StatusCode)`" } catch { `"ERR `$(`$_.Exception.Message)`"; exit 1 }") -LogFile (Get-LogFilePath -Kind 'app') -AllowFailure -TimeoutSeconds 15
            if ($probe.Status -eq 'Success') {
                $url = $candidateUrl
            }
        }
    }

    if ($probe.Status -ne 'Success') {
        return Format-StatusResult -Name 'Open Hermes Dashboard' -Status 'Error' -Message 'Hermes dashboard is running in WSL, but Windows cannot reach it.' -Details "Try opening http://localhost:9119 manually. Dashboard log: /home/admin/.hermes/logs/dashboard.log`n$($probe.Details)" -ExitCode $probe.ExitCode
    }

    try {
        Start-Process $url | Out-Null
        return Format-StatusResult -Name 'Open Hermes Dashboard' -Status 'Running' -Message 'Opened the Hermes web dashboard.' -Details "URL: $url`nProbe: $($probe.Message)`n$($dashboard.Details)"
    }
    catch {
        return Format-StatusResult -Name 'Open Hermes Dashboard' -Status 'Error' -Message 'Dashboard started, but Windows could not open the browser.' -Details "URL: $url`n$($_.Exception.Message)" -ExitCode 1
    }
}

function Open-HermesConfigFolder {
    $result = Invoke-HermesWslCommand -Command 'mkdir -p "$HOME/.hermes"; wslpath -w "$HOME/.hermes"' -AsAdminUser -TimeoutSeconds 30
    if ($result.Status -eq 'Success' -and $result.Details) {
        return Open-FolderSafe -Path (($result.Details -split "`r?`n" | Select-Object -First 1).Trim())
    }

    return Open-FolderSafe -Path (Join-Path (Get-ProjectRoot) 'config')
}

function Get-HermesReleaseCachePath {
    return Join-Path (Get-ProjectRoot) 'version-cache.json'
}

function Get-LatestHermesReleaseVersion {
    $cacheFile = Get-HermesReleaseCachePath
    $maxAge = [TimeSpan]::FromMinutes(15)

    if (Test-Path $cacheFile) {
        try {
            $cached = Get-Content -Path $cacheFile -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            $cachedAt = [DateTime]$cached.CheckedAt
            if ((Get-Date) - $cachedAt -lt $maxAge) {
                return [pscustomobject]@{
                    Status   = 'Installed'
                    Message  = "Cached latest: $($cached.LatestVersion)"
                    Details  = "Checked at $($cached.CheckedAt)"
                    ExitCode = 0
                    Version  = $cached.LatestVersion
                }
            }
        }
        catch {
        }
    }

    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $release = Invoke-RestMethod -Uri 'https://api.github.com/repos/NousResearch/hermes-agent/releases/latest' -TimeoutSec 30 -ErrorAction Stop
        $tag = $release.tag_name -replace '^v', ''

        $cacheEntry = @{
            LatestVersion = $tag
            CheckedAt     = (Get-Date).ToString('o')
        } | ConvertTo-Json -Depth 3

        try {
            [System.IO.File]::WriteAllText($cacheFile, $cacheEntry, (New-Object System.Text.UTF8Encoding($false)))
        }
        catch {
        }

        return [pscustomobject]@{
            Status   = 'Installed'
            Message  = "Latest release: v$tag"
            Details  = 'Fetched from GitHub API.'
            ExitCode = 0
            Version  = $tag
        }
    }
    catch {
        return [pscustomobject]@{
            Status   = 'Error'
            Message  = 'Could not check GitHub for latest Hermes release.'
            Details  = $_.Exception.Message
            ExitCode = 1
            Version  = ''
        }
    }
}

function Get-HermesUpdateStatus {
    $installed = Get-HermesVersion
    $latest = Get-LatestHermesReleaseVersion

    if ($installed.Status -ne 'Installed') {
        return Format-StatusResult -Name 'Hermes Update' -Status 'Missing' -Message 'Hermes Agent is not installed in WSL.' -Details 'Install Hermes Agent before checking for updates.'
    }

    $versionText = ''
    if ($installed.Details -match '(?:version[v]?[\s:]?)?(\d+\.\d+(?:\.\d+)?)') {
        $versionText = $Matches[1]
    }

    if (-not $versionText -and $installed.Message -match '(\d+\.\d+(?:\.\d+)?)') {
        $versionText = $Matches[1]
    }

    if (-not $versionText) {
        $versionText = '0.0.0'
    }

    if ($latest.Status -ne 'Installed') {
        return Format-StatusResult -Name 'Hermes Update' -Status 'Unknown' -Message 'Hermes Agent installed, but could not reach GitHub.' -Details "$($installed.Message)`n$($latest.Message)"
    }

    try {
        $currentObj = [version]$versionText
        $latestObj = [version]$latest.Version

        if ($latestObj -gt $currentObj) {
            return Format-StatusResult -Name 'Hermes Update' -Status 'Needs Update' -Message "Update available: v$($latest.Version) (installed: v$versionText)" -Details "A newer release is on GitHub. Click Update Hermes Agent to upgrade."
        }

        return Format-StatusResult -Name 'Hermes Update' -Status 'Installed' -Message "Hermes is up to date (v$versionText)" -Details "Latest: v$($latest.Version)"
    }
    catch {
        return Format-StatusResult -Name 'Hermes Update' -Status 'Unknown' -Message "Hermes installed version: v$versionText | latest: v$($latest.Version)" -Details "Version comparison failed. Check manually if needed."
    }
}

function Get-HermesAgentWindowsUpdateInfo {
    $cacheFile = Join-Path (Get-ProjectRoot) 'version-cache.json'
    if (Test-Path $cacheFile) {
        try {
            $cached = Get-Content -Path $cacheFile -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            if ($cached.LatestVersion -and $cached.CheckedAt) {
                $checkedAt = [DateTime]$cached.CheckedAt
                $age = (Get-Date) - $checkedAt
                $ageText = if ($age.TotalMinutes -lt 1) { 'just now' } elseif ($age.TotalMinutes -lt 60) { "$([math]::Round($age.TotalMinutes))m ago" } else { "$([math]::Round($age.TotalHours,1))h ago" }
                return [pscustomobject]@{
                    LatestVersion = $cached.LatestVersion
                    CheckedAt     = $cached.CheckedAt
                    AgeText       = $ageText
                }
            }
        }
        catch {
        }
    }
    return $null
}

