if (-not (Get-Command Write-Log -ErrorAction SilentlyContinue)) {
    . (Join-Path $PSScriptRoot 'utils.ps1')
}
if (-not (Get-Command Test-CommandExists -ErrorAction SilentlyContinue)) {
    . (Join-Path $PSScriptRoot 'checks.ps1')
}
if (-not (Get-Command Invoke-WslCommand -ErrorAction SilentlyContinue)) {
    . (Join-Path $PSScriptRoot 'wsl-manager.ps1')
}

if (-not (Get-Command Invoke-WslShell -ErrorAction SilentlyContinue)) {
function Invoke-WslShell {
    param(
        [Parameter(Mandatory)]
        [string]$Command,
        [int]$TimeoutSeconds = 120
    )

    return Invoke-WslCommand -Arguments @('-e', 'bash', '-lc', $Command) -LogFile (Get-LogFilePath -Kind 'app') -TimeoutSeconds $TimeoutSeconds
}
}

if (-not (Get-Command Invoke-WslRootShell -ErrorAction SilentlyContinue)) {
function Invoke-WslRootShell {
    param(
        [Parameter(Mandatory)]
        [string]$Command,
        [int]$TimeoutSeconds = 300
    )

    return Invoke-CommandSafe -FilePath 'wsl.exe' -Arguments @('-u', 'root', '-e', 'bash', '-lc', $Command) -LogFile (Get-LogFilePath -Kind 'app') -AllowFailure -TimeoutSeconds $TimeoutSeconds
}
}

function Invoke-WslAdminShell {
    param(
        [Parameter(Mandatory)]
        [string]$Command,
        [int]$TimeoutSeconds = 120
    )

    return Invoke-WslCommand -Arguments @('-u', 'admin', '-e', 'bash', '-lc', $Command) -LogFile (Get-LogFilePath -Kind 'app') -TimeoutSeconds $TimeoutSeconds
}

function Test-OllamaAvailable {
    if (-not (Test-WslExists)) {
        return $false
    }

    $result = Invoke-WslShell -Command 'command -v ollama' -TimeoutSeconds 30
    return ($result.Status -eq 'Success' -and $result.Details)
}

function Get-OllamaVersion {
    if (-not (Test-OllamaAvailable)) {
        return Format-StatusResult -Name 'Ollama Version' -Status 'Missing' -Message 'Ollama is not installed in WSL.' -Details 'The setup installs Ollama inside WSL with the official Linux installer.'
    }

    $result = Invoke-WslShell -Command 'ollama --version' -TimeoutSeconds 60
    if ($result.Status -eq 'Success') {
        return Format-StatusResult -Name 'Ollama Version' -Status 'Installed' -Message 'Ollama version retrieved from WSL.' -Details $result.Details
    }

    return Format-StatusResult -Name 'Ollama Version' -Status 'Error' -Message 'Could not read Ollama version in WSL.' -Details $result.Details -ExitCode $result.ExitCode
}

function Test-OllamaRunning {
    if (-not (Test-OllamaAvailable)) {
        return [pscustomobject]@{
            Status   = 'Missing'
            Message  = 'Ollama is not installed in WSL.'
            Details  = ''
            ExitCode = 1
        }
    }

    $result = Invoke-WslShell -Command 'pgrep -x ollama >/dev/null 2>&1 || curl -fsS http://127.0.0.1:11434/api/tags >/dev/null 2>&1' -TimeoutSeconds 30
    if ($result.Status -eq 'Success') {
        return [pscustomobject]@{
            Status   = 'Running'
            Message  = 'Ollama is running inside WSL.'
            Details  = 'The Ollama process or local API responded in WSL.'
            ExitCode = 0
        }
    }

    return [pscustomobject]@{
        Status   = 'Stopped'
        Message  = 'Ollama is installed in WSL but not running.'
        Details  = $result.Details
        ExitCode = $result.ExitCode
    }
}

function Start-Ollama {
    if (-not (Test-OllamaAvailable)) {
        return [pscustomobject]@{
            Status   = 'Missing'
            Message  = 'Ollama is not installed in WSL.'
            Details  = 'Install Ollama in WSL before starting it.'
            ExitCode = 1
        }
    }

    $running = Test-OllamaRunning
    if ($running.Status -eq 'Running') {
        return $running
    }

    $command = 'mkdir -p "$HOME/.ollama"; test -f "$HOME/.ollama-cloud.env" && . "$HOME/.ollama-cloud.env"; nohup ollama serve > "$HOME/.ollama/ollama.log" 2>&1 & sleep 3; pgrep -x ollama >/dev/null 2>&1 || curl -fsS http://127.0.0.1:11434/api/tags >/dev/null 2>&1'
    $result = Invoke-WslShell -Command $command -TimeoutSeconds 30
    if ($result.Status -eq 'Success') {
        return [pscustomobject]@{
            Status   = 'Running'
            Message  = 'Ollama started inside WSL.'
            Details  = 'Logs: ~/.ollama/ollama.log'
            ExitCode = 0
        }
    }

    return [pscustomobject]@{
        Status   = 'Error'
        Message  = 'Failed to start Ollama inside WSL.'
        Details  = $result.Details
        ExitCode = $result.ExitCode
    }
}

function Set-OllamaCloudConfig {
    param(
        [Parameter(Mandatory)]
        [string]$ApiKey,
        [string]$Model = 'kimi-k2.6:cloud'
    )

    if ([string]::IsNullOrWhiteSpace($ApiKey)) {
        return Format-StatusResult -Name 'Ollama Cloud Config' -Status 'Error' -Message 'Ollama API key is empty.' -Details 'Paste an Ollama API key before saving.' -ExitCode 1
    }

    if ([string]::IsNullOrWhiteSpace($Model)) {
        $Model = 'kimi-k2.6:cloud'
    }

    $tempFile = Join-Path ([System.IO.Path]::GetTempPath()) ("hermes-agent-windows-Ollama-{0}.env" -f ([guid]::NewGuid().ToString('N')))
    try {
        $content = @(
            '# Created by hermes-agent-windows. Do not commit this file.'
            "export OLLAMA_API_KEY='$($ApiKey.Replace("'", "'\''"))'"
            "export HERMES_AGENT_WINDOWS_OLLAMA_MODEL='$($Model.Replace("'", "'\''"))'"
        ) -join "`n"
        Set-Content -Path $tempFile -Value $content -Encoding ASCII -Force
        $source = ConvertTo-WslWindowsPath -Path $tempFile
        $result = Invoke-WslAdminShell -Command "set -e; mkdir -p `"`$HOME`" `"`$HOME/.hermes`"; cp '$source' `"`$HOME/.ollama-cloud.env`"; chmod 600 `"`$HOME/.ollama-cloud.env`"; . `"`$HOME/.ollama-cloud.env`"; env_file=`"`$HOME/.hermes/.env`"; touch `"`$env_file`"; tmp=`"`$(mktemp)`"; grep -v -E '^(OLLAMA_API_KEY|OLLAMA_BASE_URL)=' `"`$env_file`" > `"`$tmp`" || true; cat `"`$tmp`" > `"`$env_file`"; printf 'OLLAMA_API_KEY=%s\n' `"`$OLLAMA_API_KEY`" >> `"`$env_file`"; printf 'OLLAMA_BASE_URL=https://ollama.com/v1\n' >> `"`$env_file`"; chmod 600 `"`$env_file`"; rm -f `"`$tmp`"; hermes config set model.provider ollama-cloud >/dev/null 2>&1 || true; hermes config set model.default `"`$HERMES_AGENT_WINDOWS_OLLAMA_MODEL`" >/dev/null 2>&1 || true; echo 'Ollama cloud config saved for model: $Model'" -TimeoutSeconds 60
        if ($result.Status -eq 'Success') {
            return Format-StatusResult -Name 'Ollama Cloud Config' -Status 'Installed' -Message 'Ollama Cloud API key and default model were saved in WSL.' -Details "Model: $Model`nFile: /home/admin/.ollama-cloud.env"
        }

        return Format-StatusResult -Name 'Ollama Cloud Config' -Status 'Error' -Message 'Failed to save Ollama Cloud config in WSL.' -Details $result.Details -ExitCode $result.ExitCode
    }
    catch {
        return Format-StatusResult -Name 'Ollama Cloud Config' -Status 'Error' -Message 'Failed to save Ollama Cloud config.' -Details $_.Exception.Message -ExitCode 1
    }
    finally {
        Remove-Item -LiteralPath $tempFile -Force -ErrorAction SilentlyContinue
    }
}

function ConvertTo-WslWindowsPath {
    param([Parameter(Mandatory)][string]$Path)

    $resolved = (Resolve-Path -LiteralPath $Path).Path
    $drive = $resolved.Substring(0, 1).ToLowerInvariant()
    $rest = $resolved.Substring(2).Replace('\', '/')
    return "/mnt/$drive$rest"
}

function Get-OllamaCloudConfig {
    $result = Invoke-WslAdminShell -Command 'if [ -f "$HOME/.ollama-cloud.env" ]; then . "$HOME/.ollama-cloud.env"; printf "Model: %s\nAPI key: %s\n" "${HERMES_AGENT_WINDOWS_OLLAMA_MODEL:-kimi-k2.6:cloud}" "$(if [ -n "$OLLAMA_API_KEY" ]; then echo saved; else echo missing; fi)"; else echo "Model: kimi-k2.6:cloud"; echo "API key: missing"; fi' -TimeoutSeconds 30
    if ($result.Status -eq 'Success') {
        $status = if ($result.Details -match 'API key:\s+saved') { 'Installed' } else { 'Missing' }
        return Format-StatusResult -Name 'Ollama Cloud Config' -Status $status -Message 'Ollama Cloud config checked.' -Details $result.Details
    }

    return Format-StatusResult -Name 'Ollama Cloud Config' -Status 'Unknown' -Message 'Could not check Ollama Cloud config.' -Details $result.Details -ExitCode $result.ExitCode
}

function Get-OllamaCloudModels {
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $json = Invoke-RestMethod -Uri 'https://ollama.com/api/tags' -TimeoutSec 45 -ErrorAction Stop
    }
    catch {
        return [pscustomobject]@{
            Status = 'Error'
            Message = 'Could not download Ollama cloud model list.'
            Details = $_.Exception.Message
            ExitCode = 1
            Models = @()
        }
    }

    try {
        $baseModels = @($json.models | ForEach-Object { $_.name } | Where-Object { $_ } | Sort-Object -Unique)
        $models = @()
        foreach ($baseModel in $baseModels) {
            $models += $baseModel
            if ($baseModel -notmatch ':cloud$') {
                $models += "$baseModel`:cloud"
            }
        }
        $models = @($models | Sort-Object -Unique)
        if (-not ($models -contains 'kimi-k2.6:cloud')) {
            $models = @('kimi-k2.6:cloud') + $models
        }

        return [pscustomobject]@{
            Status = 'Installed'
            Message = 'Ollama cloud model list downloaded.'
            Details = "Models: $($models.Count)"
            ExitCode = 0
            Models = $models
        }
    }
    catch {
        return [pscustomobject]@{
            Status = 'Error'
            Message = 'Could not parse Ollama cloud model list.'
            Details = $_.Exception.Message
            ExitCode = 1
            Models = @('kimi-k2.6:cloud')
        }
    }
}

function Test-OllamaCloudApi {
    param(
        [string]$Model = ''
    )

    $modelArg = if ($Model) { $Model } else { 'kimi-k2.6:cloud' }
    $payload = '{"model":"MODEL_PLACEHOLDER","messages":[{"role":"user","content":"Reply with only: ok"}],"stream":false}'
    $payload = $payload.Replace('MODEL_PLACEHOLDER', $modelArg)
    $encodedPayload = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($payload))
    $command = "test -f `"`$HOME/.ollama-cloud.env`" && . `"`$HOME/.ollama-cloud.env`"; test -n `"`$OLLAMA_API_KEY`" || { echo 'OLLAMA_API_KEY is missing'; exit 2; }; printf '%s' '$encodedPayload' | base64 -d > /tmp/hermes-agent-windows-ollama-test.json; curl -4 -fsS --max-time 60 https://ollama.com/api/chat -H `"Authorization: Bearer `$OLLAMA_API_KEY`" -H 'Content-Type: application/json' -d @/tmp/hermes-agent-windows-ollama-test.json | head -c 600"
    $result = Invoke-WslAdminShell -Command $command -TimeoutSeconds 90
    if ($result.Status -eq 'Success') {
        return Format-StatusResult -Name 'Ollama Cloud API' -Status 'Installed' -Message 'Ollama Cloud API test succeeded.' -Details "Model: $modelArg`nResponse received from ollama.com."
    }

    return Format-StatusResult -Name 'Ollama Cloud API' -Status 'Error' -Message 'Ollama Cloud API test failed.' -Details $result.Details -ExitCode $result.ExitCode
}

function Open-OllamaDownloadPage {
    return [pscustomobject]@{
        Status   = 'Unknown'
        Message  = 'Manual Windows Ollama download is not used by this project.'
        Details  = 'hermes-agent-windows installs Ollama inside WSL with https://ollama.com/install.sh.'
        ExitCode = 0
    }
}

function Install-Ollama {
    param(
        [switch]$OpenPageOnFailure = $true
    )

    $logFile = Get-LogFilePath -Kind 'app'
    if (-not (Test-WslExists)) {
        return [pscustomobject]@{
            Status   = 'Missing'
            Message  = 'WSL is required before installing Ollama.'
            Details  = 'Install WSL first, then rerun setup.'
            ExitCode = 1
        }
    }

    if (Test-OllamaAvailable) {
        return [pscustomobject]@{
            Status   = 'Installed'
            Message  = 'Ollama is already installed in WSL.'
            Details  = 'No installation was needed.'
            ExitCode = 0
        }
    }

    Write-Log -Message 'Installing Ollama inside WSL with the official Linux installer.' -Level 'INFO' -LogFile $logFile | Out-Null
    $install = Invoke-WslRootShell -Command 'curl -fsSL https://ollama.com/install.sh | sh' -TimeoutSeconds 900
    if ($install.Status -eq 'Success' -and (Test-OllamaAvailable)) {
        $start = Start-Ollama
        return [pscustomobject]@{
            Status   = 'Installed'
            Message  = 'Ollama was installed inside WSL.'
            Details  = "Install: $($install.Details)`nStart: $($start.Message)"
            ExitCode = 0
        }
    }

    return [pscustomobject]@{
        Status   = 'Error'
        Message  = 'Ollama installation in WSL failed.'
        Details  = $install.Details
        ExitCode = $install.ExitCode
    }
}

