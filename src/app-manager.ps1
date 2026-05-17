if (-not (Get-Command Write-Log -ErrorAction SilentlyContinue)) {
    . (Join-Path $PSScriptRoot 'utils.ps1')
}

function Get-hermes-agent-windowsShortcutInfo {
    $projectRoot = Get-ProjectRoot
    $desktop = [Environment]::GetFolderPath('Desktop')
    $programs = [Environment]::GetFolderPath('Programs')
    $startMenuFolder = Join-Path $programs 'hermes-agent-windows'

    [pscustomobject]@{
        ProjectRoot     = $projectRoot
        LaunchScript    = Join-Path $projectRoot 'hermes-agent-windows.ps1'
        DesktopShortcut = Join-Path $desktop 'hermes-agent-windows.lnk'
        StartMenuFolder = $startMenuFolder
        StartMenuShortcut = Join-Path $startMenuFolder 'hermes-agent-windows.lnk'
        UninstallShortcut = Join-Path $startMenuFolder 'Uninstall hermes-agent-windows Shortcut.lnk'
    }
}

function New-hermes-agent-windowsShortcut {
    param(
        [Parameter(Mandatory)]
        [string]$ShortcutPath,
        [Parameter(Mandatory)]
        [string]$LaunchScript,
        [string]$Description = 'hermes-agent-windows'
    )

    $shortcutDir = Split-Path -Parent $ShortcutPath
    if ($shortcutDir -and -not (Test-Path $shortcutDir)) {
        New-Item -ItemType Directory -Path $shortcutDir -Force | Out-Null
    }

    $powershell = if (Get-Command pwsh.exe -ErrorAction SilentlyContinue) {
        (Get-Command pwsh.exe).Source
    }
    else {
        (Get-Command powershell.exe).Source
    }

    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($ShortcutPath)
    $shortcut.TargetPath = $powershell
    $shortcut.Arguments = "-STA -ExecutionPolicy Bypass -File `"$LaunchScript`""
    $shortcut.WorkingDirectory = Split-Path -Parent $LaunchScript
    $shortcut.Description = $Description
    $shortcut.IconLocation = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe,0"
    $shortcut.Save()
}

function Install-hermes-agent-windowsApp {
    $logFile = Get-LogFilePath -Kind 'app'
    $info = Get-hermes-agent-windowsShortcutInfo

    if (-not (Test-Path $info.LaunchScript)) {
        return Format-StatusResult -Name 'Install hermes-agent-windows App' -Status 'Error' -Message 'Cannot install app shortcut because hermes-agent-windows.ps1 is missing.' -Details $info.LaunchScript -ExitCode 1
    }

    try {
        Write-Log -Message 'Installing hermes-agent-windows Windows shortcuts.' -Level 'INFO' -LogFile $logFile | Out-Null
        New-hermes-agent-windowsShortcut -ShortcutPath $info.DesktopShortcut -LaunchScript $info.LaunchScript -Description 'Open hermes-agent-windows'
        New-hermes-agent-windowsShortcut -ShortcutPath $info.StartMenuShortcut -LaunchScript $info.LaunchScript -Description 'Open hermes-agent-windows'

        $uninstallCommand = Join-Path $info.ProjectRoot 'hermes-agent-windows.ps1'
        New-hermes-agent-windowsShortcut -ShortcutPath $info.UninstallShortcut -LaunchScript $uninstallCommand -Description 'Open hermes-agent-windows to uninstall shortcuts'

        return Format-StatusResult -Name 'Install hermes-agent-windows App' -Status 'Installed' -Message 'hermes-agent-windows shortcuts were installed.' -Details "Desktop: $($info.DesktopShortcut)`nStart Menu: $($info.StartMenuShortcut)"
    }
    catch {
        return Format-StatusResult -Name 'Install hermes-agent-windows App' -Status 'Error' -Message 'Failed to install hermes-agent-windows shortcuts.' -Details $_.Exception.Message -ExitCode 1
    }
}

function Uninstall-hermes-agent-windowsApp {
    $logFile = Get-LogFilePath -Kind 'app'
    $info = Get-hermes-agent-windowsShortcutInfo
    $removed = New-Object System.Collections.Generic.List[string]

    try {
        foreach ($path in @($info.DesktopShortcut, $info.StartMenuShortcut, $info.UninstallShortcut)) {
            if (Test-Path $path) {
                Remove-Item -LiteralPath $path -Force
                $removed.Add($path)
            }
        }

        if ((Test-Path $info.StartMenuFolder) -and -not (Get-ChildItem -LiteralPath $info.StartMenuFolder -Force -ErrorAction SilentlyContinue)) {
            Remove-Item -LiteralPath $info.StartMenuFolder -Force
        }

        Write-Log -Message 'hermes-agent-windows Windows shortcuts were removed.' -Level 'SUCCESS' -LogFile $logFile | Out-Null
        return Format-StatusResult -Name 'Uninstall hermes-agent-windows App' -Status 'Stopped' -Message 'hermes-agent-windows shortcuts were removed.' -Details (($removed | ForEach-Object { $_ }) -join "`n")
    }
    catch {
        return Format-StatusResult -Name 'Uninstall hermes-agent-windows App' -Status 'Error' -Message 'Failed to remove hermes-agent-windows shortcuts.' -Details $_.Exception.Message -ExitCode 1
    }
}

function Get-hermes-agent-windowsAppStatus {
    $info = Get-hermes-agent-windowsShortcutInfo
    $installed = (Test-Path $info.DesktopShortcut) -or (Test-Path $info.StartMenuShortcut)

    if ($installed) {
        return Format-StatusResult -Name 'hermes-agent-windows App' -Status 'Installed' -Message 'hermes-agent-windows shortcuts are installed.' -Details "Desktop: $([bool](Test-Path $info.DesktopShortcut))`nStart Menu: $([bool](Test-Path $info.StartMenuShortcut))"
    }

    return Format-StatusResult -Name 'hermes-agent-windows App' -Status 'Missing' -Message 'hermes-agent-windows shortcuts are not installed.' -Details 'Use Install hermes-agent-windows App to add Desktop and Start Menu shortcuts.'
}

