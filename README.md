# hermes-agent-windows

hermes-agent-windows is a Windows-friendly, PowerShell-only setup tool for Hermes Agent. Windows runs the installer and GUI, but Ollama and Hermes Agent are installed and managed inside WSL.

## What It Does

- Checks Administrator access, Windows version, PowerShell version, WSL, WSL distros, WSL accounts, Ollama, Hermes Agent, and Hermes Gateway
- Installs WSL when missing and clearly warns when Windows needs a reboot
- Creates or resets a WSL helper account named `admin` with password `admin`
- Installs and starts Ollama inside WSL, not on the Windows host
- Saves an optional Ollama Cloud API key in WSL and defaults Hermes to `kimi-k2.6:cloud`
- Installs Hermes Agent inside WSL with the official Nous Research Linux installer
- Starts Hermes Gateway in WSL-friendly foreground/background mode
- Installs or removes Windows Desktop and Start Menu shortcuts for the hermes-agent-windows GUI
- Provides a WPF GUI with live logs, status cards, and safe rerunnable actions
- Logs to `logs/install.log` and `logs/app.log`

## Requirements

- Windows 10 or Windows 11
- Administrator permission for WSL install/repair. The `.bat`, `install.ps1`, and `hermes-agent-windows.ps1` launchers request Admin permission automatically when possible.
- PowerShell 5.1 or PowerShell 7
- Internet access
- WSL 2 recommended

## One-Command Install

```powershell
irm https://jlaiii.github.io/hermes-agent-windows/install.ps1 | iex
```

You can also set `HERMES_AGENT_WINDOWS_BASE_URL` to override the default GitHub URL.

## Manual Run

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\hermes-agent-windows.ps1
```

If you downloaded the full repo as a ZIP or cloned it with Git, you can also run:

```bat
hermes-agent-windows.bat
```

Tip: `hermes-agent-windows.bat`, `install.ps1`, and `hermes-agent-windows.ps1` request Administrator permission automatically. If Windows blocks the prompt, right-click the file and choose **Run as administrator**.

## GUI Features

- Start full WSL-first setup
- Install or uninstall the hermes-agent-windows Windows shortcut
- Check status
- Install WSL
- Restart WSL
- Create/reset the WSL `admin/admin` helper account
- Install Ollama in WSL
- Start Ollama
- Save an Ollama Cloud API key
- Refresh/search Ollama Cloud models
- Test Ollama Cloud access with the selected model
- Install Hermes Agent in WSL
- Update Hermes Agent
- Run Hermes Doctor health checks
- Launch the interactive Hermes Agent CLI/chat in a normal Windows Command Prompt through WSL
- Enable Hermes Gateway
- Open the Hermes web dashboard
- Start, stop, and restart Hermes Agent
- Clean Hermes WSL files
- Reinstall or wipe the default WSL distro with warning confirmations
- Open config and logs folders

## Installing The GUI Shortcut

After the one-command setup opens the GUI, press **Install hermes-agent-windows Shortcut**. This adds:

- Desktop shortcut: `hermes-agent-windows.lnk`
- Start Menu shortcut: `hermes-agent-windows`

The shortcut launches:

```powershell
powershell.exe -STA -ExecutionPolicy Bypass -File .\hermes-agent-windows.ps1
```

Press **Uninstall Shortcut** to remove those shortcuts. This does not delete WSL, Ollama, Hermes Agent, logs, or the project folder.

## WSL Handling

WSL installation uses:

```powershell
wsl --install
```

If a reboot is likely required, hermes-agent-windows stops WSL-dependent setup and tells the user to restart Windows.

## Ollama Handling

Ollama is checked and installed inside WSL:

```bash
curl -fsSL https://ollama.com/install.sh | sh
```

The GUI starts Ollama with `ollama serve` inside WSL and checks the local WSL API.

## Ollama Cloud API

The GUI has an **Ollama API Key** password field and a searchable model picker. Press:

- **Save Cloud API** to save the key in WSL at `/home/admin/.ollama-cloud.env` and Hermes secrets at `/home/admin/.hermes/.env`
- **Refresh Models** to download the current Ollama model list from `https://ollama.com/api/tags`
- **Test Cloud API** to send a small test request to `https://ollama.com/api/chat`

The default model is:

```text
kimi-k2.6:cloud
```

The API key is not hardcoded into this project.

## Hermes Agent Handling

Hermes Agent is installed inside WSL:

```bash
curl -fsSL https://raw.githubusercontent.com/NousResearch/hermes-agent/main/scripts/install.sh | bash -s -- --skip-setup
```

The setup uses the WSL `admin` account and adds the Hermes venv path to commands so non-login WSL checks can find `hermes`.

## Gateway Handling

Hermes Gateway is started with the WSL-recommended mode:

```bash
nohup hermes gateway run --accept-hooks > ~/.hermes/logs/gateway.log 2>&1 &
```

Status is checked with:

```bash
hermes gateway status
```

The **Open Dashboard** button starts the Hermes web dashboard in WSL and opens:

```text
http://localhost:9119
```

## Logs

- Installer log: `logs/install.log`
- App log: `logs/app.log`
- WSL Ollama log: `~/.ollama/ollama.log`
- WSL Hermes Gateway log: `~/.hermes/logs/gateway.log`

## Safety Notes

- Safe to rerun: checks happen before installs.
- WSL wipe/reinstall buttons require explicit confirmation.
- Do not run the wipe/reinstall WSL actions unless you have backed up Linux files.
- The `admin/admin` WSL helper account is convenient for setup. Change or remove it later on shared machines.
- Only run scripts from sources you trust.

## Troubleshooting

- If local scripts are blocked, use the manual run command with `-ExecutionPolicy Bypass`.
- If WSL install requests a reboot, restart Windows and rerun setup.
- If Ollama is installed but stopped, press **Start Ollama**.
- If Hermes shows missing after a long install, press **Check Status**. The installer may have finished but the shell PATH may need refreshing.
- If Gateway is stopped, press **Enable Gateway** and check `~/.hermes/logs/gateway.log`.

## Author

Built by [Jay (jlaiii)](https://github.com/jlaiii)

- GitHub: [jlaiii](https://github.com/jlaiii)
- Website: [jlaiii.github.io/hermes-agent-windows](https://jlaiii.github.io/hermes-agent-windows)

