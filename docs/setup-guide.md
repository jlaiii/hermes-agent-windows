# hermes-agent-windows Setup Guide

## Step-by-Step Usage

1. Open PowerShell. Administrator PowerShell is best, but the installer will request Administrator permission automatically when possible.
2. Paste the one-command installer from the README.
3. Let the bootstrap prepare the local project files.
4. In the GUI, press **Check Status**.
5. Press **Start Hermes Agent Setup** to run the WSL-first setup.
6. Review the live log panel while it checks WSL, creates the WSL `admin` account, starts Ollama, installs Hermes, and checks the gateway.

## Downloaded ZIP Or Git Clone Usage

1. Download or clone the full `hermes-agent-windows` folder.
2. Extract it if it came from a ZIP file.
3. Double-click `hermes-agent-windows.bat`.
4. If Windows asks for permission, choose **Yes**.
5. The batch file launches the PowerShell GUI with execution-policy bypass.

## Running As Admin

- Open Start.
- Search for PowerShell.
- Right-click PowerShell.
- Select **Run as administrator**.

WSL installation and repair usually need Admin rights. `hermes-agent-windows.bat`, `install.ps1`, and `hermes-agent-windows.ps1` request Administrator permission automatically. If Windows blocks the prompt, right-click the file and choose **Run as administrator**.

## Button Guide

- **Start Hermes Agent Setup**: Runs the full WSL-first setup flow.
- **Install hermes-agent-windows Shortcut**: Adds Desktop and Start Menu shortcuts for opening the GUI later.
- **Uninstall Shortcut**: Removes the hermes-agent-windows shortcuts without deleting WSL or Hermes data.
- **Check Status**: Refreshes every status card.
- **Install WSL**: Runs `wsl --install` and warns that a reboot may be required.
- **Restart WSL**: Runs `wsl --shutdown`, then starts the default distro again.
- **Create/Reset WSL Admin**: Creates or resets WSL user `admin` with password `admin`.
- **Install Ollama in WSL**: Installs Ollama inside WSL with the official Linux installer.
- **Start Ollama**: Starts `ollama serve` inside WSL.
- **Save Cloud API**: Saves the pasted Ollama API key in WSL and configures Hermes for Ollama Cloud.
- **Refresh Models**: Downloads the current Ollama model list and fills the model picker.
- **Test Cloud API**: Verifies the selected cloud model works with the saved key.
- **Install Hermes in WSL**: Installs Hermes Agent into `/home/admin/.hermes`.
- **Update Hermes Agent**: Runs `hermes update` inside WSL.
- **Hermes Doctor**: Runs `hermes doctor` inside WSL and prints the health-check output in the live log.
- **Launch Hermes CLI**: Opens the interactive Hermes Agent chat CLI in a normal Windows Command Prompt through WSL as `admin`, using `hermes chat --accept-hooks`.
- **Enable Gateway**: Starts `hermes gateway run --accept-hooks` in the background inside WSL.
- **Open Dashboard**: Starts the Hermes web dashboard in WSL and opens `http://localhost:9119`.
- **Start Hermes Agent**: Attempts to launch Hermes from WSL.
- **Stop Hermes Agent**: Stops Hermes-related WSL processes.
- **Restart Hermes Agent**: Stops then starts Hermes from WSL.
- **Clean Hermes WSL Files**: Removes user-level Hermes files inside WSL only.
- **Reinstall WSL**: Unregisters the default WSL distro, then runs WSL install. This can delete Linux files.
- **Wipe WSL**: Unregisters the default WSL distro. This deletes that distro.
- **Open Config Folder**: Opens the WSL Hermes config folder through Windows Explorer.
- **Open Logs Folder**: Opens local hermes-agent-windows logs.
- **Exit**: Closes the GUI.

## What WSL Is

WSL means Windows Subsystem for Linux. hermes-agent-windows uses WSL as the Linux environment where Ollama and Hermes Agent live.

## What Ollama Is

Ollama is a local model runtime. In this project, Ollama runs inside WSL so Hermes can use it from the Linux side.

## Ollama Cloud Setup

1. Paste your Ollama API key into **Ollama API Key**.
2. Keep the default model `kimi-k2.6:cloud`, or press **Refresh Models** and pick another model.
3. Press **Save Cloud API**.
4. Press **Test Cloud API**.

The key is stored inside WSL under the `admin` account. It is not written into the GitHub project files.

## What Hermes Agent Setup Does

The setup installs Hermes Agent under the WSL `admin` account, checks its version, and can start Hermes Gateway. If a Hermes command is not available on PATH, the setup also checks the installed Hermes virtual environment path.

## Common Problems And Fixes

### Scripts Blocked

Run:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\hermes-agent-windows.ps1
```

### WSL Needs Reboot

Restart Windows, reopen Admin PowerShell, and run setup again.

### Ollama Installed But Stopped

Press **Start Ollama**. Logs are in WSL at `~/.ollama/ollama.log`.

### Hermes Installed But Not Running

This can be normal because Hermes is primarily a CLI. Press **Enable Gateway** if you want the background gateway.

### Gateway Stopped

Press **Enable Gateway** and review `~/.hermes/logs/gateway.log`.

### Shortcut Missing

Open the GUI manually, then press **Install hermes-agent-windows Shortcut**.

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\hermes-agent-windows.ps1
```

### Wipe Or Reinstall WSL

Use these only when you have backups. They can delete the default WSL distro's Linux filesystem.

