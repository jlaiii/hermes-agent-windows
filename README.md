# hermes-agent-windows

Run **Hermes Agent** on Windows without touching a single Linux command. This does all the heavy lifting for you -- one PowerShell command, one WPF GUI, and your entire Hermes + Ollama + WSL environment is ready to go inside WSL.

## The Short Version

**Paste this in PowerShell (as Admin):**

```powershell
irm https://jlaiii.github.io/hermes-agent-windows/install.ps1 | iex
```

That is it. It downloads everything, installs WSL if missing, sets up Ollama, installs Hermes Agent, and opens a GUI where you control the whole stack with buttons. No `apt-get`. No `systemctl`. No config files to edit by hand.

## What This Does For You

- **Checks your system** -- Windows version, PowerShell version, WSL presence, admin rights. Tells you exactly what is missing.
- **Installs WSL** when needed. Warns you if Windows needs a reboot.
- **Creates a WSL helper account** (`admin` / `admin`) so you never have to figure out WSL user setup.
- **Installs Ollama inside WSL** -- not on your Windows host. Automatically starts it, manages the API key, and lets you pick cloud models.
- **Installs Hermes Agent inside WSL** using the official Nous Research installer.
- **Starts Hermes Gateway** and opens the web dashboard at `localhost:9119`.
- **Gives you a WPF GUI** on Windows with real-time status cards and a live log panel. Start, stop, update, or wipe anything without leaving Windows.
- **Adds Desktop and Start Menu shortcuts** so you can reopen the GUI anytime.
- **Safe to re-run** -- every action checks state first. Nothing gets double-installed.

## One-Command Install

Open PowerShell as Administrator and paste:

```powershell
irm https://jlaiii.github.io/hermes-agent-windows/install.ps1 | iex
```

If Windows blocks the script, right-click PowerShell and choose **Run as administrator**, then paste again.

You can override the download URL by setting the environment variable `HERMES_AGENT_WINDOWS_BASE_URL` before running the command.

## Manual Options

Already downloaded the ZIP or cloned the repo? No problem.

**PowerShell (any terminal):**
```powershell
powershell.exe -ExecutionPolicy Bypass -File .\hermes-agent-windows.ps1
```

**Double-click (easiest):**
```bat
hermes-agent-windows.bat
```

Both auto-request admin rights and handle execution policy for you.

## The GUI At A Glance

The WPF window shows color-coded status cards and lets you:

- Run the full setup with one button
- Install or uninstall Windows shortcuts
- Check system status at any time
- Install, restart, or wipe WSL
- Create or reset the WSL `admin` helper account
- Install, start, or stop Ollama
- Save your Ollama Cloud API key and pick models
- Install, update, run doctor checks on Hermes Agent
- Start, stop, or restart Hermes Gateway
- Open the Hermes web dashboard
- Clean Hermes files or reinstall the distro
- Open config and log folders directly

Everything logs to `logs/install.log` and `logs/app.log`.

## WSL + Ollama + Hermes -- Handled For You

- **WSL install:** Uses `wsl --install`. Reboot required? It tells you instead of crashing.
- **Ollama install:** Runs `curl -fsSL https://ollama.com/install.sh | sh` inside WSL. Starts with `ollama serve`.
- **Hermes install:** Runs the official Linux installer inside WSL. Adds venv paths automatically.
- **Gateway:** Starts in WSL-friendly background mode. `hermes gateway status` checks health.

## Safety Notes

- Safe to rerun -- installs only happen when something is missing.
- WSL wipe/reinstall buttons ask twice. Back up your Linux files first.
- The `admin/admin` account is for setup convenience only -- change it on shared machines.
- Only run scripts from sources you trust.

## Troubleshooting

| Problem | Fix |
|---|---|
| Script is blocked by execution policy | Use `-ExecutionPolicy Bypass` or double-click `hermes-agent-windows.bat` |
| WSL install asks for reboot | Restart Windows, rerun the installer |
| Ollama installed but stopped | Press **Start Ollama** in the GUI |
| Hermes shows missing after long install | Press **Check Status** to refresh the PATH detection |
| Gateway stopped | Press **Enable Gateway** and check `~/.hermes/logs/gateway.log` |

## Requirements

- Windows 10 or Windows 11
- Administrator rights (auto-requested by `.bat` and installer)
- PowerShell 5.1 or PowerShell 7
- Internet connection
- WSL 2 recommended

## Author

Built by [Jay (jlaiii)](https://github.com/jlaiii)

- GitHub: [jlaiii](https://github.com/jlaiii)
- Website: [jlaiii.github.io/hermes-agent-windows](https://jlaiii.github.io/hermes-agent-windows)
