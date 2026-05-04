# Hermes Agent for Windows

[![GitHub Pages](https://img.shields.io/badge/GitHub%20Pages-Live-blue)](https://jlaiii.github.io/hermes-agent-windows/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

**One-click Windows installer for [Hermes Agent](https://hermes-agent.nousresearch.com/) by Nous Research.**

Hermes Agent officially supports Linux, macOS, and WSL2 — not native Windows. This project **bridges that gap** for Windows users by automating the entire WSL2-based setup from end to end. You run one script and everything happens automatically.

**Live Site:** https://jlaiii.github.io/hermes-agent-windows/

**GitHub Repo:** https://github.com/jlaiii/hermes-agent-windows

---

## Who Is This For?

- **Windows users** who want to run Hermes Agent but don't want to manually set up WSL, Ubuntu, Ollama, and Linux dependencies
- **Non-technical users** who just want to click a `.bat` file or run a one-liner and have it work
- **Anyone** who wants Windows-native shortcuts and launchers for a Linux-only AI agent

Hermes Agent is not natively available for Windows. Normally you would have to:
1. Manually enable WSL2 in Windows
2. Install Ubuntu-22.04 from the Microsoft Store
3. Install Ollama inside Ubuntu
4. Install Hermes Agent inside Ubuntu
5. Figure out how to launch everything from Windows

**This installer does all of that for you.** It handles the reboot, resumes automatically, and gives you proper Windows desktop shortcuts and `.bat` launchers so you never need to type a WSL command manually.

---

## What It Does

- **Installs WSL2** — Enables Windows Subsystem for Linux with a single command.
- **Installs Ubuntu-22.04** — Sets up the LTS environment required by Hermes Agent.
- **Installs Ollama** — Pulls and configures the local LLM runtime.
- **Installs Hermes Agent** — Deploys the latest Hermes Agent release inside WSL.
- **Creates .bat Launchers** — One-click Windows batch files for every major component.
- **Creates Desktop Shortcuts** — Icons placed directly on your desktop for easy access.

---

## Prerequisites

- **Windows 10 version 20H2 or later**, or **Windows 11**
- **Administrator rights** (the PowerShell script requires elevation)
- **Internet connection** (downloads WSL, Ubuntu, Ollama, and Hermes Agent)

---

## Installation

Open **PowerShell as Administrator** and run:

```powershell
irm https://raw.githubusercontent.com/jlaiii/hermes-agent-windows/main/Install-Hermes-Windows.ps1 | iex
```

Or download `Install-Hermes-Windows.ps1` manually and execute it with right-click → **Run with PowerShell**.

---

## How It Works

1. **Check Environment** — Verifies Windows version and admin privileges.
2. **Enable WSL2** — Installs the Windows Subsystem for Linux 2 kernel and sets it as default.
3. **Install Ubuntu-22.04** — Downloads and registers the Ubuntu 22.04 LTS distro from the Microsoft Store.
4. **Install Ollama** — Runs the official Ollama installer inside Ubuntu.
5. **Install Hermes Agent** — Pulls the latest Hermes Agent package and installs it within the WSL environment.
6. **Create Launchers** — Generates `.bat` files in `%USERPROFILE%\hermes-agent-windows\`.
7. **Create Shortcuts** — Places Windows desktop shortcuts pointing to each launcher.

---

## Launchers

After installation you will find the following `.bat` files and desktop shortcuts:

| Launcher | Purpose |
|----------|---------|
| **Hermes CLI** | Start an interactive Hermes Agent terminal session |
| **Gateway** | Launch the Hermes Agent web gateway |
| **ACP Server** | Start the Agent-Computer Protocol server |
| **Ollama Server** | Start the Ollama LLM backend server |
| **Setup** | Re-run initial setup / repair the installation |

All launchers are stored in `%USERPROFILE%\hermes-agent-windows\` and are also available as desktop shortcuts.

---

## System Requirements

| Component | Minimum | Recommended |
|-----------|---------|-------------|
| OS | Windows 10 20H2 | Windows 11 latest |
| RAM | 8 GB | 16 GB or more |
| Disk | 20 GB free | 50 GB free |
| CPU | 64-bit, virtualization enabled | Modern multi-core |
| Network | Broadband | Broadband |

> **Note:** Virtualization (Hyper-V / VT-x) must be enabled in BIOS for WSL2.

---

## Troubleshooting

### WSL Restart Required
If you see a message asking to restart after WSL installation, **reboot your computer** and re-run the installer. It will resume where it left off.

### WSL Not Responding
```powershell
wsl --shutdown
wsl --update
```
Then try the launcher again.

### Hermes Not Found
If `hermes` commands are not recognized inside WSL:
```bash
# Inside Ubuntu-22.04
source ~/.bashrc
hash -r
```
If the issue persists, run the **Setup** launcher or reinstall Hermes Agent.

---

## Uninstall

To remove Hermes Agent for Windows:

1. Delete the launchers folder:
   ```powershell
   Remove-Item -Recurse -Force "$env:USERPROFILE\hermes-agent-windows"
   ```
2. Remove the desktop shortcuts manually.
3. (Optional) Uninstall WSL and Ubuntu-22.04:
   ```powershell
   wsl --unregister Ubuntu-22.04
   ```
4. (Optional) Remove Hermes Agent files inside WSL by deleting its installation directory (e.g., `~/.local/share/hermes-agent/`).

---

## Contributing

Contributions are welcome! Please feel free to open an [issue](../../issues) or submit a [pull request](../../pulls) on GitHub.

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

---

## License

This project is licensed under the MIT License — see the [LICENSE](LICENSE) file for details.

Copyright (c) 2026 Jay / [jlaiii](https://github.com/jlaiii)
