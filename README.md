# Star Citizen Deposit Scanner

🚀 Read Star Citizen mining deposit codes automatically and see the results in-game or on another device.

[![GitHub release](https://img.shields.io/github/v/release/FrozenButton/Scanning-Tool)](https://github.com/FrozenButton/Scanning-Tool/releases)
[![GitHub stars](https://img.shields.io/github/stars/FrozenButton/Scanning-Tool)](https://github.com/FrozenButton/Scanning-Tool/stargazers)
[![GitHub issues](https://img.shields.io/github/issues/FrozenButton/Scanning-Tool)](https://github.com/FrozenButton/Scanning-Tool/issues)

## What it does
- Captures the HUD deposit code, reads it with the `qwen3-vl:2b` model, and shows the deposit type/quantity.

**Note:** The scanner now uses the `qwen3-vl:2b` model, which offers better OCR accuracy, faster processing, and lower VRAM usage compared to the previous `qwen2.5vl:3b` model. This upgrade is specifically beneficial for reading Star Citizen HUD codes in real-time, even on systems with limited GPU resources.
- Locks to a user-selected HUD anchor so the capture box follows head sway and ship movement.
- Hosts a small web overlay (desktop + mobile friendly) so you can view the latest scan from a browser; a GUI button opens it for you.

## Requirements (minimum that still works well)
- GPU with ~2GB free VRAM (tool uses ~1.73GB) or fast CPU fallback.
- 32GB system RAM recommended when running Star Citizen + the scanner.
- Windows 10/11 or modern 64-bit Linux.

## Quick start
### Windows (recommended)
1) Download the latest release or clone this repo.
2) Double-click `launch_windows.bat` or run `cmd.exe /c launch_windows.bat` from Git Bash.
3) Wait for the first-run downloads (Python + dependencies). The scanner will start and auto-launch Ollama with `ollama serve` if it is not already running.

### Windows using Git Bash and uv
1) Open a Git Bash shell in the repo folder.
2) Run `./launch_windows.sh` (make it executable first with `chmod +x launch_windows.sh` if needed).
3) The script uses `uv` to manage Python and install requirements.

### Linux
1) Download/clone the repo and open a terminal in the folder.
2) Run `./launch_linux.sh` (make it executable if needed: `chmod +x launch_linux.sh`).
3) Follow any prompts. Ollama is auto-started the same way as on Windows.

## Ollama setup
- **Same PC (default):** Install Ollama from [ollama.com](https://ollama.com/) and restart/log back in once. The scanner will connect to `http://127.0.0.1:11434` by default and start the service automatically when needed.
- **Remote PC:** Install and allow network access on the Ollama machine (`OLLAMA_HOST=0.0.0.0`, firewall port 11434 open). In the scanner GUI, enter the remote host (e.g., `http://192.168.1.42:11434`) in **Ollama Connection → Apply Host**. You can also set `OLLAMA_HOST` in `config.json` or as an environment variable before launching.

## Using the scanner
- **Positioning:** Use the capture sliders to place the red box over the deposit code. Toggle visibility with **8**.
- **Anchor alignment:** Place the cyan anchor frame over a stable HUD icon. Load templates from `assets/anchor_templates/`, then click **Realign Now** and tweak offsets until the capture locks on target. Auto-alignment runs before each scan when enabled.
- **Scanning:**
  - **Hotkeys:** `7` = single scan, `Ctrl+7` = start/stop auto-scan (interval slider), `8` = show/hide capture box.
  - **Buttons:** Use **Single Scan** / **Loop Toggle** if hotkeys are blocked.
- **Overlay in a browser/phone:** When the app starts it prints links like `http://127.0.0.1:5000` and `http://LAN_IP:5000`. Click **Open Mobile Overlay** in the GUI to launch your default browser. The page auto-refreshes with each scan.

## Troubleshooting (fast fixes)
- **Python not found / dependency errors:** Re-run the launch script; it reinstalls what is needed.
- **Ollama missing:** Install from [ollama.com](https://ollama.com/), then relaunch. The scanner will start `ollama serve` for you.
- **Remote Ollama unreachable:** Confirm the LAN IP/port 11434, firewall rules, and that `OLLAMA_HOST` matches the remote address.
- **Hotkeys blocked:** Use the on-screen buttons.
- **No overlays:** Click **Update Overlay** and ensure Star Citizen is focused.

## Need help or want to contribute?
- File issues or feature requests on [GitHub Issues](https://github.com/FrozenButton/Scanning-Tool/issues).
- PRs are welcome—please include a short description of your change and testing steps.

Happy mining! 🪨⛏️
