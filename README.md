# Star Citizen Deposit Scanner

🚀 **An easy-to-use tool that automatically reads and identifies Star Citizen mining deposit codes from your screen!**

[![GitHub release](https://img.shields.io/github/v/release/FrozenButton/Scanning-Tool)](https://github.com/FrozenButton/Scanning-Tool/releases)
[![GitHub stars](https://img.shields.io/github/stars/FrozenButton/Scanning-Tool)](https://github.com/FrozenButton/Scanning-Tool/stargazers)
[![GitHub issues](https://img.shields.io/github/issues/FrozenButton/Scanning-Tool)](https://github.com/FrozenButton/Scanning-Tool/issues)

## 📥 Download

**Get the latest version from GitHub:**

### Option 1: Download ZIP (Easiest)
1. Click the **green "Code" button** at the top of this page
2. Click **"Download ZIP"**
3. **Extract the ZIP file** to any folder on your computer
4. Follow the instructions below to run it

### Option 2: Git Clone (For developers)
```bash
git clone https://github.com/FrozenButton/Scanning-Tool.git
cd Scanning-Tool
```

### Option 3: Direct Release Download
- Go to **[Releases](https://github.com/FrozenButton/Scanning-Tool/releases)** and download the latest version

---

## What does this do?

This tool helps Star Citizen miners by:
- 📸 **Taking screenshots** of deposit codes on your screen
- 🤖 **Reading the numbers** using AI
- 📊 **Telling you what type of deposit** it is and how valuable the materials are
- 📍 **Showing an overlay** on your screen with the deposit information
- 🎯 **Locking onto a HUD anchor** so the capture box follows head sway and ship movement

## 💻 System Requirements

**⚠️ This tool is designed for mid to high-end gaming PCs and is NOT intended for low-spec machines.**

Since this tool runs alongside Star Citizen, you'll need:

### Recommended Requirements
- **VRAM**: 8GB+ dedicated graphics memory (tool uses ~1.73GB for AI model)
- **RAM**: 32GB+ system memory (Star Citizen + this tool)
- **CPU**: Modern multi-core processor
- **OS**: Windows 10/11 or Linux (64-bit)

### Why these requirements?
- The AI model (**qwen2.5vl:3b**) requires **~1.73GB of VRAM** to run efficiently
- Star Citizen is already demanding on system resources
- Running both simultaneously requires adequate hardware

> 💡 **Performance Tip**: If you experience lag or performance issues, close the scanner when not actively mining, or consider upgrading your graphics card or RAM.

> ⚠️ **VRAM Note**: If your graphics card doesn't have enough VRAM, Ollama will automatically fall back to using your CPU for AI processing. This will work but will be significantly slower and may impact game performance more than GPU processing.

---

## 🎮 How to use it (Super Easy!)

### For Windows Users (Most People)

1. **Download this tool** from GitHub (see Download section above)
2. **Extract the ZIP file** to any folder on your computer
3. **Double-click** the file called `launch_windows.bat`
4. **Wait** - The tool will automatically download everything it needs (this might take a few minutes the first time)
5. **That's it!** The scanner will open and be ready to use

> ⚠️ **First time setup takes 2-5 minutes** while it downloads Python and other components. After that, it starts instantly!

### Windows standalone EXE (no Python required for end users)

If you want to ship a single folder that works on Windows PCs **without requiring Python to be installed**, build the bundled EXE and share the resulting `dist/` directory:

1. On a Windows machine, open **Command Prompt** in the repository root.
2. Run `build_windows_exe.bat`.
   - The script downloads a portable Python 3.13.7, bootstraps `pip`, installs **PyInstaller** plus app requirements, and then creates the EXE.
3. After the build finishes, the packaged app lives at `dist/StarCitizenScanningTool/StarCitizenScanningTool.exe`.
4. Zip the entire `dist/StarCitizenScanningTool` folder and distribute it—recipients can run the EXE directly without setting up Python.

### For Linux Users

1. **Download this tool** from GitHub (see Download section above)
2. **Extract the files** to any folder on your computer
3. **Open a terminal** in that folder
4. **Type:** `./launch_linux.sh` and press Enter
5. **Follow any prompts** to install Python if needed
6. **That's it!** The scanner will open and be ready to use

## 🎯 What you need to install separately

### Ollama (Required for AI scanning)

**This tool needs "Ollama" to read the deposit codes with AI.**

You can run Ollama on the same computer as the game **or** on a different machine on your local network. The scanner supports both setups—choose the one that best fits your hardware.

#### Option A: Ollama and the scanner on the same PC (default)
1. Go to **https://ollama.com/** in your web browser.
2. Click the **download button for Windows** (or Linux if you use Linux).
3. **Run the installer** and follow the prompts.
4. **Restart your computer** when it's done (Windows) or log out/in (Linux) so the Ollama service starts.

This is the simplest approach. After the installer finishes, start the scanner normally (`launch_windows.bat` or `./launch_linux.sh`) and it will talk to the Ollama service that is already running on `http://127.0.0.1:11434`.
- No extra configuration is required—the scanner falls back to `http://127.0.0.1:11434` automatically when `OLLAMA_HOST` is not set.

#### Option B: Run Ollama on another PC on your LAN
Use this if you want to keep the heavy AI workload off your gaming rig. You will set up two machines:

**1. AI / Ollama PC**

- Install Ollama from **https://ollama.com/** (Windows or Linux).
- Allow it to listen on the network:
  - **Windows (PowerShell):**
    ```powershell
    setx OLLAMA_HOST "0.0.0.0"
    Stop-Service Ollama
    Start-Service Ollama
    ```
  - **Linux:** add `OLLAMA_HOST=0.0.0.0` to `/etc/systemd/system/ollama.service` or `~/.ollama/config` (create it if it doesn't exist) and restart with `sudo systemctl restart ollama`.
- Ensure port **11434** is allowed through the firewall so the gaming PC can connect.
- (Optional) Test locally with `curl http://127.0.0.1:11434/api/tags`.

**2. Gaming / Scanner PC**

- Make sure both PCs are on the same network and note the LAN IP of the Ollama PC (for example `192.168.1.42`).
- Tell the scanner where to connect. You can now do this directly inside the app or by editing the config file:
  - Launch the scanner and open the **Ollama Connection** section. Enter the LAN address (for example `192.168.1.42:11434`) and click **Apply Host**. The value is saved in `config.json` under `OLLAMA_HOST` for future runs.
  - Alternatively, edit `config.json` manually and set the `"OLLAMA_HOST"` field to the desired URL.
  - Prefer environment variables instead? Set `OLLAMA_HOST` before launching the scanner so it knows where to send requests:
    - **Windows (PowerShell):** `setx OLLAMA_HOST "http://192.168.1.42:11434"`
    - **Windows (Command Prompt - temporary for this session):** `set OLLAMA_HOST=http://192.168.1.42:11434`
    - **Linux (temporary):** `export OLLAMA_HOST=http://192.168.1.42:11434`
    - **Linux (persistent):** add the export line to your shell profile (e.g. `~/.bashrc`).
- Launch the scanner (`launch_windows.bat` or `./launch_linux.sh`). It will use the remote Ollama instance automatically.
- The scanner detects that you're using a remote host and skips the local installer prompt. You'll see log messages confirming the remote address and whether the model is present on that machine.

**Troubleshooting remote mode**

- Verify you can reach the Ollama PC from the gaming PC with `curl http://192.168.1.42:11434/api/tags`.
- Double-check firewalls and that the Ollama service is running.
- If you switch back to the single-PC setup, clear the environment variable (`setx OLLAMA_HOST ""` on Windows or remove the export line on Linux) so the scanner defaults to `127.0.0.1` again.

> 💡 **Don't worry!** If you forget to install Ollama, the scanner tool will remind you and even open the website for you.

## 🎮 How to use the scanner

### Setting up the scan area

1. **Start Star Citizen** and go to a mining area
2. **Open the scanner tool** (it will show a red capture box on your screen)
3. **Drag the sliders** in the scanner window to move the red box over where deposit codes appear
4. **Make the red box** just big enough to cover the deposit code numbers

### 🧭 Head sway compensation (anchor alignment)

The tool can now follow the in-game HUD even when head sway is enabled. It does this by
locking onto a **stable anchor icon** on your ship screen and then adjusting the red
capture box using an offset.

1. **Show the anchor overlay** – The cyan “Anchor Region” frame is always-on-top so you
   can position it over a reliable HUD element. Toggle its visibility with the
   **"Show anchor overlay"** checkbox if you need a clearer view.
2. **Position the anchor region** – Use the *Anchor Left/Top/Width/Height* sliders in the
   **Head Sway Compensation** panel to cover the icon or shape you want to track.
3. **Capture or add templates** – Click **Open Template Folder** to jump to
   `assets/anchor_templates/` and drop in one or more cropped screenshots (PNG/JPG/BMP).
   Each template should be a tight crop of the anchor icon without extra background. Use
   the Windows Snipping Tool or your favourite editor to grab these still frames from the
   game. You can organise them in subfolders—everything in this directory is loaded.
4. **Reload templates** – After adding files, click **Reload Templates**. The status bar
   will confirm how many templates were loaded. If none load, check the file extension and
   make sure the images are not empty.
5. **Calibrate offsets** – With the game open, click **Realign Now**. When the anchor is
   matched you’ll see a status message showing the template name and match score. Adjust
   the *Offset X* and *Offset Y* sliders until the red capture box snaps directly over the
   deposit code after a realign.
6. **Fine-tune the threshold** – If matching is inconsistent, tweak the detection
   threshold (default `0.82`). Lower values make matching easier but risk false positives;
   higher values require a closer match.
7. **Control the refresh rate** – The **Alignment interval (ms)** control adjusts how
   frequently the app re-detects the anchor and nudges the capture region. Lower values
   (e.g. 200–300 ms) hug the HUD more tightly, while higher values reduce GPU/CPU usage.

Tips:

- You can store multiple templates for different lighting conditions or ship displays.
- Auto alignment runs before every scan when **Enable auto alignment** is checked. Turn it
  off temporarily if you want to position the capture region manually.
- Both the capture and anchor overlays continuously lift themselves, so they stay visible
  even while Star Citizen is in focus or fullscreen.

### Scanning deposits

**Option 1 - Hotkeys (Easy):**
- Press **"7"** to scan once
- Press **"Ctrl+7"** to start auto-scanning using the interval set in the
  **Continuous capture interval (s)** control (defaults to 2.0 seconds)
- Press **"8"** to hide/show the red box

**Option 2 - Buttons (If hotkeys don't work):**
- Click **"Single Scan"** to scan once
- Click **"Loop Toggle"** to start/stop auto-scanning (uses the same adjustable interval)

### Reading the results

- The tool shows **deposit type and quantity** in the floating overlay near the top of the
  screen. Use the **Display offset X/Y** sliders in the **Result Display** panel if you need
  to nudge that overlay around.
- Only recognised deposits from the bundled database are rendered; random OCR numbers are
  ignored so the overlay stays clean until a valid match is found.
- **Green text** = High-value materials (like Quantanium, Gold)
- **Yellow text** = Medium-value materials
- **Orange text** = Lower-value materials

### View the overlay in a browser or on your phone

The scanner also hosts a tiny web server so you can check the latest scan results from any
device on the same network (for example, a tablet next to your rig):

1. Start the scanner as usual. When it launches, look in the console/log window for a
   message similar to: `Starting overlay server: http://127.0.0.1:5000 (this device) |
   http://192.168.x.x:5000 (local network)`.
2. On another device connected to the **same Wi‑Fi or LAN**, open a browser and enter the
   `http://192.168.x.x:5000` address that matches your PC's local network IP.
3. The overlay page will refresh automatically with each new scan, mirroring what the in-game
   overlay shows.

> 🔐 **Firewall tip:** If the page will not load from another device, allow Python (or port
> 5000) through your operating system firewall.

## 🛠️ If something goes wrong

### "Python not found" or similar errors
- **Solution:** Just run `launch_windows.bat` again - it will download Python automatically

### "Ollama not installed" message
- **Solution:** Go to https://ollama.com/ and download/install Ollama, then restart the scanner

### Hotkeys don't work
- **Solution:** Use the buttons in the scanner window instead - they do the same thing

### The red box doesn't appear
- **Solution:** Click "Update Overlay" button in the scanner window

### Can't see the deposit code numbers clearly
- **Solution:** Adjust your Star Citizen graphics settings to make text clearer, or make the red box bigger

## 🎯 Tips for best results

1. **Make sure Star Citizen text is clear** - adjust graphics settings if text looks blurry
2. **Position the red box precisely** over just the deposit code numbers
3. **Don't make the red box too big** - it works better when focused on just the code
4. **Wait for the code to fully appear** before scanning (don't scan while the code is still appearing)

## 🔧 For Advanced Users

If you want to set things up manually instead of using the automatic installer:

### Manual Installation
1. Install Python 3.8+ from python.org
2. Install Ollama from ollama.com  
3. Open terminal/command prompt in the tool folder
4. Run: `python -m venv venv`
5. Activate venv: `venv\Scripts\activate` (Windows) or `source venv/bin/activate` (Linux)
6. Run: `pip install -r requirements.txt`
7. Run: `python scan_deposits.py`

## 🆘 Still need help?

1. **Make sure Ollama is installed** from https://ollama.com/
2. **Try restarting your computer** after installing Ollama
3. **Run the launch script again** - it will re-download anything that's missing
4. **Check that Star Citizen is running** and you're in a mining area where deposit codes appear

## 🐛 Found a bug or need help?

- **Report issues:** [GitHub Issues](https://github.com/FrozenButton/Scanning-Tool/issues)
- **Request features:** [GitHub Issues](https://github.com/FrozenButton/Scanning-Tool/issues)
- **Get updates:** [Watch this repository](https://github.com/FrozenButton/Scanning-Tool) for notifications

**When reporting issues, please include:**
- Your operating system (Windows 10, Windows 11, Linux, etc.)
- What you were trying to do
- Any error messages you saw
- Screenshots if helpful

## ⭐ Like this tool?

If this tool helps your mining operations, please:
- ⭐ **Star this repository** on GitHub
- 🍴 **Share it** with other Star Citizen miners
- 🐛 **Report bugs** to help make it better
- 🤝 **Contribute code** via pull requests

---

**Happy mining, citizen! 🪨⛏️**

*Repository: https://github.com/FrozenButton/Scanning-Tool* 
