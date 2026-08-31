# Raspberry Pi Migration Guide

## Overview
Your SnipesBot is now ready to run on a Raspberry Pi Zero 2 W. The bot code is already platform-agnostic, so minimal setup is needed. This guide covers the migration steps and the photo persistence fix that has been implemented.

---

## Part 1: Migrating to Raspberry Pi Zero 2 W

### Prerequisites
- Raspberry Pi Zero 2 W with latest Raspberry Pi OS installed
- SSH access to the Pi (or physical access to run commands)
- Internet connection for the Pi

### Step 1: Install Python & System Dependencies

```bash
sudo apt update
sudo apt upgrade -y
sudo apt install python3 python3-pip python3-venv
```

### Step 2: Transfer Your Bot Files

**Option A: Via SSH (recommended)**
```bash
scp -r /home/peter/Coding/SnipesBot pi@<your-pi-ip>:/home/pi/
```

**Option B: Manual transfer**
- Use a USB drive or file transfer tool to copy the SnipesBot folder to `/home/pi/SnipesBot`
- Make sure your `private/` folder with your token is included

### Step 3: Set Up Virtual Environment

```bash
cd ~/SnipesBot
python3 -m venv .venv
source .venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt
```

### Step 4: Test the Bot

```bash
python3 snipes_bot.py
```

You should see:
```
Bot starting... Working Directory: /home/pi/SnipesBot
Logged in as [YourBotName]!
Synced X command(s).
```

Press `Ctrl+C` to stop once you've confirmed it works.

### Step 5: Set Up Auto-Start with Systemd

Create a new systemd service file:

```bash
sudo nano /etc/systemd/system/snipes-bot.service
```

Paste this content:

```ini
[Unit]
Description=SnipesBot Discord Bot
After=network.target

[Service]
Type=simple
User=pi
WorkingDirectory=/home/pi/SnipesBot
Environment="PATH=/home/pi/SnipesBot/.venv/bin"
ExecStart=/home/pi/SnipesBot/.venv/bin/python3 /home/pi/SnipesBot/snipes_bot.py
Restart=on-failure
RestartSec=10
StandardOutput=journal
StandardError=journal

[Install]
WantedBy=multi-user.target
```

Save with `Ctrl+X`, then `Y`, then `Enter`.

### Step 6: Enable and Start the Service

```bash
sudo systemctl daemon-reload
sudo systemctl enable snipes-bot
sudo systemctl start snipes-bot
```

### Step 7: Verify It's Running

```bash
sudo systemctl status snipes-bot
```

Check the logs anytime with:
```bash
sudo journalctl -u snipes-bot -f
```

---

## Part 2: Photo Persistence Issue - FIXED ✅

### What Was the Problem?

Your bot was storing Discord CDN URLs (like `https://cdn.discordapp.com/...`) in the Excel sheet. These URLs expire after **2-7 days** because Discord doesn't keep temporary attachments around forever. This is by design for security and storage efficiency.

### What's Fixed?

The bot now:
1. **Downloads the image** when you run `/snipe`
2. **Saves it locally** to a `proofs/` folder on your Pi
3. **Stores the file path** in Excel instead of the URL

### File Organization

```
SnipesBot/
├── proofs/                    # New folder - stores all proof images
│   ├── 20260831_143022_screenshot.png
│   ├── 20260831_153045_photo.jpg
│   └── ...
├── snipes_bot.py
├── SNIPESSTATS.xlsm
├── private/
│   ├── token
│   └── registrations.json
└── ...
```

### How It Works

When someone uses `/snipe`:
- Image is saved with a timestamp: `YYYYMMDD_HHMMSS_filename`
- Path is stored in Excel like: `/home/pi/SnipesBot/proofs/20260831_143022_screenshot.png`
- Discord message shows the filename instead of a URL

**Example:**
```
**@Alice got shot by @Bob for 2 points**
*Proof saved as: `20260831_143022_screenshot.png`*
```

### Storage Considerations

Each image takes up space on your Pi. For a Pi Zero 2 W with limited storage:

- **Monitor disk usage**: `df -h`
- **Check proofs folder size**: `du -sh ~/SnipesBot/proofs/`
- **Optional cleanup script** (run monthly):
```bash
# Keep only images from the last 90 days
find ~/SnipesBot/proofs/ -type f -mtime +90 -delete
```

---

## Troubleshooting

### Bot doesn't start on boot
```bash
sudo systemctl status snipes-bot
sudo journalctl -u snipes-bot -n 20
```

### "Module not found" error
Ensure the virtual environment is activated and all requirements installed:
```bash
cd ~/SnipesBot
source .venv/bin/activate
pip install -r requirements.txt
```

### Bot stops randomly
Check if Excel file is locked or if there's a permission issue:
```bash
sudo journalctl -u snipes-bot -f
```

### Images not saving
Verify the `proofs/` folder exists and is writable:
```bash
ls -la ~/SnipesBot/proofs/
```

---

## Windows to Pi Differences

| Aspect | Windows | Pi |
|--------|---------|-----|
| Path separator | `\` | `/` |
| Startup method | .exe or Task Scheduler | systemd service |
| Storage | Typically plenty | Limited (check `df -h`) |
| Resource usage | None, let it run | CPU/RAM minimal - good for always-on |
| Maintenance | Manual restart | Auto-restart on crash |

The code handles path differences automatically via `os.path.join()`, so no changes needed.

---

## Next Steps

1. **Migrate the bot** to your Pi using Steps 1-7 above
2. **Test a snipe** to confirm the new photo-saving workflow works
3. **Monitor the logs** for the first few days with `sudo journalctl -u snipes-bot -f`
4. **Set up optional log rotation** to prevent logs from growing too large:
   ```bash
   sudo nano /etc/logrotate.d/snipes-bot
   ```
   Add:
   ```
   /var/log/snipes-bot.log {
       daily
       missingok
       rotate 7
       compress
       delaycompress
   }
   ```

Enjoy your always-on Discord bot! 🚀
