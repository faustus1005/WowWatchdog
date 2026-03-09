# Getting Started with WoW-Watchdog

This guide walks through installing WoW-Watchdog, configuring your server paths, and validating notifications so you can start monitoring quickly.

## Prerequisites

- **Windows 10/11** with **Windows PowerShell 5.1+**.
- **Administrator rights** if you install the Windows service.
- Optional: an **NTFY** server (or a public instance) if you want push notifications.

These requirements mirror the supported platforms and features listed in the project overview.

## Installation (Recommended)

1. Download the latest installer from the **GitHub Releases** page.
2. Run the installer (administrator rights are required to install the service).
3. Launch **WoW Watchdog** from the desktop shortcut after installation completes.

WoW-Watchdog runs as a Windows service under the name **WoWWatchdog**.

## Portable Mode (No Installer)

If you prefer not to install the service, you can run the app in portable mode:

- Create a file named `portable.flag` in the same directory as the executable, **or**
- Start the app with the `-Portable` switch.

In portable mode, WoW-Watchdog stores data alongside the executable:

- `data\config.json` for configuration
- `data\secrets.json` for encrypted notification credentials
- `logs\` for logs

> Tip: Portable mode skips the automatic elevation prompt. Use it when you want to keep everything self-contained.

## First-Run Configuration

Open the app and configure these core settings:

1. **Server Paths**
   - Use the **Browse** buttons in the GUI to point to your MySQL, Authserver, and Worldserver executables or scripts.
2. **Expansion Label (Optional)**
   - Select the expansion for notification labeling; it doesn’t change monitoring behavior.
3. **Database Settings (Optional)**
   - Configure DB host/port/user/name if you want live player counts in the UI.
4. **NTFY Notifications (Optional)**
   - Enter your NTFY server and topic.
   - Choose **None**, **Basic**, or **Token** auth. Credentials are stored in `secrets.json` and encrypted on save.
   - Use the **Test** button to validate delivery.

Make sure to click **Save Configuration** so your paths and notification settings persist.

## Where Configuration and Logs Live

**Installed mode (default):**

- `%ProgramData%\WoWWatchdog\config.json`
- `%ProgramData%\WoWWatchdog\secrets.json`
- `%ProgramData%\WoWWatchdog\watchdog.log`

**Portable mode:**

- `data\config.json`
- `data\secrets.json`
- `logs\watchdog.log`

## Troubleshooting Quick Tips

- If the service won’t start, confirm the paths to your MySQL/Auth/Worldserver executables are correct.
- If notifications fail, verify the NTFY server URL and topic, then run a **Test** notification.
- For unexpected behavior, check `watchdog.log` and `crash.log` in your logs directory.

## Next Steps

- Review the full feature list and configuration details in the main README.
