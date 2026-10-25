# Raven Desktop Agent

A small Windows tray app that feeds Clerical Guru (Raven 2.0). It runs quietly next to the clock and:

- **Tracks time** – notes which app and window is in front every 5 seconds and groups it into time blocks. Idle time (no keyboard or mouse for 5 minutes, except during meetings), the lock screen, private browser windows and anything on your ignore list are never recorded.
- **Reads Outlook** – every 5 minutes it reads new mail (Inbox and Sent) and the next 14 days of calendars from **every mailbox loaded in classic Outlook**, including mailboxes on clients' systems. It only reads while Outlook is open and never starts Outlook itself.
- **Sends it to Clerical Guru** – about once a minute, to the `agentIngest` function. If the internet or Clerical Guru is unavailable, everything waits safely on disk and is sent later.

Nothing is sent anywhere except your own Clerical Guru app.

## Install

1. Go to the **Actions** tab of this repository, open the latest green **Build** run, and download **RavenAgent** under *Artifacts* (or download `RavenAgent.exe` from *Releases* once one is published).
2. Unzip it and move `RavenAgent.exe` somewhere permanent, e.g. `C:\Users\<you>\RavenAgent\`.
3. Double-click it. Windows may warn *"Windows protected your PC"* because the app isn't code-signed yet: click **More info → Run anyway**.
4. The Settings window opens. Fill in:
   - **Ingest URL** – the address of the `agentIngest` function in Base44 (Dashboard → Code → Functions → agentIngest; it looks like `https://<your-app>.base44.app/api/functions/agentIngest`).
   - **Agent key** – the same value as the `AGENT_INGEST_KEY` secret in Base44.
5. Click **Test connection**. When it says *Connected*, click **Save**.
6. Right-click the purple **R** by the clock and tick **Start with Windows**.

## Using it

Right-click the **R** icon:

| Menu item | What it does |
| --- | --- |
| Status lines | Whether it's recording, the last sync time, and what Outlook reading found |
| Pause recording (private mode) | Stops all recording until you untick it; the icon turns grey |
| Sync now | Reads Outlook and uploads immediately |
| Settings... | Ingest URL, key, Outlook on/off, ignore lists |
| Open log folder | Daily logs, kept for 7 days |
| Start with Windows | Launch automatically when you sign in |

**Outlook security prompt:** if Outlook ever shows *"A program is trying to access e-mail address information"*, your antivirus status is out of date in Windows Security. Updating it removes the prompt; you can also click **Allow**.

## Where things are stored

| What | Where |
| --- | --- |
| Settings (agent key encrypted for your Windows account) | `%APPDATA%\RavenAgent\config.json` |
| Upload queue | `%LOCALAPPDATA%\RavenAgent\queue.json` |
| Logs | `%LOCALAPPDATA%\RavenAgent\logs\` |

## For developers

- `src/RavenAgent.Core` – platform-independent logic (time segments, classification, Outlook filters, upload queue), covered by tests in `tests/`.
- `src/RavenAgent` – the Windows app (tray, window sampling, Outlook COM reader, uploader).
- Every push to `main` builds and tests on GitHub's Windows runners and publishes `RavenAgent.exe` as an artifact. Pushing a tag like `v0.1.0` also creates a release.
