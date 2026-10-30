using System.Drawing.Drawing2D;
using System.Reflection;
using Microsoft.Win32;
using RavenAgent.Core;

namespace RavenAgent;

/// <summary>The agent itself: a tray icon that samples activity, reads Outlook and uploads on timers.</summary>
sealed class TrayApp : ApplicationContext
{
    public static string Version =>
        Assembly.GetExecutingAssembly().GetName().Version?.ToString(3) ?? "0.1.0";

    readonly AgentConfig _cfg;
    readonly UploadQueue _queue;
    readonly ActivitySampler _sampler;
    readonly OutlookReader _outlook;
    readonly IngestClient _client = new();
    readonly ClientDirectory _clients = new();
    readonly ScreenContext _screen;
    readonly NotifyIcon _tray;
    readonly System.Windows.Forms.Timer _sampleTimer = new();
    readonly System.Windows.Forms.Timer _uploadTimer = new();
    readonly System.Windows.Forms.Timer _outlookTimer = new();
    readonly ToolStripMenuItem _statusItem = new() { Enabled = false };
    readonly ToolStripMenuItem _outlookItem = new() { Enabled = false };
    readonly ToolStripMenuItem _pauseItem = new("Pause recording (private mode)") { CheckOnClick = true };
    readonly ToolStripMenuItem _startupItem = new("Start with Windows") { CheckOnClick = true };

    OutlookReader.Result? _pendingCalendar;
    readonly object _calLock = new();
    bool _uploading, _readingOutlook;
    DateTime? _lastUploadOk;
    string _lastCalSig = "";
    List<string>? _mailboxes;
    readonly List<EmailIdRecord> _idBackfill = new();
    DateTime _lastCalQueued = DateTime.MinValue;
    string _lastError = "";

    public TrayApp()
    {
        _cfg = AgentConfig.Load();
        Log.Cleanup();
        Log.Info($"Agent {Version} starting on {_cfg.DeviceId}");
        _queue = new UploadQueue(Path.Combine(AgentConfig.DataFolder, "queue.json"));
        _screen = new ScreenContext(_clients);
        _sampler = new ActivitySampler(_cfg, _queue, _clients, _screen, new OutlookContext(_clients));
        _outlook = new OutlookReader(_cfg);

        _pauseItem.Checked = _cfg.PrivateMode;
        _pauseItem.CheckedChanged += (_, _) => SetPrivate(_pauseItem.Checked);
        Startup.RegisterLinkHandler();
        _startupItem.Checked = Startup.IsEnabled();
        _startupItem.CheckedChanged += (_, _) => Startup.Set(_startupItem.Checked);

        var menu = new ContextMenuStrip();
        menu.Items.Add(new ToolStripMenuItem($"Raven Desktop Agent {Version}") { Enabled = false, Font = new Font(menu.Font, FontStyle.Bold) });
        menu.Items.Add(_statusItem);
        menu.Items.Add(_outlookItem);
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add(_pauseItem);
        menu.Items.Add("Sync now", null, async (_, _) => { await ReadOutlookAsync(); await UploadAsync(); });
        menu.Items.Add("Settings...", null, (_, _) => ShowSettings());
        menu.Items.Add("Open log folder", null, (_, _) => OpenFolder(Log.Folder));
        menu.Items.Add("Capture an app's layout for support (in 5 seconds)", null, (_, _) => CaptureLayoutSoon());
        menu.Items.Add(_startupItem);
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add("Exit", null, (_, _) => Shutdown());

        _tray = new NotifyIcon { Icon = MakeIcon(false), ContextMenuStrip = menu, Visible = true, Text = "Raven" };
        _tray.DoubleClick += (_, _) => ShowSettings();

        _sampleTimer.Interval = _cfg.SampleSeconds * 1000;
        _sampleTimer.Tick += (_, _) => { _sampler.Tick(); RefreshStatus(); };
        _uploadTimer.Interval = _cfg.UploadSeconds * 1000;
        _uploadTimer.Tick += async (_, _) => await UploadAsync();
        _outlookTimer.Interval = _cfg.OutlookMinutes * 60 * 1000;
        _outlookTimer.Tick += async (_, _) => await ReadOutlookAsync();

        // Close the open segment when the PC sleeps or locks, so idle time isn't counted
        SystemEvents.PowerModeChanged += (_, e) => { if (e.Mode == PowerModes.Suspend) _sampler.Flush(); };
        SystemEvents.SessionSwitch += (_, e) => { if (e.Reason == SessionSwitchReason.SessionLock) _sampler.Flush(); };

        _sampleTimer.Start();
        _uploadTimer.Start();
        _outlookTimer.Start();

        if (!_cfg.IsConfigured)
        {
            _tray.ShowBalloonTip(5000, "Raven Desktop Agent", "Welcome! Enter your Clerical Guru address and agent key to start sending data.", ToolTipIcon.Info);
            ShowSettings();
        }
        else
        {
            // First Outlook read shortly after start, once Windows has settled
            var once = new System.Windows.Forms.Timer { Interval = 20000 };
            once.Tick += async (_, _) => { once.Stop(); once.Dispose(); await ReadOutlookAsync(); await UploadAsync(); };
            once.Start();
        }
        RefreshStatus();
    }

    void SetPrivate(bool on)
    {
        _cfg.PrivateMode = on;
        _cfg.Save();
        if (on) _sampler.Flush();
        _tray.Icon = MakeIcon(on);
        Log.Info(on ? "Private mode on" : "Private mode off");
        RefreshStatus();
    }

    async Task ReadOutlookAsync()
    {
        if (_readingOutlook || _cfg.PrivateMode || !_cfg.ReadOutlook || !_cfg.IsConfigured) return;
        _readingOutlook = true;
        try
        {
            var since = _cfg.LastMailSyncUtc ?? DateTime.UtcNow.AddDays(-_cfg.FirstRunEmailDays);
            var result = await OutlookReader.RunOnSta(() => _outlook.Read(since));
            if (result != null)
            {
                if (result.Emails.Count > 0) _queue.Add(result.Emails);
                _mailboxes = result.Mailboxes;
                // Only send the calendar when something changed (or every few hours as a safety net)
                var sig = System.Text.Json.JsonSerializer.Serialize(new { result.Events, Ids = result.Syncs.Select(x => x.ExternalEventIds) }, Json.Options);
                if (sig != _lastCalSig || DateTime.Now - _lastCalQueued > TimeSpan.FromHours(4))
                {
                    lock (_calLock) _pendingCalendar = result;
                    _lastCalSig = sig;
                    _lastCalQueued = DateTime.Now;
                }
                // Small overlap so nothing slips between reads; the server ignores duplicates
                _cfg.LastMailSyncUtc = result.ReadAtUtc.AddMinutes(-2);
                _cfg.Save();
                Log.Info(_outlook.Status);

                // Once: send Outlook's ids for emails read by older versions, so "Open in Outlook" works for them
                if (!_cfg.OutlookIdsBackfilled)
                {
                    var back = await OutlookReader.RunOnSta(() => _outlook.Read(DateTime.UtcNow.AddDays(-14)));
                    if (back != null)
                    {
                        _idBackfill.AddRange(back.Emails.Where(e => !string.IsNullOrEmpty(e.OutlookEntryId)).Select(e => new EmailIdRecord
                        {
                            ExternalId = e.ExternalId, OutlookEntryId = e.OutlookEntryId!, OutlookStoreId = e.OutlookStoreId,
                        }));
                        _cfg.OutlookIdsBackfilled = true;
                        _cfg.Save();
                        Log.Info($"Sending Outlook ids for {_idBackfill.Count} earlier emails");
                    }
                }
            }
        }
        catch (Exception ex) { Log.Error("Outlook cycle failed", ex); }
        finally { _readingOutlook = false; RefreshStatus(); }
    }

    async Task UploadAsync()
    {
        if (_uploading || !_cfg.IsConfigured) return;
        _uploading = true;
        List<PhoneSnapshotRecord> texts = new();
        try
        {
            var (logs, emails) = _queue.Peek(400);
            OutlookReader.Result? cal;
            lock (_calLock) cal = _pendingCalendar;
            // Nothing new: skip the call (each one costs Base44 credits), but check in every 30 minutes
            // so client and mailbox settings still reach this PC
            var ids = _idBackfill.Take(300).ToList();
            texts = _screen.TakePhoneSnapshots();
            if (logs.Count == 0 && emails.Count == 0 && cal == null && ids.Count == 0 && texts.Count == 0
                && _lastUploadOk is { } last && DateTime.Now - last < TimeSpan.FromMinutes(30)) return;

            var payload = new IngestPayload
            {
                DeviceId = _cfg.DeviceId,
                AgentVersion = Version,
                PrivateMode = _cfg.PrivateMode,
                TimeLogs = logs,
                Emails = emails,
                CalendarEvents = cal?.Events ?? new(),
                CalendarSyncs = cal?.Syncs ?? new(),
                Mailboxes = _mailboxes,
                EmailIds = ids.Count > 0 ? ids : null,
                PhoneMessages = texts.Count > 0 ? texts : null,
            };
            var res = await _client.SendAsync(_cfg.IngestUrl, _cfg.AgentKey, payload);
            if (res.Ok)
            {
                var directory = IngestClient.ParseClientDirectory(res.Message);
                if (directory != null) _clients.Update(directory);
                var business = IngestClient.ParseBusinessContacts(res.Message);
                if (business != null) _clients.UpdateBusinessContacts(business);
                var ignored = IngestClient.ParseIgnoredMailboxes(res.Message);
                if (ignored != null && !ignored.SequenceEqual(_cfg.IgnoredMailboxes))
                {
                    _cfg.IgnoredMailboxes = ignored;
                    _cfg.Save();
                    Log.Info($"Not reading mail from: {string.Join(", ", ignored)}");
                }
                _queue.Remove(logs, emails);
                _idBackfill.RemoveRange(0, Math.Min(ids.Count, _idBackfill.Count));
                lock (_calLock) if (ReferenceEquals(_pendingCalendar, cal)) _pendingCalendar = null;
                _lastUploadOk = DateTime.Now;
                if (_lastError != "") Log.Info("Uploads working again");
                _lastError = "";
            }
            else
            {
                _screen.ReturnPhoneSnapshots(texts);
                if (res.Message != _lastError) Log.Error($"Upload failed: {res.Message}");
                _lastError = res.Message;
            }
        }
        catch (Exception ex) { _screen.ReturnPhoneSnapshots(texts); _lastError = ex.Message; Log.Error("Upload failed", ex); }
        finally { _uploading = false; RefreshStatus(); }
    }

    void RefreshStatus()
    {
        string state = !_cfg.IsConfigured ? "Not set up - open Settings"
            : _cfg.PrivateMode ? "Paused (private mode)"
            : _lastError != "" ? $"Can't upload - {_lastError}"
            : _lastUploadOk is { } t ? $"Recording - last sync {t:t}"
            : "Recording";
        var pending = _queue.Count;
        _statusItem.Text = pending > 0 ? $"{state} ({pending} waiting)" : state;
        _outlookItem.Text = (_cfg.ReadOutlook ? $"Outlook: {_outlook.Status}" : "Outlook: off") + $" · {_clients.Count} clients known";
        var tip = $"Raven - {state}";
        _tray.Text = tip.Length > 63 ? tip[..63] : tip;
    }

    SettingsForm? _settings;
    void ShowSettings()
    {
        if (_settings != null) { _settings.Activate(); return; }
        var oldUrl = _cfg.IngestUrl;
        using (_settings = new SettingsForm(_cfg))
        {
            if (_settings.ShowDialog() == DialogResult.OK)
            {
                Log.Info("Settings saved");
                if (!string.Equals(oldUrl, _cfg.IngestUrl, StringComparison.OrdinalIgnoreCase))
                {
                    // Connected to a different Raven app: send it the full calendar, mailbox list and Outlook ids
                    Log.Info($"Now sending to {_cfg.IngestUrl}");
                    _lastCalSig = "";
                    _lastUploadOk = null;
                    _cfg.OutlookIdsBackfilled = false;
                    _cfg.Save();
                }
                _ = UploadAsync();
                _ = ReadOutlookAsync();
            }
        }
        _settings = null;
        RefreshStatus();
    }

    /// <summary>Gives Raven 5 seconds to click into the app, then saves that window's layout for support.</summary>
    void CaptureLayoutSoon()
    {
        _tray.ShowBalloonTip(4000, "Raven", "Click into the app you want Raven to read (e.g. Phone Link with a conversation open). Capturing in 5 seconds...", ToolTipIcon.Info);
        var t = new System.Windows.Forms.Timer { Interval = 5000 };
        t.Tick += (_, _) =>
        {
            t.Stop(); t.Dispose();
            var (hwnd, process, _) = Native.ForegroundWindow();
            var path = ScreenContext.DumpLayout(hwnd, string.IsNullOrEmpty(process) ? "app" : process, Log.Folder);
            Log.Info($"Saved window layout to {path}");
            _tray.ShowBalloonTip(5000, "Raven", "Layout saved. The folder is opening: send the newest 'layout' file to Claude.", ToolTipIcon.Info);
            OpenFolder(Log.Folder);
        };
        t.Start();
    }

    static void OpenFolder(string path)
    {
        Directory.CreateDirectory(path);
        System.Diagnostics.Process.Start("explorer.exe", $"\"{path}\"");
    }

    void Shutdown()
    {
        Log.Info("Agent exiting");
        _sampleTimer.Stop();
        _uploadTimer.Stop();
        _outlookTimer.Stop();
        _sampler.Flush(); // saved to the queue on disk; sent on next start
        _tray.Visible = false;
        _tray.Dispose();
        _client.Dispose();
        _screen.Dispose();
        ExitThread();
    }

    /// <summary>A purple "R" badge; grey when paused.</summary>
    static Icon MakeIcon(bool paused)
    {
        using var bmp = new Bitmap(32, 32);
        using (var g = Graphics.FromImage(bmp))
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
            using var fill = new SolidBrush(paused ? Color.FromArgb(140, 140, 150) : Color.FromArgb(88, 64, 200));
            g.FillEllipse(fill, 1, 1, 30, 30);
            using var font = new Font("Segoe UI", 15f, FontStyle.Bold, GraphicsUnit.Pixel);
            var size = g.MeasureString("R", font);
            g.DrawString("R", font, Brushes.White, (32 - size.Width) / 2 + 1, (32 - size.Height) / 2 + 1);
        }
        return Icon.FromHandle(bmp.GetHicon());
    }
}
