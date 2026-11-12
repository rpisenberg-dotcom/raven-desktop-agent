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
    readonly System.Windows.Forms.Timer _outlookRetry = new();
    readonly System.Windows.Forms.Timer _changesTimer = new() { Interval = 90_000 };
    readonly List<CalendarChangeResult> _changeResults = new();
    bool _applyingChanges;
    DateTime _lastOutlookRead = DateTime.Now, _lastFullCalendar = DateTime.MinValue;
    readonly ToolStripMenuItem _statusItem = new() { Enabled = false };
    readonly ToolStripMenuItem _outlookItem = new() { Enabled = false };
    readonly ToolStripMenuItem _pauseItem = new("Pause recording (private mode)") { CheckOnClick = true };
    readonly ToolStripMenuItem _startupItem = new("Start with Windows") { CheckOnClick = true };

    OutlookReader.Result? _pendingCalendar;
    bool _pendingIsFull;
    readonly object _calLock = new();
    bool _uploading, _readingOutlook;
    DateTime? _lastUploadOk;
    string _lastCalSig = "";
    List<string>? _mailboxes;
    readonly List<EmailIdRecord> _idBackfill = new();
    readonly List<NoticeAck> _acks = new();
    readonly HashSet<string> _noticesShown = new();
    readonly Queue<RavenNotice> _noticeQueue = new();
    NoticeWindow? _notice;
    readonly NoteQueue _notes = new();
    HotkeyWindow? _hotkey, _snapKey;
    readonly List<MeetingSnapRecord> _snaps = new();
    readonly CaptionReader _captions = new();
    readonly System.Threading.Timer _captionTimer;
    int _captionBusy;
    QuickNoteWindow? _noteWindow;
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
        menu.Items.Add("Jot a note (Ctrl+Alt+N)", null, (_, _) => ShowQuickNote());
        menu.Items.Add("Screenshot for this meeting (Ctrl+Alt+S)", null, (_, _) => TakeSnap());
        menu.Items.Add("Sync now", null, async (_, _) => { await ReadOutlookAsync(); await UploadAsync(); });
        menu.Items.Add("Settings...", null, (_, _) => ShowSettings());
        menu.Items.Add("Open log folder", null, (_, _) => OpenFolder(Log.Folder));
        menu.Items.Add("Capture an app's layout for support (in 5 seconds)", null, (_, _) => CaptureLayoutSoon());
        menu.Items.Add("Check the calendar for a meeting...", null, async (_, _) => await CheckCalendarAsync());
        menu.Items.Add(_startupItem);
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add("Exit", null, (_, _) => Shutdown());

        _hotkey = new HotkeyWindow(ShowQuickNote);
        _snapKey = new HotkeyWindow(TakeSnap, Keys.S);
        _tray = new NotifyIcon { Icon = MakeIcon(false), ContextMenuStrip = menu, Visible = true, Text = "Raven" };
        _tray.DoubleClick += (_, _) => ShowSettings();

        _sampleTimer.Interval = _cfg.SampleSeconds * 1000;
        _sampleTimer.Tick += (_, _) => { _sampler.Tick(); RefreshStatus(); };
        _uploadTimer.Interval = _cfg.UploadSeconds * 1000;
        _uploadTimer.Tick += async (_, _) => await UploadAsync();
        _outlookTimer.Interval = _cfg.OutlookMinutes * 60 * 1000;
        _outlookTimer.Tick += async (_, _) => await ReadOutlookAsync();
        _outlookRetry.Tick += async (_, _) => { _outlookRetry.Stop(); await ReadOutlookAsync(); };
        // Calendar changes made in the app (new / edited / deleted events): checked every 90 seconds
        _changesTimer.Tick += async (_, _) => { await SendSnapsAsync(); await SendCaptionsAsync(); await PollCalendarChangesAsync(); };
        _changesTimer.Start();
        // Teams captions: a look every 8 seconds, only while a call is using the microphone (or a captured meeting is ending)
        _captionTimer = new System.Threading.Timer(_ =>
        {
            if (Interlocked.Exchange(ref _captionBusy, 1) == 1) return;
            try
            {
                if (_cfg.PrivateMode || !_cfg.CaptureCaptions) return;
                if (CallDetector.ActiveCall() != null || _captions.HasSession) _captions.Tick();
            }
            catch (Exception ex) { Log.Error("Captions", ex); }
            finally { Interlocked.Exchange(ref _captionBusy, 0); }
        }, null, 15_000, 8_000);

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

    async Task ReadOutlookAsync(bool force = false)
    {
        if (_readingOutlook || _cfg.PrivateMode || !_cfg.ReadOutlook || !_cfg.IsConfigured) return;
        // Don't read while Raven is working in Outlook (it would make Outlook stutter): wait, up to 15 minutes
        var (_, front, _) = Native.ForegroundWindow();
        if (!force && front.Equals("OUTLOOK", StringComparison.OrdinalIgnoreCase) && Native.IdleTime() < TimeSpan.FromSeconds(30)
            && DateTime.Now - _lastOutlookRead < TimeSpan.FromMinutes(15))
        {
            _outlookRetry.Interval = 60_000;
            _outlookRetry.Start();
            return;
        }
        _readingOutlook = true;
        try
        {
            var since = _cfg.LastMailSyncUtc ?? DateTime.UtcNow.AddDays(-_cfg.FirstRunEmailDays);
            // The full calendar (months ahead) once an hour; in between only the next few days
            var full = DateTime.Now - _lastFullCalendar > TimeSpan.FromHours(1);
            var result = await OutlookReader.RunOnSta(() => _outlook.Read(since, full));
            _lastOutlookRead = DateTime.Now;
            if (result != null && full) _lastFullCalendar = DateTime.Now;
            if (result != null)
            {
                if (result.Emails.Count > 0) _queue.Add(result.Emails);
                _mailboxes = result.Mailboxes;
                // Only send the calendar when something changed (or every few hours as a safety net)
                var sig = System.Text.Json.JsonSerializer.Serialize(new { result.Events, Ids = result.Syncs.Select(x => x.ExternalEventIds) }, Json.Options);
                if (sig != _lastCalSig || DateTime.Now - _lastCalQueued > TimeSpan.FromHours(4))
                {
                    // A full read still waiting to be sent isn't replaced by a short one (it would lose the months ahead)
                    lock (_calLock) if (full || !_pendingIsFull || _pendingCalendar == null) { _pendingCalendar = result; _pendingIsFull = full; }
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
            var acks = _acks.ToList();
            var notes = _notes.Peek();
            if (logs.Count == 0 && emails.Count == 0 && cal == null && ids.Count == 0 && texts.Count == 0 && acks.Count == 0 && notes.Count == 0
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
                NoticeAcks = acks.Count > 0 ? acks : null,
                Notes = notes.Count > 0 ? notes : null,
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
                var hidden = IngestClient.ParseHiddenCalendars(res.Message);
                if (hidden != null && !hidden.SequenceEqual(_cfg.HiddenCalendars))
                {
                    _cfg.HiddenCalendars = hidden;
                    _cfg.Save();
                    Log.Info($"Not reading hidden calendars: {string.Join(", ", hidden)}");
                }
                _queue.Remove(logs, emails);
                _acks.RemoveAll(a => acks.Contains(a));
                if (notes.Count > 0) { _notes.Remove(notes); _tray.ShowBalloonTip(2500, "Raven", notes.Count == 1 ? "Note saved to Clerical Guru" : $"{notes.Count} notes saved to Clerical Guru", ToolTipIcon.None); }
                _ = ApplyCalendarChangesAsync(IngestClient.ParseCalendarChanges(res.Message));
                foreach (var n in IngestClient.ParseNotices(res.Message))
                    if (_noticesShown.Add(n.TaskId + n.Kind)) _noticeQueue.Enqueue(n);
                ShowNextNotice();
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

    /// <summary>Ctrl+Alt+N: jot a question, to-do or note, with the client on screen picked already.</summary>
    void ShowQuickNote()
    {
        if (_noteWindow != null) { _noteWindow.Activate(); return; }
        // The client seen on screen in the last 10 minutes (QuickBooks, an email, a call...), otherwise none
        var recent = _sampler.LastMatch != null && DateTime.UtcNow - _sampler.LastMatchUtc < TimeSpan.FromMinutes(10) ? _sampler.LastMatch.ClientId : null;
        _noteWindow = new QuickNoteWindow(_clients.All, recent, _sampler.LastContext, n =>
        {
            _notes.Add(n);
            Log.Info($"Note jotted ({n.Kind})");
            _ = UploadAsync();
        });
        _noteWindow.FormClosed += (_, _) => _noteWindow = null;
        _noteWindow.Show();
        _noteWindow.Activate();
    }

    /// <summary>One pop-up at a time, and never while Raven is in a call (she may be sharing her screen).</summary>
    void ShowNextNotice()
    {
        if (_notice != null || _noticeQueue.Count == 0 || _cfg.PrivateMode) return;
        if (CallDetector.ActiveCall() != null) { _noticesShown.Clear(); _noticeQueue.Clear(); return; } // asked again after the call
        var n = _noticeQueue.Dequeue();
        _acks.Add(new NoticeAck { TaskId = n.TaskId, Action = n.Kind == "suggest" ? "shown_suggest" : "shown_done" });
        _notice = new NoticeWindow(n, action =>
        {
            if (action is "done" or "not_yet" or "undo") _acks.Add(new NoticeAck { TaskId = n.TaskId, Action = action });
            _notice = null;
            if (action.Length > 0) _ = UploadAsync(); // tell the app right away
            ShowNextNotice();
        });
        _notice.Show();
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

    /// <summary>Ctrl+Alt+S: save what's on screen to the meeting going on (sent right away; retried with the next check).</summary>
    void TakeSnap()
    {
        if (_cfg.PrivateMode) { _tray.ShowBalloonTip(3000, "Raven", "Private mode is on: no screenshot taken.", ToolTipIcon.Info); return; }
        var recent = _sampler.LastMatch != null && DateTime.UtcNow - _sampler.LastMatchUtc < TimeSpan.FromMinutes(10) ? _sampler.LastMatch.ClientId : null;
        var snap = MeetingSnap.Take(recent);
        if (snap == null) { _tray.ShowBalloonTip(3000, "Raven", "Couldn't take the screenshot.", ToolTipIcon.Warning); return; }
        lock (_snaps) _snaps.Add(snap);
        Log.Info("Screenshot taken for the meeting");
        _tray.ShowBalloonTip(2000, "Raven", "Screenshot saved for this meeting.", ToolTipIcon.Info);
        _ = SendSnapsAsync();
    }

    /// <summary>New caption lines go to Clerical Guru every 90 seconds, so the transcript is there when the meeting ends.</summary>
    async Task SendCaptionsAsync()
    {
        var batch = _captions.TakeUnsent().Where(r => r.Lines.Count > 0 || r.Ended).ToList();
        if (batch.Count == 0 || !_cfg.IsConfigured) return;
        var res = await _client.SendAsync(_cfg.IngestUrl, _cfg.AgentKey, new IngestPayload
        {
            DeviceId = _cfg.DeviceId, AgentVersion = Version, PollOnly = true, LiveTranscripts = batch,
        });
        if (res.Ok) _captions.MarkSent(batch);
    }

    async Task SendSnapsAsync()
    {
        List<MeetingSnapRecord> batch;
        lock (_snaps) batch = _snaps.Take(3).ToList();
        if (batch.Count == 0 || !_cfg.IsConfigured) return;
        var res = await _client.SendAsync(_cfg.IngestUrl, _cfg.AgentKey, new IngestPayload
        {
            DeviceId = _cfg.DeviceId, AgentVersion = Version, PollOnly = true, MeetingSnaps = batch,
        });
        if (!res.Ok) { Log.Info("Screenshot not sent yet; will retry"); return; }
        lock (_snaps) _snaps.RemoveAll(s => batch.Contains(s));
        if (_cfg.ReadOutlook) await ApplyCalendarChangesAsync(IngestClient.ParseCalendarChanges(res.Message));
        lock (_snaps) if (_snaps.Count > 0) _ = SendSnapsAsync();
    }

    /// <summary>Asks Clerical Guru for calendar changes made in the app, and reports what was done with earlier ones.</summary>
    async Task PollCalendarChangesAsync()
    {
        if (!_cfg.IsConfigured || _cfg.PrivateMode || !_cfg.ReadOutlook || _applyingChanges) return;
        List<CalendarChangeResult> report;
        lock (_changeResults) report = _changeResults.ToList();
        var res = await _client.SendAsync(_cfg.IngestUrl, _cfg.AgentKey, new IngestPayload
        {
            DeviceId = _cfg.DeviceId, AgentVersion = Version, PollOnly = true, CalendarChangeResults = report.Count > 0 ? report : null,
        });
        if (!res.Ok) return;
        lock (_changeResults) _changeResults.RemoveAll(r => report.Contains(r));
        await ApplyCalendarChangesAsync(IngestClient.ParseCalendarChanges(res.Message));
    }

    async Task ApplyCalendarChangesAsync(List<CalendarChange> changes)
    {
        if (changes.Count == 0 || _applyingChanges) return;
        if (!OutlookReader.IsOutlookRunning()) return; // they stay queued and come back in 10 minutes
        _applyingChanges = true;
        try
        {
            var results = await OutlookReader.RunOnSta(() => _outlook.ApplyChanges(changes));
            lock (_changeResults) _changeResults.AddRange(results);
            foreach (var r in results) Log.Info($"Calendar change: {r.Status} - {r.Message}");
            var failed = results.Count(r => r.Status == "failed");
            if (failed > 0) _tray.ShowBalloonTip(4000, "Raven", $"{failed} calendar change{(failed == 1 ? "" : "s")} couldn't be made in Outlook - see the Calendar page", ToolTipIcon.Warning);
        }
        catch (Exception ex) { Log.Error("Applying calendar changes failed", ex); }
        finally { _applyingChanges = false; }
        // Report back now, and read the calendar so the app shows the change in a minute or two
        _ = PollCalendarChangesAsync();
        _ = Task.Delay(5000).ContinueWith(_ => _tray.ContextMenuStrip?.BeginInvoke(new Action(async () => { await ReadOutlookAsync(force: true); await UploadAsync(); })));
    }

    /// <summary>Support: which calendars Raven sees and whether a meeting is found, saved as a text file to send.</summary>
    async Task CheckCalendarAsync()
    {
        var word = Prompt("Check the calendar", "A word from the meeting's title (e.g. BBV):", "BBV");
        if (string.IsNullOrWhiteSpace(word)) return;
        _tray.ShowBalloonTip(3000, "Raven", "Checking your calendars - this takes a moment", ToolTipIcon.None);
        try
        {
            var report = await OutlookReader.RunOnSta(() => _outlook.CheckCalendars(word.Trim()));
            Directory.CreateDirectory(Log.Folder);
            var file = Path.Combine(Log.Folder, $"calendar-check-{DateTime.Now:yyyyMMdd-HHmm}.txt");
            File.WriteAllText(file, report);
            System.Diagnostics.Process.Start("notepad.exe", $"\"{file}\"");
        }
        catch (Exception ex) { Log.Error("Calendar check failed", ex); MessageBox.Show($"The check didn't work: {ex.Message}", "Raven"); }
    }

    static string? Prompt(string title, string label, string initial)
    {
        using var f = new Form { Text = title, FormBorderStyle = FormBorderStyle.FixedDialog, StartPosition = FormStartPosition.CenterScreen, ClientSize = new Size(360, 120),
            MinimizeBox = false, MaximizeBox = false, TopMost = true, Font = new Font("Segoe UI", 10f) };
        var l = new Label { Text = label, AutoSize = true, Location = new Point(12, 12) };
        var t = new TextBox { Text = initial, Location = new Point(12, 40), Width = 336 };
        var ok = new Button { Text = "Check", DialogResult = DialogResult.OK, Location = new Point(192, 78), Width = 75 };
        var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Location = new Point(273, 78), Width = 75 };
        f.Controls.AddRange(new Control[] { l, t, ok, cancel });
        f.AcceptButton = ok; f.CancelButton = cancel;
        return f.ShowDialog() == DialogResult.OK ? t.Text : null;
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
        _hotkey?.Dispose();
        _snapKey?.Dispose();
        _captionTimer.Dispose();
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
