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
    HotkeyWindow? _hotkey, _snapKey, _replyKey;
    bool _drafting;
    readonly List<MeetingSnapRecord> _snaps = new();
    readonly CaptionReader _captions = new();
    readonly System.Threading.Timer _captionTimer;
    int _captionBusy, _captionLooks;
    QuickNoteWindow? _noteWindow;
    SynchronizationContext? _ui;
    // Meeting recordings
    readonly System.Windows.Forms.Timer _recTimer = new() { Interval = 4000 };
    MeetingRecorder? _rec;
    List<MeetingRecorder.VideoSource> _recSources = new();
    RecordPrompt? _recPrompt;
    RecordingBar? _recBar;
    bool _inCall, _autoVideo, _userPickedVideo, _warnedCaptions, _loggedMic;
    DateTime? _callGoneSince;
    readonly ToolStripMenuItem _recStartItem = new("Record this meeting...");
    readonly ToolStripMenuItem _recStopItem = new("Stop recording") { Enabled = false };
    readonly ToolStripMenuItem _recOfferItem = new("Ask to record when a meeting starts") { CheckOnClick = true };
    readonly ToolStripMenuItem _recStatusItem = new() { Enabled = false, Visible = false };
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
        _ui = SynchronizationContext.Current; // the Windows Forms one, installed with the first control
        menu.Items.Add(new ToolStripMenuItem($"Raven Desktop Agent {Version}") { Enabled = false, Font = new Font(menu.Font, FontStyle.Bold) });
        menu.Items.Add(_statusItem);
        menu.Items.Add(_outlookItem);
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add(_pauseItem);
        menu.Items.Add("Jot a note (Ctrl+Alt+N)", null, (_, _) => ShowQuickNote());
        menu.Items.Add("Screenshot for this meeting (Ctrl+Alt+S)", null, (_, _) => TakeSnap());
        menu.Items.Add("Draft a reply or forward for the selected email (Ctrl+Alt+R)", null, (_, _) => _ = DraftReplyAsync());
        var recMenu = new ToolStripMenuItem("Meeting recording");
        _recStartItem.Click += (_, _) => ShowRecordPrompt(CallDetector.ActiveCall() ?? "Teams", manual: true);
        _recStopItem.Click += (_, _) => _ = StopRecordingAsync();
        _recOfferItem.Checked = _cfg.OfferRecording;
        _recOfferItem.CheckedChanged += (_, _) => { _cfg.OfferRecording = _recOfferItem.Checked; _cfg.Save(); };
        recMenu.DropDownItems.Add(_recStartItem);
        recMenu.DropDownItems.Add(_recStopItem);
        recMenu.DropDownItems.Add(_recStatusItem);
        recMenu.DropDownItems.Add(new ToolStripSeparator());
        recMenu.DropDownItems.Add(_recOfferItem);
        recMenu.DropDownItems.Add("Open the recordings folder", null, (_, _) => { Directory.CreateDirectory(_cfg.RecordingsFolder); OpenFolder(_cfg.RecordingsFolder); });
        recMenu.DropDownItems.Add("Change the recordings folder...", null, (_, _) => ChooseRecordingsFolder());
        menu.Items.Add(recMenu);
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
        _replyKey = new HotkeyWindow(() => _ = DraftReplyAsync(), Keys.R);
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
        _changesTimer.Tick += async (_, _) => { await SendSnapsAsync(); await SendCaptionsAsync(); await SendPendingRecordingsAsync(); await PollCalendarChangesAsync(); };
        _changesTimer.Start();
        // Teams captions: a look every 8 seconds, only while a call is using the microphone (or a captured meeting is ending)
        _captionTimer = new System.Threading.Timer(_ =>
        {
            if (Interlocked.Exchange(ref _captionBusy, 1) == 1) return;
            try
            {
                if (_cfg.PrivateMode || !_cfg.CaptureCaptions) return;
                // Windows doesn't always report a browser meeting using the microphone, so every other look also
                // checks for a Teams meeting window when no call is reported
                _captionLooks++;
                if (CallDetector.ActiveCall() != null || _captions.HasSession || _captions.Consented || _captionLooks % 2 == 0) _captions.Tick();
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
        _recTimer.Tick += (_, _) => RecordingTick();
        _recTimer.Start();
        // The free recording tools (about 400 MB) download once, a few minutes after start
        if (_cfg.OfferRecording && !MeetingTools.Ready)
            _ = Task.Run(async () => { await Task.Delay(TimeSpan.FromMinutes(3)); await MeetingTools.EnsureAsync(); });

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

    // ---------------- Meeting recordings ----------------

    /// <summary>Every 4 seconds: offers to record when a call starts, stops the recording when it ends.</summary>
    void RecordingTick()
    {
        try
        {
            if (_rec != null) { _recBar?.Tick(_rec.Elapsed); FollowScreenShare(); WarnIfCaptionsOff(); }
            var meetingOpen = DateTime.UtcNow - _captions.LastMeetingWindowUtc < TimeSpan.FromSeconds(40);
            var call = _cfg.PrivateMode ? null : CallDetector.ActiveCall() ?? (meetingOpen ? "Teams" : null);
            if (meetingOpen && !_loggedMic)
            {
                _loggedMic = true; // once per run: what Windows reports using the microphone, for support
                Log.Info("Meeting window open; microphone in use by: " + string.Join(", ", CallDetector.InUse().Select(x => $"{x.key} ({x.label ?? "?"})")));
            }
            if (call != null)
            {
                _callGoneSince = null;
                if (!_inCall)
                {
                    _inCall = true;
                    // Phone calls are never offered automatically (they can still be recorded from the menu)
                    if (_cfg.OfferRecording && _rec == null && call != "Phone Link") ShowRecordPrompt(call, manual: false);
                }
            }
            else if (_inCall)
            {
                _callGoneSince ??= DateTime.UtcNow;
                if (DateTime.UtcNow - _callGoneSince > TimeSpan.FromSeconds(45))
                {
                    _inCall = false;
                    _recPrompt?.Close();
                    if (_rec != null) { Log.Info("Call ended: stopping the recording"); _ = StopRecordingAsync(); }
                }
            }
            _recStartItem.Enabled = _rec == null;
            _recStopItem.Enabled = _rec != null;
        }
        catch (Exception ex) { Log.Error("Recording check", ex); }
    }

    void ShowRecordPrompt(string app, bool manual)
    {
        if (_rec != null || _cfg.PrivateMode) return;
        if (_recPrompt != null && !_recPrompt.IsDisposed) { _recPrompt.Activate(); return; }
        var meeting = MeetingWindows.Find(app);
        var sources = MeetingWindows.Sources(meeting);
        _recPrompt = new RecordPrompt(app, sources, _cfg.RecordVideo, _cfg.RecordingNotice, src => StartRecording(app, meeting, sources, src));
        _recPrompt.FormClosed += (_, _) => _recPrompt = null;
        _recPrompt.Show();
        if (manual) _recPrompt.Activate();
    }

    void StartRecording(string app, (IntPtr hwnd, string title)? meeting, List<MeetingRecorder.VideoSource> sources, MeetingRecorder.VideoSource video)
    {
        try
        {
            _cfg.RecordVideo = video.Kind;
            _cfg.Save();
            var title = MeetingWindows.TitleFor(meeting, app);
            Log.Info($"Recording started with notice confirmed: {title}");
            _rec = MeetingRecorder.Start(_cfg.RecordingsFolder, title, app, video);
            _recSources = sources;
            _autoVideo = false;
            _userPickedVideo = false;
            _warnedCaptions = false;
            _captions.Consented = true;
            if (!_rec.HasSound)
            {
                _tray.ShowBalloonTip(6000, "Raven", "Raven couldn't reach your microphone or speakers, so nothing is being recorded. See the log for details.", ToolTipIcon.Warning);
                _rec.Stop();
                _rec = null;
                _captions.Consented = false;
                return;
            }
            var rec = _rec;
            _recBar = new RecordingBar(sources, video, s =>
            {
                _userPickedVideo = true; // her choice: no more automatic switching for this recording
                _ = Task.Run(() =>
                {
                    rec.SwitchVideo(s);
                    OnUi(() => _recBar?.ShowSource(rec.Video));
                });
            }, () => _ = StopRecordingAsync());
            _recBar.Show();
            _tray.Text = "Raven - recording this meeting";
            if (video.Kind != "none" && !MeetingTools.HasFfmpeg)
                _tray.ShowBalloonTip(6000, "Raven", "Recording sound now. The video recorder is still downloading; pick the video again from the Raven bar in a few minutes.", ToolTipIcon.Info);
            if (video.Kind == "window" && meeting != null && MeetingWindows.IsMinimized(meeting.Value.hwnd))
                _tray.ShowBalloonTip(6000, "Raven", "The meeting window is minimized, so there's no video of it. Restore it (behind other windows is fine).", ToolTipIcon.Info);
            // If the chosen video can't be captured, the bar shows what's really being recorded
            _ = Task.Run(async () =>
            {
                await Task.Delay(6000);
                OnUi(() => { if (_rec == rec) _recBar?.ShowSource(rec.Video); });
            });
            RecordingTick();
        }
        catch (Exception ex)
        {
            Log.Error("Starting the recording", ex);
            _tray.ShowBalloonTip(6000, "Raven", $"Couldn't start recording: {ex.Message}", ToolTipIcon.Warning);
            _rec = null;
            _captions.Consented = false;
        }
    }

    async Task StopRecordingAsync()
    {
        var rec = _rec;
        if (rec == null) return;
        _rec = null;
        _captions.Consented = false;
        _recBar?.Close();
        _recBar = null;
        _tray.Text = "Raven";
        RecordingTick();
        var result = await Task.Run(rec.Stop);
        _tray.ShowBalloonTip(6000, "Raven", "Recording saved. Raven is writing up the transcript (a few minutes, in the background).", ToolTipIcon.Info);
        var captions = _captions.Snapshot(result.StartedUtc);
        void Status(string s) => OnUi(() => { _recStatusItem.Text = s; _recStatusItem.Visible = true; });
        var record = await Task.Run(() => MeetingTranscriber.ProcessAsync(result, captions, _cfg.MyName, Status));
        if (record == null)
        {
            _tray.ShowBalloonTip(8000, "Raven", $"The recording is saved in {result.Folder}, but Raven couldn't write up the transcript. See the log.", ToolTipIcon.Warning);
            return;
        }
        SavePendingRecording(record);
        await SendPendingRecordingsAsync();
    }

    /// <summary>
    /// While Raven shares her screen, the meeting window only shows "you're presenting", so the video follows the
    /// shared screen and goes back to the meeting window when she stops (unless she picked the video herself).
    /// </summary>
    void FollowScreenShare()
    {
        var rec = _rec;
        if (rec == null || _recBar == null || _userPickedVideo) return;
        if (_captions.Sharing && rec.Video.Kind == "window")
        {
            var screen = SharedScreen(rec);
            if (screen == null) return;
            _autoVideo = true;
            Log.Info($"Screen sharing started: video follows {screen.Label}");
            SwitchVideoTo(rec, screen);
        }
        else if (!_captions.Sharing && _autoVideo && rec.Video.Kind == "screen")
        {
            var window = _recSources.FirstOrDefault(s => s.Kind == "window");
            _autoVideo = false;
            if (window == null) return;
            Log.Info("Screen sharing stopped: video back to the meeting window");
            SwitchVideoTo(rec, window);
        }
    }

    void SwitchVideoTo(MeetingRecorder rec, MeetingRecorder.VideoSource s)
    {
        _recBar?.ShowSource(s);
        _ = Task.Run(() =>
        {
            rec.SwitchVideo(s);
            OnUi(() => { if (_rec == rec) _recBar?.ShowSource(rec.Video); });
        });
    }

    /// <summary>The screen being shared: where the browser's "is sharing your screen" bar is, otherwise the screen
    /// the meeting window isn't on (with two screens), otherwise the only screen.</summary>
    MeetingRecorder.VideoSource? SharedScreen(MeetingRecorder rec)
    {
        var screens = _recSources.Where(s => s.Kind == "screen").ToList();
        if (screens.Count == 0) return null;
        if (screens.Count == 1) return screens[0];
        var bar = MeetingWindows.SharingBar();
        var meeting = _captions.MeetingWindow;
        var onScreen = bar != IntPtr.Zero ? MeetingWindows.MonitorOf(bar) : IntPtr.Zero;
        if (onScreen != IntPtr.Zero) return screens.FirstOrDefault(s => s.Handle == onScreen) ?? screens[0];
        var meetingScreen = meeting != IntPtr.Zero ? MeetingWindows.MonitorOf(meeting) : IntPtr.Zero;
        return screens.FirstOrDefault(s => s.Handle != meetingScreen) ?? screens[0];
    }

    /// <summary>Other people can only be named when Teams' live captions are on: one reminder per recording.</summary>
    void WarnIfCaptionsOff()
    {
        if (_warnedCaptions || _rec == null || _rec.Elapsed < TimeSpan.FromSeconds(40) || !_captions.CaptionsOff) return;
        _warnedCaptions = true;
        _tray.ShowBalloonTip(8000, "Raven", "Turn on Live captions in Teams (More → Language and speech) so Raven can tell who's speaking.", ToolTipIcon.Info);
    }

    /// <summary>Runs a small screen update on the app's own thread (from background work).</summary>
    void OnUi(Action a)
    {
        if (_ui == null) return;
        _ui.Post(_ => { try { a(); } catch (Exception ex) { Log.Error("Screen update", ex); } }, null);
    }

    static string PendingFolder => Path.Combine(AgentConfig.DataFolder, "pending-recordings");

    /// <summary>Kept on disk until Clerical Guru has it, so a transcript is never lost to a dropped connection.</summary>
    static void SavePendingRecording(LiveTranscriptRecord r)
    {
        try
        {
            Directory.CreateDirectory(PendingFolder);
            File.WriteAllText(Path.Combine(PendingFolder, r.MeetingKey + ".json"), System.Text.Json.JsonSerializer.Serialize(r, Json.Options));
        }
        catch (Exception ex) { Log.Error("Saving the transcript for upload", ex); }
    }

    async Task SendPendingRecordingsAsync()
    {
        if (!_cfg.IsConfigured || !Directory.Exists(PendingFolder)) return;
        foreach (var file in Directory.GetFiles(PendingFolder, "*.json").Take(2))
        {
            try
            {
                var r = System.Text.Json.JsonSerializer.Deserialize<LiveTranscriptRecord>(File.ReadAllText(file), Json.Options);
                if (r == null) { File.Delete(file); continue; }
                var res = await _client.SendAsync(_cfg.IngestUrl, _cfg.AgentKey, new IngestPayload
                {
                    DeviceId = _cfg.DeviceId, AgentVersion = Version, PollOnly = true, LiveTranscripts = new() { r },
                }, slow: true);
                if (!res.Ok) { Log.Info($"Transcript upload will be retried: {res.Message}"); return; }
                File.Delete(file);
                Log.Info($"Transcript sent: {r.Title} ({r.Lines.Count} lines)");
                _tray.ShowBalloonTip(6000, "Raven", $"The transcript of \"{r.Title}\" is in Clerical Guru: open the meeting to see it and make the summary.", ToolTipIcon.Info);
            }
            catch (Exception ex) { Log.Error("Sending a recorded transcript", ex); }
        }
    }

    void ChooseRecordingsFolder()
    {
        using var dlg = new FolderBrowserDialog { Description = "Where should Raven save meeting recordings?", SelectedPath = _cfg.RecordingsFolder, UseDescriptionForTitle = true };
        if (dlg.ShowDialog() != DialogResult.OK) return;
        _cfg.RecordingsFolder = dlg.SelectedPath;
        _cfg.Save();
    }

    void SetPrivate(bool on)
    {
        _cfg.PrivateMode = on;
        _cfg.Save();
        if (on) _sampler.Flush();
        if (on) { _recPrompt?.Close(); if (_rec != null) _ = StopRecordingAsync(); }
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

    /// <summary>
    /// Ctrl+Alt+R: Raven drafts a reply to the email selected (or open) in Outlook, from what she knows about the
    /// client, and opens Outlook's Reply All window with it typed in. Nothing is sent; you review and press Send.
    /// </summary>
    async Task DraftReplyAsync()
    {
        if (_drafting) return;
        if (!_cfg.IsConfigured) { _tray.ShowBalloonTip(3000, "Raven", "Set up the agent first (Settings).", ToolTipIcon.Info); return; }
        if (!OutlookReader.IsOutlookRunning()) { _tray.ShowBalloonTip(3000, "Raven", "Open Outlook and select the email first.", ToolTipIcon.Info); return; }
        object? item = null;
        DraftRequest? req = null;
        try
        {
            dynamic app = Activator.CreateInstance(Type.GetTypeFromProgID("Outlook.Application")!)!;
            dynamic? window = app.ActiveWindow();
            if (window != null)
            {
                try { item = (object)window.CurrentItem; } catch { }
                if (item == null) { try { dynamic sel = window.Selection; if ((int)sel.Count > 0) item = (object)sel.Item(1); } catch { } }
            }
            if (item == null) { try { dynamic sel = app.ActiveExplorer().Selection; if ((int)sel.Count > 0) item = (object)sel.Item(1); } catch { } }
            dynamic? m = item;
            if (m == null || (int)m.Class != 43) // 43 = an email
            {
                _tray.ShowBalloonTip(3000, "Raven", "Select an email in Outlook first, then press Ctrl+Alt+R.", ToolTipIcon.Info);
                return;
            }
            string S(Func<object?> f) { try { return Convert.ToString(f()) ?? ""; } catch { return ""; } }
            string body = S(() => m.Body);
            req = new DraftRequest
            {
                EntryId = S(() => m.EntryID), Subject = S(() => m.Subject), FromName = S(() => m.SenderName),
                FromEmail = OutlookContext.SenderAddress((object)m) ?? S(() => m.SenderEmailAddress),
                To = S(() => m.To), Cc = S(() => m.CC), Body = body.Length > 20000 ? body[..20000] : body,
                ReceivedAt = DateTime.TryParse(S(() => m.ReceivedTime), out var r) ? Json.Iso(r.ToUniversalTime()) : null,
            };
        }
        catch (Exception ex) { Log.Error("Reading the selected email", ex); _tray.ShowBalloonTip(3000, "Raven", "Couldn't read the selected email.", ToolTipIcon.Warning); return; }

        var subject = req.Subject ?? "";
        var ask = AskGuidance(subject, req.FromName ?? req.FromEmail ?? "");
        if (ask == null) return; // cancelled
        var (guidance, how, forwardTo) = ask.Value;
        _drafting = true;
        _tray.ShowBalloonTip(4000, "Raven", "Drafting your reply… (about 20 seconds)", ToolTipIcon.Info);
        try
        {
            var res = await _client.SendAsync(_cfg.IngestUrl, _cfg.AgentKey, new IngestPayload
            {
                DeviceId = _cfg.DeviceId, AgentVersion = Version, PollOnly = true, DraftRequest = req with { Guidance = guidance, Mode = how == "forward" ? "forward" : "reply", ForwardTo = forwardTo },
            }, slow: true);
            string draft = "", error = res.Ok ? "" : res.Message;
            if (res.Ok)
            {
                using var doc = System.Text.Json.JsonDocument.Parse(res.Message);
                if (doc.RootElement.TryGetProperty("draft", out var d)) draft = d.GetString() ?? "";
                if (doc.RootElement.TryGetProperty("error", out var e)) error = e.GetString() ?? "";
            }
            if (draft.Length == 0)
            {
                Log.Info($"Draft failed: {error}");
                var why = string.IsNullOrWhiteSpace(error) ? "no draft came back" : error.Length > 150 ? error[..150] : error;
                _tray.ShowBalloonTip(8000, "Raven couldn't draft that reply", why, ToolTipIcon.Warning);
                return;
            }
            OutlookOpener.OpenReply(item!, draft, all: how == "all", forwardTo: how == "forward" ? forwardTo : null);
            // What Raven looked up online first (tax rates, rules...), with sources, to confirm before sending
            try
            {
                using var doc2 = System.Text.Json.JsonDocument.Parse(res.Message);
                if (doc2.RootElement.TryGetProperty("checks", out var ch) && ch.ValueKind == System.Text.Json.JsonValueKind.Array && ch.GetArrayLength() > 0)
                    ShowChecks(ch);
            }
            catch (Exception ex) { Log.Error("Showing checks", ex); }
            if (draft.Contains("[[")) _tray.ShowBalloonTip(5000, "Raven", "Draft ready. Fill in the [[blanks]] before you send.", ToolTipIcon.Info);
        }
        catch (Exception ex) { Log.Error("Draft reply", ex); _tray.ShowBalloonTip(8000, "Raven couldn't draft that reply", ex.Message.Length > 150 ? ex.Message[..150] : ex.Message, ToolTipIcon.Warning); }
        finally { _drafting = false; }
    }

    /// <summary>
    /// A roomy box for telling Raven what to write, plus Reply all / Reply / Forward (with who to forward to).
    /// Enter = new line, Ctrl+Enter = Draft. Null if cancelled.
    /// </summary>
    static (string guidance, string how, string forwardTo)? AskGuidance(string subject, string from)
    {
        using var f = new Form
        {
            Text = "Draft with Raven", StartPosition = FormStartPosition.CenterScreen, ClientSize = new Size(560, 380),
            MinimumSize = new Size(440, 330), MinimizeBox = false, MaximizeBox = false, TopMost = true, Font = new Font("Segoe UI", 10f),
            KeyPreview = true,
        };
        var head = new Label
        {
            Text = $"{(from.Length > 40 ? from[..40] + "…" : from)}: {(subject.Length > 60 ? subject[..60] + "…" : subject)}",
            AutoEllipsis = true, Location = new Point(12, 10), Size = new Size(536, 22), Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            Font = new Font("Segoe UI", 10f, FontStyle.Bold),
        };
        var replyAll = new RadioButton { Text = "Reply all", Checked = true, AutoSize = true, Location = new Point(12, 38) };
        var reply = new RadioButton { Text = "Reply", AutoSize = true, Location = new Point(110, 38) };
        var forward = new RadioButton { Text = "Forward to:", AutoSize = true, Location = new Point(190, 38) };
        var to = new TextBox
        {
            Location = new Point(300, 36), Size = new Size(248, 26), Enabled = false, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            PlaceholderText = "name or email (optional)",
        };
        var hint = new Label
        {
            Text = "What should it say? Points to make, answers, tone, what you need. Leave empty to let Raven decide from what she knows.",
            Location = new Point(12, 70), Size = new Size(536, 40), Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right, ForeColor = Color.DimGray,
        };
        var box = new TextBox
        {
            Multiline = true, AcceptsReturn = true, ScrollBars = ScrollBars.Vertical, WordWrap = true,
            Location = new Point(12, 114), Size = new Size(536, 210), Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            PlaceholderText = "e.g. Thank Darin. I'll sign the BBV acceptance block today and send it back once White Dog signs.",
        };
        void Mode()
        {
            to.Enabled = forward.Checked;
            box.PlaceholderText = forward.Checked
                ? "e.g. Josh, see Darin's signed one pager below. Can you sign the BBV block by Monday?"
                : "e.g. Thank Darin. I'll sign the BBV acceptance block today and send it back once White Dog signs.";
            if (forward.Checked) to.Focus();
        }
        forward.CheckedChanged += (_, _) => Mode();
        var tip = new Label { Text = "Ctrl+Enter to draft", AutoSize = true, ForeColor = Color.Gray, Location = new Point(12, 342), Anchor = AnchorStyles.Bottom | AnchorStyles.Left };
        var ok = new Button { Text = "Draft", DialogResult = DialogResult.OK, Size = new Size(90, 32), Location = new Point(362, 336), Anchor = AnchorStyles.Bottom | AnchorStyles.Right };
        var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Size = new Size(90, 32), Location = new Point(458, 336), Anchor = AnchorStyles.Bottom | AnchorStyles.Right };
        f.Controls.AddRange(new Control[] { head, replyAll, reply, forward, to, hint, box, tip, ok, cancel });
        f.CancelButton = cancel;
        f.KeyDown += (_, e) => { if (e.Control && e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; f.DialogResult = DialogResult.OK; f.Close(); } };
        f.Shown += (_, _) => { f.Activate(); box.Focus(); };
        if (f.ShowDialog() != DialogResult.OK) return null;
        return (box.Text.Trim(), forward.Checked ? "forward" : reply.Checked ? "reply" : "all", forward.Checked ? to.Text.Trim() : "");
    }

    /// <summary>A small window listing what Raven checked online (question, answer, sources). Stays until closed.</summary>
    static void ShowChecks(System.Text.Json.JsonElement checks)
    {
        var f = new Form { Text = "Raven checked these before drafting", StartPosition = FormStartPosition.CenterScreen, ClientSize = new Size(560, 300),
            TopMost = true, Font = new Font("Segoe UI", 10f), MinimizeBox = false, MaximizeBox = false };
        var box = new RichTextBox { Dock = DockStyle.Fill, ReadOnly = true, BorderStyle = BorderStyle.None, DetectUrls = true, BackColor = Color.White };
        foreach (var c in checks.EnumerateArray())
        {
            string S(string k) => c.TryGetProperty(k, out var v) && v.ValueKind == System.Text.Json.JsonValueKind.String ? v.GetString() ?? "" : "";
            box.SelectionFont = new Font("Segoe UI", 10f, FontStyle.Bold);
            box.AppendText(S("question") + "\n");
            box.SelectionFont = new Font("Segoe UI", 10f);
            box.AppendText(S("answer"));
            if (c.TryGetProperty("ravens_claim_ok", out var ok) && ok.ValueKind == System.Text.Json.JsonValueKind.False) box.AppendText("  (differs from what you said)");
            if (S("confidence") == "low") box.AppendText("  (not sure - please verify)");
            box.AppendText("\n");
            if (c.TryGetProperty("sources", out var src) && src.ValueKind == System.Text.Json.JsonValueKind.Array)
                foreach (var u in src.EnumerateArray().Take(3)) box.AppendText(u.GetString() + "\n");
            box.AppendText("\n");
        }
        box.LinkClicked += (_, e) => { try { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(e.LinkText!) { UseShellExecute = true }); } catch { } };
        f.Controls.Add(box);
        f.Padding = new Padding(12);
        f.FormClosed += (_, _) => f.Dispose();
        f.Show();
    }

    static string? Prompt(string title, string label, string initial, string okText = "Check", int height = 120)
    {
        using var f = new Form { Text = title, FormBorderStyle = FormBorderStyle.FixedDialog, StartPosition = FormStartPosition.CenterScreen, ClientSize = new Size(360, height),
            MinimizeBox = false, MaximizeBox = false, TopMost = true, Font = new Font("Segoe UI", 10f) };
        var l = new Label { Text = label, AutoSize = true, Location = new Point(12, 12) };
        var t = new TextBox { Text = initial, Location = new Point(12, height - 80), Width = 336 };
        var ok = new Button { Text = okText, DialogResult = DialogResult.OK, Location = new Point(192, height - 42), Width = 75 };
        var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Location = new Point(273, height - 42), Width = 75 };
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
        if (_rec != null)
        {
            if (MessageBox.Show("A meeting is being recorded. Stop the recording and exit?\n\n(The recording is kept, but Raven won't write up the transcript.)",
                    "Raven", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;
            try { _rec.Stop(); } catch (Exception ex) { Log.Error("Stopping the recording", ex); }
            _rec = null;
        }
        _recTimer.Stop();
        Log.Info("Agent exiting");
        _sampleTimer.Stop();
        _uploadTimer.Stop();
        _outlookTimer.Stop();
        _sampler.Flush(); // saved to the queue on disk; sent on next start
        _hotkey?.Dispose();
        _snapKey?.Dispose();
        _replyKey?.Dispose();
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
