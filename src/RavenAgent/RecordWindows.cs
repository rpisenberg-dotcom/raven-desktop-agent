using System.Runtime.InteropServices;

namespace RavenAgent;

/// <summary>Raven's own little windows are left out of screen shares and of her own screen recordings.</summary>
static class CaptureHide
{
    [DllImport("user32.dll")] static extern bool SetWindowDisplayAffinity(IntPtr hwnd, uint affinity);
    const uint WDA_EXCLUDEFROMCAPTURE = 0x11;
    public static void Apply(Form f) { try { SetWindowDisplayAffinity(f.Handle, WDA_EXCLUDEFROMCAPTURE); } catch { } }
}

/// <summary>
/// "Teams meeting started - record it?" Recording can only start once Raven ticks that everyone has been told.
/// Closes by itself after 3 minutes (nothing is recorded).
/// </summary>
sealed class RecordPrompt : Form
{
    readonly System.Windows.Forms.Timer _timeout = new() { Interval = 180_000 };

    public RecordPrompt(string app, List<MeetingRecorder.VideoSource> sources, string defaultKind, string notice, Action<MeetingRecorder.VideoSource> start)
    {
        Text = "Raven";
        FormBorderStyle = FormBorderStyle.FixedToolWindow;
        TopMost = true;
        ShowInTaskbar = true;
        StartPosition = FormStartPosition.Manual;
        BackColor = Color.White;
        Font = new Font("Segoe UI", 9.5f);
        ClientSize = new Size(440, 268);

        var panel = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(16, 12, 16, 8) };
        var head = new Label { Text = $"{app} meeting started. Record it?", Font = new Font("Segoe UI", 11.5f, FontStyle.Bold), ForeColor = Color.FromArgb(15, 23, 42), AutoSize = true, Margin = new Padding(0, 0, 0, 4) };
        var info = new Label
        {
            Text = "Raven records the sound (and video if you like) to your Raven Recordings folder, then writes up a transcript with names, the summary and the follow-up tasks.",
            AutoSize = false, Size = new Size(405, 40), ForeColor = Color.FromArgb(71, 85, 105),
        };
        var told = new CheckBox { Text = "I've told everyone this meeting is being recorded and transcribed", AutoSize = false, Size = new Size(405, 26), Font = new Font("Segoe UI", 9.5f, FontStyle.Bold), Margin = new Padding(0, 6, 0, 0) };
        var copy = new LinkLabel { Text = "Copy a notice to paste in the meeting chat", AutoSize = true, Margin = new Padding(18, 0, 0, 8) };
        copy.LinkClicked += (_, _) => { try { Clipboard.SetText(notice); copy.Text = "Copied - paste it in the meeting chat ✓"; } catch { } };

        var row = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, WrapContents = false, Margin = new Padding(0, 2, 0, 8) };
        row.Controls.Add(new Label { Text = "Video:", AutoSize = true, Margin = new Padding(0, 6, 6, 0) });
        var combo = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 230 };
        foreach (var s in sources) combo.Items.Add(s.Label);
        var pick = sources.FindIndex(s => s.Kind == defaultKind);
        combo.SelectedIndex = pick >= 0 ? pick : 0;
        row.Controls.Add(combo);
        var tip = new Label
        {
            Text = "Keep the meeting window open (behind other windows is fine, minimized isn't).",
            AutoSize = false, Size = new Size(405, 20), ForeColor = Color.FromArgb(100, 116, 139), Font = new Font("Segoe UI", 8.5f),
        };

        var buttons = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 6, 0, 0) };
        var go = new Button { Text = "Start recording", AutoSize = true, Enabled = false, BackColor = Color.FromArgb(220, 38, 38), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Padding = new Padding(8, 2, 8, 2) };
        go.FlatAppearance.BorderSize = 0;
        var skip = new Button { Text = "Not this time", AutoSize = true, FlatStyle = FlatStyle.Flat, Padding = new Padding(8, 2, 8, 2), Margin = new Padding(8, 0, 0, 0) };
        told.CheckedChanged += (_, _) => go.Enabled = told.Checked;
        go.Click += (_, _) => { var s = sources[Math.Max(0, combo.SelectedIndex)]; Close(); start(s); };
        skip.Click += (_, _) => Close();
        buttons.Controls.Add(go);
        buttons.Controls.Add(skip);

        panel.Controls.AddRange(new Control[] { head, info, told, copy, row, tip, buttons });
        Controls.Add(panel);

        _timeout.Tick += (_, _) => Close();
        _timeout.Start();
        FormClosed += (_, _) => _timeout.Dispose();
        var wa = Screen.PrimaryScreen!.WorkingArea;
        Location = new Point(wa.Right - Width - 16, wa.Bottom - Height - 16);
    }

    protected override void OnHandleCreated(EventArgs e) { base.OnHandleCreated(e); CaptureHide.Apply(this); }
}

/// <summary>A small always-on-top bar while recording: time, what the video shows (can be switched), Stop.</summary>
sealed class RecordingBar : Form
{
    readonly Label _time;
    readonly ComboBox _video;
    readonly List<MeetingRecorder.VideoSource> _sources;
    bool _quiet;

    [DllImport("user32.dll")] static extern bool ReleaseCapture();
    [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr h, int msg, IntPtr w, IntPtr l);

    public RecordingBar(List<MeetingRecorder.VideoSource> sources, MeetingRecorder.VideoSource current, Action<MeetingRecorder.VideoSource> switchTo, Action stop)
    {
        _sources = sources;
        FormBorderStyle = FormBorderStyle.None;
        TopMost = true;
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.Manual;
        BackColor = Color.FromArgb(30, 41, 59);
        Font = new Font("Segoe UI", 9f);
        ClientSize = new Size(390, 36);

        var row = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, WrapContents = false, Padding = new Padding(8, 5, 6, 4) };
        _time = new Label { Text = "● Recording 0:00", ForeColor = Color.FromArgb(248, 113, 113), Font = new Font("Segoe UI", 9.5f, FontStyle.Bold), AutoSize = true, Margin = new Padding(0, 4, 8, 0), Cursor = Cursors.SizeAll };
        _video = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 150, FlatStyle = FlatStyle.Flat };
        foreach (var s in sources) _video.Items.Add(s.Label);
        ShowSource(current);
        _video.SelectedIndexChanged += (_, _) => { if (!_quiet && _video.SelectedIndex >= 0) switchTo(_sources[_video.SelectedIndex]); };
        var stopBtn = new Button { Text = "Stop", AutoSize = true, BackColor = Color.FromArgb(220, 38, 38), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Margin = new Padding(8, 0, 0, 0) };
        stopBtn.FlatAppearance.BorderSize = 0;
        stopBtn.Click += (_, _) => stop();
        row.Controls.AddRange(new Control[] { _time, _video, stopBtn });
        Controls.Add(row);

        // Drag it anywhere by the "Recording" label
        _time.MouseDown += (_, e) => { if (e.Button == MouseButtons.Left) { ReleaseCapture(); SendMessage(Handle, 0xA1, (IntPtr)2, IntPtr.Zero); } };
        var wa = Screen.PrimaryScreen!.WorkingArea;
        Location = new Point(wa.Left + (wa.Width - Width) / 2, wa.Top + 8);
    }

    /// <summary>Shows what the video is really recording (it falls back to sound only if a source can't be captured).</summary>
    public void ShowSource(MeetingRecorder.VideoSource current)
    {
        _quiet = true;
        var i = _sources.FindIndex(s => s.Kind == current.Kind && s.Handle == current.Handle);
        _video.SelectedIndex = i >= 0 ? i : _sources.Count - 1;
        _quiet = false;
    }

    public void Tick(TimeSpan elapsed) =>
        _time.Text = $"● Recording {(elapsed.TotalHours >= 1 ? elapsed.ToString(@"h\:mm\:ss") : elapsed.ToString(@"m\:ss"))}";

    protected override void OnHandleCreated(EventArgs e) { base.OnHandleCreated(e); CaptureHide.Apply(this); }
}
