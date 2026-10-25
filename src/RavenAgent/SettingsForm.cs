using RavenAgent.Core;

namespace RavenAgent;

/// <summary>Where to send data, the agent key, and what never to record.</summary>
sealed class SettingsForm : Form
{
    readonly AgentConfig _cfg;
    readonly TextBox _url = new() { Width = 460 };
    readonly TextBox _key = new() { Width = 460, UseSystemPasswordChar = true };
    readonly CheckBox _outlook = new() { Text = "Read my Outlook mailboxes and calendars", AutoSize = true };
    readonly TextBox _ignoreApps = new() { Width = 460, Height = 60, Multiline = true, ScrollBars = ScrollBars.Vertical };
    readonly TextBox _ignoreWords = new() { Width = 460, Height = 60, Multiline = true, ScrollBars = ScrollBars.Vertical };
    readonly Label _status = new() { AutoSize = true, MaximumSize = new Size(460, 0), ForeColor = Color.DimGray };

    public SettingsForm(AgentConfig cfg)
    {
        _cfg = cfg;
        Text = "Raven Desktop Agent - Settings";
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        StartPosition = FormStartPosition.CenterScreen;
        AutoSize = true;
        AutoSizeMode = AutoSizeMode.GrowAndShrink;
        Font = new Font("Segoe UI", 9.5f);

        _url.Text = cfg.IngestUrl;
        _key.Text = cfg.AgentKey;
        _outlook.Checked = cfg.ReadOutlook;
        _ignoreApps.Text = string.Join(Environment.NewLine, cfg.IgnoreApps);
        _ignoreWords.Text = string.Join(Environment.NewLine, cfg.IgnoreTitleWords);

        var test = new Button { Text = "Test connection", AutoSize = true };
        var save = new Button { Text = "Save", AutoSize = true, DialogResult = DialogResult.OK };
        var cancel = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
        test.Click += async (_, _) => await TestAsync();
        save.Click += (_, _) => SaveSettings();
        AcceptButton = save;
        CancelButton = cancel;

        var grid = new TableLayoutPanel { ColumnCount = 1, AutoSize = true, Padding = new Padding(16), Dock = DockStyle.Fill };
        void Add(Control c, int top = 0) { c.Margin = new Padding(0, top, 0, 2); grid.Controls.Add(c); }
        Label L(string t) => new() { Text = t, AutoSize = true, MaximumSize = new Size(460, 0) };

        Add(L("Ingest URL (the agentIngest function address from Base44)"));
        Add(_url);
        Add(L("Agent key (the AGENT_INGEST_KEY secret in Base44)"), 10);
        Add(_key);
        Add(_outlook, 10);
        Add(L("Never record these apps (one per line, e.g. Spotify)"), 10);
        Add(_ignoreApps);
        Add(L("Never record windows whose title contains (one per line, e.g. your bank's name)"), 10);
        Add(_ignoreWords);
        var buttons = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 14, 0, 0) };
        buttons.Controls.AddRange(new Control[] { test, save, cancel });
        Add(buttons, 14);
        Add(_status, 6);
        Controls.Add(grid);
    }

    void SaveSettings()
    {
        _cfg.IngestUrl = _url.Text.Trim();
        _cfg.AgentKey = _key.Text.Trim();
        _cfg.ReadOutlook = _outlook.Checked;
        _cfg.IgnoreApps = Lines(_ignoreApps.Text);
        _cfg.IgnoreTitleWords = Lines(_ignoreWords.Text);
        _cfg.Save();
    }

    static List<string> Lines(string s) =>
        s.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).ToList();

    async Task TestAsync()
    {
        _status.ForeColor = Color.DimGray;
        _status.Text = "Testing...";
        using var client = new IngestClient();
        var res = await client.SendAsync(_url.Text.Trim(), _key.Text.Trim(), new IngestPayload
        {
            DeviceId = _cfg.DeviceId,
            AgentVersion = TrayApp.Version,
        });
        _status.ForeColor = res.Ok ? Color.SeaGreen : Color.Firebrick;
        _status.Text = res.Ok ? "Connected. Clerical Guru accepted the test." : res.Message;
    }
}
