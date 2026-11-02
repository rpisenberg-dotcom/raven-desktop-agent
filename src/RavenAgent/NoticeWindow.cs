using RavenAgent.Core;

namespace RavenAgent;

/// <summary>
/// A small box in the bottom-right corner: "Raven thinks this is done - Done / Not yet", or
/// "Raven marked this done - Undo". It never takes the focus from what Raven is typing in, and closes
/// on its own after a while (the question then waits on the Today page).
/// </summary>
sealed class NoticeWindow : Form
{
    readonly Action<string> _answer;
    bool _answered;

    public NoticeWindow(RavenNotice n, Action<string> answer)
    {
        _answer = answer;
        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar = false;
        TopMost = true;
        StartPosition = FormStartPosition.Manual;
        BackColor = Color.White;
        Padding = new Padding(14);
        Width = 380;
        Font = new Font("Segoe UI", 9.5f);

        var suggest = n.Kind == "suggest";
        var header = new Label
        {
            Text = suggest ? "Raven thinks this is done" : "Raven marked this done ✓",
            Font = new Font("Segoe UI", 9f, FontStyle.Bold), ForeColor = Color.FromArgb(109, 40, 217),
            AutoSize = false, Dock = DockStyle.Top, Height = 20,
        };
        var title = new Label { Text = n.Title, Font = new Font("Segoe UI", 10.5f, FontStyle.Bold), AutoSize = false, Dock = DockStyle.Top, Height = 44, ForeColor = Color.FromArgb(15, 23, 42) };
        var why = new Label { Text = n.Why, AutoSize = false, Dock = DockStyle.Top, Height = 36, ForeColor = Color.FromArgb(100, 116, 139) };
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 36, FlowDirection = FlowDirection.RightToLeft };
        Button B(string text, string action, bool primary = false)
        {
            var b = new Button { Text = text, AutoSize = true, Height = 30, FlatStyle = FlatStyle.Flat, Margin = new Padding(6, 0, 0, 0) };
            b.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225);
            if (primary) { b.BackColor = Color.FromArgb(22, 163, 74); b.ForeColor = Color.White; b.FlatAppearance.BorderColor = b.BackColor; }
            b.Click += (_, _) => Answer(action);
            return b;
        }
        if (suggest) { buttons.Controls.Add(B("Done", "done", true)); buttons.Controls.Add(B("Not yet", "not_yet")); }
        else { buttons.Controls.Add(B("OK", "ok", true)); buttons.Controls.Add(B("Undo", "undo")); }
        Controls.Add(why);
        Controls.Add(title);
        Controls.Add(header);
        Controls.Add(buttons);
        Height = 150;
        var area = Screen.PrimaryScreen!.WorkingArea;
        Location = new Point(area.Right - Width - 16, area.Bottom - Height - 16);
        Paint += (_, e) => e.Graphics.DrawRectangle(new Pen(Color.FromArgb(196, 181, 253), 2), 1, 1, Width - 3, Height - 3);

        var timer = new System.Windows.Forms.Timer { Interval = 45000 };
        timer.Tick += (_, _) => { timer.Stop(); Answer(""); };
        timer.Start();
    }

    void Answer(string action)
    {
        if (_answered) return;
        _answered = true;
        try { _answer(action); } catch (Exception ex) { Log.Error("Notice answer failed", ex); }
        Close();
    }

    protected override bool ShowWithoutActivation => true;
    protected override CreateParams CreateParams
    {
        get
        {
            var cp = base.CreateParams;
            cp.ExStyle |= 0x08000000 /* WS_EX_NOACTIVATE */ | 0x00000080 /* WS_EX_TOOLWINDOW */;
            return cp;
        }
    }
}
