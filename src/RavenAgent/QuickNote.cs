using System.Runtime.InteropServices;
using System.Text.Json;
using RavenAgent.Core;

namespace RavenAgent;

/// <summary>Ctrl+Alt+N anywhere: a hidden window that receives the system-wide hotkey.</summary>
sealed class HotkeyWindow : NativeWindow, IDisposable
{
    [DllImport("user32.dll")] static extern bool RegisterHotKey(IntPtr hWnd, int id, uint mods, uint vk);
    [DllImport("user32.dll")] static extern bool UnregisterHotKey(IntPtr hWnd, int id);
    const int WM_HOTKEY = 0x0312, Id = 0x5241;
    const uint MOD_ALT = 1, MOD_CONTROL = 2, MOD_NOREPEAT = 0x4000;
    readonly Action _pressed;
    public bool Registered { get; }

    public HotkeyWindow(Action pressed)
    {
        _pressed = pressed;
        CreateHandle(new CreateParams());
        Registered = RegisterHotKey(Handle, Id, MOD_CONTROL | MOD_ALT | MOD_NOREPEAT, (uint)Keys.N);
        if (!Registered) Log.Info("Ctrl+Alt+N is already used by another program; use the tray menu to jot a note");
    }

    protected override void WndProc(ref Message m)
    {
        if (m.Msg == WM_HOTKEY && (int)m.WParam == Id) _pressed();
        base.WndProc(ref m);
    }

    public void Dispose()
    {
        if (Registered) UnregisterHotKey(Handle, Id);
        DestroyHandle();
    }
}

/// <summary>Notes waiting to be sent, kept on disk so nothing is lost if the PC is offline.</summary>
sealed class NoteQueue
{
    static string FilePath => Path.Combine(AgentConfig.DataFolder, "notes.json");
    readonly object _lock = new();
    List<QuickNoteRecord> _notes = new();

    public NoteQueue()
    {
        try { if (File.Exists(FilePath)) _notes = JsonSerializer.Deserialize<List<QuickNoteRecord>>(File.ReadAllText(FilePath), Json.Options) ?? new(); }
        catch (Exception ex) { Log.Error("Couldn't load saved notes", ex); }
    }
    public void Add(QuickNoteRecord n) { lock (_lock) { _notes.Add(n); Save(); } }
    public List<QuickNoteRecord> Peek() { lock (_lock) return _notes.Take(50).ToList(); }
    public void Remove(IEnumerable<QuickNoteRecord> sent)
    {
        var keys = sent.Select(n => n.AgentKey).ToHashSet();
        lock (_lock) { if (_notes.RemoveAll(n => keys.Contains(n.AgentKey)) > 0) Save(); }
    }
    void Save()
    {
        try { Directory.CreateDirectory(AgentConfig.DataFolder); File.WriteAllText(FilePath, JsonSerializer.Serialize(_notes, Json.Options)); }
        catch (Exception ex) { Log.Error("Couldn't save notes", ex); }
    }
}

/// <summary>A small window to jot a question, to-do or note for a client. Enter saves, Esc closes.</summary>
sealed class QuickNoteWindow : Form
{
    sealed record Choice(string Label, string? ClientId, string? Category) { public override string ToString() => Label; }

    readonly TextBox _text;
    readonly ComboBox _client;
    readonly RadioButton _question, _todo, _note;
    readonly string _context;
    readonly Action<QuickNoteRecord> _save;

    public QuickNoteWindow(List<ClientEntry> clients, string? preselectClientId, string context, Action<QuickNoteRecord> save)
    {
        _context = context;
        _save = save;
        Text = "Raven - jot a note";
        FormBorderStyle = FormBorderStyle.FixedToolWindow;
        StartPosition = FormStartPosition.CenterScreen;
        TopMost = true;
        ShowInTaskbar = true;
        KeyPreview = true;
        Font = new Font("Segoe UI", 10f);
        BackColor = Color.White;
        ClientSize = new Size(460, 250);
        Padding = new Padding(12);

        var clientLabel = new Label { Text = "For:", AutoSize = true, Location = new Point(12, 16) };
        _client = new ComboBox { DropDownStyle = ComboBoxStyle.DropDown, AutoCompleteMode = AutoCompleteMode.SuggestAppend, AutoCompleteSource = AutoCompleteSource.ListItems, Location = new Point(52, 12), Width = 396 };
        _client.Items.Add(new Choice("No client yet (sort it later)", null, null));
        _client.Items.Add(new Choice("CG Business", null, "business"));
        _client.Items.Add(new Choice("Personal", null, "personal"));
        foreach (var c in clients) _client.Items.Add(new Choice(c.Name, c.Id, "client"));
        _client.SelectedIndex = 0;
        if (preselectClientId != null)
            for (var i = 0; i < _client.Items.Count; i++)
                if (((Choice)_client.Items[i]!).ClientId == preselectClientId) { _client.SelectedIndex = i; break; }

        _question = new RadioButton { Text = "Question", AutoSize = true, Location = new Point(12, 48), Checked = true };
        _todo = new RadioButton { Text = "To-do", AutoSize = true, Location = new Point(112, 48) };
        _note = new RadioButton { Text = "Note", AutoSize = true, Location = new Point(196, 48) };

        _text = new TextBox { Multiline = true, AcceptsReturn = false, ScrollBars = ScrollBars.Vertical, Location = new Point(12, 80), Size = new Size(436, 110) };
        var hint = new Label { Text = "Enter saves · Shift+Enter for a new line · Esc closes", AutoSize = true, ForeColor = Color.FromArgb(100, 116, 139), Font = new Font("Segoe UI", 8.5f), Location = new Point(12, 212) };
        var save = new Button { Text = "Save", Location = new Point(368, 204), Size = new Size(80, 32), FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(37, 99, 235), ForeColor = Color.White };
        save.FlatAppearance.BorderSize = 0;
        save.Click += (_, _) => Save();

        Controls.AddRange(new Control[] { clientLabel, _client, _question, _todo, _note, _text, hint, save });
        _text.KeyDown += (_, e) =>
        {
            if (e.KeyCode == Keys.Enter && e.Shift) { _text.SelectedText = Environment.NewLine; e.SuppressKeyPress = true; }
            else if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; Save(); }
        };
        KeyDown += (_, e) => { if (e.KeyCode == Keys.Escape) Close(); };
        Shown += (_, _) => { Activate(); BringToFront(); _text.Focus(); };
    }

    Choice? SelectedChoice()
    {
        if (_client.SelectedItem is Choice c) return c;
        // Typed a name: pick the matching client, or none
        var typed = _client.Text.Trim();
        foreach (Choice x in _client.Items) if (string.Equals(x.Label, typed, StringComparison.OrdinalIgnoreCase)) return x;
        return null;
    }

    void Save()
    {
        var text = _text.Text.Trim();
        if (text.Length == 0) { Close(); return; }
        var c = SelectedChoice();
        _save(new QuickNoteRecord
        {
            AgentKey = "note-" + Guid.NewGuid().ToString("N"),
            Text = text,
            Kind = _todo.Checked ? "todo" : _note.Checked ? "note" : "question",
            ClientId = c?.ClientId,
            Category = c?.Category,
            Context = _context.Length > 300 ? _context[..300] : _context,
        });
        Close();
    }
}
