using System.Reflection;
using System.Text;
using BISTMatriks.Services;

namespace BISTMatriks.Forms;

public sealed class MainForm : Form
{
    private readonly WatcherService _watcher = new();
    private AppSettings _settings;

    private readonly RichTextBox _logBox;
    private readonly Button      _btnToggle;
    private readonly Button      _btnSettings;
    private readonly Label       _lblStatus;

    // ── Renkler ───────────────────────────────────────────────────────────
    private static readonly Color BgDark    = Color.FromArgb(22, 22, 30);
    private static readonly Color BgMid     = Color.FromArgb(30, 30, 40);
    private static readonly Color BgLight   = Color.FromArgb(40, 40, 52);
    private static readonly Color AccentGrn = Color.FromArgb(0, 115, 55);
    private static readonly Color AccentRed = Color.FromArgb(150, 35, 35);

    public MainForm()
    {
        _settings = AppSettings.Load();

        Text          = "Bayyustek Software";
        Size          = new Size(900, 580);
        MinimumSize   = new Size(600, 420);
        StartPosition = FormStartPosition.CenterScreen;
        BackColor     = BgDark;
        ForeColor     = Color.FromArgb(210, 210, 220);
        Font          = new Font("Segoe UI", 9f);

        // ── Durum etiketi ──────────────────────────────────────────────────
        _lblStatus = new Label
        {
            Text      = "● Durdu",
            ForeColor = Color.OrangeRed,
            Font      = new Font("Segoe UI", 10f, FontStyle.Bold),
            AutoSize  = true,
            Location  = new Point(14, 14)
        };

        // ── Durdur / Başlat butonu ─────────────────────────────────────────
        _btnToggle = new Button
        {
            Text      = "▶  Başlat",
            BackColor = AccentGrn,
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font      = new Font("Segoe UI", 9f, FontStyle.Bold),
            Size      = new Size(110, 32),
            Cursor    = Cursors.Hand
        };
        _btnToggle.FlatAppearance.BorderSize = 0;
        _btnToggle.Click += ToggleWatcher;

        // ── Ayarlar butonu ─────────────────────────────────────────────────
        _btnSettings = new Button
        {
            Text      = "⚙  Ayarlar",
            BackColor = Color.FromArgb(55, 55, 75),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font      = new Font("Segoe UI", 9f),
            Size      = new Size(100, 32),
            Cursor    = Cursors.Hand
        };
        _btnSettings.FlatAppearance.BorderSize = 0;
        _btnSettings.Click += (_, _) => OpenSettings();

        // ── Log kutusu ─────────────────────────────────────────────────────
        _logBox = new RichTextBox
        {
            ReadOnly    = true,
            BackColor   = Color.FromArgb(14, 14, 20),
            ForeColor   = Color.LightGreen,
            Font        = new Font("Consolas", 9f),
            BorderStyle = BorderStyle.None,
            ScrollBars  = RichTextBoxScrollBars.Vertical,
            WordWrap    = false,
            Dock        = DockStyle.Fill
        };

        // ── Uyarı banner'ı ────────────────────────────────────────────────
        var pnlUyari = new Panel
        {
            BackColor = Color.FromArgb(80, 50, 0),
            Dock      = DockStyle.Bottom,
            Height    = 28
        };
        var lblUyariBanner = new Label
        {
            Text      = "⚠  MatriksIQ kodlarını çalıştırmadan uygulamayı başlatmayın!  —  Önce her iki scripti MatriksIQ AlgoTrader'da çalıştırın.",
            ForeColor = Color.Gold,
            Font      = new Font("Segoe UI", 8.5f, FontStyle.Bold),
            Dock      = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleCenter
        };
        pnlUyari.Controls.Add(lblUyariBanner);

        // ── TabControl ─────────────────────────────────────────────────────
        var tabs = BuildTabControl();

        Controls.AddRange(new Control[] { _lblStatus, _btnToggle, _btnSettings, tabs, pnlUyari });

        Resize += (_, _) => ArrangeControls(tabs);
        Shown  += (_, _) => OnShown();
        ArrangeControls(tabs);

        Console.SetOut(new LogWriter(AppendLog));
    }

    // ── Sekme kontrolü ────────────────────────────────────────────────────
    private TabControl BuildTabControl()
    {
        var tc = new TabControl
        {
            DrawMode  = TabDrawMode.OwnerDrawFixed,
            SizeMode  = TabSizeMode.Fixed,
            ItemSize  = new Size(200, 30),
            BackColor = BgDark,
            ForeColor = Color.White,
            Padding   = new Point(10, 4),
            Font      = new Font("Segoe UI", 9f)
        };
        tc.DrawItem += DrawTab;

        var tabIzleme    = new TabPage("📊  İzleme")           { BackColor = BgDark, BorderStyle = BorderStyle.None };
        var tabCanli     = new TabPage("📡  Canlı Veri Kodu")  { BackColor = BgDark, BorderStyle = BorderStyle.None };
        var tabTarihsel  = new TabPage("📈  Tarihsel Veri Kodu"){ BackColor = BgDark, BorderStyle = BorderStyle.None };

        tabIzleme.Controls.Add(_logBox);
        tabCanli.Controls.Add(BuildCodePanel("canli"));
        tabTarihsel.Controls.Add(BuildCodePanel("tarihsel"));

        tc.TabPages.AddRange(new[] { tabIzleme, tabCanli, tabTarihsel });
        return tc;
    }

    private static void DrawTab(object? sender, DrawItemEventArgs e)
    {
        var tc       = (TabControl)sender!;
        bool sel     = e.Index == tc.SelectedIndex;
        var bgColor  = sel ? Color.FromArgb(45, 45, 62) : Color.FromArgb(28, 28, 38);
        var fgColor  = sel ? Color.White : Color.FromArgb(150, 150, 170);

        e.Graphics.FillRectangle(new SolidBrush(bgColor), e.Bounds);

        if (sel)
            e.Graphics.FillRectangle(new SolidBrush(Color.FromArgb(0, 115, 55)),
                new Rectangle(e.Bounds.X, e.Bounds.Bottom - 3, e.Bounds.Width, 3));

        TextRenderer.DrawText(e.Graphics, tc.TabPages[e.Index].Text, tc.Font,
            e.Bounds, fgColor,
            TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
    }

    // ── Kod görüntüleyici panel ────────────────────────────────────────────
    private Panel BuildCodePanel(string tip)
    {
        string code      = LoadEmbeddedCode(tip);
        string csvAdı    = tip == "canli" ? "bist_canli.csv" : "bist_tarihsel.csv";
        string sinifAdı  = tip == "canli" ? "BistCanliVeri" : "BistTarihselVeri";
        string dogruYol  = Path.Combine(AppContext.BaseDirectory, csvAdı);

        var panel = new Panel { Dock = DockStyle.Fill, BackColor = BgDark };

        string calismaUyarisi = tip == "canli"
            ? "⏱  Bu script MatriksIQ'da sürekli açık kalmalıdır — borsa açık olduğu sürece kapatmayın."
            : "⏱  Bu script günde bir kez çalıştırılması yeterlidir — çalıştırıp durdurabilirsiniz.";
        Color calismaRenk = tip == "canli" ? Color.Tomato : Color.MediumAquamarine;

        // Uyarı kutusu
        var pnlInfo = new Panel
        {
            BackColor = Color.FromArgb(60, 40, 10),
            Height    = 132,
            Dock      = DockStyle.Top,
        };

        var lblUyari1 = new Label
        {
            Text      = "⚠  Sarı ile işaretli CIKTI_YOLU satırını değiştirmeniz gerekiyor.",
            ForeColor = Color.Gold,
            Font      = new Font("Segoe UI", 9f, FontStyle.Bold),
            AutoSize  = true,
            Location  = new Point(12, 8)
        };
        var lblUyari2 = new Label
        {
            Text      = "      Size gönderilen \"Bayyustek Software.exe\" dosyasının bulunduğu klasörün yolunu kullanın.",
            ForeColor = Color.White,
            Font      = new Font("Segoe UI", 9f),
            AutoSize  = true,
            Location  = new Point(12, 26)
        };
        var lblUyari3 = new Label
        {
            Text      = $"      Dosya adı da tam olarak şu şekilde olmalıdır:  {csvAdı}",
            ForeColor = Color.LightGreen,
            Font      = new Font("Segoe UI", 9f),
            AutoSize  = true,
            Location  = new Point(12, 44)
        };
        var lblUyari4 = new Label
        {
            Text      = $"      MatriksIQ'da strateji oluştururken sınıf adı olarak tam olarak şunu yazın:  {sinifAdı}",
            ForeColor = Color.SkyBlue,
            Font      = new Font("Segoe UI", 9f),
            AutoSize  = true,
            Location  = new Point(12, 62)
        };
        var lblUyari5 = new Label
        {
            Text      = calismaUyarisi,
            ForeColor = calismaRenk,
            Font      = new Font("Segoe UI", 9f, FontStyle.Bold),
            AutoSize  = true,
            Location  = new Point(12, 82)
        };

        var btnCopyPath = new Button
        {
            Text      = "📋 Doğru Yolu Kopyala",
            ForeColor = Color.White,
            BackColor = Color.FromArgb(80, 65, 15),
            FlatStyle = FlatStyle.Flat,
            Font      = new Font("Segoe UI", 8f),
            AutoSize  = true,
            Location  = new Point(12, 106),
            Cursor    = Cursors.Hand
        };
        btnCopyPath.FlatAppearance.BorderSize = 0;
        btnCopyPath.Click += (_, _) =>
        {
            Clipboard.SetText(dogruYol);
            btnCopyPath.Text = "✓ Kopyalandı!";
            Task.Delay(1500).ContinueWith(_ => Invoke(() => btnCopyPath.Text = "📋 Doğru Yolu Kopyala"));
        };

        pnlInfo.Controls.AddRange(new Control[] { lblUyari1, lblUyari2, lblUyari3, lblUyari4, lblUyari5, btnCopyPath });

        // Kod kutusu
        var rtbCode = new RichTextBox
        {
            ReadOnly    = true,
            BackColor   = Color.FromArgb(14, 14, 20),
            Font        = new Font("Consolas", 9f),
            BorderStyle = BorderStyle.None,
            ScrollBars  = RichTextBoxScrollBars.Both,
            WordWrap    = false,
            Dock        = DockStyle.Fill
        };

        // Kodu sat satır yaz, CIKTI_YOLU satırını vurgula
        foreach (var line in code.Split('\n'))
        {
            rtbCode.SelectionStart  = rtbCode.TextLength;
            rtbCode.SelectionLength = 0;

            if (line.TrimStart().StartsWith("private const string CIKTI_YOLU"))
            {
                rtbCode.SelectionBackColor = Color.FromArgb(70, 60, 0);
                rtbCode.SelectionColor     = Color.Yellow;
            }
            else
            {
                rtbCode.SelectionBackColor = Color.FromArgb(14, 14, 20);
                rtbCode.SelectionColor     = Color.FromArgb(200, 200, 210);
            }
            rtbCode.AppendText(line + "\n");
        }

        // Tüm kodu kopyala butonu
        var pnlBottom = new Panel
        {
            Height    = 44,
            Dock      = DockStyle.Bottom,
            BackColor = BgMid,
            Padding   = new Padding(8)
        };

        var btnCopyAll = new Button
        {
            Text      = "📋 Tüm Kodu Kopyala",
            ForeColor = Color.White,
            BackColor = Color.FromArgb(40, 80, 130),
            FlatStyle = FlatStyle.Flat,
            Font      = new Font("Segoe UI", 9f),
            Height    = 28,
            Width     = 180,
            Location  = new Point(8, 8),
            Cursor    = Cursors.Hand
        };
        btnCopyAll.FlatAppearance.BorderSize = 0;
        btnCopyAll.Click += (_, _) =>
        {
            Clipboard.SetText(code);
            btnCopyAll.Text = "✓ Kopyalandı!";
            Task.Delay(1500).ContinueWith(_ => Invoke(() => btnCopyAll.Text = "📋 Tüm Kodu Kopyala"));
        };

        pnlBottom.Controls.Add(btnCopyAll);
        panel.Controls.AddRange(new Control[] { pnlBottom, rtbCode, pnlInfo });
        return panel;
    }

    private static string LoadEmbeddedCode(string tip)
    {
        string resName = tip == "canli"
            ? "BISTMatriks.Resources.Canli_Veri_MatriksIQ.cs"
            : "BISTMatriks.Resources.Tarihsel_Veri_MatriksIQ.cs";

        var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resName);
        if (stream == null) return $"// Kaynak bulunamadı: {resName}";
        using var sr = new StreamReader(stream, Encoding.UTF8);
        return sr.ReadToEnd();
    }

    // ── Düzen ─────────────────────────────────────────────────────────────
    private void ArrangeControls(TabControl tabs)
    {
        const int pad = 12, topH = 52, botH = 28;
        _btnSettings.Location = new Point(ClientSize.Width - pad - _btnSettings.Width, pad);
        _btnToggle.Location   = new Point(_btnSettings.Left - 6 - _btnToggle.Width, pad);
        _lblStatus.Location   = new Point(pad, pad + 6);
        tabs.SetBounds(0, topH, ClientSize.Width, ClientSize.Height - topH - botH);
    }

    private void OnShown()
    {
        if (!_settings.IsConfigured) OpenSettings();
        // Otomatik başlatma yok — kullanıcı Başlat'a basmalı
    }

    private void ToggleWatcher(object? sender, EventArgs e)
    {
        if (_watcher.IsRunning)
            StopWatcher();
        else
        {
            if (!_settings.IsConfigured) { OpenSettings(); return; }
            StartWatcher();
        }
    }

    private void StartWatcher()
    {
        _btnToggle.Enabled   = false;
        _btnSettings.Enabled = false;

        Task.Run(() =>
        {
            try
            {
                _watcher.Start(_settings);
                Invoke(() =>
                {
                    _btnToggle.Text      = "■  Durdur";
                    _btnToggle.BackColor = AccentRed;
                    _lblStatus.Text      = "● Çalışıyor";
                    _lblStatus.ForeColor = Color.LimeGreen;
                });
            }
            catch (Exception ex)
            {
                Invoke(() =>
                    Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] [HATA] Servis başlatılamadı: {ex.Message}"));
            }
            finally
            {
                Invoke(() =>
                {
                    _btnToggle.Enabled   = true;
                    _btnSettings.Enabled = true;
                });
            }
        });
    }

    private void StopWatcher()
    {
        _watcher.Stop();
        _btnToggle.Text      = "▶  Başlat";
        _btnToggle.BackColor = AccentGrn;
        _lblStatus.Text      = "● Durdu";
        _lblStatus.ForeColor = Color.OrangeRed;
    }

    private void OpenSettings()
    {
        bool wasRunning = _watcher.IsRunning;
        if (wasRunning) StopWatcher();
        using var dlg = new SettingsForm(_settings);
        if (dlg.ShowDialog(this) == DialogResult.OK)
        {
            _settings = dlg.Result;
            _settings.Save();
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] Ayarlar kaydedildi. Başlatmak için ▶ Başlat butonuna basın.");
        }
        else if (wasRunning) StartWatcher();
    }

    private void AppendLog(string msg)
    {
        if (InvokeRequired) { Invoke(() => AppendLog(msg)); return; }
        Color color = msg.Contains("[HATA]") ? Color.Tomato
                    : msg.Contains("UYARI")  ? Color.Gold
                    : msg.Contains("COM")    ? Color.Cyan
                    : Color.FromArgb(140, 220, 140);
        _logBox.SelectionStart  = _logBox.TextLength;
        _logBox.SelectionLength = 0;
        _logBox.SelectionColor  = color;
        _logBox.AppendText(msg + "\n");
        _logBox.ScrollToCaret();
    }

    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        _watcher.Dispose();
        base.OnFormClosed(e);
    }
}

internal sealed class LogWriter : TextWriter
{
    private readonly Action<string> _onLine;
    private readonly StringBuilder  _buf = new();
    public LogWriter(Action<string> onLine) => _onLine = onLine;
    public override Encoding Encoding => Encoding.UTF8;
    public override void Write(char value)
    {
        if (value == '\n') Flush();
        else if (value != '\r') _buf.Append(value);
    }
    public override void Write(string? value) { if (value != null) foreach (char c in value) Write(c); }
    public override void WriteLine(string? value) { if (value != null) _buf.Append(value); Flush(); }
    public override void Flush() { if (_buf.Length > 0) { _onLine(_buf.ToString()); _buf.Clear(); } }
}
