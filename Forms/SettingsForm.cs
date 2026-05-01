using BISTMatriks.Services;

namespace BISTMatriks.Forms;

public sealed class SettingsForm : Form
{
    public AppSettings Result { get; private set; }

    private readonly TextBox _txtWorkbook;

    public SettingsForm(AppSettings current)
    {
        Result = current;

        Text            = "Bayyustek Software — Ayarlar";
        Size            = new Size(620, 160);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox     = false;
        StartPosition   = FormStartPosition.CenterParent;
        BackColor       = Color.FromArgb(28, 28, 36);
        ForeColor       = Color.FromArgb(210, 210, 220);
        Font            = new Font("Segoe UI", 9f);

        var lbl = new Label
        {
            Text      = "Excel Dosyası (.xlsm):",
            Location  = new Point(14, 22),
            Size      = new Size(160, 20),
            ForeColor = Color.FromArgb(160, 160, 180)
        };

        _txtWorkbook = new TextBox
        {
            Text        = current.WorkbookPath,
            Location    = new Point(178, 19),
            Size        = new Size(330, 24),
            BackColor   = Color.FromArgb(40, 40, 52),
            ForeColor   = Color.White,
            BorderStyle = BorderStyle.FixedSingle
        };

        var btnBrowse = new Button
        {
            Text      = "...",
            Location  = new Point(516, 19),
            Size      = new Size(72, 24),
            BackColor = Color.FromArgb(60, 60, 80),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor    = Cursors.Hand
        };
        btnBrowse.FlatAppearance.BorderSize = 0;
        btnBrowse.Click += (_, _) => Browse(_txtWorkbook);

        var btnSave = new Button
        {
            Text      = "Kaydet",
            Location  = new Point(426, 68),
            Size      = new Size(80, 28),
            BackColor = Color.FromArgb(0, 110, 55),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor    = Cursors.Hand
        };
        btnSave.FlatAppearance.BorderSize = 0;
        btnSave.Click += BtnSave_Click;

        var btnCancel = new Button
        {
            Text         = "İptal",
            Location     = new Point(514, 68),
            Size         = new Size(74, 28),
            BackColor    = Color.FromArgb(70, 40, 40),
            ForeColor    = Color.White,
            FlatStyle    = FlatStyle.Flat,
            DialogResult = DialogResult.Cancel,
            Cursor       = Cursors.Hand
        };
        btnCancel.FlatAppearance.BorderSize = 0;

        Controls.AddRange(new Control[] { lbl, _txtWorkbook, btnBrowse, btnSave, btnCancel });
        AcceptButton = btnSave;
        CancelButton = btnCancel;
    }

    private void Browse(TextBox txt)
    {
        var thread = new Thread(() =>
        {
            using var dlg = new OpenFileDialog
            {
                Filter             = "Excel Makro Kitabı (*.xlsm)|*.xlsm|Tüm dosyalar (*.*)|*.*",
                Title              = "Excel Dosyası Seç",
                AutoUpgradeEnabled = true
            };

            string? current = null;
            Invoke(() => current = txt.Text);
            if (!string.IsNullOrWhiteSpace(current))
            {
                string? dir = Path.GetDirectoryName(current);
                if (dir != null && Directory.Exists(dir))
                    dlg.InitialDirectory = dir;
            }

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                string path = dlg.FileName;
                Invoke(() => txt.Text = path);
            }
        });
        thread.SetApartmentState(ApartmentState.STA);
        thread.IsBackground = true;
        thread.Start();
    }

    private void BtnSave_Click(object? sender, EventArgs e)
    {
        string path = _txtWorkbook.Text.Trim();

        if (string.IsNullOrWhiteSpace(path))
        {
            MessageBox.Show("Lütfen Excel dosyasını seçin.", "Eksik Bilgi",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (!File.Exists(path))
        {
            MessageBox.Show(
                $"Dosya bulunamadı:\n{path}\n\nLütfen geçerli bir .xlsm dosyası seçin.",
                "Dosya Bulunamadı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        Result = new AppSettings { WorkbookPath = path };
        DialogResult = DialogResult.OK;
        Close();
    }
}
