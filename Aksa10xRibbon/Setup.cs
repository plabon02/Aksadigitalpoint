using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Windows.Forms;

class SetupForm : Form
{
    private Panel panelPage;
    private RadioButton rbAccept, rbReject;
    private RadioButton rbTrial, rbLicense;
    private TextBox txtLicenseKey;
    private Button btnBack, btnNext, btnInstall, btnCancel;
    private Label lblStatus;
    private ProgressBar progressBar;
    private int currentPage;

    public SetupForm()
    {
        Text = "AKSA 10X FASTER - Setup";
        Size = new Size(520, 420);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        StartPosition = FormStartPosition.CenterScreen;
        Font = new Font("Segoe UI", 10);
        currentPage = 0;

        panelPage = new Panel { Location = new Point(12, 12), Size = new Size(480, 310) };

        btnBack = new Button { Text = "Back", Location = new Point(190, 340), Size = new Size(90, 32), FlatStyle = FlatStyle.Flat, Enabled = false };
        btnNext = new Button { Text = "Next >", Location = new Point(290, 340), Size = new Size(90, 32), FlatStyle = FlatStyle.Flat };
        btnInstall = new Button { Text = "Install", Location = new Point(290, 340), Size = new Size(90, 32), FlatStyle = FlatStyle.Flat, Visible = false };
        btnCancel = new Button { Text = "Cancel", Location = new Point(390, 340), Size = new Size(90, 32), FlatStyle = FlatStyle.Flat };

        btnBack.Click += (s, e) => { currentPage--; ShowPage(currentPage); };
        btnNext.Click += (s, e) => { if (ValidatePage()) { currentPage++; ShowPage(currentPage); } };
        btnInstall.Click += BtnInstall_Click;
        btnCancel.Click += (s, e) => Close();

        lblStatus = new Label { Text = "", Location = new Point(12, 320), Size = new Size(490, 20), ForeColor = Color.Gray, TextAlign = ContentAlignment.MiddleLeft };
        progressBar = new ProgressBar { Location = new Point(12, 345), Size = new Size(160, 20), Style = ProgressBarStyle.Blocks, Minimum = 0, Maximum = 4, Visible = false };

        Controls.Add(panelPage);
        Controls.Add(lblStatus);
        Controls.Add(progressBar);
        Controls.Add(btnBack);
        Controls.Add(btnNext);
        Controls.Add(btnInstall);
        Controls.Add(btnCancel);

        // App icon
        try
        {
            var pngPath = Path.Combine(Application.StartupPath, "Aksa10xfaster.png");
            if (File.Exists(pngPath))
                Icon = Icon.FromHandle(new Bitmap(pngPath).GetHicon());
        }
        catch { }

        ShowPage(0);
    }

    private void ShowPage(int page)
    {
        panelPage.Controls.Clear();
        currentPage = page;
        btnBack.Visible = true;
        btnNext.Visible = false;
        btnInstall.Visible = false;

        if (page == 0)
        {
            ShowLicensePage();
            btnBack.Enabled = false;
            btnNext.Visible = true;
        }
        else if (page == 1)
        {
            ShowActivationPage();
            btnBack.Enabled = true;
            btnNext.Visible = false;
            btnInstall.Visible = true;
        }
    }

    private void ShowLicensePage()
    {
        int y = 10;
        var lblTitle = new Label
        {
            Text = "License Agreement",
            Font = new Font("Segoe UI", 14, FontStyle.Bold),
            Location = new Point(0, y),
            Size = new Size(480, 30)
        };
        y += 35;

        var lblAgree = new Label
        {
            Text = "Please read the following license agreement carefully.",
            Location = new Point(0, y),
            Size = new Size(480, 20)
        };
        y += 25;

        var tbLicense = new TextBox
        {
            Text = "AKSA 10X FASTER END USER LICENSE AGREEMENT\n\n"
                 + "1. GRANT OF LICENSE: This software is licensed, not sold. You are granted a non-exclusive, non-transferable "
                 + "license to use the software on a single computer.\n\n"
                 + "2. TRIAL PERIOD: You may use the software for a 30-day trial period without payment. After the trial, "
                 + "some premium features (JPG save, Math tools, AI features) will be locked until a valid license key is purchased.\n\n"
                 + "3. LICENSE KEY: Purchasing a license grants you continued access to all features for the chosen duration "
                 + "(6 months, 1 year, 2 years, or lifetime).\n\n"
                 + "4. RESTRICTIONS: You may not distribute, modify, reverse engineer, or rent this software.\n\n"
                 + "5. DISCLAIMER: This software is provided 'as is' without warranty of any kind. The author is not liable "
                 + "for any damages arising from its use.",
            Location = new Point(0, y),
            Size = new Size(480, 180),
            Multiline = true,
            ReadOnly = true,
            BackColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle,
            ScrollBars = ScrollBars.Vertical
        };
        y += 190;

        rbAccept = new RadioButton { Text = "I accept the agreement", Location = new Point(0, y), Size = new Size(300, 22), Checked = true };
        y += 24;
        rbReject = new RadioButton { Text = "I do not accept the agreement", Location = new Point(0, y), Size = new Size(300, 22) };

        rbReject.CheckedChanged += (s, e) => btnNext.Enabled = rbAccept.Checked;
        rbAccept.CheckedChanged += (s, e) => btnNext.Enabled = rbAccept.Checked;

        panelPage.Controls.Add(lblTitle);
        panelPage.Controls.Add(lblAgree);
        panelPage.Controls.Add(tbLicense);
        panelPage.Controls.Add(rbAccept);
        panelPage.Controls.Add(rbReject);
    }

    private void ShowActivationPage()
    {
        int y = 10;
        var lblTitle = new Label
        {
            Text = "Activation",
            Font = new Font("Segoe UI", 14, FontStyle.Bold),
            Location = new Point(0, y),
            Size = new Size(480, 30)
        };
        y += 35;

        lblStatus.Text = "Choose how to activate AKSA 10X FASTER.";
        lblStatus.ForeColor = Color.Gray;

        rbTrial = new RadioButton
        {
            Text = "30 Days Free Trial",
            Location = new Point(0, y),
            Size = new Size(480, 22),
            Checked = true
        };
        y += 26;

        var lblTrialNote = new Label
        {
            Text = "Full features for 30 days. No credit card required. Some features (JPG, Math, AI) will be locked after trial.",
            Location = new Point(18, y),
            Size = new Size(460, 30),
            ForeColor = Color.Gray,
            Font = new Font("Segoe UI", 9)
        };
        y += 36;

        rbLicense = new RadioButton
        {
            Text = "Activate License Key",
            Location = new Point(0, y),
            Size = new Size(480, 22)
        };
        y += 26;

        txtLicenseKey = new TextBox
        {
            Location = new Point(18, y),
            Size = new Size(460, 24),
            Font = new Font("Consolas", 11),
            CharacterCasing = CharacterCasing.Upper,
            Enabled = false,
            Text = ""
        };
        y += 30;

        var lblKeyNote = new Label
        {
            Text = "Enter your 25-character license key (e.g. XXXXX-XXXXX-XXXXX-XXXXX-XXXXX)",
            Location = new Point(18, y),
            Size = new Size(460, 20),
            ForeColor = Color.Gray,
            Font = new Font("Segoe UI", 9)
        };

        rbLicense.CheckedChanged += (s, e) => txtLicenseKey.Enabled = rbLicense.Checked;
        rbTrial.CheckedChanged += (s, e) => txtLicenseKey.Enabled = rbLicense.Checked;

        // Validation status
        var lblKeyValid = new Label
        {
            Text = "",
            Location = new Point(18, y + 20),
            Size = new Size(460, 20),
            ForeColor = Color.Green,
            Font = new Font("Segoe UI", 9)
        };

        txtLicenseKey.TextChanged += (s, e) =>
        {
            var key = txtLicenseKey.Text.Trim().ToUpper();
            if (key.Length > 0)
            {
                string clean = key.Replace("-", "").Replace(" ", "");
                bool valid = clean.Length == 25;
                if (valid)
                {
                    for (int i = 0; i < 25; i++)
                    {
                        char c = clean[i];
                        if (!((c >= 'A' && c <= 'Z') || (c >= '0' && c <= '9')))
                        { valid = false; break; }
                    }
                }
                lblKeyValid.Text = valid ? "✓ Valid license key format" : "Invalid license key format";
                lblKeyValid.ForeColor = valid ? Color.Green : Color.Red;
            }
            else
            {
                lblKeyValid.Text = "";
            }
        };

        panelPage.Controls.Add(lblTitle);
        panelPage.Controls.Add(rbTrial);
        panelPage.Controls.Add(lblTrialNote);
        panelPage.Controls.Add(rbLicense);
        panelPage.Controls.Add(txtLicenseKey);
        panelPage.Controls.Add(lblKeyNote);
        panelPage.Controls.Add(lblKeyValid);
    }

    private bool ValidatePage()
    {
        if (currentPage == 0)
        {
            if (rbReject.Checked)
            {
                MessageBox.Show("You must accept the license agreement to continue.", "License Agreement",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            return true;
        }
        return true;
    }

    private void BtnInstall_Click(object sender, EventArgs e)
    {
        btnInstall.Enabled = false;
        btnCancel.Enabled = false;
        lblStatus.Visible = true;
        progressBar.Visible = true;
        string targetDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Aksa10xFaster");

        try
        {
            SetStatus("Extracting files...", 1);
            if (Directory.Exists(targetDir))
                RemoveDirectoryContents(targetDir);
            Directory.CreateDirectory(targetDir);
            ExtractPackage(targetDir);

            SetStatus("Copying icons...", 2);
            string iconsDest = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "Aksa10xFaster", "Icons");
            Directory.CreateDirectory(iconsDest);
            string iconsSrc = Path.Combine(targetDir, "icons");
            if (Directory.Exists(iconsSrc))
            {
                foreach (var file in Directory.GetFiles(iconsSrc, "*.png"))
                {
                    File.Copy(file, Path.Combine(iconsDest, Path.GetFileName(file)), true);
                }
            }

            SetStatus("Copying document templates...", 3);
            string docTmplSrc = Path.Combine(targetDir, "Document Template");
            string docTmplDest = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "Aksa10xFaster", "Document Template");
            if (Directory.Exists(docTmplSrc))
                CopyDirectory(docTmplSrc, docTmplDest);

            SetStatus("Registering add-in...", 4);
            string dllPath = Path.Combine(targetDir, "bin", "WordAddIn.dll");
            if (!File.Exists(dllPath))
                throw new Exception("WordAddIn.dll not found.");
            RegisterCom(dllPath);

            if (rbLicense.Checked && !string.IsNullOrWhiteSpace(txtLicenseKey.Text))
            {
                string key = txtLicenseKey.Text.Trim().ToUpper();
                Microsoft.Win32.Registry.SetValue(
                    @"HKEY_CURRENT_USER\Software\Aksa10xFaster",
                    "LicenseKey", key);
                SetStatus("License key saved!", 4);
            }

            SetStatus("Installation complete!", 4);
            MessageBox.Show(
                "AKSA 10X FASTER installed successfully!\n\n" +
                (rbLicense.Checked && !string.IsNullOrWhiteSpace(txtLicenseKey.Text)
                    ? "License key has been saved.\n"
                    : "30-day free trial will start when you open Word.\n") +
                "\nPlease restart Microsoft Word.",
                "Setup Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            DialogResult = DialogResult.OK;
            Close();
        }
        catch (Exception ex)
        {
            SetStatus("Installation failed!", 0);
            MessageBox.Show("Error:\n" + ex.Message, "Setup Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            btnInstall.Enabled = true;
            btnCancel.Enabled = true;
        }
    }

    private void ExtractPackage(string targetDir)
    {
        var asm = Assembly.GetExecutingAssembly();
        string zipRes = null;
        foreach (var r in asm.GetManifestResourceNames())
        {
            if (r.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
            { zipRes = r; break; }
        }
        if (zipRes == null)
            throw new Exception("Package not found in resources.");

        string zipPath = Path.Combine(Path.GetTempPath(), "aksapkg.zip");
        using (var stream = asm.GetManifestResourceStream(zipRes))
        using (var fs = new FileStream(zipPath, FileMode.Create, FileAccess.Write))
        {
            stream.CopyTo(fs);
        }
        ZipFile.ExtractToDirectory(zipPath, targetDir);
        File.Delete(zipPath);
    }

    private void RegisterCom(string dllPath)
    {
        string guid = "{33F39564-802B-412B-A713-E947F98DCBB8}";
        string progId = "Aksa10xFaster.Connect";
        string codeBase = "file:///" + dllPath.Replace("\\", "/");

        Microsoft.Win32.Registry.SetValue(@"HKEY_CURRENT_USER\Software\Classes\CLSID\" + guid, "", progId);
        Microsoft.Win32.Registry.SetValue(@"HKEY_CURRENT_USER\Software\Classes\CLSID\" + guid + @"\InprocServer32", "", "mscoree.dll");
        Microsoft.Win32.Registry.SetValue(@"HKEY_CURRENT_USER\Software\Classes\CLSID\" + guid + @"\InprocServer32", "ThreadingModel", "Both");
        Microsoft.Win32.Registry.SetValue(@"HKEY_CURRENT_USER\Software\Classes\CLSID\" + guid + @"\InprocServer32", "Assembly", "WordAddIn, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null");
        Microsoft.Win32.Registry.SetValue(@"HKEY_CURRENT_USER\Software\Classes\CLSID\" + guid + @"\InprocServer32", "Class", "WordToExeAddin.WordAddIn");
        Microsoft.Win32.Registry.SetValue(@"HKEY_CURRENT_USER\Software\Classes\CLSID\" + guid + @"\InprocServer32", "RuntimeVersion", "v4.0.30319");
        Microsoft.Win32.Registry.SetValue(@"HKEY_CURRENT_USER\Software\Classes\CLSID\" + guid + @"\InprocServer32", "CodeBase", codeBase);
        Microsoft.Win32.Registry.SetValue(@"HKEY_CURRENT_USER\Software\Classes\" + progId, "", "AKSA 10X FASTER");
        Microsoft.Win32.Registry.SetValue(@"HKEY_CURRENT_USER\Software\Classes\" + progId + @"\CLSID", "", guid);
        Microsoft.Win32.Registry.SetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Office\Word\Addins\" + progId, "Description", "AKSA 10X FASTER - Word shortcut tools");
        Microsoft.Win32.Registry.SetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Office\Word\Addins\" + progId, "FriendlyName", "AKSA 10X FASTER");
        Microsoft.Win32.Registry.SetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Office\Word\Addins\" + progId, "LoadBehavior", 3);
    }

    private void RemoveDirectoryContents(string dir)
    {
        foreach (var file in Directory.GetFiles(dir))
            File.Delete(file);
        foreach (var sub in Directory.GetDirectories(dir))
            RemoveDirectoryContents(sub);
    }

    private void CopyDirectory(string src, string dst)
    {
        Directory.CreateDirectory(dst);
        foreach (var file in Directory.GetFiles(src))
            File.Copy(file, Path.Combine(dst, Path.GetFileName(file)), true);
        foreach (var dir in Directory.GetDirectories(src))
            CopyDirectory(dir, Path.Combine(dst, Path.GetFileName(dir)));
    }

    private void SetStatus(string text, int progress)
    {
        lblStatus.Text = text;
        progressBar.Value = progress;
        Application.DoEvents();
    }

    [STAThread]
    static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new SetupForm());
    }
}
