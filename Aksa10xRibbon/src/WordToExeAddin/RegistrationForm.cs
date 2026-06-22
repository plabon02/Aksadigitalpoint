using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

public class RegistrationForm : Form
{
    private Label lblStatus;
    private TextBox txtKey;
    private Button btnRegister;
    private Button btnCancel;
    private LinkLabel lnkBuy;
    private WordToExeAddin.LicenseService _license;

    public RegistrationForm(WordToExeAddin.LicenseService license)
    {
        _license = license;
        InitializeComponents();
    }

    private void InitializeComponents()
    {
        Text = "AKSA 10X FASTER - License";
        Size = new Size(440, 300);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        StartPosition = FormStartPosition.CenterParent;
        Font = new Font("Segoe UI", 10);

        int xCtrl = 20, wCtrl = 380, h = 26, gap = 6, y = 20;

        var status = _license.GetStatus();
        string statusText;
        Color statusColor;
        string iconFile;
        bool canActivate = true;

        if (status == WordToExeAddin.LicenseService.LicenseStatus.Licensed)
        {
            string expiry = _license.GetLicenseExpiryDateString() ?? "Lifetime";
            string duration = _license.GetLicenseDurationLabel() ?? "";
            statusText = "Licensed (" + duration + ") - Expires: " + expiry;
            statusColor = Color.Green;
            iconFile = "unlock_keys.png";
            canActivate = false;
        }
        else if (status == WordToExeAddin.LicenseService.LicenseStatus.LicenseExpired)
        {
            string expiry = _license.GetLicenseExpiryDateString() ?? "Unknown";
            statusText = "License expired on " + expiry + "! Enter a new key.";
            statusColor = Color.Red;
            iconFile = "lock_key.png";
        }
        else if (status == WordToExeAddin.LicenseService.LicenseStatus.TrialActive)
        {
            int days = _license.GetTrialDaysLeft();
            string startD = _license.GetTrialStartDateString();
            string endD = _license.GetTrialEndDateString();
            statusText = "Trial: " + days + " day" + (days > 1 ? "s" : "") + " left  (" + startD + " \u2192 " + endD + ")";
            statusColor = Color.DarkOrange;
            iconFile = "lock_key.png";
        }
        else
        {
            string endD = _license.GetTrialEndDateString();
            statusText = "Trial expired on " + endD + "! Purchase to continue.";
            statusColor = Color.Red;
            iconFile = "lock_key.png";
        }

        string iconsPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "Aksa10xFaster", "Icons");
        string iconPath = Path.Combine(iconsPath, iconFile);
        var picIcon = new PictureBox
        {
            Location = new Point(xCtrl, y),
            Size = new Size(48, 48),
            SizeMode = PictureBoxSizeMode.Zoom
        };
        if (File.Exists(iconPath))
        {
            using (var img = Image.FromFile(iconPath))
                picIcon.Image = new Bitmap(img);
        }

        lblStatus = new Label
        {
            Text = statusText,
            Location = new Point(xCtrl + 56, y + 6),
            Size = new Size(wCtrl - 56, 40),
            ForeColor = statusColor,
            Font = new Font("Segoe UI", 11, FontStyle.Bold)
        };

        y = 80;
        var lblInfo = new Label
        {
            Text = "Enter your license key to unlock all features:",
            Location = new Point(xCtrl, y),
            Size = new Size(wCtrl, h)
        };

        y += h + gap;
        txtKey = new TextBox
        {
            Location = new Point(xCtrl, y),
            Size = new Size(wCtrl, h),
            Font = new Font("Consolas", 11),
            CharacterCasing = CharacterCasing.Upper
        };

        y += h + gap + 4;
        btnRegister = new Button
        {
            Text = "Activate License",
            Location = new Point(xCtrl, y),
            Size = new Size(140, 32),
            FlatStyle = FlatStyle.Flat
        };
        btnRegister.Click += BtnRegister_Click;

        btnCancel = new Button
        {
            Text = "Cancel",
            Location = new Point(xCtrl + 150, y),
            Size = new Size(100, 32),
            FlatStyle = FlatStyle.Flat
        };
        btnCancel.Click += (s, e) => Close();

        y += 42;
        lnkBuy = new LinkLabel
        {
            Text = "Don't have a license? Click here to purchase",
            Location = new Point(xCtrl, y),
            Size = new Size(wCtrl, h),
            LinkBehavior = LinkBehavior.HoverUnderline
        };
        lnkBuy.LinkClicked += (s, e) =>
        {
            System.Diagnostics.Process.Start("https://example.com/buy");
        };

        Controls.Add(picIcon);
        Controls.Add(lblStatus);
        Controls.Add(lblInfo);
        Controls.Add(txtKey);
        Controls.Add(btnRegister);
        Controls.Add(btnCancel);
        Controls.Add(lnkBuy);

        if (!canActivate)
        {
            btnRegister.Enabled = false;
            btnRegister.Text = "Already Activated";
            txtKey.Enabled = false;
            lnkBuy.Visible = false;
        }
    }

    private void BtnRegister_Click(object sender, EventArgs e)
    {
        string key = txtKey.Text.Trim();
        if (string.IsNullOrEmpty(key))
        {
            MessageBox.Show("Please enter a license key.", "License", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (_license.Register(key))
        {
            string expiry = _license.GetLicenseExpiryDateString() ?? "Lifetime";
            string duration = _license.GetLicenseDurationLabel() ?? "";
            MessageBox.Show("License activated successfully!\n\nDuration: " + duration + "\nExpires: " + expiry +
                "\n\nThank you for purchasing AKSA 10X FASTER!",
                "License", MessageBoxButtons.OK, MessageBoxIcon.Information);
            DialogResult = DialogResult.OK;
            Close();
        }
        else
        {
            MessageBox.Show("Invalid or expired license key. Please check and try again.\n\nKey format: XXXXX-XXXXX-XXXXX-XXXXX-XXXXX",
                "License", MessageBoxButtons.OK, MessageBoxIcon.Error);
            txtKey.SelectAll();
            txtKey.Focus();
        }
    }
}
