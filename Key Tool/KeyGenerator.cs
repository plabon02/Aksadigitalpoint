using System;
using System.Drawing;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;

class KeyGeneratorForm : Form
{
    private ComboBox cmbDuration;
    private TextBox txtSeed;
    private TextBox txtKey;
    private Button btnGenerate;
    private Button btnCopy;
    private Button btnClose;

    private static readonly DateTime BaseDate = new DateTime(2026, 1, 1);
    private const string SecretSalt = "AKSA10X_V2_2026";
    private const string CharSet = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ";

    public KeyGeneratorForm()
    {
        Text = "AKSA 10X FASTER - License Key Generator";
        Size = new Size(540, 260);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        StartPosition = FormStartPosition.CenterScreen;
        Font = new Font("Segoe UI", 10);
        BackColor = Color.FromArgb(245, 245, 245);

        int x = 20, w = 490, h = 26, gap = 6, y = 20;

        var lblSeed = new Label { Text = "Customer / Seed:", Location = new Point(x, y), Size = new Size(120, h) };
        txtSeed = new TextBox
        {
            Location = new Point(x + 125, y), Size = new Size(w - 125, h),
            Text = "Aksa10x",
            Font = new Font("Consolas", 10)
        };

        y += h + gap;
        var lblDuration = new Label { Text = "Duration:", Location = new Point(x, y), Size = new Size(120, h) };
        cmbDuration = new ComboBox
        {
            Location = new Point(x + 125, y), Size = new Size(200, h),
            DropDownStyle = ComboBoxStyle.DropDownList
        };
        cmbDuration.Items.AddRange(new[] {
            "6 Months",
            "1 Year",
            "2 Years",
            "Lifetime"
        });
        cmbDuration.SelectedIndex = 1;

        y += h + gap + 4;
        btnGenerate = new Button
        {
            Text = "Generate Key",
            Location = new Point(x + 125, y),
            Size = new Size(130, 34),
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(0, 120, 215),
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };
        btnGenerate.Click += BtnGenerate_Click;

        y += 44;
        var lblKey = new Label { Text = "License Key:", Location = new Point(x, y), Size = new Size(120, h) };
        txtKey = new TextBox
        {
            Location = new Point(x + 125, y), Size = new Size(w - 125, h + 4),
            Font = new Font("Consolas", 12, FontStyle.Bold),
            ReadOnly = true,
            BackColor = Color.White,
            TextAlign = HorizontalAlignment.Center
        };

        y += h + gap + 6;
        btnCopy = new Button
        {
            Text = "Copy to Clipboard",
            Location = new Point(x + 125, y),
            Size = new Size(130, 30),
            FlatStyle = FlatStyle.Flat,
            Enabled = false
        };
        btnCopy.Click += (s, e) =>
        {
            if (!string.IsNullOrEmpty(txtKey.Text))
            {
                Clipboard.SetText(txtKey.Text);
                MessageBox.Show("License key copied to clipboard!", "Copied", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        };

        btnClose = new Button
        {
            Text = "Close",
            Location = new Point(x + 270, y),
            Size = new Size(100, 30),
            FlatStyle = FlatStyle.Flat
        };
        btnClose.Click += (s, e) => Close();

        Controls.AddRange(new Control[] {
            lblSeed, txtSeed, lblDuration, cmbDuration,
            btnGenerate, lblKey, txtKey, btnCopy, btnClose
        });
    }

    private void BtnGenerate_Click(object sender, EventArgs e)
    {
        string seed = txtSeed.Text.Trim();
        if (string.IsNullOrEmpty(seed))
        {
            MessageBox.Show("Please enter a customer name or seed.", "Input Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        int months;
        switch (cmbDuration.SelectedIndex)
        {
            case 0: months = 6; break;
            case 1: months = 12; break;
            case 2: months = 24; break;
            case 3: months = 999999; break;
            default: months = 12; break;
        }

        string key = GenerateKey(seed, months);
        txtKey.Text = key;
        btnCopy.Enabled = true;
    }

    public static string GenerateKey(string seed, int durationMonths)
    {
        using (var sha = SHA256.Create())
        {
            byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes(seed + SecretSalt + durationMonths));
            char[] result = new char[25];

            DateTime expiry;
            if (durationMonths >= 100000)
            {
                expiry = DateTime.MaxValue;
            }
            else
            {
                expiry = DateTime.Now.Date.AddMonths(durationMonths);
            }

            string expiryCode;
            if (durationMonths >= 100000)
            {
                expiryCode = "ZZZ";
            }
            else
            {
                int monthOffset = (expiry.Year - BaseDate.Year) * 12 + (expiry.Month - BaseDate.Month);
                expiryCode = EncodeBase36(monthOffset, 3);
            }

            for (int i = 0; i < 3; i++)
                result[i] = expiryCode[i];

            for (int i = 3; i < 23; i++)
                result[i] = CharSet[hash[(i - 3) % hash.Length] % 36];

            char durChar;
            if (durationMonths >= 100000) durChar = 'L';
            else if (durationMonths >= 24) durChar = '2';
            else if (durationMonths >= 12) durChar = '1';
            else durChar = '6';
            result[23] = durChar;

            int checksum = 0;
            for (int i = 0; i < 24; i++)
            {
                char c = result[i];
                checksum += (c >= '0' && c <= '9') ? (c - '0') : (c - 'A' + 10);
            }
            result[24] = CharSet[checksum % 36];

            return new string(result, 0, 5) + "-" +
                   new string(result, 5, 5) + "-" +
                   new string(result, 10, 5) + "-" +
                   new string(result, 15, 5) + "-" +
                   new string(result, 20, 5);
        }
    }

    private static string EncodeBase36(int value, int minLen)
    {
        string result = "";
        do
        {
            result = CharSet[value % 36] + result;
            value /= 36;
        } while (value > 0);
        while (result.Length < minLen)
            result = "0" + result;
        return result;
    }

    [STAThread]
    static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new KeyGeneratorForm());
    }
}
