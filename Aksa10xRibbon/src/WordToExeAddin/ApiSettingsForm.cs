using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

public class ApiSettingsForm : Form
{
    private ComboBox cmbProvider;
    private TextBox txtApiKey;
    private ComboBox cmbModel;
    private Button btnSave;
    private Button btnTest;
    private Label lblStatus;
    private CheckBox chkShowKey;
    private LinkLabel lnkGetKey;

    private static readonly string ConfigPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "Aksa10xFaster", "config.json");

    private static readonly Dictionary<string, string[]> Models = new Dictionary<string, string[]>
    {
        { "gemini", new[] { "gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash" } },
        { "deepseek", new[] { "deepseek-chat", "deepseek-reasoner" } },
        { "openai", new[] { "gpt-4o-mini", "gpt-4o", "gpt-4-turbo" } },
        { "claude", new[] { "claude-3-5-sonnet-20241022", "claude-3-5-haiku-20241022" } }
    };

    private static readonly string[] ProviderKeys = { "gemini", "deepseek", "openai", "claude" };
    private static readonly string[] ProviderLabels = { "Gemini (Free)", "DeepSeek (Free/Cheap)", "OpenAI (Paid)", "Claude (Paid)" };
    private static readonly string[] ProviderKeyUrls = {
        "https://makersuite.google.com/app/apikey",
        "https://platform.deepseek.com/api_keys",
        "https://platform.openai.com/api-keys",
        "https://console.anthropic.com/settings/keys"
    };

    public ApiSettingsForm()
    {
        InitializeComponents();
        LoadConfig();
    }

    private void InitializeComponents()
    {
        Text = "AI API Settings";
        Size = new Size(480, 310);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        StartPosition = FormStartPosition.CenterParent;
        Font = new Font("Segoe UI", 10);

        int xLbl = 20, xCtrl = 130, wCtrl = 310, h = 26, gap = 6, y = 20;

        var lblProvider = new Label { Text = "Provider:", Location = new Point(xLbl, y + 2), Size = new Size(100, h) };
        cmbProvider = new ComboBox { Location = new Point(xCtrl, y), Size = new Size(wCtrl, h), DropDownStyle = ComboBoxStyle.DropDownList };
        cmbProvider.Items.AddRange(ProviderLabels);
        cmbProvider.SelectedIndexChanged += (s, e) => { UpdateModels(); UpdateGetKeyLink(); };

        y += h + gap;
        lnkGetKey = new LinkLabel { Text = "Get API Key →", Location = new Point(xCtrl, y), Size = new Size(wCtrl, h), LinkBehavior = LinkBehavior.HoverUnderline, Visible = false };
        lnkGetKey.LinkClicked += (s, e) => System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(lnkGetKey.Tag?.ToString() ?? "") { UseShellExecute = true });

        y += h + gap;
        var lblKey = new Label { Text = "API Key:", Location = new Point(xLbl, y + 2), Size = new Size(100, h) };
        txtApiKey = new TextBox { Location = new Point(xCtrl, y), Size = new Size(wCtrl, h), PasswordChar = '*' };

        y += h + gap;
        var lblModel = new Label { Text = "Model:", Location = new Point(xLbl, y + 2), Size = new Size(100, h) };
        cmbModel = new ComboBox { Location = new Point(xCtrl, y), Size = new Size(wCtrl, h), DropDownStyle = ComboBoxStyle.DropDownList };

        y += h + gap + 4;
        chkShowKey = new CheckBox { Text = "Show API Key", Location = new Point(xCtrl, y), Size = new Size(150, h) };
        chkShowKey.CheckedChanged += (s, e) => { txtApiKey.PasswordChar = chkShowKey.Checked ? '\0' : '*'; };

        y += h + gap + 4;
        btnTest = new Button { Text = "Test Connection", Location = new Point(xCtrl, y), Size = new Size(140, 32) };
        btnTest.Click += BtnTest_Click;

        btnSave = new Button { Text = "Save", Location = new Point(xCtrl + 150, y), Size = new Size(100, 32) };
        btnSave.Click += BtnSave_Click;

        y += 40;
        lblStatus = new Label { Text = "", Location = new Point(xCtrl, y), Size = new Size(wCtrl, 40), ForeColor = Color.Gray };

        Controls.AddRange(new Control[] {
            lblProvider, cmbProvider, lnkGetKey, lblKey, txtApiKey, lblModel, cmbModel,
            chkShowKey, btnTest, btnSave, lblStatus
        });
    }

        private void UpdateModels()
        {
            int idx = cmbProvider.SelectedIndex;
            if (idx < 0) return;
            string key = ProviderKeys[idx];
            cmbModel.Items.Clear();
            cmbModel.Items.AddRange(Models[key]);
            cmbModel.SelectedIndex = 0;
        }

        private void UpdateGetKeyLink()
        {
            int idx = cmbProvider.SelectedIndex;
            if (idx >= 0 && idx < ProviderKeyUrls.Length)
            {
                lnkGetKey.Tag = ProviderKeyUrls[idx];
                lnkGetKey.Visible = true;
            }
            else
            {
                lnkGetKey.Visible = false;
            }
        }

    private void LoadConfig()
    {
        try
        {
            if (!File.Exists(ConfigPath)) return;
            string json = File.ReadAllText(ConfigPath);
            var ser = new System.Web.Script.Serialization.JavaScriptSerializer();
            var obj = ser.DeserializeObject(json) as Dictionary<string, object>;
            if (obj == null) return;

            string provider = GetStr(obj, "provider");
            for (int i = 0; i < ProviderKeys.Length; i++)
            {
                if (ProviderKeys[i] == provider) { cmbProvider.SelectedIndex = i; break; }
            }

            if (obj.ContainsKey(provider))
            {
                var pObj = obj[provider] as Dictionary<string, object>;
                if (pObj != null)
                {
                    txtApiKey.Text = GetStr(pObj, "apiKey") ?? "";
                    string model = GetStr(pObj, "model") ?? "";
                    UpdateModels();
                    for (int i = 0; i < cmbModel.Items.Count; i++)
                    {
                        if (cmbModel.Items[i].ToString() == model) { cmbModel.SelectedIndex = i; break; }
                    }
                }
            }
        }
        catch { }
    }

    private void BtnTest_Click(object sender, EventArgs e)
    {
        int idx = cmbProvider.SelectedIndex;
        if (idx < 0) { lblStatus.Text = "Please select a provider."; return; }
        string key = txtApiKey.Text.Trim();
        if (string.IsNullOrEmpty(key)) { lblStatus.Text = "Please enter an API key."; return; }

        lblStatus.Text = "Testing...";
        lblStatus.ForeColor = Color.Gray;
        btnTest.Enabled = false;

        string result = TestConnection(ProviderKeys[idx], key);
        if (result == "OK")
        {
            lblStatus.Text = "Connection successful!";
            lblStatus.ForeColor = Color.Green;
        }
        else
        {
            lblStatus.Text = "Failed: " + result;
            lblStatus.ForeColor = Color.Red;
        }
        btnTest.Enabled = true;
    }

    private string TestConnection(string provider, string apiKey)
    {
        try
        {
            var wc = new System.Net.WebClient();
            wc.Encoding = System.Text.Encoding.UTF8;
            wc.Headers["Content-Type"] = "application/json";

            if (provider == "gemini")
            {
                string url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + apiKey;
                string body = "{\"contents\":[{\"parts\":[{\"text\":\"Hi\"}]}]}";
                string resp = wc.UploadString(url, body);
                if (resp.Contains("\"text\"")) return "OK";
                return "Unexpected response";
            }
            else if (provider == "deepseek" || provider == "openai")
            {
                string url = provider == "deepseek" ? "https://api.deepseek.com/v1/chat/completions" : "https://api.openai.com/v1/chat/completions";
                wc.Headers["Authorization"] = "Bearer " + apiKey;
                string body = "{\"model\":\"gpt-4o-mini\",\"messages\":[{\"role\":\"user\",\"content\":\"Hi\"}],\"max_tokens\":5}";
                if (provider == "deepseek")
                    body = "{\"model\":\"deepseek-chat\",\"messages\":[{\"role\":\"user\",\"content\":\"Hi\"}],\"max_tokens\":5}";
                string resp = wc.UploadString(url, body);
                if (resp.Contains("\"content\"")) return "OK";
                return "Unexpected response";
            }
            else if (provider == "claude")
            {
                string url = "https://api.anthropic.com/v1/messages";
                wc.Headers["x-api-key"] = apiKey;
                wc.Headers["anthropic-version"] = "2023-06-01";
                string body = "{\"model\":\"claude-3-5-haiku-20241022\",\"max_tokens\":5,\"messages\":[{\"role\":\"user\",\"content\":\"Hi\"}]}";
                string resp = wc.UploadString(url, body);
                if (resp.Contains("\"text\"")) return "OK";
                return "Unexpected response";
            }
            return "Unknown provider";
        }
        catch (Exception ex)
        {
            return ex.Message;
        }
    }

    private void BtnSave_Click(object sender, EventArgs e)
    {
        int idx = cmbProvider.SelectedIndex;
        if (idx < 0) { MessageBox.Show("Please select a provider.", "Error"); return; }
        string key = txtApiKey.Text.Trim();
        if (string.IsNullOrEmpty(key)) { MessageBox.Show("Please enter an API key.", "Error"); return; }

        string provider = ProviderKeys[idx];
        string model = cmbModel.SelectedItem != null ? cmbModel.SelectedItem.ToString() : Models[provider][0];

        var config = new Dictionary<string, object>
        {
            { "provider", provider },
            { provider, new Dictionary<string, object>
                {
                    { "apiKey", key },
                    { "model", model }
                }
            }
        };

        try
        {
            string dir = Path.GetDirectoryName(ConfigPath);
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
            string json = new System.Web.Script.Serialization.JavaScriptSerializer().Serialize(config);
            File.WriteAllText(ConfigPath, json);
            DialogResult = DialogResult.OK;
            Close();
        }
        catch (Exception ex)
        {
            MessageBox.Show("Save failed: " + ex.Message, "Error");
        }
    }

    private static string GetStr(Dictionary<string, object> d, string k)
    {
        object v;
        if (d.TryGetValue(k, out v) && v != null) return v.ToString();
        return null;
    }
}
