using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Extensibility;
using Microsoft.Win32;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using stdole;



namespace WordToExeAddin
{
    [ComVisible(true)]
    [Guid("33F39564-802B-412B-A713-E947F98DCBB8")]
    [ProgId("Aksa10xFaster.Connect")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class WordAddIn : IDTExtensibility2, Office.IRibbonExtensibility
    {
        private Word.Application _app;
        private Office.IRibbonUI _ribbon;
        private LicenseService _license;

        public WordAddIn()
        {
            try { File.WriteAllText(@"C:\Users\Plabon\AppData\Local\Temp\aksadb.txt", "Constructor called\r\n"); } catch { }
        }

        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            try
            {
                try { File.AppendAllText(@"C:\Users\Plabon\AppData\Local\Temp\aksadb.txt", "OnConnection start\r\n"); } catch { }
                _app = application as Word.Application;
                try { File.AppendAllText(@"C:\Users\Plabon\AppData\Local\Temp\aksadb.txt", "OnConnection _app=" + (_app != null ? "OK" : "NULL") + "\r\n"); } catch { }
                _license = new LicenseService();
                try { File.AppendAllText(@"C:\Users\Plabon\AppData\Local\Temp\aksadb.txt", "OnConnection _license OK\r\n"); } catch { }
            }
            catch (Exception ex)
            {
                try { File.WriteAllText(@"C:\Users\Plabon\AppData\Local\Temp\aksadb.txt", "OnConnection ERROR: " + ex.ToString() + "\r\n"); } catch { }
            }
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            _app = null;
            _license = null;
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        public string OnGetEmptyLabel(Office.IRibbonControl c)
        {
            return " ";
        }

        public string GetCustomUI(string ribbonId)
        {
            try
            {
                try { File.AppendAllText(@"C:\Users\Plabon\AppData\Local\Temp\aksadb.txt", "GetCustomUI called, ribbonId=" + ribbonId + "\r\n"); } catch { }
                using (var s = Assembly.GetExecutingAssembly()
                    .GetManifestResourceStream("WordToExeAddin.WordToExeAddin_Ribbon.xml"))
                {
                    if (s == null) { try { File.AppendAllText(@"C:\Users\Plabon\AppData\Local\Temp\aksadb.txt", "GetCustomUI: resource not found\r\n"); } catch { } return ""; }
                    using (var r = new StreamReader(s))
                    {
                        string xml = r.ReadToEnd();
                        try { File.AppendAllText(@"C:\Users\Plabon\AppData\Local\Temp\aksadb.txt", "GetCustomUI OK, length=" + xml.Length + "\r\n"); } catch { }
                        return xml;
                    }
                }
            }
            catch (Exception ex)
            {
                try { File.WriteAllText(@"C:\Users\Plabon\AppData\Local\Temp\aksadb.txt", "GetCustomUI ERROR: " + ex.ToString() + "\r\n"); } catch { }
                return "";
            }
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUi) { _ribbon = ribbonUi; try { File.AppendAllText(@"C:\Users\Plabon\AppData\Local\Temp\aksadb.txt", "Ribbon_Load called\r\n"); } catch { } }

        // ===================== LICENSE =====================

        private bool IsTrialExpired()
        {
            if (_license == null) return true;
            var status = _license.GetStatus();
            return status == LicenseService.LicenseStatus.TrialExpired ||
                   status == LicenseService.LicenseStatus.LicenseExpired;
        }

        private bool CheckLicense()
        {
            if (_license == null)
                return false;
            var status = _license.GetStatus();
            if (status == LicenseService.LicenseStatus.Licensed ||
                status == LicenseService.LicenseStatus.TrialActive)
                return true;

            string msg = status == LicenseService.LicenseStatus.LicenseExpired
                ? "Your license has expired.\n\nPurchase a new license to continue using all features."
                : "Your 30-day trial has expired.\n\nPurchase a license to continue using all features.";

            if (MessageBox.Show(msg, "License Expired", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
            {
                using (var form = new RegistrationForm(_license))
                {
                    form.ShowDialog();
                }
                if (_ribbon != null) _ribbon.Invalidate();
            }
            return false;
        }

        public void OnRegister(Office.IRibbonControl c)
        {
            if (_license == null) return;
            using (var form = new RegistrationForm(_license))
            {
                form.ShowDialog();
            }
            if (_ribbon != null) _ribbon.Invalidate();
        }

        private Word.Document Doc
        {
            get
            {
                if (_app == null) { MessageBox.Show("Word application not connected.", "AKSA 10X FASTER"); return null; }
                try { return _app.ActiveDocument; }
                catch { MessageBox.Show("No document open.", "AKSA 10X FASTER"); return null; }
            }
        }

        private Word.Selection Sel
        {
            get { try { return _app.Selection; } catch { return null; } }
        }

        // ===================== PAGE GROUP =====================

        public void OnPageSetup(Office.IRibbonControl c)
        {
            _app.Dialogs[Word.WdWordDialog.wdDialogFilePageSetup].Show();
        }

        private void SetMargins(float top, float bottom, float left, float right)
        {
            var d = Doc;
            if (d == null) return;
            d.PageSetup.TopMargin = _app.CentimetersToPoints(top);
            d.PageSetup.BottomMargin = _app.CentimetersToPoints(bottom);
            d.PageSetup.LeftMargin = _app.CentimetersToPoints(left);
            d.PageSetup.RightMargin = _app.CentimetersToPoints(right);
        }

        public void OnMarginNormal(Office.IRibbonControl c) { SetMargins(2.54f, 2.54f, 2.54f, 2.54f); }
        public void OnMarginNarrow(Office.IRibbonControl c) { SetMargins(1.27f, 1.27f, 1.27f, 1.27f); }
        public void OnMarginModerate(Office.IRibbonControl c) { SetMargins(2.54f, 2.54f, 1.91f, 1.91f); }
        public void OnMarginWide(Office.IRibbonControl c) { SetMargins(2.54f, 2.54f, 5.08f, 5.08f); }
        public void OnMarginMirrored(Office.IRibbonControl c)
        {
            var d = Doc;
            if (d == null) return;
            d.PageSetup.MirrorMargins = (d.PageSetup.MirrorMargins == 0) ? -1 : 0;
        }

        public void OnOrientationPortrait(Office.IRibbonControl c)
        {
            var d = Doc; if (d == null) return;
            d.PageSetup.Orientation = Word.WdOrientation.wdOrientPortrait;
        }

        public void OnOrientationLandscape(Office.IRibbonControl c)
        {
            var d = Doc; if (d == null) return;
            d.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;
        }

        private void SetPageSize(float width, float height)
        {
            var d = Doc; if (d == null) return;
            d.PageSetup.PageWidth = _app.CentimetersToPoints(width);
            d.PageSetup.PageHeight = _app.CentimetersToPoints(height);
        }

        public void OnPageSizeA4(Office.IRibbonControl c) { SetPageSize(21f, 29.7f); }
        public void OnPageSizeLetter(Office.IRibbonControl c) { SetPageSize(21.6f, 27.9f); }
        public void OnPageSizeLegal(Office.IRibbonControl c) { SetPageSize(21.6f, 35.6f); }
        public void OnPageSizeA3(Office.IRibbonControl c) { SetPageSize(29.7f, 42f); }

        public void OnPageBreak(Office.IRibbonControl c)
        {
            var d = Doc; if (d == null) return;
            d.Content.InsertBreak(Word.WdBreakType.wdPageBreak);
        }

        public void OnSectionBreak(Office.IRibbonControl c)
        {
            var d = Doc; if (d == null) return;
            d.Content.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage);
        }

        public void OnLineNumbers(Office.IRibbonControl c)
        {
            var d = Doc; if (d == null) return;
            var ln = d.PageSetup.LineNumbering;
            if (ln.Active == 0 || ln.Active == -1)
                ln.Active = (ln.Active == 0) ? -1 : 0;
            else
                ln.Active = -1;
        }

        // ===================== PAGE PRESETS =====================

        private void SetPagePreset(float top, float bottom, float left, float right,
            string fontName, float fontSize, float pageW, float pageH)
        {
            var d = Doc;
            if (d == null) return;
            d.PageSetup.PageWidth = _app.CentimetersToPoints(pageW);
            d.PageSetup.PageHeight = _app.CentimetersToPoints(pageH);
            d.PageSetup.TopMargin = _app.InchesToPoints(top);
            d.PageSetup.BottomMargin = _app.InchesToPoints(bottom);
            d.PageSetup.LeftMargin = _app.InchesToPoints(left);
            d.PageSetup.RightMargin = _app.InchesToPoints(right);
            d.Content.Font.Name = fontName;
            d.Content.Font.Size = fontSize;
            d.Content.Font.Bold = 0;
        }

        public void OnAgreementPage(Office.IRibbonControl c)
        {
            try { SetPagePreset(5f, 1.5f, 1f, 1f, "SutonnyMJ", 16, 21.6f, 35.6f); }
            catch (System.Exception ex) { MessageBox.Show(ex.Message, "Agreements Page Error"); }
        }

        public void OnCvPage(Office.IRibbonControl c)
        {
            try { SetPagePreset(0.5f, 0.5f, 1f, 1f, "Arial", 16, 21f, 29.7f); }
            catch (System.Exception ex) { MessageBox.Show(ex.Message, "CV Page Error"); }
        }

        public void OnApplicationPage(Office.IRibbonControl c)
        {
            try { SetPagePreset(1f, 0.5f, 1f, 0.5f, "SutonnyMJ", 16, 21f, 29.7f); }
            catch (System.Exception ex) { MessageBox.Show(ex.Message, "Application Page Error"); }
        }

        public void OnQuestionPage(Office.IRibbonControl c)
        {
            try
            {
                var d = _app.ActiveDocument;
                d.PageSetup.TopMargin = _app.InchesToPoints(0.5f);
                d.PageSetup.BottomMargin = _app.InchesToPoints(0.5f);
                d.PageSetup.LeftMargin = _app.InchesToPoints(0.5f);
                d.PageSetup.RightMargin = _app.InchesToPoints(0.5f);
                d.Content.Font.Name = "SutonnyMJ";
                d.Content.Font.Size = 14;
                d.Content.Font.Bold = 0;
                d.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;
                d.PageSetup.PageWidth = _app.CentimetersToPoints(29.7f);
                d.PageSetup.PageHeight = _app.CentimetersToPoints(21f);
                d.PageSetup.TextColumns.SetCount(2);
            }
            catch (System.Exception ex) { MessageBox.Show(ex.Message, "Question Page Error"); }
        }

        public void OnNumberToWord(Office.IRibbonControl c)
        {
            try
            {
                var s = Sel; if (s == null) return;
                var text = s.Text.Trim();
                if (string.IsNullOrEmpty(text)) { MessageBox.Show("Select a number first.", "N→W"); return; }
                decimal number;
                if (!decimal.TryParse(text, out number)) { MessageBox.Show("Selected text is not a valid number.", "N→W"); return; }
                s.Text = NumberToWords(number);
            }
            catch (System.Exception ex) { MessageBox.Show(ex.Message, "N→W Error"); }
        }

        public void OnNumberToWordBn(Office.IRibbonControl c)
        {
            try
            {
                var s = Sel; if (s == null) return;
                var text = s.Text.Trim();
                if (string.IsNullOrEmpty(text)) { MessageBox.Show("Select a number first.", "BN N→W"); return; }
                text = BengaliDigitsToArabic(text);
                decimal number;
                if (!decimal.TryParse(text, out number)) { MessageBox.Show("Selected text is not a valid number.", "BN N→W"); return; }
                s.Text = NumberToWordsBangla(number);
            }
            catch (System.Exception ex) { MessageBox.Show(ex.Message, "BN N→W Error"); }
        }

        public void OnAiWrite(Office.IRibbonControl c)
        {
            try
            {
                using (var form = new Form())
                {
                    form.Text = "AI Write";
                    form.Size = new Size(440, 260);
                    form.StartPosition = FormStartPosition.CenterScreen;
                    form.FormBorderStyle = FormBorderStyle.FixedDialog;
                    form.MaximizeBox = false;
                    form.MinimizeBox = false;

                    var lbl = new Label { Text = "Describe what to write:", Left = 12, Top = 12, Width = 400 };
                    var txt = new TextBox { Left = 12, Top = 32, Width = 400, Height = 60, Multiline = true, AcceptsReturn = false };

                    var cbAgree = new CheckBox { Text = "Agreements", Left = 12, Top = 100, Width = 120, Checked = true };
                    var cbApp = new CheckBox { Text = "Application", Left = 140, Top = 100, Width = 120 };
                    var cbCustom = new CheckBox { Text = "Custom", Left = 270, Top = 100, Width = 120 };

                    var btnOk = new Button { Text = "Write", Left = 120, Top = 150, Width = 80, DialogResult = DialogResult.OK };
                    var btnCancel = new Button { Text = "Cancel", Left = 210, Top = 150, Width = 80, DialogResult = DialogResult.Cancel };
                    form.AcceptButton = btnOk;
                    form.CancelButton = btnCancel;

                    form.Controls.AddRange(new Control[] { lbl, txt, cbAgree, cbApp, cbCustom, btnOk, btnCancel });

                    if (form.ShowDialog() == DialogResult.OK)
                    {
                        string instruction = txt.Text.Trim();
                        if (string.IsNullOrEmpty(instruction)) return;

                        string systemPrompt;
                        if (cbAgree.Checked)
                            systemPrompt = "You are a professional Bengali deed writer and lawyer. Write a very concise legal agreement in Bengali that fits on 3-page stamp paper. Use blanks (______) where the user fills in details. Each clause max 2-3 lines, max 7-8 clauses total. Output ONLY the agreement in Bengali, no explanations or translations.";
                        else if (cbApp.Checked)
                            systemPrompt = "You are a professional Bengali letter/application writer. Write a formal application or letter in Bengali based on the user's request. Keep it concise and professional. Output ONLY the letter in Bengali, no explanations.";
                        else
                            systemPrompt = "You are a helpful Bengali writing assistant. Write in Bengali based on the user's request. Output ONLY the content in Bengali, no explanations.";

                        string result = CallAi(systemPrompt, instruction);
                        if (!string.IsNullOrEmpty(result))
                            InsertWithTypingEffect(result);
                    }
                }
            }
            catch (System.Exception ex) { File.WriteAllText("C:\\Users\\apnan\\AppData\\Local\\Temp\\ai_error.txt", ex.ToString()); MessageBox.Show("Error: " + ex.Message, "AI Write Error"); }
        }

        public void OnAiFix(Office.IRibbonControl c)
        {
            try
            {
                var s = Sel; if (s == null) return;
                var text = s.Text.Trim();
                if (string.IsNullOrEmpty(text)) { MessageBox.Show("Select some text first.", "AI Fix"); return; }
                var unicode = ConvertBijoyToUnicode(text);
                string result = CallAi(
                    "You are a Bengali grammar expert. Fix spelling and grammar in the following Bengali text. Output only the corrected text in Bengali, no explanations.",
                    unicode);
                if (!string.IsNullOrEmpty(result))
                    InsertWithTypingEffect(result);
            }
            catch (System.Exception ex) { MessageBox.Show("Error: " + ex.Message, "AI Fix Error"); }
        }

        public void OnAiSummarize(Office.IRibbonControl c)
        {
            try
            {
                var s = Sel; if (s == null) return;
                var text = s.Text.Trim();
                if (string.IsNullOrEmpty(text)) { MessageBox.Show("Select some text first.", "AI Summarize"); return; }
                var unicode = ConvertBijoyToUnicode(text);
                string result = CallAi(
                    "You are a Bengali summarizer. Write a concise summary (3-4 lines) of the following Bengali text. Output only the summary in Bengali, no explanations.",
                    unicode);
                if (!string.IsNullOrEmpty(result))
                    InsertWithTypingEffect(result);
            }
            catch (System.Exception ex) { MessageBox.Show("Error: " + ex.Message, "AI Summarize Error"); }
        }

        public void OnAboutMe(Office.IRibbonControl c)
        {
            var form = new Form();
            form.Text = "About Me";
            form.ClientSize = new Size(540, 420);
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.MaximizeBox = false;
            form.MinimizeBox = false;
            form.StartPosition = FormStartPosition.CenterScreen;
            form.BackColor = Color.White;

            var pic = new PictureBox();
            pic.Size = new Size(120, 145);
            pic.Location = new Point(20, 25);
            pic.SizeMode = PictureBoxSizeMode.Zoom;
            string photoPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "Aksa10xFaster", "Icons", "Apnan Ahmed (Plabon).png");
            if (File.Exists(photoPath))
                pic.Image = Image.FromFile(photoPath);
            else
                pic.BackColor = Color.SteelBlue;
            form.Controls.Add(pic);

            var lblName = new Label();
            lblName.Text = "Apnan Ahmed (Plabon)";
            lblName.Font = new System.Drawing.Font("Arial", 14, FontStyle.Bold);
            lblName.Location = new Point(160, 25);
            lblName.Size = new Size(350, 30);
            form.Controls.Add(lblName);

            var lblTitle = new Label();
            lblTitle.Text = "Software Engineer & Developer";
            lblTitle.Font = new System.Drawing.Font("Arial", 10);
            lblTitle.Location = new Point(160, 58);
            lblTitle.Size = new Size(350, 20);
            form.Controls.Add(lblTitle);

            var lblLoc = new Label();
            lblLoc.Text = "Bangladesh";
            lblLoc.Font = new System.Drawing.Font("Arial", 10);
            lblLoc.Location = new Point(160, 80);
            lblLoc.Size = new Size(350, 20);
            form.Controls.Add(lblLoc);

            var lblPhone = new Label();
            lblPhone.Text = "+880 1670-201266";
            lblPhone.Font = new System.Drawing.Font("Arial", 9);
            lblPhone.Location = new Point(160, 103);
            lblPhone.Size = new Size(350, 18);
            form.Controls.Add(lblPhone);

            var lblEmail = new Label();
            lblEmail.Text = "apnan2010@gmail.com";
            lblEmail.Font = new System.Drawing.Font("Arial", 9);
            lblEmail.Location = new Point(160, 122);
            lblEmail.Size = new Size(350, 18);
            form.Controls.Add(lblEmail);

            var sep = new Label();
            sep.Text = "─────────────────────────────────";
            sep.Location = new Point(20, 192);
            sep.Size = new Size(490, 20);
            form.Controls.Add(sep);

            var lblApp = new Label();
            lblApp.Text = "AKSA 10X FASTER  v1.0";
            lblApp.Font = new System.Drawing.Font("Arial", 10, FontStyle.Italic);
            lblApp.Location = new Point(20, 215);
            lblApp.Size = new Size(490, 20);
            form.Controls.Add(lblApp);

            var lblSocial = new Label();
            lblSocial.Text = "Follow me on";
            lblSocial.Font = new System.Drawing.Font("Arial", 9, FontStyle.Bold);
            lblSocial.Location = new Point(20, 240);
            lblSocial.Size = new Size(490, 20);
            form.Controls.Add(lblSocial);

            var socialPlatforms = new (string id, string name, Color color)[]
            {
                ("whatsapp", "WhatsApp", Color.FromArgb(0x25, 0xD3, 0x66)),
                ("facebook", "Facebook", Color.FromArgb(0x18, 0x78, 0xF2)),
                ("messenger", "Messenger", Color.FromArgb(0x00, 0x7F, 0xFF)),
                ("instagram", "Instagram", Color.FromArgb(0xE4, 0x40, 0x5F)),
                ("threads", "Threads", Color.FromArgb(0x00, 0x00, 0x00)),
                ("pinterest", "Pinterest", Color.FromArgb(0xBD, 0x08, 0x1C)),
                ("telegram", "Telegram", Color.FromArgb(0x00, 0x86, 0xCB)),
                ("imo", "Imo", Color.FromArgb(0x5D, 0xB9, 0xF4)),
                ("twitter", "Twitter", Color.FromArgb(0x1D, 0xA1, 0xF2)),
                ("linkedin", "LinkedIn", Color.FromArgb(0x0A, 0x66, 0xC2)),
            };

            var socialUrls = new Dictionary<string, string>
            {
                { "facebook", "https://www.facebook.com/A.A.Plabon" },
                { "threads", "https://www.threads.net/@a.a.plabon" },
                { "instagram", "https://www.instagram.com/a.a.plabon?igsh=MTNmZ2x0NXA2aDYzYw==" },
                { "pinterest", "https://pin.it/3GkdSkqlX" },
                { "imo", "https://s.imoim.net/wl6MPC" },
            };

            int btnSize = 40;
            int gap = 6;
            int startX = 20;
            int startY = 265;
            int rowCount = 0;

            foreach (var (id, name, color) in socialPlatforms)
            {
                int x = startX + (btnSize + gap) * (rowCount % 9);
                int y = startY + (btnSize + gap) * (rowCount / 9);

                var btn = new PictureBox();
                btn.Size = new Size(btnSize, btnSize);
                btn.Location = new Point(x, y);
                btn.SizeMode = PictureBoxSizeMode.Zoom;
                btn.Cursor = Cursors.Hand;

                string asmDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string[] iconDirs =
                {
                    Path.Combine(asmDir, "Icons"),
                    Path.Combine(asmDir, "icon"),
                    Path.Combine(Directory.GetParent(asmDir).FullName, "icon"),
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Aksa10xFaster", "Icons"),
                };
                string[] tryNames = { id + ".png", name + ".png", name.ToLower() + ".png" };
                string found = null;
                foreach (var dir in iconDirs)
                    foreach (var fn in tryNames)
                    {
                        string fp = Path.Combine(dir, fn);
                        if (File.Exists(fp)) { found = fp; break; }
                    }
                if (found != null)
                {
                    btn.Image = Image.FromFile(found);
                }
                else
                {
                    var bmp = new Bitmap(btnSize, btnSize);
                    using (var g = Graphics.FromImage(bmp))
                    {
                        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                        g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
                        g.Clear(color);
                        using (var f = new System.Drawing.Font("Segoe UI", 12f, FontStyle.Bold))
                        using (var tb = new SolidBrush(Color.White))
                        {
                            var ms = g.MeasureString(name[0].ToString(), f);
                            g.DrawString(name[0].ToString(), f, tb, (btnSize - ms.Width) / 2f, (btnSize - ms.Height) / 2f);
                        }
                    }
                    btn.Image = bmp;
                    btn.BackColor = color;
                }

                string url;
                socialUrls.TryGetValue(id, out url);
                if (url != null)
                {
                    btn.Click += (s, e) => Process.Start(url);
                }

                var tooltip = new ToolTip();
                tooltip.SetToolTip(btn, name);

                form.Controls.Add(btn);
                rowCount++;
            }

            var btnOk = new Button();
            btnOk.Text = "OK";
            btnOk.Location = new Point(230, 360);
            btnOk.Size = new Size(80, 28);
            btnOk.DialogResult = DialogResult.OK;
            form.Controls.Add(btnOk);

            form.AcceptButton = btnOk;
            form.ShowDialog();
            if (pic.Image != null) pic.Image.Dispose();
            form.Dispose();
        }

        public void OnApiSettings(Office.IRibbonControl c)
        {
            using (var form = new ApiSettingsForm())
                form.ShowDialog();
        }

        public void OnInsertPageNumber(Office.IRibbonControl c)
        {
            try
            {
                var doc = _app.ActiveDocument;
                if (doc == null) return;
                int pageCount = doc.ComputeStatistics(Word.WdStatistic.wdStatisticPages);
                if (pageCount <= 1)
                {
                    MessageBox.Show("Page numbers are added when the document has multiple pages.", "Page Number");
                    return;
                }
                bool isBangla = c.Id == "btnPageNumBn";
                object missing = System.Reflection.Missing.Value;
                object collapseEnd = Word.WdCollapseDirection.wdCollapseEnd;

                foreach (Word.Section section in doc.Sections)
                {
                    var footer = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                    footer.LinkToPrevious = false;
                    string prefix = isBangla ? "পাতা- " : "Page- ";
                    string separator = "/";

                    Word.Range rng = footer.Range;
                    rng.Find.ClearFormatting();
                    rng.Find.Text = ".*";
                    rng.Find.MatchWildcards = true;
                    rng.Find.Execute();
                    if (rng.Find.Found) rng.Text = "";
                    rng = footer.Range;
                    rng.ParagraphFormat.SpaceBefore = 0;
                    rng.ParagraphFormat.SpaceAfter = 0;
                    rng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                    rng.Text = prefix + "<pg>" + separator + "<np>";
                    if (isBangla) try { rng.Font.Name = "SolaimanLipi"; } catch { }

                    rng = footer.Range;
                    rng.Find.ClearFormatting();
                    rng.Find.Text = "<pg>";
                    rng.Find.Execute();
                    if (rng.Find.Found) rng.Fields.Add(rng, Word.WdFieldType.wdFieldPage, ref missing, false);

                    rng = footer.Range;
                    rng.Find.Text = "<np>";
                    rng.Find.Execute();
                    if (rng.Find.Found) rng.Fields.Add(rng, Word.WdFieldType.wdFieldNumPages, ref missing, false);
                }
            }
            catch (System.Exception ex) { MessageBox.Show(ex.Message, "Page Number Error"); }
        }

        public void OnShare(Office.IRibbonControl c)
        {
            try
            {
                var doc = _app.ActiveDocument;
                if (doc == null) return;
                doc.Save();
                string docName = Uri.EscapeDataString(doc.Name);
                string docPath = Uri.EscapeDataString(doc.FullName);
                switch (c.Id)
                {
                    case "btnShareEmail":
                        Process.Start("mailto:?subject=" + docName);
                        break;
                    case "btnShareWhatsapp":
                        Process.Start("https://wa.me/?text=" + docName);
                        break;
                    case "btnShareImo":
                        Process.Start("https://imo.im/?text=" + docName);
                        break;
                    case "btnShareTelegram":
                        Process.Start("https://t.me/share/url?url=" + docPath + "&text=" + docName);
                        break;
                }
            }
            catch (System.Exception ex) { MessageBox.Show(ex.Message, "Share Error"); }
        }

        private static string NumberToWords(decimal number)
        {
            if (number == 0) return "Zero";
            var parts = number.ToString("F2").Split('.');
            var intPart = long.Parse(parts[0]);
            var decPart = parts.Length > 1 && parts[1] != "00" ? int.Parse(parts[1].TrimEnd('0')) : 0;
            var result = IntToWords(intPart).Trim();
            if (decPart > 0)
                result += " point " + IntToWords(decPart).Trim();
            return result;
        }

        private static string IntToWords(long n)
        {
            if (n == 0) return "Zero";
            if (n < 0) return "Minus " + IntToWords(-n);
            string[] ones = { "", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine",
                             "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen",
                             "Seventeen", "Eighteen", "Nineteen" };
            string[] tens = { "", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety" };
            string[] groups = { "", "Thousand", "Million", "Billion", "Trillion" };
            var result = "";
            long[] groupVals = { 0, 0, 0, 0, 0 };
            long remaining = n;
            for (int i = 0; i < 5 && remaining > 0; i++)
            {
                groupVals[i] = remaining % 1000;
                remaining /= 1000;
            }
            for (int i = 4; i >= 0; i--)
            {
                if (groupVals[i] == 0) continue;
                var g = groupVals[i];
                if (g >= 100) { result += " " + ones[g / 100] + " Hundred"; g %= 100; }
                if (g >= 20) { result += " " + tens[g / 10]; g %= 10; if (g > 0) result += "-" + ones[g]; }
                else if (g > 0) result += " " + ones[g];
                if (!string.IsNullOrEmpty(groups[i])) result += " " + groups[i];
            }
            return result.Trim();
        }

        private static string NumberToWordsBangla(decimal number)
        {
            if (number == 0) return "শূন্য";
            var parts = number.ToString("F2").Split('.');
            var intPart = long.Parse(parts[0]);
            var decPart = parts.Length > 1 && parts[1] != "00" ? int.Parse(parts[1].TrimEnd('0')) : 0;
            var result = IntToWordsBangla(intPart).Trim();
            if (decPart > 0)
                result += " দশমিক " + DigitsToBangla(decPart);
            return result;
        }

        private static string IntToWordsBangla(long n)
        {
            if (n == 0) return "শূন্য";
            if (n < 0) return "ঋণাত্মক " + IntToWordsBangla(-n);
            string result = "";
            long crore = n / 10000000;
            if (crore > 0) { result += NumberBelow100Bangla(crore) + " কোটি "; n %= 10000000; }
            long lakh = n / 100000;
            if (lakh > 0) { result += NumberBelow100Bangla(lakh) + " লক্ষ "; n %= 100000; }
            long thousand = n / 1000;
            if (thousand > 0) { result += NumberBelow100Bangla(thousand) + " হাজার "; n %= 1000; }
            long hundred = n / 100;
            if (hundred > 0) { result += NumberBelow100Bangla(hundred) + "শ "; n %= 100; }
            if (n > 0) result += NumberBelow100Bangla(n);
            return result.Trim();
        }

        private static string NumberBelow100Bangla(long n)
        {
            string[] words = {
                "", "এক", "দুই", "তিন", "চার", "পাঁচ", "ছয়", "সাত", "আট", "নয়",
                "দশ", "এগারো", "বারো", "তেরো", "চৌদ্দ", "পনেরো", "ষোল", "সতেরো", "আঠারো", "উনিশ",
                "বিশ", "একুশ", "বাইশ", "তেইশ", "চব্বিশ", "পঁচিশ", "ছাব্বিশ", "সাতাশ", "আঠাশ", "উনত্রিশ",
                "ত্রিশ", "একত্রিশ", "বত্রিশ", "তেত্রিশ", "চৌত্রিশ", "পঁয়ত্রিশ", "ছত্রিশ", "সাঁইত্রিশ", "আটত্রিশ", "উনচল্লিশ",
                "চল্লিশ", "একচল্লিশ", "বিয়াল্লিশ", "তেতাল্লিশ", "চুয়াল্লিশ", "পঁয়তাল্লিশ", "ছেচাল্লিশ", "সাতচল্লিশ", "আটচল্লিশ", "উনপঞ্চাশ",
                "পঞ্চাশ", "একান্ন", "বাহান্ন", "তিপ্পান্ন", "চুয়ান্ন", "পঞ্চান্ন", "ছাপ্পান্ন", "সাতান্ন", "আটান্ন", "উনষাট",
                "ষাট", "একষট্টি", "বাষট্টি", "তেষট্টি", "চৌষট্টি", "পঁয়ষট্টি", "ছেষট্টি", "সাতষট্টি", "আটষট্টি", "উনসত্তর",
                "সত্তর", "একাত্তর", "বাহাত্তর", "তিয়াত্তর", "চুয়াত্তর", "পঁচাত্তর", "ছিয়াত্তর", "সাতাত্তর", "আটাত্তর", "উনআশি",
                "আশি", "একাশি", "বিরাশি", "তিরাশি", "চুরাশি", "পঁচাশি", "ছিয়াশি", "সাতাশি", "আটাশি", "উননব্বই",
                "নব্বই", "একানব্বই", "বিরানব্বই", "তিরানব্বই", "চুরানব্বই", "পঁচানব্বই", "ছিয়ানব্বই", "সাতানব্বই", "আটানব্বই", "নিরানব্বই"
            };
            if (n >= 0 && n <= 99) return words[n];
            return "";
        }

        private static string DigitsToBangla(long n)
        {
            string s = n.ToString();
            string[] bd = { "০", "১", "২", "৩", "৪", "৫", "৬", "৭", "৮", "৯" };
            var sb = new System.Text.StringBuilder();
            foreach (char c in s) sb.Append(bd[c - '0']);
            return sb.ToString();
        }

        private static string BengaliDigitsToArabic(string text)
        {
            string[] bd = { "০", "১", "২", "৩", "৪", "৫", "৬", "৭", "৮", "৯" };
            var sb = new System.Text.StringBuilder();
            foreach (char c in text)
            {
                int idx = -1;
                for (int k = 0; k < 10; k++) { if (c == bd[k][0]) { idx = k; break; } }
                sb.Append(idx >= 0 ? (char)('0' + idx) : c);
            }
            return sb.ToString();
        }

        private static string ArabicToBengaliDigits(int number)
        {
            string s = number.ToString();
            char[] result = new char[s.Length];
            for (int i = 0; i < s.Length; i++)
            {
                if (s[i] >= '0' && s[i] <= '9')
                    result[i] = (char)(s[i] - '0' + 0x09E6);
                else
                    result[i] = s[i];
            }
            return new string(result);
        }

        // SutonnyMJ/Bijoy52 ASCII → Unicode mapping (Kontho Keyboard standard based)
        private static readonly Dictionary<string, string> _bijoyToUnicode = new Dictionary<string, string>()
        {
            // Vowels
            { "Av", "\u0986" }, { "A", "\u0985" },
            { "B", "\u0987" }, { "C", "\u0988" },
            { "D", "\u0989" }, { "E", "\u098A" },
            { "F", "\u098B" }, { "G", "\u098F" },
            { "H", "\u0990" }, { "I", "\u0993" },
            { "J", "\u0994" },
            // Consonants
            { "K", "\u0995" }, { "L", "\u0996" },
            { "M", "\u0997" }, { "N", "\u0998" },
            { "O", "\u0999" }, { "P", "\u099A" },
            { "Q", "\u099B" }, { "R", "\u099C" },
            { "S", "\u099D" }, { "T", "\u099E" },
            { "U", "\u099F" }, { "V", "\u09A0" },
            { "W", "\u09A1" }, { "X", "\u09A2" },
            { "Y", "\u09A3" }, { "Z", "\u09A4" },
            { "_", "\u09A5" }, { "`", "\u09A6" },
            { "a", "\u09A7" }, { "b", "\u09A8" },
            { "c", "\u09AA" }, { "d", "\u09AB" },
            { "e", "\u09AC" }, { "f", "\u09AD" },
            { "g", "\u09AE" }, { "h", "\u09AF" },
            { "i", "\u09B0" }, { "j", "\u09B2" },
            { "k", "\u09B6" }, { "l", "\u09B7" },
            { "m", "\u09B8" }, { "n", "\u09B9" },
            { "o", "\u09A1\u09BC" }, { "p", "\u09A2\u09BC" },
            { "q", "\u09AF\u09BC" }, { "r", "\u09CE" },
            { "s", "\u0982" }, { "t", "\u0983" },
            { "u", "\u0981" },
            // Digits
            { "0", "\u09E6" }, { "1", "\u09E7" },
            { "2", "\u09E8" }, { "3", "\u09E9" },
            { "4", "\u09EA" }, { "5", "\u09EB" },
            { "6", "\u09EC" }, { "7", "\u09ED" },
            { "8", "\u09EE" }, { "9", "\u09EF" },
            // Kars (vowel signs)
            { "v", "\u09BE" }, { "w", "\u09BF" },
            { "x", "\u09C0" }, { "y", "\u09C1" },
            { "z", "\u09C1" }, { "~", "\u09C2" },
            // Special
            { "|", "\u0964" }, { "&", "\u09CD" },
            { "^", "\u09CD\u09AC" },
            // Extended ASCII (0x80-0xFF) per SutonnyMJ standard
            { "\x80", "\u0985" }, { "\x81", "\u0986" },
            { "\x82", "\u0987" }, { "\x83", "\u0988" },
            { "\x84", "\u09C3" }, { "\x85", "\u098A" }, { "\x86", "\u09C7" },
            { "\x87", "\u098F" }, { "\x88", "\u0990" },
            { "\x89", "\u0993" }, { "\x8A", "\u09D7" },
            { "\x8B", "\u0995" }, { "\x8C", "\u0996" },
            { "\x8D", "\u0997" }, { "\x8E", "\u0998" },
            { "\x8F", "\u0999" }, { "\x90", "\u099A" },
            { "\x91", "\u099B" }, { "\x92", "\u099C" },
            { "\x93", "\u099D" }, { "\x94", "\u099A\u09CD" },
            { "\x95", "\u0999\u09CD" }, { "\x96", "\u09A0" },
            { "\x97", "\u09A1" }, { "\x98", "\u09A2" },
            { "\x99", "\u09A3" }, { "\x9A", "\u09A4" },
            { "\x9B", "\u09A5" }, { "\x9C", "\u09A6" },
            { "\x9D", "\u09A7" }, { "\x9E", "\u09A8" },
            { "\x9F", "\u09AA" }, { "\xA0", "\u09AB" },
            { "\xA1", "\u09CD\u09AC" }, { "\xA2", "\u09CD\u09AD" },
            { "\xA3", "\u09AD\u09CD\u09B0" }, { "\xA4", "\u09AE\u09CD" },
            { "\xA5", "\u09CD\u09AE" }, { "\xA6", "\u09CD\u09AC" },
            { "\xA7", "\u09CD\u09AE" }, { "\xA8", "\u09CD\u09AF" },
            { "\xA9", "\u09B0\u09CD" }, { "\xAA", "\u09CD\u09B0" },
            { "\xAB", "\u09CD\u09B0" }, { "\xAC", "\u09CD\u09B2" },
            { "\xAD", "\u09CD\u09B2" }, { "\xAE", "\u09B7\u09CD" },
            { "\xAF", "\u09B8\u09CD" }, { "\xB0", "\u0995\u09CD\u0995" },
            { "\xB1", "\u0995\u09CD\u099F" }, { "\xB2", "\u0995\u09CD\u09B7\u09CD\u09AE" },
            { "\xB3", "\u0995\u09CD\u09A4" }, { "\xB4", "\u0995\u09CD\u09AE" },
            { "\xB5", "\u0995\u09CD\u09B0" }, { "\xB6", "\u0995\u09CD\u09B7" },
            { "\xB7", "\u0995\u09CD\u09B8" }, { "\xB8", "\u0997\u09C1" },
            { "\xB9", "\u0997\u09CD\u0997" }, { "\xBA", "\u0997\u09CD\u09A6" },
            { "\xBB", "\u0997\u09CD\u09A7" }, { "\xBC", "\u0999\u09CD\u0995" },
            { "\xBD", "\u0999\u09CD\u0997" }, { "\xBE", "\u099C\u09CD\u099C" },
            { "\xBF", "\u09A4\u09CD\u09B0" }, { "\xC0", "\u099C\u09CD\u099D" },
            { "\xC1", "\u099C\u09CD\u099E" }, { "\xC2", "\u099E\u09CD\u099A" },
            { "\xC3", "\u099E\u09CD\u099B" }, { "\xC4", "\u099E\u09CD\u099C" },
            { "\xC5", "\u099E\u09CD\u099D" }, { "\xC6", "\u099F\u09CD\u099F" },
            { "\xC7", "\u09A1\u09CD\u09A1" }, { "\xC8", "\u09A3\u09CD\u099F" },
            { "\xC9", "\u09A3\u09CD\u09A0" }, { "\xCA", "\u09A3\u09CD\u09A1" },
            { "\xCB", "\u09A4\u09CD\u09A4" }, { "\xCC", "\u09A4\u09CD\u09A5" },
            { "\xCD", "\u09A4\u09CD\u09AE" }, { "\xCE", "\u09A4\u09CD\u09B0" },
            { "\xCF", "\u09A6\u09CD\u09A6" }, { "\xD0", "\u002D" },
            { "\xD1", "\u002D" }, { "\xD2", "\u201C" },
            { "\xD3", "\u201D" }, { "\xD4", "\u2018" },
            { "\xD5", "\u2019" }, { "\xD6", "\u09CD\u09B0" },
            { "\xD7", "\u09A6\u09CD\u09A7" }, { "\xD8", "\u09A6\u09CD\u09AC" },
            { "\xD9", "\u09A6\u09CD\u09AE" }, { "\xDA", "\u09A8\u09CD\u09A0" },
            { "\xDB", "\u09A8\u09CD\u09A1" }, { "\xDC", "\u09A8\u09CD\u09A7" },
            { "\xDD", "\u09A8\u09CD\u09B8" }, { "\xDE", "\u09AA\u09CD\u099F" },
            { "\xDF", "\u09AA\u09CD\u09A4" }, { "\xE0", "\u09AA\u09CD\u09AA" },
            { "\xE1", "\u09AA\u09CD\u09B8" }, { "\xE2", "\u09AC\u09CD\u099C" },
            { "\xE3", "\u09AC\u09CD\u09A6" }, { "\xE4", "\u09AC\u09CD\u09A7" },
            { "\xE5", "\u09AD\u09CD\u09B0" }, { "\xE6", "\u09C1" },
            { "\xE7", "\u09AE\u09CD\u09AB" }, { "\xE8", "\u09CD\u09A8" },
            { "\xE9", "\u09B2\u09CD\u0995" }, { "\xEA", "\u09B2\u09CD\u0997" },
            { "\xEB", "\u09B2\u09CD\u099F" }, { "\xEC", "\u09B2\u09CD\u09A1" },
            { "\xED", "\u09B2\u09CD\u09AA" }, { "\xEE", "\u09B2\u09CD\u09AB" },
            { "\xEF", "\u09B6\u09C1" }, { "\xF0", "\u09B6\u09CD\u099A" },
            { "\xF1", "\u09B6\u09CD\u099B" }, { "\xF2", "\u09B7\u09CD\u09A3" },
            { "\xF3", "\u09B7\u09CD\u099F" }, { "\xF4", "\u09B7\u09CD\u09A0" },
            { "\xF5", "\u09B7\u09CD\u09AB" }, { "\xF6", "\u09B8\u09CD\u0996" },
            { "\xF7", "\u09B8\u09CD\u099F" }, { "\xF8", "\u09CD\u09B2" },
            { "\xF9", "\u09B8\u09CD\u09AB" }, { "\xFA", "\u09CD\u09AA" },
            { "\xFB", "\u09B9\u09C1" }, { "\xFC", "\u09B9\u09C3" },
            { "\xFD", "\u09B9\u09CD\u09A8" }, { "\xFE", "\u09B9\u09CD\u09AE" },
            { "\xFF", "\u0995\u09CD\u09B7" },
            // Multi-byte conjunct sequences (longest match first)
            // 4-byte conjunct sequences
            { "\x4C\xAA\xA8\xA9", "\u09B0\u09CD\u0996\u09CD\u09B0\u09CD\u09AF" }, { "\x4D\x9C\xA8\x79", "\u0997\u09CD\u09A8\u09CD\u09AF\u09C1" },
            { "\x4D\xD6\xA8\xA9", "\u09B0\u09CD\u0997\u09CD\u09B0\u09CD\u09AF" }, { "\x4E\xAA\xA8\xA9", "\u09B0\u09CD\u0998\u09CD\u09B0\u09CD\u09AF" },
            { "\x50\xAA\xA8\xA9", "\u09B0\u09CD\u099A\u09CD\u09B0\u09CD\u09AF" }, { "\x51\xAB\xA8\xA9", "\u09B0\u09CD\u099B\u09CD\u09B0\u09CD\u09AF" },
            { "\x52\xAA\xA8\xA9", "\u09B0\u09CD\u099C\u09CD\u09B0\u09CD\u09AF" }, { "\x53\xAA\xA8\xA9", "\u09B0\u09CD\u099D\u09CD\u09B0\u09CD\u09AF" },
            { "\x54\xAA\xA8\xA9", "\u09B0\u09CD\u099E\u09CD\u09B0\u09CD\u09AF" }, { "\x55\xAA\xA8\xA9", "\u09B0\u09CD\u099F\u09CD\u09B0\u09CD\u09AF" },
            { "\x56\xAA\xA8\xA9", "\u09B0\u09CD\u09A0\u09CD\u09B0\u09CD\u09AF" }, { "\x57\xAA\xA8\xA9", "\u09B0\u09CD\u09A1\u09CD\u09B0\u09CD\u09AF" },
            { "\x58\xAA\xA8\xA9", "\u09B0\u09CD\u09A2\u09CD\u09B0\u09CD\u09AF" }, { "\x59\xAA\xA8\xA9", "\u09B0\u09CD\u09A3\u09CD\u09B0\u09CD\u09AF" },
            { "\x5F\xAA\xA8\xA9", "\u09B0\u09CD\u09A5\u09CD\u09B0\u09CD\u09AF" }, { "\x60\xAA\xA8\xA9", "\u09B0\u09CD\u09A6\u09CD\u09B0\u09CD\u09AF" },
            { "\x61\xAA\xA8\xA9", "\u09B0\u09CD\u09A7\u09CD\u09B0\u09CD\u09AF" }, { "\x62\xAA\xA8\xA9", "\u09B0\u09CD\u09A8\u09CD\u09B0\u09CD\u09AF" },
            { "\x63\xD6\xA8\xA9", "\u09B0\u09CD\u09AA\u09CD\u09B0\u09CD\u09AF" }, { "\x64\xAB\xA8\xA9", "\u09B0\u09CD\u09AB\u09CD\u09B0\u09CD\u09AF" },
            { "\x65\xAA\xA8\xA9", "\u09B0\u09CD\u09AC\u09CD\u09B0\u09CD\u09AF" }, { "\x66\xAA\xA8\xA9", "\u09B0\u09CD\u09AD\u09CD\u09B0\u09CD\u09AF" },
            { "\x67\xAA\xA8\xA9", "\u09B0\u09CD\u09AE\u09CD\u09B0\u09CD\u09AF" }, { "\x68\xAA\xA8\xA9", "\u09B0\u09CD\u09AF\u09CD\u09B0\u09CD\u09AF" },
            { "\x69\xAA\xA8\xA9", "\u09B0\u09CD\u09B0\u09CD\u09B0\u09CD\u09AF" }, { "\x6A\xAA\xA8\xA9", "\u09B0\u09CD\u09B2\u09CD\u09B0\u09CD\u09AF" },
            { "\x6B\xD6\xA8\xA9", "\u09B0\u09CD\u09B6\u09CD\u09B0\u09CD\u09AF" }, { "\x6C\xAA\xA8\xA9", "\u09B0\u09CD\u09B7\u09CD\u09B0\u09CD\u09AF" },
            { "\x6D\xAA\xA8\xA9", "\u09B0\u09CD\u09B8\u09CD\u09B0\u09CD\u09AF" }, { "\x6E\xAB\xA8\xA9", "\u09B0\u09CD\u09B9\u09CD\u09B0\u09CD\u09AF" },
            { "\x6F\xA9\xAA\xA8", "\u09B0\u09CD\u09DC\u09CD\u09B0\u09CD\u09AF" }, { "\x9A\x97\xAA\xA8", "\u09A8\u09CD\u09A4\u09CD\u09B0\u09CD\u09AF" },
            { "\x9B\x60\xAA\x83", "\u09A8\u09CD\u09A6\u09CD\u09B0\u09C2" }, { "\x9B\x60\xAA\x93", "\u09A8\u09CD\u09A6\u09CD\u09B0\u09C1" },
            { "\xA4\xFA\xD6\x7E", "\u09AE\u09CD\u09AA\u09CD\u09B0\u09C2" }, { "\xAE\xFA\xD6\x79", "\u09B7\u09CD\u09AA\u09CD\u09B0\u09C1" },
            { "\xAE\xFA\xD6\x7E", "\u09B7\u09CD\u09AA\u09CD\u09B0\u09C2" }, { "\xAF\xFA\xD6\x7E", "\u09B8\u09CD\u09AA\u09CD\u09B0\u09C2" },
            // 3-byte conjunct sequences
            { "\x4B\xA9\x82", "\u0995\u09B0\u09CD\u09C2" }, { "\x4B\xAC\x79", "\u0995\u09CD\u09B2\u09C1" },
            { "\x4B\xAC\x7E", "\u0995\u09CD\u09B2\u09C2" }, { "\x4B\xE8\x79", "\u0995\u09CD\u09A8\u09C1" },
            { "\x4B\xE8\x7E", "\u0995\u09CD\u09A8\u09C2" }, { "\x4B\xE8\x84", "\u0995\u09CD\u09A8\u09C3" },
            { "\x4C\xAA\x79", "\u0996\u09CD\u09B0\u09C1" }, { "\x4C\xAA\x7E", "\u0996\u09CD\u09B0\u09C2" },
            { "\x4C\xAA\xA8", "\u0996\u09CD\u09B0\u09CD\u09AF" }, { "\x4C\xAA\xA9", "\u09B0\u09CD\u0996\u09CD\u09B0" },
            { "\x4D\x9C\x79", "\u0997\u09CD\u09A8\u09C1" }, { "\x4D\xA8\xA9", "\u09B0\u09CD\u0997\u09CD\u09AF" },
            { "\x4D\xD6\x83", "\u0997\u09CD\u09B0\u09C2" }, { "\x4D\xD6\x93", "\u0997\u09CD\u09B0\u09C1" },
            { "\x4D\xD6\xA8", "\u0997\u09CD\u09B0\u09CD\u09AF" }, { "\x4D\xD6\xA9", "\u09B0\u09CD\u0997\u09CD\u09B0" },
            { "\x4D\xE8\x79", "\u0997\u09CD\u09A3\u09C1" }, { "\x4E\xAA\x79", "\u0998\u09CD\u09B0\u09C1" },
            { "\x4E\xAA\x7E", "\u0998\u09CD\u09B0\u09C2" }, { "\x4E\xAA\xA8", "\u0998\u09CD\u09B0\u09CD\u09AF" },
            { "\x4E\xAA\xA9", "\u09B0\u09CD\u0998\u09CD\u09B0" }, { "\x4F\xAA\xA8", "\u0999\u09CD\u09B0\u09CD\u09AF" },
            { "\x4F\xAA\xA9", "\u09B0\u09CD\u0999\u09CD\u09B0" }, { "\x50\xA9\x82", "\u099A\u09B0\u09CD\u09C2" },
            { "\x50\xA9\x85", "\u099A\u09B0\u09CD\u09C3" }, { "\x50\xAA\xA8", "\u099A\u09CD\u09B0\u09CD\u09AF" },
            { "\x50\xAA\xA9", "\u09B0\u09CD\u099A\u09CD\u09B0" }, { "\x51\xA9\x82", "\u099B\u09B0\u09CD\u09C2" },
            { "\x51\xAB\xA8", "\u099B\u09CD\u09B0\u09CD\u09AF" }, { "\x51\xAB\xA9", "\u09B0\u09CD\u099B\u09CD\u09B0" },
            { "\x52\xAA\xA8", "\u099C\u09CD\u09B0\u09CD\u09AF" }, { "\x52\xAA\xA9", "\u09B0\u09CD\u099C\u09CD\u09B0" },
            { "\x53\xA9\x82", "\u099D\u09B0\u09CD\u09C2" }, { "\x53\xA9\x85", "\u099D\u09B0\u09CD\u09C3" },
            { "\x53\xAA\xA8", "\u099D\u09CD\u09B0\u09CD\u09AF" }, { "\x53\xAA\xA9", "\u09B0\u09CD\u099D\u09CD\u09B0" },
            { "\x54\xA9\x82", "\u099E\u09B0\u09CD\u09C2" }, { "\x54\xAA\xA8", "\u099E\u09CD\u09B0\u09CD\u09AF" },
            { "\x54\xAA\xA9", "\u09B0\u09CD\u099E\u09CD\u09B0" }, { "\x55\xA9\x82", "\u099F\u09B0\u09CD\u09C2" },
            { "\x55\xA9\x85", "\u099F\u09B0\u09CD\u09C3" }, { "\x55\xAA\x79", "\u099F\u09CD\u09B0\u09C1" },
            { "\x55\xAA\xA8", "\u099F\u09CD\u09B0\u09CD\u09AF" }, { "\x55\xAA\xA9", "\u09B0\u09CD\u099F\u09CD\u09B0" },
            { "\x56\xA9\x82", "\u09A0\u09B0\u09CD\u09C2" }, { "\x56\xA9\x85", "\u09A0\u09B0\u09CD\u09C3" },
            { "\x56\xAA\xA8", "\u09A0\u09CD\u09B0\u09CD\u09AF" }, { "\x56\xAA\xA9", "\u09B0\u09CD\u09A0\u09CD\u09B0" },
            { "\x57\xA9\x82", "\u09A1\u09B0\u09CD\u09C2" }, { "\x57\xA9\x85", "\u09A1\u09B0\u09CD\u09C3" },
            { "\x57\xAA\xA8", "\u09A1\u09CD\u09B0\u09CD\u09AF" }, { "\x57\xAA\xA9", "\u09B0\u09CD\u09A1\u09CD\u09B0" },
            { "\x58\xA9\x82", "\u09A2\u09B0\u09CD\u09C2" }, { "\x58\xA9\x85", "\u09A2\u09B0\u09CD\u09C3" },
            { "\x58\xAA\xA8", "\u09A2\u09CD\u09B0\u09CD\u09AF" }, { "\x58\xAA\xA9", "\u09B0\u09CD\u09A2\u09CD\u09B0" },
            { "\x59\x58\x82", "\u09A3\u09CD\u09A2\u09C2" }, { "\x59\xAA\xA8", "\u09A3\u09CD\u09B0\u09CD\u09AF" },
            { "\x59\xAA\xA9", "\u09B0\u09CD\u09A3\u09CD\u09B0" }, { "\x5A\xA9\x82", "\u09A4\u09B0\u09CD\u09C2" },
            { "\x5A\xA9\x85", "\u09A4\u09B0\u09CD\u09C3" }, { "\x5F\xAA\x93", "\u09A5\u09CD\u09B0\u09C1" },
            { "\x5F\xAA\xA8", "\u09A5\u09CD\u09B0\u09CD\u09AF" }, { "\x5F\xAA\xA9", "\u09B0\u09CD\u09A5\u09CD\u09B0" },
            { "\x60\xA8\x79", "\u09A6\u09CD\u09AF\u09C1" }, { "\x60\xAA\x83", "\u09A6\u09CD\u09B0\u09C2" },
            { "\x60\xAA\x93", "\u09A6\u09CD\u09B0\u09C1" }, { "\x60\xAA\xA8", "\u09A6\u09CD\u09B0\u09CD\u09AF" },
            { "\x60\xAA\xA9", "\u09B0\u09CD\u09A6\u09CD\u09B0" }, { "\x61\xAA\x83", "\u09A7\u09CD\u09B0\u09C1" },
            { "\x61\xAA\xA8", "\u09A7\u09CD\u09B0\u09CD\u09AF" }, { "\x61\xAA\xA9", "\u09B0\u09CD\u09A7\u09CD\u09B0" },
            { "\x62\xAA\xA8", "\u09A8\u09CD\u09B0\u09CD\u09AF" }, { "\x62\xAA\xA9", "\u09B0\u09CD\u09A8\u09CD\u09B0" },
            { "\x63\xD6\x83", "\u09AA\u09CD\u09B0\u09C2" }, { "\x63\xD6\x93", "\u09AA\u09CD\u09B0\u09C1" },
            { "\x63\xD6\xA8", "\u09AA\u09CD\u09B0\u09CD\u09AF" }, { "\x63\xD6\xA9", "\u09B0\u09CD\u09AA\u09CD\u09B0" },
            { "\x64\xA9\x82", "\u09AB\u09B0\u09CD\u09C2" }, { "\x64\xAB\x79", "\u09AB\u09CD\u09B0\u09C1" },
            { "\x64\xAB\x7E", "\u09AB\u09CD\u09B0\u09C2" }, { "\x64\xAB\x84", "\u09AB\u09CD\u09B0\u09C3" },
            { "\x64\xAB\xA8", "\u09AB\u09CD\u09B0\u09CD\u09AF" }, { "\x64\xAB\xA9", "\u09B0\u09CD\u09AB\u09CD\u09B0" },
            { "\x64\xAC\x79", "\u09AB\u09CD\u09B2\u09C1" }, { "\x64\xAC\x7E", "\u09AB\u09CD\u09B2\u09C2" },
            { "\x65\xAA\x83", "\u09AC\u09CD\u09B0\u09C2" }, { "\x65\xAA\x93", "\u09AC\u09CD\u09B0\u09C1" },
            { "\x65\xAA\xA8", "\u09AC\u09CD\u09B0\u09CD\u09AF" }, { "\x65\xAA\xA9", "\u09B0\u09CD\u09AC\u09CD\u09B0" },
            { "\x66\xA9\x82", "\u09AD\u09B0\u09CD\u09C2" }, { "\x66\xA9\x85", "\u09AD\u09B0\u09CD\u09C3" },
            { "\x66\xAA\xA8", "\u09AD\u09CD\u09B0\u09CD\u09AF" }, { "\x67\xAA\x83", "\u09AE\u09CD\u09B0\u09C2" },
            { "\x67\xAA\x93", "\u09AE\u09CD\u09B0\u09C1" }, { "\x67\xAA\xA8", "\u09AE\u09CD\u09B0\u09CD\u09AF" },
            { "\x67\xAA\xA9", "\u09B0\u09CD\u09AE\u09CD\u09B0" }, { "\x68\xAA\xA8", "\u09AF\u09CD\u09B0\u09CD\u09AF" },
            { "\x6A\xAA\xA8", "\u09B2\u09CD\u09B0\u09CD\u09AF" }, { "\x6A\xAA\xA9", "\u09B0\u09CD\u09B2\u09CD\u09B0" },
            { "\x6B\xAA\x93", "\u09B6\u09CD\u09B0\u09C1" }, { "\x6B\xD6\xA8", "\u09B6\u09CD\u09B0\u09CD\u09AF" },
            { "\x6B\xD6\xA9", "\u09B0\u09CD\u09B6\u09CD\u09B0" }, { "\x6C\xAA\xA8", "\u09B7\u09CD\u09B0\u09CD\u09AF" },
            { "\x6C\xAA\xA9", "\u09B0\u09CD\u09B7\u09CD\u09B0" }, { "\x6D\xAA\x83", "\u09B8\u09CD\u09B0\u09C2" },
            { "\x6D\xAA\x93", "\u09B8\u09CD\u09B0\u09C1" }, { "\x6D\xAA\xA8", "\u09B8\u09CD\u09B0\u09CD\u09AF" },
            { "\x6D\xAA\xA9", "\u09B0\u09CD\u09B8\u09CD\u09B0" }, { "\x6E\xAB\xA8", "\u09B9\u09CD\u09B0\u09CD\u09AF" },
            { "\x6E\xAB\xA9", "\u09B0\u09CD\u09B9\u09CD\u09B0" }, { "\x6F\xAA\xA8", "\u09DC\u09CD\u09B0\u09CD\u09AF" },
            { "\x70\xAA\xA8", "\u09DD\u09CD\u09B0\u09CD\u09AF" }, { "\x71\xAA\xA8", "\u09DF\u09CD\u09B0\u09CD\u09AF" },
            { "\x94\x50\x7A", "\u099A\u09CD\u099A\u09C1" }, { "\x94\x50\x82", "\u099A\u09CD\u099A\u09C2" },
            { "\x94\x50\x85", "\u099A\u09CD\u099A\u09C3" }, { "\x94\x51\x7A", "\u099A\u09CD\u099B\u09C1" },
            { "\x94\x51\x82", "\u099A\u09CD\u099B\u09C2" }, { "\x94\x51\x85", "\u099A\u09CD\u099B\u09C3" },
            { "\x94\x51\xA1", "\u099A\u09CD\u099B\u09CD\u09AC" }, { "\x94\x51\xAA", "\u099A\u09CD\u099B\u09CD\u09B0" },
            { "\x94\x54\x7A", "\u099A\u09CD\u099E\u09C1" }, { "\x94\x54\x82", "\u099A\u09CD\u099E\u09C2" },
            { "\x94\x54\x85", "\u099A\u09CD\u099E\u09C3" }, { "\x95\x8C\x79", "\u0999\u09CD\u0995\u09CD\u09B0\u09C1" },
            { "\x95\x8C\x7E", "\u0999\u09CD\u0995\u09CD\u09B0\u09C2" }, { "\x95\xB6\x7A", "\u0999\u09CD\u0995\u09CD\u09B7\u09C1" },
            { "\x95\xB6\x7E", "\u0999\u09CD\u0995\u09CD\u09B7\u09C2" }, { "\x95\xB6\x84", "\u0999\u09CD\u0995\u09CD\u09B7\u09C3" },
            { "\x99\xA2\x7A", "\u09A6\u09CD\u09AD\u09C1" }, { "\x99\xA2\x82", "\u09A6\u09CD\u09AD\u09C2" },
            { "\x99\xA2\x85", "\u09A6\u09CD\u09AD\u09C3" }, { "\x99\xA3\x83", "\u09A6\u09CD\u09AD\u09CD\u09B0\u09C2" },
            { "\x99\xA3\x93", "\u09A6\u09CD\u09AD\u09CD\u09B0\u09C1" }, { "\x9A\x92\x79", "\u09A8\u09CD\u09A5\u09C1" },
            { "\x9A\x92\x7E", "\u09A8\u09CD\u09A5\u09C2" }, { "\x9A\x92\x84", "\u09A8\u09CD\u09A5\u09C3" },
            { "\x9A\x97\x82", "\u09A8\u09CD\u09A4\u09C2" }, { "\x9A\x97\x85", "\u09A8\u09CD\u09A4\u09C3" },
            { "\x9A\x97\xA1", "\u09A8\u09CD\u09A4\u09CD\u09AC" }, { "\x9B\x55\x7A", "\u09A8\u09CD\u099F\u09C1" },
            { "\x9B\x55\x82", "\u09A8\u09CD\u099F\u09C2" }, { "\x9B\x55\x85", "\u09A8\u09CD\u099F\u09C3" },
            { "\xA4\x63\xAD", "\u09AE\u09CD\u09AA\u09CD\u09B2" }, { "\xA4\xFA\xD6", "\u09AE\u09CD\u09AA\u09CD\u09B0" },
            { "\xAE\x8B\x7A", "\u09B7\u09CD\u0995\u09C1" }, { "\xAE\x8B\x82", "\u09B7\u09CD\u0995\u09C2" },
            { "\xAE\x8B\x85", "\u09B7\u09CD\u0995\u09C3" }, { "\xAE\x8C\x79", "\u09B7\u09CD\u0995\u09CD\u09B0\u09C1" },
            { "\xAE\x8C\x7E", "\u09B7\u09CD\u0995\u09CD\u09B0\u09C2" }, { "\xAE\x8C\x84", "\u09B7\u09CD\u0995\u09CD\u09B0\u09C3" },
            { "\xAE\xA7\xAA", "\u09B7\u09CD\u09AE\u09CD\u09B0" }, { "\xAE\xFA\xD6", "\u09B7\u09CD\u09AA\u09CD\u09B0" },
            { "\xAF\x63\xAD", "\u09B8\u09CD\u09AA\u09CD\u09B2" }, { "\xAF\x8B\x7A", "\u09B8\u09CD\u0995\u09C1" },
            { "\xAF\x8B\x82", "\u09B8\u09CD\u0995\u09C2" }, { "\xAF\x8B\x85", "\u09B8\u09CD\u0995\u09C3" },
            { "\xAF\x8C\x79", "\u09B8\u09CD\u0995\u09CD\u09B0\u09C1" }, { "\xAF\x8C\x7E", "\u09B8\u09CD\u0995\u09CD\u09B0\u09C2" },
            { "\xAF\x8C\x84", "\u09B8\u09CD\u0995\u09CD\u09B0\u09C3" }, { "\xAF\x92\x79", "\u09B8\u09CD\u09A5\u09C1" },
            { "\xAF\x92\x7E", "\u09B8\u09CD\u09A5\u09C2" }, { "\xAF\x92\x84", "\u09B8\u09CD\u09A5\u09C3" },
            { "\xAF\x97\x82", "\u09B8\u09CD\u09A4\u09C2" }, { "\xAF\x97\x85", "\u09B8\u09CD\u09A4\u09C3" },
            { "\xAF\x97\xA1", "\u09B8\u09CD\u09A4\u09CD\u09AC" }, { "\xAF\xBF\x83", "\u09B8\u09CD\u09A4\u09CD\u09B0\u09C2" },
            { "\xAF\xFA\xD6", "\u09B8\u09CD\u09AA\u09CD\u09B0" }, { "\xB5\xA8\xA9", "\u09B0\u09CD\u0995\u09CD\u09B0\u09CD\u09AF" },
            { "\xB6\x9C\x79", "\u0995\u09CD\u09B7\u09CD\u09A8\u09C1" }, { "\xB6\x9C\x7E", "\u0995\u09CD\u09B7\u09CD\u09A8\u09C2" },
            { "\xB6\x9C\x84", "\u0995\u09CD\u09B7\u09CD\u09A8\u09C3" }, { "\xB6\xA9\x85", "\u0995\u09CD\u09B7\u09B0\u09CD\u09C3" },
            { "\xCE\xA8\xA9", "\u09B0\u09CD\u09A4\u09CD\u09B0\u09CD\u09AF" }, { "\xDB\xAA\x79", "\u09A8\u09CD\u09A1\u09CD\u09B0\u09C1" },
            { "\xDC\xAA\x83", "\u09A8\u09CD\u09A7\u09CD\u09B0\u09C2" }, { "\xDC\xAA\x93", "\u09A8\u09CD\u09A7\u09CD\u09B0\u09C1" },
            { "\xDC\xAB\xA8", "\u09A8\u09CD\u09A7\u09CD\u09B0\u09CD\u09AF" }, { "\xF9\xAB\x79", "\u09B8\u09CD\u09AB\u09CD\u09B0\u09C1" },
            { "\xF9\xAB\x7E", "\u09B8\u09CD\u09AB\u09CD\u09B0\u09C2" },
            // 2-byte conjunct sequences
            { "\x26\x26", "\u09CD\u200C" },
            { "\x4B\x7A", "\u0995\u09C1" }, { "\x4B\x82", "\u0995\u09C2" },
            { "\x4B\x85", "\u0995\u09C3" }, { "\x4B\x91", "\u0995\u09CD\u09A4\u09C1" },
            { "\x4B\xA1", "\u0995\u09CD\u09AC" }, { "\x4B\xAC", "\u0995\u09CD\u09B2" },
            { "\x4B\xE8", "\u0995\u09CD\u09A8" }, { "\x4D\x9C", "\u0997\u09CD\u09A8" },
            { "\x4D\xA5", "\u0997\u09CD\u09AE" }, { "\x4D\xA6", "\u0997\u09CD\u09AC" },
            { "\x4D\xAD", "\u0997\u09CD\u09B2" }, { "\x4D\xD6", "\u0997\u09CD\u09B0" },
            { "\x4D\xE8", "\u0997\u09CD\u09A3" }, { "\x4E\x9C", "\u0998\u09CD\u09A8" },
            { "\x4F\x7A", "\u0999\u09C1" }, { "\x4F\x82", "\u0999\u09C2" },
            { "\x4F\x85", "\u0999\u09C3" }, { "\x50\x7A", "\u099A\u09C1" },
            { "\x50\x82", "\u099A\u09C2" }, { "\x50\x85", "\u099A\u09C3" },
            { "\x51\x79", "\u099B\u09C1" }, { "\x51\x7E", "\u099B\u09C2" },
            { "\x51\x84", "\u099B\u09C3" }, { "\x52\xA1", "\u099C\u09CD\u09AC" },
            { "\x53\x7A", "\u099D\u09C1" }, { "\x53\x82", "\u099D\u09C2" },
            { "\x53\x85", "\u099D\u09C3" }, { "\x53\xA1", "\u099D\u09CD\u09AC" },
            { "\x54\x7A", "\u099E\u09C1" }, { "\x54\x82", "\u099E\u09C2" },
            { "\x54\x85", "\u099E\u09C3" }, { "\x55\x7A", "\u099F\u09C1" },
            { "\x55\x82", "\u099F\u09C2" }, { "\x55\x85", "\u099F\u09C3" },
            { "\x55\xA1", "\u099F\u09CD\u09AC" }, { "\x55\xA5", "\u099F\u09CD\u09AE" },
            { "\x56\x7A", "\u09A0\u09C1" }, { "\x56\x82", "\u09A0\u09C2" },
            { "\x56\x85", "\u09A0\u09C3" }, { "\x57\x7A", "\u09A1\u09C1" },
            { "\x57\x82", "\u09A1\u09C2" }, { "\x57\x85", "\u09A1\u09C3" },
            { "\x57\xA1", "\u09A1\u09CD\u09AC" }, { "\x57\xAA", "\u09A1\u09CD\u09B0" },
            { "\x58\x7A", "\u09A2\u09C1" }, { "\x58\x82", "\u09A2\u09C2" },
            { "\x58\x85", "\u09A2\u09C3" }, { "\x58\xAA", "\u09A2\u09CD\u09B0" },
            { "\x59\x58", "\u09A3\u09CD\u09A2" }, { "\x59\x9C", "\u09A3\u09CD\u09A8" },
            { "\x59\xA1", "\u09A3\u09CD\u09AC" }, { "\x59\xA5", "\u09A3\u09CD\u09AE" },
            { "\x59\xE8", "\u09A3\u09CD\u09A3" }, { "\x5A\x7A", "\u09A4\u09C1" },
            { "\x5A\x82", "\u09A4\u09C2" }, { "\x5A\x85", "\u09A4\u09C3" },
            { "\x5A\x91", "\u09A4\u09CD\u09A4\u09C1" }, { "\x5A\x9C", "\u09A4\u09CD\u09A8" },
            { "\x5A\xA1", "\u09A4\u09CD\u09AC" }, { "\x5F\xA1", "\u09A5\u09CD\u09AC" },
            { "\x61\x9C", "\u09A7\u09CD\u09A8" }, { "\x61\x9F", "\u09A7\u09CD\u09AC" },
            { "\x61\xA5", "\u09A7\u09CD\u09AE" }, { "\x62\x9C", "\u09A8\u09CD\u09A8" },
            { "\x62\xA5", "\u09A8\u09CD\u09AE" }, { "\x63\x9C", "\u09AA\u09CD\u09A8" },
            { "\x63\xAD", "\u09AA\u09CD\u09B2" }, { "\x63\xD6", "\u09AA\u09CD\u09B0" },
            { "\x64\x7A", "\u09AB\u09C1" }, { "\x64\x82", "\u09AB\u09C2" },
            { "\x64\x85", "\u09AB\u09C3" }, { "\x64\xAB", "\u09AB\u09CD\u09B0" },
            { "\x64\xAC", "\u09AB\u09CD\u09B2" }, { "\x65\x9F", "\u09AC\u09CD\u09AC" },
            { "\x65\xAD", "\u09AC\u09CD\u09B2" }, { "\x66\x7A", "\u09AD\u09C1" },
            { "\x66\x82", "\u09AD\u09C2" }, { "\x66\x85", "\u09AD\u09C3" },
            { "\x66\xA1", "\u09AD\u09CD\u09AC" }, { "\x66\xAC", "\u09AD\u09CD\u09B2" },
            { "\x67\xAA", "\u09AE\u09CD\u09B0" }, { "\x69\x83", "\u09B0\u09C2" },
            { "\x69\x93", "\u09B0\u09C1" }, { "\x69\xA8", "\u09B0\u200C\u09CD\u09AF" },
            { "\x6A\xA1", "\u09B2\u09CD\u09AC" }, { "\x6A\xA5", "\u09B2\u09CD\u09AE" },
            { "\x6A\xAD", "\u09B2\u09CD\u09B2" }, { "\x6B\x9C", "\u09B6\u09CD\u09A8" },
            { "\x6B\xA5", "\u09B6\u09CD\u09AE" }, { "\x6B\xA6", "\u09B6\u09CD\u09AC" },
            { "\x6B\xAA", "\u09B6\u09CD\u09B0" }, { "\x6B\xAD", "\u09B6\u09CD\u09B2" },
            { "\x6D\xAA", "\u09B8\u09CD\u09B0" }, { "\x6E\x9C", "\u09B9\u09CD\u09A3" },
            { "\x6E\x9F", "\u09B9\u09CD\u09AC" }, { "\x6E\xAD", "\u09B9\u09CD\u09B2" },
            { "\x6F\x96", "\u09DC\u09C1" }, { "\x75\x76", "\u09BE\u0981" },
            { "\x75\x78", "\u09C0\u0981" }, { "\x75\x79", "\u09C1\u0981" },
            { "\x75\x8A", "\u09D7\u0981" }, { "\x79\xA8", "\u09CD\u09AF\u09C1" },
            { "\x94\x50", "\u099A\u09CD\u099A" }, { "\x94\x51", "\u099A\u09CD\u099B" },
            { "\x94\x54", "\u099A\u09CD\u099E" }, { "\x94\x9C", "\u099A\u09CD\u09A8" },
            { "\x94\xA1", "\u099A\u09CD\u09AC" }, { "\x95\x4C", "\u0999\u09CD\u0996" },
            { "\x95\x4E", "\u0999\u09CD\u0998" }, { "\x95\x67", "\u0999\u09CD\u09AE" },
            { "\x95\x8C", "\u0999\u09CD\u0995\u09CD\u09B0" }, { "\x95\xB3", "\u0999\u09CD\u0995\u09CD\u09A4" },
            { "\x95\xB6", "\u0999\u09CD\u0995\u09CD\u09B7" }, { "\x98\x4D", "\u09A6\u09CD\u0997" },
            { "\x98\x4E", "\u09A6\u09CD\u0998" }, { "\x99\x9C", "\u09A6\u09CD\u09A8" },
            { "\x99\xA2", "\u09A6\u09CD\u09AD" }, { "\x99\xA3", "\u09A6\u09CD\u09AD\u09CD\u09B0" },
            { "\x9A\x91", "\u09A8\u09CD\u09A4\u09C1" }, { "\x9A\x92", "\u09A8\u09CD\u09A5" },
            { "\x9A\x97", "\u09A8\u09CD\u09A4" }, { "\x9A\xBF", "\u09A8\u09CD\u09A4\u09CD\u09B0" },
            { "\x9B\x55", "\u09A8\u09CD\u099F" }, { "\x9B\x60", "\u09A8\u09CD\u09A6" },
            { "\x9B\xD8", "\u09A8\u09CD\u09A6\u09CD\u09AC" }, { "\xA4\xA2", "\u09AE\u09CD\u09AD" },
            { "\xA4\xA3", "\u09AE\u09CD\u09AD\u09CD\u09B0" }, { "\xA4\xA7", "\u09AE\u09CD\u09AE" },
            { "\xA4\xAC", "\u09AE\u09CD\u09B2" }, { "\xA4\xFA", "\u09AE\u09CD\u09AA" },
            { "\xA8\x79", "\u09CD\u09AF\u09C1" }, { "\xA8\x7E", "\u09CD\u09AF\u09C2" },
            { "\xA8\xA9", "\u09B0\u09CD\u09AF" }, { "\xAA\x7E", "\u09CD\u09B0\u09C2" },
            { "\xAE\x8B", "\u09B7\u09CD\u0995" }, { "\xAE\x8C", "\u09B7\u09CD\u0995\u09CD\u09B0" },
            { "\xAE\xA7", "\u09B7\u09CD\u09AE" }, { "\xAE\xFA", "\u09B7\u09CD\u09AA" },
            { "\xAF\x8B", "\u09B8\u09CD\u0995" }, { "\xAF\x8C", "\u09B8\u09CD\u0995\u09CD\u09B0" },
            { "\xAF\x91", "\u09B8\u09CD\u09A4\u09C1" }, { "\xAF\x92", "\u09B8\u09CD\u09A5" },
            { "\xAF\x97", "\u09B8\u09CD\u09A4" }, { "\xAF\xA7", "\u09B8\u09CD\u09AE" },
            { "\xAF\xAC", "\u09B8\u09CD\u09B2" }, { "\xAF\xBF", "\u09B8\u09CD\u09A4\u09CD\u09B0" },
            { "\xAF\xFA", "\u09B8\u09CD\u09AA" }, { "\xB0\x7A", "\u0995\u09CD\u0995\u09C1" },
            { "\xB0\x82", "\u0995\u09CD\u0995\u09C2" }, { "\xB0\x85", "\u0995\u09CD\u0995\u09C3" },
            { "\xB1\x7A", "\u0995\u09CD\u099F\u09C1" }, { "\xB1\x82", "\u0995\u09CD\u099F\u09C2" },
            { "\xB1\x85", "\u0995\u09CD\u099F\u09C3" }, { "\xB2\xA8", "\u0995\u09CD\u09B7\u09CD\u09AE\u09CD\u09AF" },
            { "\xB3\x82", "\u0995\u09CD\u09A4\u09C2" }, { "\xB3\x85", "\u0995\u09CD\u09A4\u09C3" },
            { "\xB3\xA1", "\u0995\u09CD\u09A4\u09CD\u09AC" }, { "\xB3\xAB", "\u0995\u09CD\u09A4\u09CD\u09B0" },
            { "\xB5\x79", "\u0995\u09CD\u09B0\u09C1" }, { "\xB5\x7E", "\u0995\u09CD\u09B0\u09C2" },
            { "\xB5\x84", "\u0995\u09CD\u09B0\u09C3" }, { "\xB5\xA8", "\u0995\u09CD\u09B0\u09CD\u09AF" },
            { "\xB5\xA9", "\u09B0\u09CD\u0995\u09CD\u09B0" }, { "\xB6\x7A", "\u0995\u09CD\u09B7\u09C1" },
            { "\xB6\x7E", "\u0995\u09CD\u09B7\u09C2" }, { "\xB6\x84", "\u0995\u09CD\u09B7\u09C3" },
            { "\xB6\x9C", "\u0995\u09CD\u09B7\u09CD\u09A8" }, { "\xB6\xA1", "\u0995\u09CD\u09B7\u09CD\u09AC" },
            { "\xB6\xA8", "\u0995\u09CD\u09B7\u09CD\u09AF" }, { "\xB6\xE8", "\u0995\u09CD\u09B7\u09CD\u09A3" },
            { "\xBB\x79", "\u0997\u09CD\u09A7\u09C1" }, { "\xBB\x7E", "\u0997\u09CD\u09A7\u09C2" },
            { "\xBB\x84", "\u0997\u09CD\u09A7\u09C3" }, { "\xBB\xAA", "\u0997\u09CD\u09A7\u09CD\u09B0" },
            { "\xBC\x7A", "\u0999\u09CD\u0995\u09C1" }, { "\xBC\x85", "\u0999\u09CD\u0995\u09C3" },
            { "\xBE\xA1", "\u099C\u09CD\u099C\u09CD\u09AC" }, { "\xC0\x82", "\u099C\u09CD\u099D\u09C2" },
            { "\xC0\x85", "\u099C\u09CD\u099D\u09C3" }, { "\xC1\x7A", "\u099C\u09CD\u099E\u09C1" },
            { "\xC1\x82", "\u099C\u09CD\u099E\u09C2" }, { "\xC1\x85", "\u099C\u09CD\u099E\u09C3" },
            { "\xC2\x7A", "\u099E\u09CD\u099A\u09C1" }, { "\xC2\x82", "\u099E\u09CD\u099A\u09C2" },
            { "\xC2\x85", "\u099E\u09CD\u099A\u09C3" }, { "\xC3\x7A", "\u099E\u09CD\u099B\u09C1" },
            { "\xC3\x82", "\u099E\u09CD\u099B\u09C2" }, { "\xC3\x85", "\u099E\u09CD\u099B\u09C3" },
            { "\xC6\x7A", "\u099F\u09CD\u099F\u09C1" }, { "\xC6\x82", "\u099F\u09CD\u099F\u09C2" },
            { "\xC6\x85", "\u099F\u09CD\u099F\u09C3" }, { "\xC7\x7A", "\u09A1\u09CD\u09A1\u09C1" },
            { "\xC7\x82", "\u09A1\u09CD\u09A1\u09C2" }, { "\xC7\x85", "\u09A1\u09CD\u09A1\u09C3" },
            { "\xC8\x7A", "\u09A3\u09CD\u099F\u09C1" }, { "\xC8\x82", "\u09A3\u09CD\u099F\u09C2" },
            { "\xC8\x85", "\u09A3\u09CD\u099F\u09C3" }, { "\xC9\x7A", "\u09A3\u09CD\u09A0\u09C1" },
            { "\xC9\x82", "\u09A3\u09CD\u09A0\u09C2" }, { "\xC9\x85", "\u09A3\u09CD\u09A0\u09C3" },
            { "\xCA\x7A", "\u09A3\u09CD\u09A1\u09C1" }, { "\xCA\x82", "\u09A3\u09CD\u09A1\u09C2" },
            { "\xCA\x85", "\u09A3\u09CD\u09A1\u09C3" }, { "\xCA\xAA", "\u09A3\u09CD\u09A1\u09CD\u09B0" },
            { "\xCB\x82", "\u09A4\u09CD\u09A4\u09C2" }, { "\xCB\x85", "\u09A4\u09CD\u09A4\u09C3" },
            { "\xCB\xA1", "\u09A4\u09CD\u09A4\u09CD\u09AC" }, { "\xCB\xAA", "\u09A4\u09CD\u09A4\u09CD\u09B0" },
            { "\xCE\x83", "\u09A4\u09CD\u09B0\u09C2" }, { "\xCE\x93", "\u09A4\u09CD\u09B0\u09C1" },
            { "\xCE\xA8", "\u09A4\u09CD\u09B0\u09CD\u09AF" }, { "\xCE\xA9", "\u09B0\u09CD\u09A4\u09CD\u09B0" },
            { "\xCF\xA1", "\u09A6\u09CD\u09A6\u09CD\u09AC" }, { "\xD7\x79", "\u09A6\u09CD\u09A7\u09C1" },
            { "\xD7\x7E", "\u09A6\u09CD\u09A7\u09C2" }, { "\xD7\x84", "\u09A6\u09CD\u09A7\u09C3" },
            { "\xDA\x7A", "\u09A8\u09CD\u09A0\u09C1" }, { "\xDA\x82", "\u09A8\u09CD\u09A0\u09C2" },
            { "\xDA\x85", "\u09A8\u09CD\u09A0\u09C3" }, { "\xDB\x7A", "\u09A8\u09CD\u09A1\u09C1" },
            { "\xDB\x82", "\u09A8\u09CD\u09A1\u09C2" }, { "\xDB\xAA", "\u09A8\u09CD\u09A1\u09CD\u09B0" },
            { "\xDC\x79", "\u09A8\u09CD\u09A7\u09C1" }, { "\xDC\x7E", "\u09A8\u09CD\u09A7\u09C2" },
            { "\xDC\x84", "\u09A8\u09CD\u09A7\u09C3" }, { "\xDE\x7A", "\u09AA\u09CD\u099F\u09C1" },
            { "\xDE\x82", "\u09AA\u09CD\u099F\u09C2" }, { "\xDE\x85", "\u09AA\u09CD\u099F\u09C3" },
            { "\xDF\x7A", "\u09AA\u09CD\u09A4\u09C1" }, { "\xDF\x82", "\u09AA\u09CD\u09A4\u09C2" },
            { "\xDF\x85", "\u09AA\u09CD\u09A4\u09C3" }, { "\xE4\x79", "\u09AC\u09CD\u09A7\u09C1" },
            { "\xE4\x7E", "\u09AC\u09CD\u09A7\u09C2" }, { "\xE4\x84", "\u09AC\u09CD\u09A7\u09C3" },
            { "\xE5\x83", "\u09AD\u09CD\u09B0\u09C2" }, { "\xE5\x93", "\u09AD\u09CD\u09B0\u09C1" },
            { "\xE5\xA9", "\u09B0\u09CD\u09AD\u09CD\u09B0" }, { "\xE7\x7A", "\u09AE\u09CD\u09AB\u09C1" },
            { "\xE7\x82", "\u09AE\u09CD\u09AB\u09C2" }, { "\xE7\x85", "\u09AE\u09CD\u09AB\u09C3" },
            { "\xE9\x7A", "\u09B2\u09CD\u0995\u09C1" }, { "\xE9\x82", "\u09B2\u09CD\u0995\u09C2" },
            { "\xE9\x85", "\u09B2\u09CD\u0995\u09C3" }, { "\xEB\x7A", "\u09B2\u09CD\u099F\u09C1" },
            { "\xEB\x82", "\u09B2\u09CD\u099F\u09C2" }, { "\xEB\x85", "\u09B2\u09CD\u099F\u09C3" },
            { "\xEE\x7A", "\u09B2\u09CD\u09AB\u09C1" }, { "\xEE\x82", "\u09B2\u09CD\u09AB\u09C2" },
            { "\xEE\x85", "\u09B2\u09CD\u09AB\u09C3" }, { "\xEE\xAB", "\u09B2\u09CD\u09AB\u09CD\u09B0" },
            { "\xF0\x7A", "\u09B6\u09CD\u099A\u09C1" }, { "\xF0\x82", "\u09B6\u09CD\u099A\u09C2" },
            { "\xF0\x85", "\u09B6\u09CD\u099A\u09C3" }, { "\xF1\x7A", "\u09B6\u09CD\u099B\u09C1" },
            { "\xF1\x82", "\u09B6\u09CD\u099B\u09C2" }, { "\xF1\x85", "\u09B6\u09CD\u099B\u09C3" },
            { "\xF2\x7A", "\u09B7\u09CD\u09A3\u09C1" }, { "\xF2\x7E", "\u09B7\u09CD\u09A3\u09C2" },
            { "\xF2\x84", "\u09B7\u09CD\u09A3\u09C3" }, { "\xF3\x7A", "\u09B7\u09CD\u099F\u09C1" },
            { "\xF3\x82", "\u09B7\u09CD\u099F\u09C2" }, { "\xF3\x85", "\u09B7\u09CD\u099F\u09C3" },
            { "\xF4\x7A", "\u09B7\u09CD\u09A0\u09C1" }, { "\xF4\x82", "\u09B7\u09CD\u09A0\u09C2" },
            { "\xF4\x85", "\u09B7\u09CD\u09A0\u09C3" }, { "\xF5\x7A", "\u09B7\u09CD\u09AB\u09C1" },
            { "\xF5\x82", "\u09B7\u09CD\u09AB\u09C2" }, { "\xF5\x85", "\u09B7\u09CD\u09AB\u09C3" },
            { "\xF7\x7A", "\u09B8\u09CD\u099F\u09C1" }, { "\xF7\x82", "\u09B8\u09CD\u099F\u09C2" },
            { "\xF7\x85", "\u09B8\u09CD\u099F\u09C3" }, { "\xF9\x7A", "\u09B8\u09CD\u09AB\u09C1" },
            { "\xF9\x82", "\u09B8\u09CD\u09AB\u09C2" }, { "\xF9\x85", "\u09B8\u09CD\u09AB\u09C3" },
            { "\xF9\xAB", "\u09B8\u09CD\u09AB\u09CD\u09B0" }, { "\xFE\x7A", "\u09B9\u09CD\u09AE\u09C1" },
            { "\xFE\x7E", "\u09B9\u09CD\u09AE\u09C2" }, { "\xFE\x85", "\u09B9\u09CD\u09AE\u09C3" },
            // CP1252 Unicode-key entries (for Windows-1252 interpreted text)
            { "\u20AC", "\u0985" }, { "\u201A", "\u09C2" }, { "\u0192", "\u09C2" }, { "\u201E", "\u09C3" },
            { "\u2026", "\u09C3" }, { "\u2020", "\u09C7" }, { "\u2021", "\u09C7" }, { "\u02C6", "\u09C8" },
            { "\u2030", "\u09C8" }, { "\u0160", "\u09D7" }, { "\u2039", "\u0995\u09CD" }, { "\u0152", "\u0995\u09CD\u09B0" },
            { "\u017D", "\u09AC\u09CD" }, { "\u2018", "\u09CD\u09A4\u09C1" }, { "\u2019", "\u09CD\u09A5" }, { "\u201C", "\u09C1" },
            { "\u201D", "\u099A\u09CD" }, { "\u2022", "\u0999\u09CD" }, { "\u2013", "\u09C1" }, { "\u2014", "\u09A4\u09CD" },
            { "\u02DC", "\u09A6\u09CD" }, { "\u2122", "\u09A6\u09CD" }, { "\u0161", "\u09A8\u09CD" }, { "\u203A", "\u09A8\u09CD" },
            { "\u0153", "\u09CD\u09A8" }, { "\u0178", "\u09CD\u09AC" },
            // Missing 2-byte conjuncts added for U2B conversion
            { "\x60\xA8", "\u09A6\u09CD\u09AF" }, { "\x61\xA8", "\u09A7\u09CD\u09AF" },
            { "\x9B\x60\x7A", "\u09A8\u09CD\u09A6\u09C1" },
            { "\xA4\x9C", "\u09AE\u09CD\u09A8" }, { "\xAF\xA1", "\u09B8\u09CD\u09AC" },
        };
        private static readonly Dictionary<string, string> _unicodeToBijoy;
        private static readonly HashSet<char> _halantFirstBijoyChars;

        private static bool IsAsciiKey(string key)
        {
            return key.Length == 1 && key[0] >= 0x20 && key[0] <= 0x7E;
        }
        private static bool IsExtendedKey(string key)
        {
            return key.Length == 1 && key[0] >= '\x80' && key[0] <= '\xFF';
        }

        private static Dictionary<string, string> BuildUnicodeToBijoy()
        {
            var d = new Dictionary<string, string>();
            foreach (var kv in _bijoyToUnicode)
            {
                // Skip the generic "র্" entry — reph should only be applied
                // via the reph handler (Step 5b), not the dictionary.
                // This prevents ALL র্ from being blindly mapped to \xA9.
                if (kv.Value == "\u09B0\u09CD")
                    continue;
                if (!d.ContainsKey(kv.Value))
                    d[kv.Value] = kv.Key;
                else
                {
                    var existing = d[kv.Value];
                    bool existingAscii = IsAsciiKey(existing);
                    bool newAscii = IsAsciiKey(kv.Key);
                    if (newAscii && !existingAscii)
                        d[kv.Value] = kv.Key;
                    else if (existingAscii && !newAscii)
                        { }
                    else if (IsExtendedKey(kv.Key) && !IsExtendedKey(existing))
                        d[kv.Value] = kv.Key;
                    else if (IsExtendedKey(existing) && !IsExtendedKey(kv.Key))
                        { }
                    else if (kv.Key.Length > existing.Length)
                        d[kv.Value] = kv.Key;
                    else if (kv.Key.Length == existing.Length && string.CompareOrdinal(kv.Key, existing) < 0)
                        d[kv.Value] = kv.Key;
                }
            }
            return d;
        }
        private static HashSet<char> BuildHalantFirstSet()
        {
            var set = new HashSet<char>();
            foreach (var kv in _bijoyToUnicode)
            {
                if (kv.Key.Length == 1 && kv.Value.StartsWith("\u09CD"))
                    set.Add(kv.Key[0]);
            }
            return set;
        }
        static WordAddIn() { _unicodeToBijoy = BuildUnicodeToBijoy(); _halantFirstBijoyChars = BuildHalantFirstSet(); }

        public void OnFontConvert(Office.IRibbonControl c)
        {
            try
            {
                var sel = _app.Selection;
                if (sel == null) return;
                var text = sel.Text;
                if (string.IsNullOrEmpty(text)) { MessageBox.Show("Select some text first.", "Font Convert"); return; }
                bool toUnicode = c.Id == "btnB2U";
                var result = toUnicode ? ConvertBijoyToUnicode(text) : ConvertUnicodeToBijoy(text);
                if (result == text) return;
                var doc = _app.ActiveDocument;
                if (doc == null) return;
                Word.Range rng = sel.Range;
                int startPos = rng.Start;
                rng.Text = result;
                Word.Range convRng = doc.Range(startPos, startPos + result.Length);
                if (toUnicode)
                {
                    try { convRng.Font.Name = "SolaimanLipi"; } catch { }
                }
                else { convRng.Font.Name = "SutonnyMJ"; }
            }
            catch (System.Exception ex) { MessageBox.Show(ex.Message, "Font Convert Error"); }
        }

        private static string ConvertBijoyToUnicode(string text)
        {
            var sb = new System.Text.StringBuilder();
            int i = 0;
            // First pass: map characters (longest match first)
            while (i < text.Length)
            {
                bool found = false;
                for (int len = 4; len >= 1; len--)
                {
                    if (i + len <= text.Length)
                    {
                        var sub = text.Substring(i, len);
                        string u;
                        if (_bijoyToUnicode.TryGetValue(sub, out u))
                        {
                            sb.Append(u);
                            i += len;
                            found = true;
                            break;
                        }
                    }
                }
                if (!found) { sb.Append(text[i]); i++; }
            }
            // Second pass: reorder pre-kars (ি, ে, ৈ) to after their consonant
            var result = sb.ToString();
            sb.Clear();

            // Special pre-process for 2+ pre-kars in wk order:
            // Move the first pre-kar (before first consonant) to the end
            // so that remaining pre-kars can be processed correctly.
            // Example: ে+ন+ি+ম্+ন → ন+ি+ম্+ন+ে
            string prepped = result;
            if (result.Length > 1
                && (result[0] == '\u09BF' || result[0] == '\u09C7' || result[0] == '\u09C8'))
            {
                bool hasLater = false;
                for (int j = 1; j < result.Length; j++)
                {
                    if (result[j] == '\u09BF' || result[j] == '\u09C7' || result[j] == '\u09C8')
                    { hasLater = true; break; }
                }
                if (hasLater)
                    prepped = result.Substring(1) + result[0];
            }

            i = 0;
            while (i < prepped.Length)
            {
                char c = prepped[i];
                if ((c == '\u09BF' || c == '\u09C7' || c == '\u09C8') && i + 1 < prepped.Length)
                {
                    // If preceded by a consonant, check if the following cluster
                    // is a conjunct (has halant). If it IS a conjunct AND
                    // there's no other pre-kar later, this was typed in wk
                    // order (pre-kar before the whole cluster) → swap.
                    // If no conjunct follows, this is kw order → keep.
                    if (i > 0 && IsConsonant(prepped[i - 1]))
                    {
                        int scan = i + 1;
                        while (scan < prepped.Length && IsPostKar(prepped[scan]))
                            scan++;
                        bool hasConjunct = false;
                        if (scan < prepped.Length && prepped[scan] == '\u09CD')
                            hasConjunct = true;
                        else if (scan < prepped.Length && IsConsonant(prepped[scan])
                            && scan + 1 < prepped.Length && prepped[scan + 1] == '\u09CD')
                            hasConjunct = true;
                        // Only swap if no other pre-kar follows (avoid double-swap
                        // in multi-prekar words like ন+ি+ম্+ন+ে)
                        bool hasLaterPreKar = false;
                        for (int j = i + 1; j < prepped.Length; j++)
                        {
                            if (prepped[j] == '\u09BF' || prepped[j] == '\u09C7' || prepped[j] == '\u09C8')
                            { hasLaterPreKar = true; break; }
                        }
                        if (!hasConjunct || hasLaterPreKar)
                        {
                            sb.Append(c); i++;
                            continue;
                        }
                    }
                    // Swap: move pre-kar past the immediately following cluster
                    int consStart = i + 1;
                    while (consStart < prepped.Length && IsPostKar(prepped[consStart]))
                        consStart++;
                    while (consStart < prepped.Length && prepped[consStart] == '\u09CD')
                        consStart++;
                    if (consStart < prepped.Length && IsConsonant(prepped[consStart]))
                    {
                        int clusterEnd = consStart + 1;
                        while (clusterEnd < prepped.Length && prepped[clusterEnd] == '\u09CD')
                        {
                            clusterEnd++;
                            if (clusterEnd < prepped.Length && IsConsonant(prepped[clusterEnd]))
                                clusterEnd++;
                            else break;
                        }
                        sb.Append(prepped.Substring(i + 1, clusterEnd - i - 1));
                        sb.Append(c);
                        i = clusterEnd;
                    }
                    else { sb.Append(c); i++; }
                }
                else
                {
                    sb.Append(c); i++;
                }
            }
            // Pass 2.5: reph and ra-phala positioning
            //   Rule 1:  র্ mid-word, preceded by post-kar, followed by C
            //            → insert halant before র (ও+য+া+র্+ড → ও+য+া+্+র+ড)
            //   Rule 2:  র্ at word end
            //            → move original র্ before preceding consonant (ধ+ম+র্ → ধ+র্+ম)
            var rephPass = sb.ToString();
            sb.Clear();
            int lastCopied = 0;
            i = 0;
            while (i < rephPass.Length)
            {
                bool isEndOfWord = (i + 2 >= rephPass.Length)
                    || rephPass[i + 2] == ' ' || rephPass[i + 2] == '\t'
                    || rephPass[i + 2] == ',' || rephPass[i + 2] == ';'
                    || rephPass[i + 2] == '.' || rephPass[i + 2] == '!'
                    || rephPass[i + 2] == '?' || rephPass[i + 2] == ':'
                    || rephPass[i + 2] == '\u0964' || rephPass[i + 2] == '\u0965'
                    || rephPass[i + 2] == '\r' || rephPass[i + 2] == '\n';
                // Rule 1: র্ mid-word, preceded by post-kar, followed by C
                // → insert halant before র (remove original halant after র)
                //   (ও+য+া+র্+ড → ও+য+া+্+র+ড = ওয়ার্ড)
                if (i + 1 < rephPass.Length && rephPass[i] == '\u09B0' && rephPass[i + 1] == '\u09CD'
                    && !isEndOfWord && i + 2 < rephPass.Length && IsConsonant(rephPass[i + 2])
                    && i > 0 && IsPostKar(rephPass[i - 1]))
                {
                    if (lastCopied < i)
                        sb.Append(rephPass.Substring(lastCopied, i - lastCopied));
                    sb.Append('\u09CD');
                    sb.Append(rephPass[i]);
                    lastCopied = i + 2;
                    i += 2;
                    continue;
                }
                // Rule 2: র্ at word end → move before preceding consonant (keep ra-phala form র্)
                if (i + 1 < rephPass.Length && rephPass[i] == '\u09B0' && rephPass[i + 1] == '\u09CD'
                    && isEndOfWord)
                {
                    int consPos = i - 1;
                    while (consPos > 0 && IsPostKar(rephPass[consPos]))
                        consPos--;
                    if (consPos >= 0 && IsConsonant(rephPass[consPos]))
                    {
                        if (lastCopied < consPos)
                            sb.Append(rephPass.Substring(lastCopied, consPos - lastCopied));
                        sb.Append(rephPass[i]);
                        sb.Append(rephPass[i + 1]);
                        sb.Append(rephPass.Substring(consPos, i - consPos));
                        lastCopied = i + 2;
                        i += 2;
                        continue;
                    }
                }
                i++;
            }
            if (lastCopied < rephPass.Length)
                sb.Append(rephPass.Substring(lastCopied));
            // Third pass: remove consecutive halants (proConversionMap)
            var final = sb.ToString();
            sb.Clear();
            i = 0;
            while (i < final.Length)
            {
                if (i + 2 <= final.Length && final[i] == '\u09CD' && final[i + 1] == '\u09CD')
                {
                    sb.Append(final[i]);
                    i += 2;
                }
                else { sb.Append(final[i]); i++; }
            }
            return sb.ToString();
        }

        private static bool IsConsonant(char c)
        {
            return (c >= '\u0995' && c <= '\u09B9') || c == '\u09CE' ||
                   (c >= '\u09DC' && c <= '\u09DF');
        }
        private static bool IsPostKar(char c)
        {
            return c == '\u09BE' || c == '\u09C0' || c == '\u09C1' ||
                   c == '\u09C2' || c == '\u09C3' || c == '\u09D7';
        }
        private static string ConvertUnicodeToBijoy(string text)
        {
            // Step 1: NFC normalization
            text = text.Normalize(System.Text.NormalizationForm.FormC);

            // Step 2: decompose o-kar (U+09CB) → ে + া, au-kar (U+09CC) → ে + ৗ
            var sb1 = new System.Text.StringBuilder(text.Length);
            for (int j = 0; j < text.Length; j++)
            {
                if (text[j] == '\u09CB')
                { sb1.Append('\u09C7'); sb1.Append('\u09BE'); }
                else if (text[j] == '\u09CC')
                { sb1.Append('\u09C7'); sb1.Append('\u09D7'); }
                else
                    sb1.Append(text[j]);
            }
            text = sb1.ToString();

            // Step 3: decompose ড়/ঢ়/য়
            text = text.Replace("\u09DC", "\u09A1\u09BC")
                       .Replace("\u09DD", "\u09A2\u09BC")
                       .Replace("\u09DF", "\u09AF\u09BC");

            // Step 4: swap pre-kar (ি ে ৈ) with preceding consonant
            // Ensures pre-kar code comes before its consonant in output for SutonnyMJ
            var sb4 = new System.Text.StringBuilder(text);
            for (int j = 1; j < sb4.Length; j++)
            {
                if ((sb4[j] == '\u09BF' || sb4[j] == '\u09C7' || sb4[j] == '\u09C8')
                    && IsConsonant(sb4[j - 1]))
                {
                    char tmp = sb4[j];
                    sb4[j] = sb4[j - 1];
                    sb4[j - 1] = tmp;
                }
            }
            text = sb4.ToString();

            // Step 5: main pass — dictionary lookup (longest-match, reph, single char)
            var sb = new System.Text.StringBuilder();
            int i = 0;
            while (i < text.Length)
            {
                bool found = false;

                // 5a: try dictionary longest-match
                for (int len = Math.Min(8, text.Length - i); len >= 2; len--)
                {
                    var sub = text.Substring(i, len);
                    string b;
                    if (_unicodeToBijoy.TryGetValue(sub, out b))
                    {
                        sb.Append(b);
                        i += len;
                        found = true;
                        break;
                    }
                }
                if (found) continue;

                // 5b: reph + consonant → convert the consonant then append reph
                // Rule: র্ + consonant-cluster is reph ONLY when:
                //   - Cluster has "substance" (has post-kar or is multi-consonant), OR
                //   - Preceded by a consonant (not vowel sign) or start of word/independent vowel
                // NOT reph when: preceded by a vowel sign (any kar) and cluster is bare single consonant.
                bool isReph = false;
                if (IsRephAt(text, i))
                {
                    int clusterStart = i + 2;
                    if (clusterStart < text.Length && (IsConsonant(text[clusterStart]) || text[clusterStart] == '\u09DC' || text[clusterStart] == '\u09DD' || text[clusterStart] == '\u09DF'))
                    {
                        // Find end of the cluster (consonant + optional kars + optional halant+consonant...)
                        int clusterEnd = clusterStart;
                        bool hasSubstance = false;
                        while (clusterEnd < text.Length)
                        {
                            if (IsConsonant(text[clusterEnd]) || text[clusterEnd] == '\u09DC' || text[clusterEnd] == '\u09DD' || text[clusterEnd] == '\u09DF')
                            {
                                clusterEnd++;
                                // Optional kars after consonant
                                while (clusterEnd < text.Length && IsKar(text[clusterEnd]))
                                {
                                    clusterEnd++;
                                    hasSubstance = true;
                                }
                                // Check for halant + next consonant
                                if (clusterEnd + 1 < text.Length && text[clusterEnd] == '\u09CD' && 
                                    (IsConsonant(text[clusterEnd + 1]) || text[clusterEnd + 1] == '\u09DC' || text[clusterEnd + 1] == '\u09DD' || text[clusterEnd + 1] == '\u09DF'))
                                {
                                    clusterEnd += 2;
                                    hasSubstance = true;
                                }
                            }
                            break;
                        }
                        // Determine if this is truly reph:
                        bool precededByKar = i > 0 && IsKar(text[i - 1]);
                        bool precededByConsonant = i > 0 && IsConsonant(text[i - 1]);
                        // Reph when: cluster has substance, OR preceded by consonant
                        // (not vowel sign — like "কর্ম"), OR at start of text,
                        // OR preceded by independent vowel.
                        // NOT reph when: preceded by vowel sign (kar) and cluster is simple
                        // (like র্+জ in "ইনচার্জ").
                        if (hasSubstance || i == 0 || precededByConsonant ||
                            (i > 0 && (text[i - 1] == '\u0985' || text[i - 1] == '\u0986' ||
                             text[i - 1] == '\u0987' || text[i - 1] == '\u0988' ||
                             text[i - 1] == '\u0989' || text[i - 1] == '\u098A' ||
                             text[i - 1] == '\u098B' || text[i - 1] == '\u09E0' ||
                             text[i - 1] == '\u098C' || text[i - 1] == '\u09E1' ||
                             text[i - 1] == '\u098F' || text[i - 1] == '\u0990' ||
                             text[i - 1] == '\u0993' || text[i - 1] == '\u0994')))
                        {
                            // Convert the cluster (without reph)
                            string clusterText = text.Substring(clusterStart, clusterEnd - clusterStart);
                            string clusterResult = ConvertSimpleCluster(clusterText);
                            sb.Append(clusterResult);
                            sb.Append('\xA9');
                            i = clusterEnd;
                            found = true;
                            isReph = true;
                        }
                        // else: preceded by kar (like া in ইনচার্জ) + simple cluster → not reph
                    }
                }
                if (isReph) continue;

                // 5c: single char lookup
                string b1;
                if (_unicodeToBijoy.TryGetValue(text[i].ToString(), out b1))
                {
                    sb.Append(b1);
                    i++;
                    found = true;
                    continue;
                }

                // 5d: halant
                if (text[i] == '\u09CD')
                {
                    sb.Append('&');
                    i++;
                    found = true;
                    continue;
                }

                // 5e: fallback
                sb.Append(CharToBijoy(text[i]));
                i++;
            }

            // Step 6: Post-process - move reph (\xA9) to after its consonant cluster
            var resChars = sb.ToString().ToCharArray();
            bool moved;
            do
            {
                moved = false;
                for (int j = 0; j < resChars.Length - 1; j++)
                {
                    if (resChars[j] == '\xA9')
                    {
                        char next = resChars[j + 1];
                        // Don't swap past halant, punctuation, space, or non-Bengali chars
                        bool isPunct = next == ' ' || next == '\t' || next == ',' || next == '.' ||
                                       next == ';' || next == '!' || next == '?' || next == ':' ||
                                       next == '"' || next == '\'' || next == '(' || next == ')' ||
                                       next == '|' || next == '\r' || next == '\n' || next == '\xA9';
                        bool isHalant = next == '&';
                        if (!isPunct && !isHalant)
                        {
                            resChars[j] = next;
                            resChars[j + 1] = '\xA9';
                            moved = true;
                            break;
                        }
                    }
                }
            } while (moved);

            var result = new string(resChars);
            result = result.Replace('\u00D0', '-').Replace('\u00D1', '-');
            return result;
        }

        // Convert a simple consonant cluster (without reph) to Bijoy
        private static string ConvertSimpleCluster(string text)
        {
            var sb = new System.Text.StringBuilder();
            int i = 0;
            while (i < text.Length)
            {
                bool found = false;
                // Try multi-char dictionary match
                for (int len = Math.Min(6, text.Length - i); len >= 2; len--)
                {
                    var sub = text.Substring(i, len);
                    string b;
                    if (_unicodeToBijoy.TryGetValue(sub, out b))
                    {
                        sb.Append(b);
                        i += len;
                        found = true;
                        break;
                    }
                }
                if (found) continue;

                string b1;
                if (_unicodeToBijoy.TryGetValue(text[i].ToString(), out b1))
                {
                    sb.Append(b1);
                    i++;
                    continue;
                }
                if (text[i] == '\u09CD')
                {
                    sb.Append('&');
                    i++;
                    continue;
                }
                sb.Append(CharToBijoy(text[i]));
                i++;
            }
            return sb.ToString();
        }

        private static string CharToBijoy(char c)
        {
            string b;
            if (_unicodeToBijoy.TryGetValue(c.ToString(), out b))
                return b;
            if (c >= '\u09E6' && c <= '\u09EF')
                return ((char)('0' + (c - '\u09E6'))).ToString();
            if (c == '\u0982') return "s";
            if (c == '\u0983') return "t";
            if (c == '\u0981') return "u";
            if (c == '\u09BE') return "v";
            if (c == '\u09BF') return "w";
            if (c == '\u09C0') return "x";
            if (c == '\u09C1') return "y";
            if (c == '\u09C2') return "~";
            if (c == '\u09C3') return "\x84";
            if (c == '\u09C7') return "\x86";
            if (c == '\u09C8') return "\x88";
            if (c == '\u09D7') return "\x8A";
            if (c == '\u0964') return "|";
            if (c == '\u09CE') return "r";
            return c.ToString();
        }

        // Reph pattern: detects if position i starts with র্ (U+09B0 U+09CD)
        private static bool IsRephAt(string text, int i)
        {
            return i + 1 < text.Length && text[i] == '\u09B0' && text[i + 1] == '\u09CD';
        }

        // Returns true if c is a Bengali vowel sign (post-kar)
        private static bool IsKar(char c)
        {
            return c == '\u09BE' || c == '\u09BF' || c == '\u09C0' ||
                   c == '\u09C1' || c == '\u09C2' || c == '\u09C3' ||
                   c == '\u09C7' || c == '\u09C8' || c == '\u09D7';
        }

        private string TemplateFolder
        {
            get
            {
                try
                {
                    using (var key = Registry.CurrentUser.OpenSubKey(@"Software\Aksa10xFaster"))
                    {
                        if (key == null) return "";
                        return (string)key.GetValue("TemplateFolder", "");
                    }
                }
                catch { return ""; }
            }
            set
            {
                try
                {
                    using (var key = Registry.CurrentUser.CreateSubKey(@"Software\Aksa10xFaster"))
                    {
                        key.SetValue("TemplateFolder", value ?? "");
                    }
                }
                catch { }
            }
        }

        public void OnSelectFolder(Office.IRibbonControl c)
        {
            using (var dlg = new FolderBrowserDialog())
            {
                dlg.Description = "Select the folder containing agreement templates";
                dlg.SelectedPath = TemplateFolder;
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    TemplateFolder = dlg.SelectedPath;
                    MessageBox.Show("Template folder set to:\n" + dlg.SelectedPath, "AKSA 10X FASTER");
                }
            }
        }

        private static readonly Dictionary<string, string> _templateNames = new Dictionary<string, string>
        {
            { "btnLandSale", "Land Sale" },
            { "btnLandMortgage", "Land Mortgage" },
            { "btnLandRent", "Land Rent" },
            { "btnHouseRent", "House Rent" },
            { "btnHouseSale", "House Sale" },
            { "btnShopRent", "Shop Rent" },
            { "btnShopSale", "Shop Sale" },
            { "btnGardenMortgage", "Garden Mortgage" },
            { "btnGardenHouseMortgage", "Garden House Mortgage" },
            { "btnCarSale", "Car Sale" },
            { "btnMotorcycleSale", "Motorcycle Sale" },
            { "btnStorage", "Produce Storage" },
            { "btnGoingAbroad", "Going Abroad" },
            { "btnConstruction", "Construction" },
            { "btnBabyTaxiSale", "Baby Taxi Auto Sale" },
            { "btnMahr", "Mahr" },
            { "btnCngSale", "CNG Sale" },
            { "btnLandInstallment", "Land Installment" },
            { "btnPondMortgage", "Pond Mortgage" },
            { "btnPartnership", "Partnership Business" },
            { "btnLoanRepayment", "Loan Repayment" },
            { "btnBuildingMaterials", "Building Materials" },
            { "btnMoneyTransaction", "Money Transaction" },
            { "btnBoatSale", "Boat Trawler Sale" },
            { "btnPoultryFarmRent", "Poultry Farm Rent" },
            { "btnCattleFarmSale", "Cattle Farm Sale" },
            { "btnFosterCare", "Foster Care" },
        };

        private static readonly string[] AnnaChars = new string[]
        {
            "\u2044",              // 1 anna
            "\u09F5",              // 2 anna
            "\u09F6",              // 3 anna
            "\u09F7",              // 4 anna
            "\u09F7\u2044",        // 5 anna
            "\u09F7\u09F5",        // 6 anna
            "\u09F7\u09F6",        // 7 anna
            "\u09F7\u09F7",        // 8 anna
            "\u09F7\u09F7\u2044",  // 9 anna
            "\u09F7\u09F7\u09F5",  // 10 anna
            "\u09F7\u09F7\u09F6",  // 11 anna
            "\u09F8",              // 12 anna
            "\u09F8\u2044",        // 13 anna
            "\u09F8\u09F5",        // 14 anna
            "\u09F8\u09F6",        // 15 anna
            "\u09E7",              // 16 anna
        };

        public void OnAgreementPreset(Office.IRibbonControl c)
        {
            try
            {
                var id = c.Id;
                var page = "1";
                if (id.EndsWith("2")) page = "2";
                else if (id.EndsWith("3")) page = "3";
                var baseId = id.EndsWith("2") || id.EndsWith("3") ? id.Substring(0, id.Length - 1) : id;
                string templateName;
                if (!_templateNames.TryGetValue(baseId, out templateName))
                {
                    MessageBox.Show("Unknown template for: " + id, "AKSA 10X FASTER");
                    return;
                }
                var baseDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                var projRoot = Path.GetDirectoryName(baseDir);
                var subFolder = Path.Combine(projRoot, "Document Template", "Agreements", "Page " + page);
                var filePath = Path.Combine(subFolder, templateName + ".docx");
                if (!File.Exists(filePath))
                {
                    MessageBox.Show("File not found:\n" + filePath, "AKSA 10X FASTER");
                    return;
                }
                object templatePath = filePath;
                object missing = Type.Missing;
                _app.Documents.Add(ref templatePath, ref missing, ref missing, ref missing);
            }
            catch (System.Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }

        public void OnInsertAnna(Office.IRibbonControl c)
        {
            try
            {
                var s = Sel; if (s == null) return;
                var id = c.Id;
                int n;
                if (int.TryParse(id.Replace("btnAnna", ""), out n) && n >= 1 && n <= 16)
                {
                    s.Text = AnnaChars[n - 1];
                    s.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                }
            }
            catch (System.Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }

        // ===================== SAVE GROUP =====================

        // ===================== REGISTRY HELPERS =====================
        private static string GetRegStr(string name, string def)
        {
            using (var key = Registry.CurrentUser.OpenSubKey("Software\\Aksa10xFaster"))
                return key != null ? (key.GetValue(name, def) as string ?? def) : def;
        }
        private static void SetRegStr(string name, string value)
        {
            using (var key = Registry.CurrentUser.CreateSubKey("Software\\Aksa10xFaster"))
                key.SetValue(name, value);
        }

        // ===================== FOLDER SELECT DIALOG =====================
        public void OnFolderSelect(Office.IRibbonControl c)
        {
            string docFolder = GetRegStr("DocFolder", "");
            string pdfFolder = GetRegStr("PdfFolder", "");
            string jpgFolder = GetRegStr("JpgFolder", "");
            using (var form = new Form())
            {
                form.Text = "AKSA 10X FASTER - Save Folders";
                form.Size = new Size(620, 275);
                form.StartPosition = FormStartPosition.CenterScreen;
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.MaximizeBox = false;
                form.MinimizeBox = false;
                form.ShowIcon = false;
                form.ShowInTaskbar = false;
                var lblDoc = new Label { Text = "DOC Save Folder:", Left = 12, Top = 18, Width = 120 };
                var txtDoc = new TextBox { Text = docFolder, Left = 130, Top = 16, Width = 340 };
                var btnDoc = new Button { Text = "Browse...", Left = 475, Top = 15, Width = 80, Height = 23 };
                btnDoc.Click += delegate {
                    using (var fb = new FolderBrowserDialog { SelectedPath = txtDoc.Text })
                        if (fb.ShowDialog() == DialogResult.OK) txtDoc.Text = fb.SelectedPath;
                };
                var lblPdf = new Label { Text = "PDF Save Folder:", Left = 12, Top = 56, Width = 120 };
                var txtPdf = new TextBox { Text = pdfFolder, Left = 130, Top = 54, Width = 340 };
                var btnPdf = new Button { Text = "Browse...", Left = 475, Top = 53, Width = 80, Height = 23 };
                btnPdf.Click += delegate {
                    using (var fb = new FolderBrowserDialog { SelectedPath = txtPdf.Text })
                        if (fb.ShowDialog() == DialogResult.OK) txtPdf.Text = fb.SelectedPath;
                };
                var lblJpg = new Label { Text = "JPG Save Folder:", Left = 12, Top = 94, Width = 120 };
                var txtJpg = new TextBox { Text = jpgFolder, Left = 130, Top = 92, Width = 340 };
                var btnJpg = new Button { Text = "Browse...", Left = 475, Top = 91, Width = 80, Height = 23 };
                btnJpg.Click += delegate {
                    using (var fb = new FolderBrowserDialog { SelectedPath = txtJpg.Text })
                        if (fb.ShowDialog() == DialogResult.OK) txtJpg.Text = fb.SelectedPath;
                };
                var btnOk = new Button { Text = "Save", Left = 220, Top = 150, Width = 80, DialogResult = DialogResult.OK };
                var btnCancel = new Button { Text = "Cancel", Left = 310, Top = 150, Width = 80, DialogResult = DialogResult.Cancel };
                form.AcceptButton = btnOk;
                form.CancelButton = btnCancel;
                form.Controls.AddRange(new Control[] { lblDoc, txtDoc, btnDoc, lblPdf, txtPdf, btnPdf, lblJpg, txtJpg, btnJpg, btnOk, btnCancel });
                if (form.ShowDialog() == DialogResult.OK)
                {
                    SetRegStr("DocFolder", txtDoc.Text.Trim());
                    SetRegStr("PdfFolder", txtPdf.Text.Trim());
                    SetRegStr("JpgFolder", txtJpg.Text.Trim());
                    if (!string.IsNullOrEmpty(txtDoc.Text.Trim()) || !string.IsNullOrEmpty(txtPdf.Text.Trim()) || !string.IsNullOrEmpty(txtJpg.Text.Trim()))
                        MessageBox.Show("Folder settings saved.", "AKSA 10X FASTER");
                }
            }
        }

        // ===================== SAVE DOC =====================
        public void OnSaveDoc(Office.IRibbonControl c)
        {
            try
            {
                var d = _app.ActiveDocument;
                if (d == null) { MessageBox.Show("No document open.", "AKSA 10X FASTER"); return; }
                string folder = GetRegStr("DocFolder", "");
                if (string.IsNullOrEmpty(folder))
                {
                    if (MessageBox.Show("DOC folder not configured. Configure now?", "AKSA 10X FASTER", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        OnFolderSelect(c);
                    return;
                }
                string defName = Path.GetFileNameWithoutExtension(d.FullName);
                if (string.IsNullOrEmpty(defName)) defName = "Document";
                using (var dlg = new SaveFileDialog())
                {
                    dlg.InitialDirectory = folder;
                    dlg.FileName = defName + ".docx";
                    dlg.Filter = "Word documents|*.docx";
                    if (dlg.ShowDialog() == DialogResult.OK)
                    {
                        d.SaveAs2(dlg.FileName, Word.WdSaveFormat.wdFormatXMLDocument);
                        string savedPath = dlg.FileName;
                        MessageBox.Show("Saved:\n" + savedPath, "AKSA 10X FASTER");
                    }
                }
            }
            catch (System.Exception ex) { MessageBox.Show(ex.Message, "Save Doc Error"); }
        }

        // ===================== SAVE PDF =====================
        public void OnSavePdfFast(Office.IRibbonControl c)
        {
            try
            {
                var d = _app.ActiveDocument;
                if (d == null) { MessageBox.Show("No document open.", "AKSA 10X FASTER"); return; }
                string folder = GetRegStr("PdfFolder", "");
                if (string.IsNullOrEmpty(folder))
                {
                    if (MessageBox.Show("PDF folder not configured. Configure now?", "AKSA 10X FASTER", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        OnFolderSelect(c);
                    return;
                }
                string defName = Path.GetFileNameWithoutExtension(d.FullName);
                if (string.IsNullOrEmpty(defName)) defName = "Document";
                using (var dlg = new SaveFileDialog())
                {
                    dlg.InitialDirectory = folder;
                    dlg.FileName = defName + ".pdf";
                    dlg.Filter = "PDF files|*.pdf";
                    if (dlg.ShowDialog() == DialogResult.OK)
                    {
                        object extClassPtr = Type.Missing;
                        d.ExportAsFixedFormat(dlg.FileName, Word.WdExportFormat.wdExportFormatPDF, false, Word.WdExportOptimizeFor.wdExportOptimizeForPrint, Word.WdExportRange.wdExportAllDocument, 1, 1, Word.WdExportItem.wdExportDocumentContent, true, true, Word.WdExportCreateBookmarks.wdExportCreateWordBookmarks, true, true, false, ref extClassPtr);
                        MessageBox.Show("Saved:\n" + dlg.FileName, "AKSA 10X FASTER");
                    }
                }
            }
            catch (System.Exception ex) { MessageBox.Show(ex.Message, "Save PDF Error"); }
        }

        // ===================== SAVE JPG =====================
        public void OnSaveJpgFast(Office.IRibbonControl c)
        {
            try
            {
                var d = _app.ActiveDocument;
                if (d == null) { MessageBox.Show("No document open.", "AKSA 10X FASTER"); return; }
                string folder = GetRegStr("JpgFolder", "");
                if (string.IsNullOrEmpty(folder))
                {
                    if (MessageBox.Show("JPG folder not configured. Configure now?", "AKSA 10X FASTER", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        OnFolderSelect(c);
                    return;
                }
                if (!Directory.Exists(folder))
                {
                    if (MessageBox.Show("JPG folder does not exist:\n" + folder + "\nCreate it?", "AKSA 10X FASTER", MessageBoxButtons.YesNo) != DialogResult.Yes) return;
                    Directory.CreateDirectory(folder);
                }
                int pageCount = d.ComputeStatistics(Word.WdStatistic.wdStatisticPages);
                object miss = Type.Missing;
                string baseName = Path.GetFileNameWithoutExtension(d.FullName);
                if (string.IsNullOrEmpty(baseName)) baseName = "Document";
                float pageW = d.PageSetup.PageWidth;
                float pageH = d.PageSetup.PageHeight;
                float marginL = d.PageSetup.LeftMargin;
                float marginT = d.PageSetup.TopMargin;
                float marginR = d.PageSetup.RightMargin;
                float marginB = d.PageSetup.BottomMargin;
                float dpiScale = 300f / 72f;
                int bmpW = (int)(pageW * dpiScale);
                int bmpH = (int)(pageH * dpiScale);
                if (bmpW > 6000) { float s = 6000f / bmpW; bmpW = 6000; bmpH = (int)(bmpH * s); }
                if (bmpH > 6000) { float s = 6000f / bmpH; bmpH = 6000; bmpW = (int)(bmpW * s); }
                if (bmpW < 1) bmpW = 1;
                if (bmpH < 1) bmpH = 1;
                object what = Word.WdGoToItem.wdGoToPage;
                object which = Word.WdGoToDirection.wdGoToAbsolute;
                int saved = 0;
                for (int pg = 1; pg <= pageCount; pg++)
                {
                    object count = pg;
                    Word.Range pgStart = _app.Selection.GoTo(ref what, ref which, ref count, ref miss) as Word.Range;
                    if (pgStart == null) continue;
                    int sPos = pgStart.Start;
                    int ePos;
                    if (pg < pageCount)
                    {
                        object countNext = pg + 1;
                        Word.Range nextStart = _app.Selection.GoTo(ref what, ref which, ref countNext, ref miss) as Word.Range;
                        if (nextStart == null) continue;
                        ePos = Math.Max(sPos, nextStart.Start - 1);
                    }
                    else
                    {
                        ePos = Math.Max(sPos, d.Content.End - 1);
                    }
                    Word.Range pgRange = d.Range(sPos, ePos);
                    var emfBytes = pgRange.EnhMetaFileBits as byte[];
                    if (emfBytes == null || emfBytes.Length == 0) continue;
                    using (var ms = new MemoryStream(emfBytes))
                    {
                        var emf = new System.Drawing.Imaging.Metafile(ms);
                        var dpiX = emf.HorizontalResolution;
                        var dpiY = emf.VerticalResolution;
                        if (dpiX < 1) dpiX = 96;
                        if (dpiY < 1) dpiY = 96;
                        int emfPixW = Math.Max(1, (int)(emf.Width * 300f / dpiX));
                        int emfPixH = Math.Max(1, (int)(emf.Height * 300f / dpiY));
                        using (var bmp = new Bitmap(bmpW, bmpH))
                        {
                            using (var g = Graphics.FromImage(bmp))
                            {
                                g.Clear(Color.White);
                                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
                                int marginLPix = (int)(marginL * dpiScale);
                                int marginTPix = (int)(marginT * dpiScale);
                                int marginRPix = (int)(marginR * dpiScale);
                                int marginBPix = (int)(marginB * dpiScale);
                                int contentAreaW = bmpW - marginLPix - marginRPix;
                                int contentAreaH = bmpH - marginTPix - marginBPix;
                                float scale = (float)contentAreaW / emfPixW;
                                if (emfPixH * scale > contentAreaH) scale = (float)contentAreaH / emfPixH;
                                if (scale > 1f) scale = 1f;
                                int drawW = Math.Max(1, (int)(emfPixW * scale));
                                int drawH = Math.Max(1, (int)(emfPixH * scale));
                                int drawX = marginLPix + (contentAreaW - drawW) / 2;
                                int drawY = marginTPix + (contentAreaH - drawH) / 2;
                                g.DrawImage(emf, drawX, drawY, drawW, drawH);
                            }
                            string fileName = pageCount > 1
                                ? string.Format("{0}_page{1}.jpg", baseName, pg)
                                : string.Format("{0}.jpg", baseName);
                            bmp.Save(Path.Combine(folder, fileName), System.Drawing.Imaging.ImageFormat.Jpeg);
                        }
                    }
                    saved++;
                }
                Process.Start("explorer.exe", folder);
                MessageBox.Show(string.Format("Saved {0} page(s) as JPG.", saved), "AKSA 10X FASTER");
            }
            catch (System.Exception ex) { MessageBox.Show(ex.Message, "Save JPG Error"); }
        }

        public void OnSaveAsJpg(Office.IRibbonControl c)
        {
            try
            {
                var d = _app.ActiveDocument;
                if (d == null) { MessageBox.Show("No document open.", "AKSA 10X FASTER"); return; }
                int pageCount = d.ComputeStatistics(Word.WdStatistic.wdStatisticPages);
                object miss = Type.Missing;
                using (var dlg = new FolderBrowserDialog())
                {
                    dlg.Description = "Select folder to save JPG images";
                    if (dlg.ShowDialog() != DialogResult.OK) return;
                    string folder = dlg.SelectedPath;
                    string baseName = Path.GetFileNameWithoutExtension(d.FullName);
                    if (string.IsNullOrEmpty(baseName)) baseName = "Document";
                    float pageW = d.PageSetup.PageWidth;
                    float pageH = d.PageSetup.PageHeight;
                    float marginL = d.PageSetup.LeftMargin;
                    float marginT = d.PageSetup.TopMargin;
                    float marginR = d.PageSetup.RightMargin;
                    float marginB = d.PageSetup.BottomMargin;
                    float dpiScale = 300f / 72f;
                    int bmpW = (int)(pageW * dpiScale);
                    int bmpH = (int)(pageH * dpiScale);
                    if (bmpW > 6000) { float s = 6000f / bmpW; bmpW = 6000; bmpH = (int)(bmpH * s); }
                    if (bmpH > 6000) { float s = 6000f / bmpH; bmpH = 6000; bmpW = (int)(bmpW * s); }
                    if (bmpW < 1) bmpW = 1;
                    if (bmpH < 1) bmpH = 1;
                    object what = Word.WdGoToItem.wdGoToPage;
                    object which = Word.WdGoToDirection.wdGoToAbsolute;
                    for (int pg = 1; pg <= pageCount; pg++)
                    {
                        object count = pg;
                        Word.Range pgStart = _app.Selection.GoTo(ref what, ref which, ref count, ref miss) as Word.Range;
                        if (pgStart == null) continue;
                        int sPos = pgStart.Start;
                        int ePos;
                        if (pg < pageCount)
                        {
                            object countNext = pg + 1;
                            Word.Range nextStart = _app.Selection.GoTo(ref what, ref which, ref countNext, ref miss) as Word.Range;
                            if (nextStart == null) continue;
                            ePos = Math.Max(sPos, nextStart.Start - 1);
                        }
                        else
                        {
                            ePos = Math.Max(sPos, d.Content.End - 1);
                        }
                        Word.Range pgRange = d.Range(sPos, ePos);
                        var emfBytes = pgRange.EnhMetaFileBits as byte[];
                        if (emfBytes == null || emfBytes.Length == 0) continue;
                        using (var ms = new MemoryStream(emfBytes))
                        {
                            var emf = new System.Drawing.Imaging.Metafile(ms);
                            var dpiX = emf.HorizontalResolution;
                            var dpiY = emf.VerticalResolution;
                            if (dpiX < 1) dpiX = 96;
                            if (dpiY < 1) dpiY = 96;
                            int emfPixW = Math.Max(1, (int)(emf.Width * 300f / dpiX));
                            int emfPixH = Math.Max(1, (int)(emf.Height * 300f / dpiY));
                            using (var bmp = new Bitmap(bmpW, bmpH))
                            {
                                using (var g = Graphics.FromImage(bmp))
                                {
                                    g.Clear(Color.White);
                                    g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                    g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                                    g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
                                    int marginLPix = (int)(marginL * dpiScale);
                                    int marginTPix = (int)(marginT * dpiScale);
                                    int marginRPix = (int)(marginR * dpiScale);
                                    int marginBPix = (int)(marginB * dpiScale);
                                    int contentAreaW = bmpW - marginLPix - marginRPix;
                                    int contentAreaH = bmpH - marginTPix - marginBPix;
                                    float scale = (float)contentAreaW / emfPixW;
                                    if (emfPixH * scale > contentAreaH) scale = (float)contentAreaH / emfPixH;
                                    if (scale > 1f) scale = 1f;
                                    int drawW = Math.Max(1, (int)(emfPixW * scale));
                                    int drawH = Math.Max(1, (int)(emfPixH * scale));
                                    int drawX = marginLPix + (contentAreaW - drawW) / 2;
                                    int drawY = marginTPix + (contentAreaH - drawH) / 2;
                                    g.DrawImage(emf, drawX, drawY, drawW, drawH);
                                }
                                string fileName = pageCount > 1
                                    ? string.Format("{0}_page{1}.jpg", baseName, pg)
                                    : string.Format("{0}.jpg", baseName);
                                var encParams = new System.Drawing.Imaging.EncoderParameters(1);
                                encParams.Param[0] = new System.Drawing.Imaging.EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 100L);
                                var jpgCodec = System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders()
                                    .FirstOrDefault(codec => codec.FormatID == System.Drawing.Imaging.ImageFormat.Jpeg.Guid);
                                if (jpgCodec != null)
                                    bmp.Save(Path.Combine(folder, fileName), jpgCodec, encParams);
                                else
                                    bmp.Save(Path.Combine(folder, fileName), System.Drawing.Imaging.ImageFormat.Jpeg);
                            }
                        }
                    }
                    MessageBox.Show(string.Format("{0} page(s) saved as JPG.", pageCount), "AKSA 10X FASTER");
                }
            }
            catch (System.Exception ex) { MessageBox.Show(ex.Message, "Save as JPG Error"); }
        }
        public void OnSaveCopy(Office.IRibbonControl c)
        {
            var d = Doc; if (d == null) return;
            using (var diag = new SaveFileDialog())
            {
                diag.Filter = "Word documents|*.docx";
                diag.FileName = Path.GetFileNameWithoutExtension(d.FullName) + "_copy.docx";
                if (diag.ShowDialog() == DialogResult.OK)
                    d.SaveAs2(diag.FileName, Word.WdSaveFormat.wdFormatXMLDocument);
            }
        }

        public void OnSaveCloseAll(Office.IRibbonControl c)
        {
            int count = 0;
            while (_app.Documents.Count > 0)
            {
                var doc = _app.Documents[1] as Word._Document;
                if (doc != null) { if (doc.Saved == false) doc.Save(); doc.Close(); }
                count++;
            }
            MessageBox.Show("Saved and closed " + count + " document(s).", "AKSA 10X FASTER");
        }

        public void OnSearchChanged(Office.IRibbonControl c, string text)
        {
            try
            {
                string folder = GetRegStr("DocFolder", "");
                if (string.IsNullOrEmpty(folder))
                {
                    MessageBox.Show("Please configure Save Doc folder first using the 'Folder' button.", "AKSA 10X FASTER");
                    return;
                }
                if (!Directory.Exists(folder))
                {
                    MessageBox.Show("Doc folder does not exist:\n" + folder, "AKSA 10X FASTER");
                    return;
                }
                var files = Directory.GetFiles(folder, "*.docx")
                    .Where(f => Path.GetFileNameWithoutExtension(f).IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0)
                    .OrderBy(f => f)
                    .ToArray();
                if (files.Length == 0)
                {
                    MessageBox.Show("No matching files found for:\n" + text, "Search");
                    return;
                }
                using (var dlg = new SearchDialog(files, text))
                {
                    if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK && dlg.SelectedFile != null)
                    {
                        object filePath = dlg.SelectedFile;
                        object missing = Type.Missing;
                        _app.Documents.Open(ref filePath, ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing, ref missing, ref missing,
                            ref missing, ref missing);
                    }
                }
            }
            catch (System.Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }

        private class SearchDialog : System.Windows.Forms.Form
        {
            private System.Windows.Forms.TextBox _txtSearch;
            private System.Windows.Forms.ListBox _lstFiles;
            private System.Windows.Forms.Button _btnOpen;
            private System.Windows.Forms.Button _btnCancel;
            private string[] _allFiles;
            public string SelectedFile { get; private set; }

            public SearchDialog(string[] files, string initialText)
            {
                _allFiles = files;
                SelectedFile = null;
                Text = "Search Documents";
                StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
                Width = 500;
                Height = 400;
                MinimumSize = new System.Drawing.Size(300, 200);

                _txtSearch = new System.Windows.Forms.TextBox();
                _txtSearch.Dock = System.Windows.Forms.DockStyle.Top;
                _txtSearch.Height = 30;
                _txtSearch.Font = new System.Drawing.Font("Segoe UI", 11);
                _txtSearch.Text = initialText;
                _txtSearch.TextChanged += OnTextChanged;

                _lstFiles = new System.Windows.Forms.ListBox();
                _lstFiles.Dock = System.Windows.Forms.DockStyle.Fill;
                _lstFiles.Font = new System.Drawing.Font("Segoe UI", 10);
                _lstFiles.DoubleClick += OnItemDoubleClick;

                var bottomPanel = new System.Windows.Forms.Panel();
                bottomPanel.Dock = System.Windows.Forms.DockStyle.Bottom;
                bottomPanel.Height = 40;
                bottomPanel.Padding = new System.Windows.Forms.Padding(5);

                _btnOpen = new System.Windows.Forms.Button();
                _btnOpen.Text = "Open";
                _btnOpen.DialogResult = System.Windows.Forms.DialogResult.OK;
                _btnOpen.Click += OnOpenClick;
                _btnOpen.Anchor = System.Windows.Forms.AnchorStyles.Right;
                _btnOpen.Size = new System.Drawing.Size(80, 30);
                _btnOpen.Location = new System.Drawing.Point(bottomPanel.Width - 170, 5);

                _btnCancel = new System.Windows.Forms.Button();
                _btnCancel.Text = "Cancel";
                _btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
                _btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Right;
                _btnCancel.Size = new System.Drawing.Size(80, 30);
                _btnCancel.Location = new System.Drawing.Point(bottomPanel.Width - 85, 5);

                bottomPanel.Controls.Add(_btnOpen);
                bottomPanel.Controls.Add(_btnCancel);

                Controls.Add(_lstFiles);
                Controls.Add(_txtSearch);
                Controls.Add(bottomPanel);

                FilterList(initialText);
                if (_lstFiles.Items.Count > 0) _lstFiles.SelectedIndex = 0;
            }

            private void OnTextChanged(object sender, EventArgs e)
            {
                FilterList(_txtSearch.Text);
            }

            private void FilterList(string filter)
            {
                _lstFiles.Items.Clear();
                foreach (var f in _allFiles)
                {
                    if (Path.GetFileNameWithoutExtension(f).IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0)
                        _lstFiles.Items.Add(f);
                }
                if (_lstFiles.Items.Count > 0) _lstFiles.SelectedIndex = 0;
            }

            private void OnItemDoubleClick(object sender, EventArgs e)
            {
                if (_lstFiles.SelectedItem != null)
                {
                    SelectedFile = _lstFiles.SelectedItem.ToString();
                    DialogResult = System.Windows.Forms.DialogResult.OK;
                    Close();
                }
            }

            private void OnOpenClick(object sender, EventArgs e)
            {
                if (_lstFiles.SelectedItem != null)
                {
                    SelectedFile = _lstFiles.SelectedItem.ToString();
                }
            }
        }

        // ===================== AGREEMENTS GROUP =====================

        public void OnSignatureLine(Office.IRibbonControl c)
        {
            var d = Doc; if (d == null) return;
            var s = Sel;
            s.TypeParagraph();
            s.TypeText("__________________________");
            s.TypeParagraph();
            s.TypeText("Signature");
            s.TypeParagraph();
            s.TypeText("Date: " + DateTime.Now.ToString("dd/MM/yyyy"));
            s.TypeParagraph();
        }

        public void OnInsertDate(Office.IRibbonControl c)
        {
            try
            {
                var sel = _app.Selection;
                if (sel == null) return;
                var now = DateTime.Now;
                string dateStr;
                if (c.Id == "btnDateBn")
                {
                    // Bangla calendar: year starts April 14
                    int[] monthDays = { 31, 31, 31, 31, 31, 31, 30, 30, 30, 30, 30, 30 };
                    string[] bnMonths = { "বৈশাখ", "জ্যৈষ্ঠ", "আষাঢ়", "শ্রাবণ", "ভাদ্র", "আশ্বিন", "কার্তিক", "অগ্রহায়ণ", "পৌষ", "মাঘ", "ফাল্গুন", "চৈত্র" };
                    var banglaNewYear = new DateTime(now.Year, 4, 14);
                    if (now < banglaNewYear)
                        banglaNewYear = new DateTime(now.Year - 1, 4, 14);
                    int banglaYear = banglaNewYear.Year - 593;
                    int dayOfYear = (int)(now - banglaNewYear).TotalDays;
                    int banglaMonth = 0;
                    while (dayOfYear >= monthDays[banglaMonth])
                    { dayOfYear -= monthDays[banglaMonth]; banglaMonth++; }
                    int banglaDay = dayOfYear + 1;
                    dateStr = ArabicToBengaliDigits(banglaDay) + " " + bnMonths[banglaMonth] + " " + ArabicToBengaliDigits(banglaYear);
                }
                else
                {
                    dateStr = now.ToString("dddd, MMMM dd, yyyy");
                }
                sel.Text = dateStr;
            }
            catch (System.Exception ex) { MessageBox.Show(ex.Message, "Date Error"); }
        }

        private void ToggleWatermark(string text)
        {
            var d = Doc; if (d == null) return;
            foreach (Word.Shape shp in d.Shapes)
            {
                if (shp.Type == Microsoft.Office.Core.MsoShapeType.msoTextBox && shp.TextFrame.HasText != 0)
                {
                    if (shp.TextFrame.TextRange.Text.Trim().ToUpper() == text)
                    { shp.Delete(); return; }
                }
            }
            var range = d.Content;
            var sh = d.Shapes.AddTextbox(
                Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                0, 0, _app.CentimetersToPoints(15), _app.CentimetersToPoints(3));
            sh.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            sh.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            sh.TextFrame.TextRange.Text = text;
            sh.TextFrame.TextRange.Font.Size = 72;
            sh.TextFrame.TextRange.Font.Color = Word.WdColor.wdColorGray25;
            sh.TextFrame.TextRange.Font.Name = "Arial";
            sh.Rotation = -45;
            sh.RelativeHorizontalPosition = Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin;
            sh.RelativeVerticalPosition = Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionMargin;
            sh.Left = (float)Word.WdShapePosition.wdShapeCenter;
            sh.Top = (float)Word.WdShapePosition.wdShapeCenter;
        }
        public void OnDraftWatermark(Office.IRibbonControl c) { ToggleWatermark("DRAFT"); }

        public void OnConfidentialWatermark(Office.IRibbonControl c) { ToggleWatermark("CONFIDENTIAL"); }

        public void OnClauseNumbering(Office.IRibbonControl c)
        {
            var d = Doc; if (d == null) return;
            var s = Sel;
            s.Range.ListFormat.ApplyNumberDefault(Word.WdDefaultListBehavior.wdWord10ListBehavior);
        }

        public void OnInsertDisclaimer(Office.IRibbonControl c)
        {
            var s = Sel;
            s.TypeParagraph();
            s.TypeText("DISCLAIMER");
            s.TypeParagraph();
            s.TypeText("This document is intended for informational purposes only and does not " +
                "constitute legal, financial, or professional advice. The author makes no " +
                "representations as to the accuracy or completeness of any information contained " +
                "herein and will not be liable for any errors or omissions.");
            s.TypeParagraph();
            var rng = s.Range;
            rng.Font.Size = 8;
            rng.Font.Italic = 1;
            rng.Font.Color = Word.WdColor.wdColorGray50;
        }

        // ===================== CV & RESUME GROUP =====================

        private void ApplyStyle(Word.WdBuiltinStyle style)
        {
            var d = Doc; if (d == null) return;
            Sel.set_Style(style);
        }

        private void InsertTemplateText(string title, string[] sections)
        {
            var d = Doc; if (d == null) return;
            d.Content.Select();
            _app.Selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);

            var rng = d.Content;
            rng.Collapse(Word.WdCollapseDirection.wdCollapseStart);

            rng.Text = title;
            rng.Font.Size = 26;
            rng.Font.Bold = 1;
            rng.Font.Name = "Calibri";
            rng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            rng.InsertParagraphAfter();
            rng.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            foreach (var section in sections)
            {
                rng.InsertParagraphAfter();
                rng.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                var parts = section.Split('|');
                rng.Text = parts[0];
                rng.Font.Size = 16;
                rng.Font.Bold = 1;
                rng.Font.Name = "Calibri";
                rng.Font.Color = Word.WdColor.wdColorDarkBlue;
                rng.InsertParagraphAfter();
                rng.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                if (parts.Length > 1)
                {
                    rng.Text = parts[1];
                    rng.Font.Size = 11;
                    rng.Font.Bold = 0;
                    rng.Font.Name = "Calibri";
                    rng.Font.Color = Word.WdColor.wdColorAutomatic;
                    rng.InsertParagraphAfter();
                    rng.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                }
            }
        }

        public void OnCvProfessional(Office.IRibbonControl c)
        {
            InsertTemplateText("CURRICULUM VITAE", new[] {
                "Personal Information|Name  |  Email  |  Phone  |  Address",
                "Professional Summary|Write a brief summary of your professional background, key skills, and career objectives.",
                "Work Experience|Job Title  |  Company Name  |  Dates\nDescribe your responsibilities and achievements.",
                "Education|Degree  |  Institution  |  Year",
                "Skills|List your key technical and professional skills.",
                "Certifications|Certification Name  |  Issuing Organization  |  Year",
                "Languages|Language  |  Proficiency Level",
                "References|Available upon request."
            });
        }

        public void OnCvModern(Office.IRibbonControl c)
        {
            InsertTemplateText("RESUME", new[] {
                "Contact|Name  |  Phone  |  Email  |  LinkedIn  |  Portfolio",
                "Profile|A brief professional summary highlighting your unique value proposition.",
                "Core Competencies|Skill 1  |  Skill 2  |  Skill 3  |  Skill 4  |  Skill 5",
                "Professional Experience|Job Title - Company (Year-Year)\nKey achievement 1\nKey achievement 2\nKey achievement 3",
                "Education|Degree - Institution (Year)",
                "Projects|Project Name - Brief description of the project and your role.",
                "Awards & Recognition|Award 1  |  Award 2",
                "Interests|Interest 1  |  Interest 2  |  Interest 3"
            });
        }

        public void OnCoverLetter(Office.IRibbonControl c)
        {
            InsertTemplateText("COVER LETTER", new[] {
                "Contact Information|Your Name\nPhone | Email | LinkedIn",
                "Date|" + DateTime.Now.ToString("dd MMMM yyyy"),
                "Recipient|Hiring Manager\nCompany Name\nCompany Address",
                "Subject|Application for [Position Name]",
                "Dear Hiring Manager,|",
                "Body Paragraph 1|I am writing to express my strong interest in the [Position Name] role at [Company Name]. With my background in [field] and proven track record of [key achievement], I am confident in my ability to contribute effectively to your team.",
                "Body Paragraph 2|In my previous role at [Previous Company], I successfully [specific accomplishment]. This experience has equipped me with [key skills] that I believe would be valuable to [Company Name].",
                "Body Paragraph 3|I am particularly drawn to this opportunity because [reason for interest]. I am eager to bring my expertise to your organization and help drive [company goal].",
                "Closing|Thank you for considering my application. I look forward to the opportunity to discuss how my skills and experience align with the needs of your team.\n\nSincerely,\n[Your Name]"
            });
        }

        public void OnInsertSkills(Office.IRibbonControl c)
        {
            var d = Doc; if (d == null) return;
            var s = Sel;
            s.TypeText("Skills");
            s.TypeParagraph();
            var tbl = d.Tables.Add(s.Range, 5, 3);
            tbl.set_Style("Grid Table 4 - Accent 1");
            tbl.Borders.Enable = 1;
            var headers = new[] { "Category", "Skill", "Proficiency" };
            for (int i = 0; i < 3; i++)
            {
                tbl.Cell(1, i + 1).Range.Text = headers[i];
                tbl.Cell(1, i + 1).Range.Font.Bold = 1;
            }
            for (int r = 2; r <= 5; r++)
            {
                tbl.Cell(r, 1).Range.Text = "Category " + (r - 1);
                tbl.Cell(r, 2).Range.Text = "Skill " + (r - 1);
                tbl.Cell(r, 3).Range.Text = "Beginner / Intermediate / Advanced";
            }
        }

        public void OnInsertExperience(Office.IRibbonControl c)
        {
            var s = Sel;
            s.TypeText("Professional Experience");
            s.TypeParagraph();
            s.set_Style(Word.WdBuiltinStyle.wdStyleHeading2);
            s.TypeParagraph();
            s.TypeText("Job Title  |  Company Name  |  Start Date - End Date");
            s.TypeParagraph();
            s.TypeText("• Key responsibility or achievement 1");
            s.TypeParagraph();
            s.TypeText("• Key responsibility or achievement 2");
            s.TypeParagraph();
            s.TypeText("• Key responsibility or achievement 3");
            s.TypeParagraph();
        }

        public void OnInsertEducation(Office.IRibbonControl c)
        {
            var s = Sel;
            s.TypeText("Education");
            s.TypeParagraph();
            s.set_Style(Word.WdBuiltinStyle.wdStyleHeading2);
            s.TypeParagraph();
            s.TypeText("Degree Name  |  Institution  |  Year of Graduation");
            s.TypeParagraph();
            s.TypeText("• Relevant coursework or achievement");
            s.TypeParagraph();
            s.TypeText("• GPA or academic honor (if applicable)");
            s.TypeParagraph();
        }

        // ===================== IMAGES =====================

        private Dictionary<string, IPictureDisp> _iconCache = new Dictionary<string, IPictureDisp>();

        private static Color ColorForKey(string key)
        {
            if (key == null) return Color.FromArgb(128, 128, 128);
            switch (key)
            {
                case "page_setup":
                case "btnPageSetup":
                case "agree_page":
                case "btnAgreementPage":
                case "cv_page":
                case "btnCvPage":
                case "app_page":
                case "btnApplicationPage":
                case "question_size":
                case "btnQuestionPage":
                case "number_word":
                case "btnNumberToWord":
                case "nb_word":
                case "btnNumberToWordBn":
                case "font_conv":
                case "mnuConvertFont":
                case "b2u":
                case "btnB2U":
                case "u2b":
                case "btnU2B":
                case "margins":
                case "btnMarginNormal":
                case "btnMarginNarrow":
                case "btnMarginModerate":
                case "btnMarginWide":
                case "btnMarginMirrored":
                case "orientation":
                case "btnPortrait":
                case "btnLandscape":
                case "page_size":
                case "btnA4":
                case "btnLetter":
                case "btnLegal":
                case "btnA3":
                case "page_break":
                case "btnPageBreak":
                case "section_break":
                case "btnSectionBreak":
                case "line_numbers":
                case "btnLineNumbers":
                    return Color.FromArgb(68, 114, 196);
                case "pdf":
                case "btnSavePdf":
                case "docx":
                case "btnSaveDocx":
                case "txt":
                case "btnSaveTxt":
                case "jpg":
                case "btnSaveJpg":
                case "save_copy":
                case "btnSaveCopy":
                case "save_doc":
                case "save_pdf":
                case "save_jpg":
                case "close_all":
                case "save_close":
                case "btnSaveCloseAll":
                case "btnSavePdfFast":
                case "btnSaveDocFast":
                case "btnSaveJpgFast":
                case "save_search":
                case "btnSearchDocFolder":
                    return Color.FromArgb(112, 173, 71);
                case "folder_sel":
                case "save_folder":
                case "btnFolderSelect":
                    return Color.FromArgb(255, 192, 0);
                case "bengali":
                case "mnuBengali":
                case "english":
                case "mnuEnglish":
                case "select_folder":
                case "template_folder":
                case "list":
                case "btnSelectFolder":
                case "page1":
                case "mnuPage1":
                case "page2":
                case "mnuPage2":
                case "page3":
                case "mnuPage3":
                case "land_sale":
                case "btnLandSale":
                case "btnLandSale2":
                case "btnLandSale3":
                case "land_mortgage":
                case "btnLandMortgage":
                case "btnLandMortgage2":
                case "btnLandMortgage3":
                case "land_rent":
                case "btnLandRent":
                case "btnLandRent2":
                case "btnLandRent3":
                case "house_rent":
                case "btnHouseRent":
                case "btnHouseRent2":
                case "btnHouseRent3":
                case "house_sale":
                case "btnHouseSale":
                case "btnHouseSale2":
                case "btnHouseSale3":
                case "shop_rent":
                case "btnShopRent":
                case "btnShopRent2":
                case "btnShopRent3":
                case "shop_sale":
                case "btnShopSale":
                case "btnShopSale2":
                case "btnShopSale3":
                case "garden_mortgage":
                case "btnGardenMortgage":
                case "btnGardenMortgage2":
                case "btnGardenMortgage3":
                case "garden_house_mortgage":
                case "btnGardenHouseMortgage":
                case "btnGardenHouseMortgage2":
                case "btnGardenHouseMortgage3":
                case "car_sale":
                case "btnCarSale":
                case "btnCarSale2":
                case "btnCarSale3":
                case "motorcycle_sale":
                case "btnMotorcycleSale":
                case "btnMotorcycleSale2":
                case "btnMotorcycleSale3":
                case "storage":
                case "btnStorage":
                case "btnStorage2":
                case "btnStorage3":
                case "going_abroad":
                case "btnGoingAbroad":
                case "btnGoingAbroad2":
                case "btnGoingAbroad3":
                case "construction":
                case "btnConstruction":
                case "btnConstruction2":
                case "btnConstruction3":
                case "baby_taxi_sale":
                case "btnBabyTaxiSale":
                case "btnBabyTaxiSale2":
                case "btnBabyTaxiSale3":
                case "mahr":
                case "btnMahr":
                case "btnMahr2":
                case "btnMahr3":
                case "cng_sale":
                case "btnCngSale":
                case "btnCngSale2":
                case "btnCngSale3":
                case "land_installment":
                case "btnLandInstallment":
                case "btnLandInstallment2":
                case "btnLandInstallment3":
                case "pond_mortgage":
                case "btnPondMortgage":
                case "btnPondMortgage2":
                case "btnPondMortgage3":
                case "partnership":
                case "btnPartnership":
                case "btnPartnership2":
                case "btnPartnership3":
                case "loan_repayment":
                case "btnLoanRepayment":
                case "btnLoanRepayment2":
                case "btnLoanRepayment3":
                case "building_materials":
                case "btnBuildingMaterials":
                case "btnBuildingMaterials2":
                case "btnBuildingMaterials3":
                case "money_transaction":
                case "btnMoneyTransaction":
                case "btnMoneyTransaction2":
                case "btnMoneyTransaction3":
                case "boat_sale":
                case "btnBoatSale":
                case "btnBoatSale2":
                case "btnBoatSale3":
                case "poultry_farm_rent":
                case "btnPoultryFarmRent":
                case "btnPoultryFarmRent2":
                case "btnPoultryFarmRent3":
                case "cattle_farm_sale":
                case "btnCattleFarmSale":
                case "btnCattleFarmSale2":
                case "btnCattleFarmSale3":
                case "foster_care":
                case "btnFosterCare":
                case "btnFosterCare2":
                case "btnFosterCare3":
                case "anna_menu":
                case "anna_item":
                case "signature":
                case "btnSignatureLine":
                case "date":
                case "btnInsertDate":
                case "draft":
                case "btnDraftWatermark":
                case "confidential":
                case "btnConfidentialWatermark":
                case "clause":
                case "btnClauseNumbering":
                case "disclaimer":
                case "btnDisclaimer":
                    return Color.FromArgb(237, 125, 49);
                case "professional":
                case "btnProfessional":
                case "modern":
                case "btnModern":
                case "cover":
                case "btnCoverLetter":
                case "skills":
                case "btnSkills":
                case "experience":
                case "btnExperience":
                case "education":
                case "btnEducation":
                    return Color.FromArgb(112, 48, 160);
                case "ai_write":
                case "btnAiWrite":
                case "ai_fix":
                case "btnAiFix":
                case "ai_summarize":
                case "btnAiSummarize":
                case "aksa_ai":
                case "mnuAksaAI":
                    return Color.FromArgb(0, 150, 136);
                case "register":
                case "btnRegister":
                    return Color.FromArgb(255, 185, 0);
                default:
                    return Color.FromArgb(128, 128, 128);
            }
        }

        private static string LabelForKey(string key)
        {
            if (key == null) return "";
            switch (key)
            {
                case "page_setup":
                case "btnPageSetup": return "PS";
                case "agree_page":
                case "btnAgreementPage": return "AG";
                case "cv_page":
                case "btnCvPage": return "CV";
                case "app_page":
                case "btnApplicationPage": return "AJ";
                case "question_size":
                case "btnQuestionPage": return "QN";
                case "number_word":
                case "btnNumberToWord": return "NW";
                case "nb_word":
                case "btnNumberToWordBn": return "NB";
                case "ai_write":
                case "btnAiWrite": return "W";
                case "ai_fix":
                case "btnAiFix": return "F";
                case "ai_summarize":
                case "btnAiSummarize": return "S";
                case "aksa_ai":
                case "mnuAksaAI": return "AI";
                case "font_conv":
                case "mnuConvertFont": return "FN";
                case "b2u":
                case "btnB2U": return "BU";
                case "u2b":
                case "btnU2B": return "UB";
                case "margins":
                case "btnMarginNormal":
                case "btnMarginNarrow":
                case "btnMarginModerate":
                case "btnMarginWide":
                case "btnMarginMirrored": return "M";
                case "orientation":
                case "btnPortrait":
                case "btnLandscape": return "O";
                case "page_size":
                case "btnA4":
                case "btnLetter":
                case "btnLegal":
                case "btnA3": return "SZ";
                case "page_break":
                case "btnPageBreak": return "PB";
                case "section_break":
                case "btnSectionBreak": return "SB";
                case "line_numbers":
                case "btnLineNumbers": return "LN";
                case "pdf":
                case "btnSavePdf": return "PDF";
                case "docx":
                case "btnSaveDocx": return "DX";
                case "txt":
                case "btnSaveTxt": return "TXT";
                case "jpg":
                case "btnSaveJpg": return "JPG";
                case "save_copy":
                case "btnSaveCopy": return "CP";
                case "close_all":
                case "save_close":
                case "btnSaveCloseAll": return "CA";
                case "folder_sel":
                case "save_folder":
                case "btnFolderSelect": return "CF";
                case "save_doc":
                case "btnSaveDocFast": return "DC";
                case "save_pdf":
                case "btnSavePdfFast": return "PD";
                case "save_jpg":
                case "btnSaveJpgFast": return "JG";
                case "save_search":
                case "btnSearchDocFolder": return "SR";
                case "signature":
                case "btnSignatureLine": return "SG";
                case "date":
                case "btnInsertDate": return "DT";
                case "bengali":
                case "mnuBengali": return "BN";
                case "english":
                case "mnuEnglish": return "EN";
                case "list":
                case "select_folder":
                case "template_folder":
                case "btnSelectFolder": return "TF";
                case "anna_menu": return "16";
                case "anna_item": return "AN";
                case "page1":
                case "mnuPage1": return "P1";
                case "page2":
                case "mnuPage2": return "P2";
                case "page3":
                case "mnuPage3": return "P3";
                case "land_sale":
                case "btnLandSale":
                case "btnLandSale2":
                case "btnLandSale3": return "LS";
                case "land_mortgage":
                case "btnLandMortgage":
                case "btnLandMortgage2":
                case "btnLandMortgage3": return "LM";
                case "land_rent":
                case "btnLandRent":
                case "btnLandRent2":
                case "btnLandRent3": return "LR";
                case "house_rent":
                case "btnHouseRent":
                case "btnHouseRent2":
                case "btnHouseRent3": return "HR";
                case "house_sale":
                case "btnHouseSale":
                case "btnHouseSale2":
                case "btnHouseSale3": return "HS";
                case "shop_rent":
                case "btnShopRent":
                case "btnShopRent2":
                case "btnShopRent3": return "SR";
                case "shop_sale":
                case "btnShopSale":
                case "btnShopSale2":
                case "btnShopSale3": return "SS";
                case "garden_mortgage":
                case "btnGardenMortgage":
                case "btnGardenMortgage2":
                case "btnGardenMortgage3": return "GM";
                case "garden_house_mortgage":
                case "btnGardenHouseMortgage":
                case "btnGardenHouseMortgage2":
                case "btnGardenHouseMortgage3": return "GH";
                case "car_sale":
                case "btnCarSale":
                case "btnCarSale2":
                case "btnCarSale3": return "CS";
                case "motorcycle_sale":
                case "btnMotorcycleSale":
                case "btnMotorcycleSale2":
                case "btnMotorcycleSale3": return "MS";
                case "storage":
                case "btnStorage":
                case "btnStorage2":
                case "btnStorage3": return "ST";
                case "going_abroad":
                case "btnGoingAbroad":
                case "btnGoingAbroad2":
                case "btnGoingAbroad3": return "GA";
                case "construction":
                case "btnConstruction":
                case "btnConstruction2":
                case "btnConstruction3": return "CN";
                case "baby_taxi_sale":
                case "btnBabyTaxiSale":
                case "btnBabyTaxiSale2":
                case "btnBabyTaxiSale3": return "BT";
                case "mahr":
                case "btnMahr":
                case "btnMahr2":
                case "btnMahr3": return "MR";
                case "cng_sale":
                case "btnCngSale":
                case "btnCngSale2":
                case "btnCngSale3": return "CG";
                case "land_installment":
                case "btnLandInstallment":
                case "btnLandInstallment2":
                case "btnLandInstallment3": return "LI";
                case "pond_mortgage":
                case "btnPondMortgage":
                case "btnPondMortgage2":
                case "btnPondMortgage3": return "PM";
                case "partnership":
                case "btnPartnership":
                case "btnPartnership2":
                case "btnPartnership3": return "PT";
                case "loan_repayment":
                case "btnLoanRepayment":
                case "btnLoanRepayment2":
                case "btnLoanRepayment3": return "LP";
                case "building_materials":
                case "btnBuildingMaterials":
                case "btnBuildingMaterials2":
                case "btnBuildingMaterials3": return "BM";
                case "money_transaction":
                case "btnMoneyTransaction":
                case "btnMoneyTransaction2":
                case "btnMoneyTransaction3": return "MT";
                case "boat_sale":
                case "btnBoatSale":
                case "btnBoatSale2":
                case "btnBoatSale3": return "BS";
                case "poultry_farm_rent":
                case "btnPoultryFarmRent":
                case "btnPoultryFarmRent2":
                case "btnPoultryFarmRent3": return "PF";
                case "cattle_farm_sale":
                case "btnCattleFarmSale":
                case "btnCattleFarmSale2":
                case "btnCattleFarmSale3": return "CF";
                case "foster_care":
                case "btnFosterCare":
                case "btnFosterCare2":
                case "btnFosterCare3": return "FC";
                case "draft":
                case "btnDraftWatermark": return "DR";
                case "confidential":
                case "btnConfidentialWatermark": return "CF";
                case "clause":
                case "btnClauseNumbering": return "CL";
                case "disclaimer":
                case "btnDisclaimer": return "DM";
                case "professional":
                case "btnProfessional": return "PR";
                case "modern":
                case "btnModern": return "MD";
                case "cover":
                case "btnCoverLetter": return "CV";
                case "skills":
                case "btnSkills": return "SK";
                case "experience":
                case "btnExperience": return "EX";
                case "education":
                case "btnEducation": return "ED";
                case "register":
                case "btnRegister": return "LC";
                default: return "";
            }
        }

        private sealed class PicHelper : System.Windows.Forms.AxHost
        {
            private PicHelper() : base("00000000-0000-0000-0000-000000000000") { }
            public static IPictureDisp FromImage(System.Drawing.Image img)
            {
                return (IPictureDisp)GetIPictureDispFromPicture(img);
            }
        }

        private static string IconsFolderPath
        {
            get
            {
                string folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Aksa10xFaster", "Icons");
                if (!Directory.Exists(folder)) Directory.CreateDirectory(folder);
                return folder;
            }
        }

        private static readonly Dictionary<string, string> IconFileMap = new Dictionary<string, string>
        {
            { "register", "software-license.png" },
            { "aksa_ai", "AKSA AI.png" },
            { "ai_fix", "Correction_AI.png" },
            { "ai_summarize", "Summerize_AI.png" },
            { "ai_write", "Agreement_AI.png" },
            { "agree_page", "agreement_list.png" },
            { "app_page", "application_size.png" },
            { "cv_page", "cv_size.png" },
            { "question_size", "question_size.png" },
            { "api_settings", "api_setting.png" },
            { "b2u", "bangla_ntow.png" },
            { "u2b", "conveter.png" },
            { "font_conv", "conveter.png" },
            { "bengali", "Bengali_cv.png" },
            { "english", "English_cv.png" },
            { "nb_word", "bangla_ntow.png" },
            { "number_word", "number_word.png" },
            { "about_me", "about_me.png" },
            { "anna_menu", "Ana_Symble.png" },
            { "date_bn", "date_bn.png" },
            { "date_en", "date_en.png" },
            { "page_number", "page_no.png" },
            { "page_num_bn", "bangla_num.png" },
            { "page_num_en", "english_num.png" },
            { "page1", "page_1.png" },
            { "page2", "page_2.png" },
            { "page3", "page_3.png" },
            { "share_email", "gmail.png" },
            { "share_whatsapp", "whatsapp.png" },
            { "share_imo", "imo.png" },
            { "share_telegram", "telegram.png" },
            { "cover", "about_me.png" },
            { "professional", "about_me.png" },
            { "modern", "about_me.png" },
            { "education", "about_me.png" },
            { "experience", "about_me.png" },
            { "skills", "about_me.png" },
            { "list", "agreement_list.png" },
        };

        private IPictureDisp TryLoadIcon(string key)
        {
            string fileName;
            if (!IconFileMap.TryGetValue(key, out fileName))
                fileName = key + ".png";
            string path = Path.Combine(IconsFolderPath, fileName);
            if (File.Exists(path))
            {
                using (var img = Image.FromFile(path))
                    return PicHelper.FromImage(img);
            }
            return null;
        }

        public bool GetEnabled(Office.IRibbonControl c)
        {
            try
            {
                if (c == null) return true;
                bool expired = IsTrialExpired();
                switch (c.Id)
                {
                    // JPG button (Save group)
                    case "btnSaveJpgFast":
                    // Math group controls
                    case "btnNumberToWord":
                    case "btnNumberToWordBn":
                    case "btnDateBn":
                    case "btnDateEn":
                    case "mnuConvertFont":
                    case "btnB2U":
                    case "btnU2B":
                    case "mnuAnna":
                    case "btnAnna1": case "btnAnna2": case "btnAnna3":
                    case "btnAnna4": case "btnAnna5": case "btnAnna6":
                    case "btnAnna7": case "btnAnna8": case "btnAnna9":
                    case "btnAnna10": case "btnAnna11": case "btnAnna12":
                    case "btnAnna13": case "btnAnna14": case "btnAnna15":
                    case "btnAnna16":
                    // AI group controls
                    case "mnuAksaAI":
                    case "btnAiWrite":
                    case "btnAiFix":
                    case "btnAiSummarize":
                        return !expired;
                    default:
                        return true;
                }
            }
            catch { return true; }
        }

        public void OnOpenIconsFolder(Office.IRibbonControl c)
        {
            Process.Start("explorer.exe", IconsFolderPath);
        }

        private IPictureDisp BuildIcon(string label, Color color)
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
                g.Clear(Color.Transparent);
                var path = new System.Drawing.Drawing2D.GraphicsPath();
                int sz = 22;
                int x = (32 - sz) / 2;
                int y = (32 - sz) / 2;
                int r = 4;
                path.AddArc(x, y, r, r, 180, 90);
                path.AddArc(x + sz - r, y, r, r, 270, 90);
                path.AddArc(x + sz - r, y + sz - r, r, r, 0, 90);
                path.AddArc(x, y + sz - r, r, r, 90, 90);
                path.CloseFigure();
                using (path)
                using (var brush = new SolidBrush(color))
                    g.FillPath(brush, path);
                using (var f = new System.Drawing.Font("Segoe UI", 9f, FontStyle.Bold))
                using (var tb = new SolidBrush(Color.White))
                {
                    var ms = g.MeasureString(label, f);
                    g.DrawString(label, f, tb, (32f - ms.Width) / 2f, (32f - ms.Height) / 2f);
                }
            }
            return PicHelper.FromImage(bmp);
        }

        public IPictureDisp GetImage(Office.IRibbonControl control)
        {
            try
            {
                var key = control.Tag;
                if (key == null) key = control.Id;
                if (key == null) return null;
                if (_iconCache.ContainsKey(key))
                    return _iconCache[key];
                var pic = TryLoadIcon(key);
                if (pic == null)
                {
                    var label = LabelForKey(key);
                    if (label == "") label = "!";
                    var color = ColorForKey(key);
                    pic = BuildIcon(label, color);
                }
                _iconCache[key] = pic;
                return pic;
            }
            catch
            {
                return null;
            }
        }

        // ===================== AI API =====================

        private static string AiConfigPath
        {
            get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Aksa10xFaster", "config.json"); }
        }

        private Dictionary<string, object> GetAiConfig()
        {
            try
            {
                if (!File.Exists(AiConfigPath)) return null;
                string json = File.ReadAllText(AiConfigPath);
                var ser = new System.Web.Script.Serialization.JavaScriptSerializer();
                return ser.DeserializeObject(json) as Dictionary<string, object>;
            }
            catch { return null; }
        }

        private string GetApiKey(Dictionary<string, object> config)
        {
            if (config == null) return null;
            string provider = config.ContainsKey("provider") ? config["provider"] as string : null;
            if (provider == null) return null;
            if (!config.ContainsKey(provider)) return null;
            var pObj = config[provider] as Dictionary<string, object>;
            if (pObj == null) return null;
            return pObj.ContainsKey("apiKey") ? pObj["apiKey"] as string : null;
        }

        private string GetAiModel(Dictionary<string, object> config)
        {
            if (config == null) return null;
            string provider = config.ContainsKey("provider") ? config["provider"] as string : null;
            if (provider == null) return null;
            if (!config.ContainsKey(provider)) return null;
            var pObj = config[provider] as Dictionary<string, object>;
            if (pObj == null) return null;
            return pObj.ContainsKey("model") ? pObj["model"] as string : null;
        }

        private string GetAiProvider(Dictionary<string, object> config)
        {
            if (config == null) return "gemini";
            return config.ContainsKey("provider") ? config["provider"] as string ?? "gemini" : "gemini";
        }

        private string CallAi(string systemPrompt, string userText)
        {
            var config = GetAiConfig();
            if (config == null || GetApiKey(config) == null)
                throw new System.Exception("API key not set. Go to About → API Settings to set your API key.");
            string key = GetApiKey(config);
            string provider = GetAiProvider(config);
            string model = GetAiModel(config) ?? "";
            switch (provider)
            {
                case "gemini": return CallGemini(key, model, systemPrompt, userText);
                case "openai": return CallOpenAi(key, "https://api.openai.com/v1/chat/completions", model, systemPrompt, userText);
                case "deepseek": return CallOpenAi(key, "https://api.deepseek.com/v1/chat/completions", model, systemPrompt, userText);
                case "claude": return CallClaude(key, model, systemPrompt, userText);
                default: throw new System.Exception("Unknown AI provider: " + provider);
            }
        }

        private static string JsonEscape(string s)
        {
            return "\"" + s.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\r\n", "\\n").Replace("\n", "\\n").Replace("\r", "\\n") + "\"";
        }

        private static string CallGemini(string apiKey, string model, string systemPrompt, string userText)
        {
            if (string.IsNullOrEmpty(model)) model = "gemini-2.5-flash";
            string url = "https://generativelanguage.googleapis.com/v1beta/models/" + model + ":generateContent?key=" + apiKey;
            string prompt = JsonEscape(systemPrompt + "\n\n" + userText);
            string body = "{\"contents\":[{\"parts\":[{\"text\":" + prompt + "}]}]}";

            int maxRetries = 3;
            for (int attempt = 1; attempt <= maxRetries; attempt++)
            {
                var wc = new System.Net.WebClient();
                wc.Encoding = System.Text.Encoding.UTF8;
                wc.Headers[System.Net.HttpRequestHeader.ContentType] = "application/json";
                try
                {
                    string resp = wc.UploadString(url, body);
                    var ser = new System.Web.Script.Serialization.JavaScriptSerializer();
                    var obj = ser.DeserializeObject(resp) as System.Collections.Generic.Dictionary<string, object>;
                    if (obj == null) throw new System.Exception("Cannot parse Gemini response: " + resp);

                    object temp;
                    if (!obj.TryGetValue("candidates", out temp))
                    {
                        string errText = "";
                        if (obj.TryGetValue("error", out temp))
                        {
                            var errObj = temp as System.Collections.Generic.Dictionary<string, object>;
                            if (errObj != null && errObj.TryGetValue("message", out temp))
                                errText = "error: " + temp.ToString();
                        }
                        else if (obj.TryGetValue("promptFeedback", out temp))
                        {
                            var fb = temp as System.Collections.Generic.Dictionary<string, object>;
                            if (fb != null && fb.TryGetValue("blockReason", out temp))
                                errText = "blocked: " + temp.ToString();
                        }
                        if (string.IsNullOrEmpty(errText)) errText = "raw: " + resp;
                        throw new System.Exception("Gemini API " + errText);
                    }

                    int count = 0;
                    var candidatesList = temp as System.Collections.ArrayList;
                    if (candidatesList != null) count = candidatesList.Count;
                    else { var arr = temp as object[]; if (arr != null) count = arr.Length; }
                    if (count == 0)
                        throw new System.Exception("Gemini candidates empty. type=" + temp.GetType().Name + " raw=" + resp);
                    var first = candidatesList != null ? candidatesList[0] : (temp as object[])[0];
                    var firstObj = first as System.Collections.Generic.Dictionary<string, object>;
                    if (firstObj == null) throw new System.Exception("Gemini candidate not an object");
                    if (!firstObj.TryGetValue("content", out temp)) throw new System.Exception("Gemini no content");
                    var content = temp as System.Collections.Generic.Dictionary<string, object>;
                    if (content == null) throw new System.Exception("Gemini content not an object");
                    if (!content.TryGetValue("parts", out temp)) throw new System.Exception("Gemini no parts");
                    var partsList = temp as System.Collections.ArrayList;
                    int partCount = 0;
                    if (partsList != null) partCount = partsList.Count;
                    else { var arr = temp as object[]; if (arr != null) partCount = arr.Length; }
                    if (partCount == 0) throw new System.Exception("Gemini parts empty");
                    var p0 = partsList != null ? partsList[0] : (temp as object[])[0];
                    var firstPart = p0 as System.Collections.Generic.Dictionary<string, object>;
                    if (firstPart == null) throw new System.Exception("Gemini part not an object");
                    if (!firstPart.TryGetValue("text", out temp)) throw new System.Exception("Gemini no text");
                    return temp as string;
                }
                catch (System.Net.WebException ex)
                {
                    string errBody = "";
                    try { using (var r = new System.IO.StreamReader(ex.Response.GetResponseStream())) errBody = r.ReadToEnd(); } catch { }
                    if (errBody.Contains("\"code\": 503") && attempt < maxRetries)
                    {
                        System.Threading.Thread.Sleep(3000 * attempt);
                        continue;
                    }
                    throw new System.Exception("Gemini API responded: " + errBody);
                }
            }
            throw new System.Exception("Gemini API failed after " + maxRetries + " attempts.");
        }

        private static string CallOpenAi(string apiKey, string baseUrl, string model, string systemPrompt, string userText)
        {
            if (string.IsNullOrEmpty(model)) model = "gpt-4o-mini";
            string body = "{\"model\":" + JsonEscape(model) + ",\"messages\":[{\"role\":\"system\",\"content\":" + JsonEscape(systemPrompt) + "},{\"role\":\"user\",\"content\":" + JsonEscape(userText) + "}],\"max_tokens\":4096}";
            var wc = new System.Net.WebClient();
            wc.Encoding = System.Text.Encoding.UTF8;
            wc.Headers["Content-Type"] = "application/json";
            wc.Headers["Authorization"] = "Bearer " + apiKey;
            string resp = wc.UploadString(baseUrl, body);
            var ser = new System.Web.Script.Serialization.JavaScriptSerializer();
            var obj = ser.DeserializeObject(resp) as Dictionary<string, object>;
            if (obj == null) throw new System.Exception("Cannot parse response: " + resp);
            object temp;
            if (!obj.TryGetValue("choices", out temp)) throw new System.Exception("API no choices: " + resp);
            var choices = temp as System.Collections.ArrayList;
            if (choices == null || choices.Count == 0) throw new System.Exception("API no choices: " + resp);
            var first = choices[0] as Dictionary<string, object>;
            if (first == null || !first.TryGetValue("message", out temp)) throw new System.Exception("API no message");
            var msg = temp as Dictionary<string, object>;
            if (msg == null || !msg.TryGetValue("content", out temp)) throw new System.Exception("API no content");
            return temp as string;
        }

        private static string CallClaude(string apiKey, string model, string systemPrompt, string userText)
        {
            if (string.IsNullOrEmpty(model)) model = "claude-3-5-sonnet-20241022";
            string body = "{\"model\":" + JsonEscape(model) + ",\"max_tokens\":4096,\"system\":" + JsonEscape(systemPrompt) + ",\"messages\":[{\"role\":\"user\",\"content\":" + JsonEscape(userText) + "}]}";
            var wc = new System.Net.WebClient();
            wc.Encoding = System.Text.Encoding.UTF8;
            wc.Headers["Content-Type"] = "application/json";
            wc.Headers["x-api-key"] = apiKey;
            wc.Headers["anthropic-version"] = "2023-06-01";
            string resp = wc.UploadString("https://api.anthropic.com/v1/messages", body);
            var ser = new System.Web.Script.Serialization.JavaScriptSerializer();
            var obj = ser.DeserializeObject(resp) as Dictionary<string, object>;
            if (obj == null) throw new System.Exception("Cannot parse response: " + resp);
            object temp;
            if (!obj.TryGetValue("content", out temp)) throw new System.Exception("Claude no content: " + resp);
            var parts = temp as System.Collections.ArrayList;
            if (parts == null || parts.Count == 0) throw new System.Exception("Claude empty content");
            var first = parts[0] as Dictionary<string, object>;
            if (first == null || !first.TryGetValue("text", out temp)) throw new System.Exception("Claude no text");
            return temp as string;
        }

        private void InsertWithTypingEffect(string text)
        {
            if (string.IsNullOrEmpty(text)) return;
            var doc = _app.ActiveDocument;
            if (doc == null) return;
            var sel = _app.Selection;
            if (sel == null) return;
            int pos = sel.Range.Start;
            sel.Text = "";
            var enumerator = System.Globalization.StringInfo.GetTextElementEnumerator(text);
            var chars = new System.Collections.Generic.List<string>();
            while (enumerator.MoveNext())
                chars.Add(enumerator.GetTextElement());
            for (int i = 0; i < chars.Count; i++)
            {
                var r = doc.Range(pos, pos);
                r.Text = chars[i];
                if (i == 0) try { r.Font.Name = "SolaimanLipi"; } catch { }
                pos = r.End;
                System.Windows.Forms.Application.DoEvents();
                System.Threading.Thread.Sleep(30);
            }
        }

        // ===================== REGISTRATION =====================

        [ComRegisterFunction]
        public static void RegisterAddIn(Type type)
        {
            try
            {
                var key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(
                    @"Software\Microsoft\Office\Word\Addins\Aksa10xFaster.Connect");
                key.SetValue("Description", "AKSA 10X FASTER - Word shortcut tools");
                key.SetValue("FriendlyName", "AKSA 10X FASTER");
                key.SetValue("LoadBehavior", 3);
                key.Close();
            }
            catch { }
        }

        [ComUnregisterFunction]
        public static void UnregisterAddIn(Type type)
        {
            try
            {
                Microsoft.Win32.Registry.CurrentUser.DeleteSubKey(
                    @"Software\Microsoft\Office\Word\Addins\Aksa10xFaster.Connect", false);
            }
            catch { }
        }
    }
}
