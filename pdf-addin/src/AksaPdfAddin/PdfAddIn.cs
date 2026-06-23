using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Extensibility;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace AksaPdfAddin
{
    [ComVisible(true)]
    [Guid("33F39564-802B-412B-A713-E947F98DCBB9")]
    [ProgId("AksaPdfTools.Connect")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class PdfAddIn : IDTExtensibility2, Office.IRibbonExtensibility
    {
        private Word.Application _app;
        private Office.IRibbonUI _ribbon;
        private PdfService _pdf;

        public PdfAddIn() { }

        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            _app = application as Word.Application;
            _pdf = new PdfService();
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            _app = null;
            _pdf = null;
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        public string GetCustomUI(string ribbonId)
        {
            using (var s = Assembly.GetExecutingAssembly()
                .GetManifestResourceStream("AksaPdfAddin.AksaPdfAddin_Ribbon.xml"))
            {
                if (s == null) return "";
                using (var r = new StreamReader(s))
                    return r.ReadToEnd();
            }
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUi) { _ribbon = ribbonUi; }

        public string GetImage(Office.IRibbonControl c) { return ""; }

        private Word.Document Doc
        {
            get
            {
                if (_app == null) { MessageBox.Show("Word application not connected.", "AKSA PDF Tools"); return null; }
                try { return _app.ActiveDocument; }
                catch { MessageBox.Show("No document open.", "AKSA PDF Tools"); return null; }
            }
        }

        private bool ConfirmOpen(string msg)
        {
            return MessageBox.Show(msg + "\n\nOpen the file now?", "AKSA PDF Tools",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;
        }

        private string PickFile(string title, string filter, bool save, string defaultName)
        {
            if (save)
            {
                var dlg = new SaveFileDialog { Title = title, Filter = filter, FileName = defaultName };
                return dlg.ShowDialog() == DialogResult.OK ? dlg.FileName : null;
            }
            else
            {
                var dlg = new OpenFileDialog { Title = title, Filter = filter };
                return dlg.ShowDialog() == DialogResult.OK ? dlg.FileName : null;
            }
        }

        // ===================== OPEN PDF =====================

        public void OnOpenPdf(Office.IRibbonControl c)
        {
            try
            {
                string path = PickFile("Open PDF File", "PDF Files (*.pdf)|*.pdf|All Files (*.*)|*.*", false, null);
                if (path == null) return;

                object missing = System.Reflection.Missing.Value;
                object confirmConversions = false;
                object readOnly = false;
                object addToRecentFiles = false;
                object format = Word.WdOpenFormat.wdOpenFormatAuto;

                var doc = _app.Documents.Open(path,
                    ref confirmConversions, ref readOnly, ref addToRecentFiles,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref format, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing);

                // Set optimal viewing layout
                _app.ActiveWindow.View.Type = Word.WdViewType.wdPrintView;
                _app.ActiveWindow.View.Zoom.PageFit = Word.WdPageFit.wdPageFitBestFit;
            }
            catch (Exception ex) { MessageBox.Show("Error: " + ex.Message, "Open PDF Error"); }
        }

        // ===================== WORD TO PDF =====================

        public void OnWordToPdf(Office.IRibbonControl c)
        {
            try
            {
                var doc = Doc;
                if (doc == null) return;

                string path = PickFile("Save as PDF", "PDF Files (*.pdf)|*.pdf", true,
                    Path.GetFileNameWithoutExtension(doc.FullName) + ".pdf");
                if (path == null) return;

                _pdf.WordToPdf(doc, path);
                if (ConfirmOpen("PDF saved successfully."))
                    Process.Start(path);
            }
            catch (Exception ex) { MessageBox.Show("Error: " + ex.Message, "Word to PDF Error"); }
        }

        // ===================== PDF TO WORD =====================

        public void OnPdfToWord(Office.IRibbonControl c)
        {
            try
            {
                string input = PickFile("Select PDF file", "PDF Files (*.pdf)|*.pdf", false, null);
                if (input == null) return;

                MessageBox.Show(
                    "PDF to Word conversion uses Word's built-in converter.\n\n" +
                    "Note: Complex PDF layouts may not convert perfectly.\n" +
                    "Text, images, and shapes may shift between pages.\n\n" +
                    "For best results, use simple text-based PDFs.",
                    "PDF to Word", MessageBoxButtons.OK, MessageBoxIcon.Information);

                string output = PickFile("Save Word document as", "Word Documents (*.docx)|*.docx", true,
                    Path.GetFileNameWithoutExtension(input) + ".docx");
                if (output == null) return;

                _pdf.PdfToWord(_app, input, output);
                if (ConfirmOpen("Word document saved successfully."))
                    Process.Start(output);
            }
            catch (Exception ex) { MessageBox.Show("Error: " + ex.Message, "PDF to Word Error"); }
        }

        // ===================== BATCH CONVERT =====================

        public void OnBatchWordToPdf(Office.IRibbonControl c)
        {
            try
            {
                var dlg = new OpenFileDialog
                {
                    Title = "Select Word files",
                    Filter = "Word Documents (*.docx;*.doc)|*.docx;*.doc",
                    Multiselect = true
                };
                if (dlg.ShowDialog() != DialogResult.OK || dlg.FileNames.Length == 0) return;

                using (var fb = new FolderBrowserDialog { Description = "Select destination folder" })
                {
                    if (fb.ShowDialog() != DialogResult.OK) return;

                    int done = 0;
                    var errors = new List<string>();
                    foreach (var file in dlg.FileNames)
                    {
                        try
                        {
                            string pdfPath = Path.Combine(fb.SelectedPath,
                                Path.GetFileNameWithoutExtension(file) + ".pdf");
                            _pdf.BatchWordToPdf(_app, file, pdfPath);
                            done++;
                        }
                        catch (Exception ex) { errors.Add(Path.GetFileName(file) + ": " + ex.Message); }
                    }

                    string msg = $"Converted {done} of {dlg.FileNames.Length} files.";
                    if (errors.Count > 0) msg += "\n\nErrors:\n" + string.Join("\n", errors);
                    MessageBox.Show(msg, "Batch Convert Complete");
                }
            }
            catch (Exception ex) { MessageBox.Show("Error: " + ex.Message, "Batch Convert Error"); }
        }

        public void OnBatchPdfToWord(Office.IRibbonControl c)
        {
            try
            {
                var dlg = new OpenFileDialog
                {
                    Title = "Select PDF files",
                    Filter = "PDF Files (*.pdf)|*.pdf",
                    Multiselect = true
                };
                if (dlg.ShowDialog() != DialogResult.OK || dlg.FileNames.Length == 0) return;

                using (var fb = new FolderBrowserDialog { Description = "Select destination folder" })
                {
                    if (fb.ShowDialog() != DialogResult.OK) return;

                    int done = 0;
                    var errors = new List<string>();
                    foreach (var file in dlg.FileNames)
                    {
                        try
                        {
                            string docxPath = Path.Combine(fb.SelectedPath,
                                Path.GetFileNameWithoutExtension(file) + ".docx");
                            _pdf.PdfToWord(_app, file, docxPath);
                            done++;
                        }
                        catch (Exception ex) { errors.Add(Path.GetFileName(file) + ": " + ex.Message); }
                    }

                    string msg = $"Converted {done} of {dlg.FileNames.Length} files.";
                    if (errors.Count > 0) msg += "\n\nErrors:\n" + string.Join("\n", errors);
                    MessageBox.Show(msg, "Batch Convert Complete");
                }
            }
            catch (Exception ex) { MessageBox.Show("Error: " + ex.Message, "Batch Convert Error"); }
        }

        // ===================== MERGE PDF =====================

        public void OnMergePdf(Office.IRibbonControl c)
        {
            try
            {
                var dlg = new OpenFileDialog
                {
                    Title = "Select PDF files to merge",
                    Filter = "PDF Files (*.pdf)|*.pdf",
                    Multiselect = true
                };
                if (dlg.ShowDialog() != DialogResult.OK || dlg.FileNames.Length < 2)
                {
                    MessageBox.Show("Please select at least 2 PDF files.", "Merge PDF");
                    return;
                }

                string output = PickFile("Save merged PDF as", "PDF Files (*.pdf)|*.pdf", true, "Merged.pdf");
                if (output == null) return;

                _pdf.MergePdf(dlg.FileNames.ToList(), output);
                if (ConfirmOpen("PDFs merged successfully."))
                    Process.Start(output);
            }
            catch (Exception ex) { MessageBox.Show("Error: " + ex.Message, "Merge PDF Error"); }
        }

        // ===================== SPLIT PDF =====================

        public void OnSplitPdf(Office.IRibbonControl c)
        {
            try
            {
                string input = PickFile("Select PDF file to split", "PDF Files (*.pdf)|*.pdf", false, null);
                if (input == null) return;

                using (var fb = new FolderBrowserDialog { Description = "Select destination folder" })
                {
                    if (fb.ShowDialog() != DialogResult.OK) return;
                    _pdf.SplitPdf(input, fb.SelectedPath);
                    MessageBox.Show("PDF split successfully.\nPages saved to: " + fb.SelectedPath, "AKSA PDF Tools");
                }
            }
            catch (Exception ex) { MessageBox.Show("Error: " + ex.Message, "Split PDF Error"); }
        }

        // ===================== EXTRACT PAGES =====================

        public void OnExtractPages(Office.IRibbonControl c)
        {
            try
            {
                string input = PickFile("Select PDF file", "PDF Files (*.pdf)|*.pdf", false, null);
                if (input == null) return;

                var form = new Form
                {
                    Text = "Extract Pages",
                    Size = new System.Drawing.Size(400, 160),
                    FormBorderStyle = FormBorderStyle.FixedDialog,
                    MaximizeBox = false,
                    MinimizeBox = false,
                    StartPosition = FormStartPosition.CenterScreen
                };

                var lbl = new Label { Text = "Enter page numbers (e.g. 1,3,5-8):", Left = 12, Top = 12, Width = 360 };
                var txt = new TextBox { Text = "1", Left = 12, Top = 40, Width = 360 };
                var btnOk = new Button { Text = "OK", Left = 100, Top = 80, Width = 80, DialogResult = DialogResult.OK };
                var btnCancel = new Button { Text = "Cancel", Left = 200, Top = 80, Width = 80, DialogResult = DialogResult.Cancel };
                form.AcceptButton = btnOk;
                form.CancelButton = btnCancel;
                form.Controls.AddRange(new Control[] { lbl, txt, btnOk, btnCancel });

                if (form.ShowDialog() != DialogResult.OK || string.IsNullOrWhiteSpace(txt.Text)) return;

                var pages = ParsePageRange(txt.Text);
                if (pages.Count == 0) { MessageBox.Show("No valid page numbers.", "Extract Pages"); return; }

                string output = PickFile("Save extracted pages as", "PDF Files (*.pdf)|*.pdf", true, "Extracted.pdf");
                if (output == null) return;

                _pdf.ExtractPages(input, output, pages);
                if (ConfirmOpen("Pages extracted successfully."))
                    Process.Start(output);
            }
            catch (Exception ex) { MessageBox.Show("Error: " + ex.Message, "Extract Pages Error"); }
        }

        private static List<int> ParsePageRange(string input)
        {
            var pages = new List<int>();
            foreach (var part in input.Split(','))
            {
                var t = part.Trim();
                if (t.Contains('-'))
                {
                    var range = t.Split('-');
                    if (int.TryParse(range[0], out int s) && int.TryParse(range[1], out int e))
                        for (int i = s; i <= e; i++) pages.Add(i);
                }
                else if (int.TryParse(t, out int n)) pages.Add(n);
            }
            return pages.Distinct().OrderBy(p => p).ToList();
        }

        // ===================== PROTECT PDF =====================

        public void OnProtectPdf(Office.IRibbonControl c)
        {
            try
            {
                string input = PickFile("Select PDF file to protect", "PDF Files (*.pdf)|*.pdf", false, null);
                if (input == null) return;

                string password = RequestInput("Protect PDF", "Enter password:", "");
                if (string.IsNullOrWhiteSpace(password)) return;

                string output = PickFile("Save protected PDF as", "PDF Files (*.pdf)|*.pdf", true,
                    Path.GetFileNameWithoutExtension(input) + "_protected.pdf");
                if (output == null) return;

                _pdf.ProtectPdf(input, output, password);
                MessageBox.Show("PDF protected successfully with password.", "AKSA PDF Tools");
            }
            catch (Exception ex) { MessageBox.Show("Error: " + ex.Message, "Protect PDF Error"); }
        }

        // ===================== UNLOCK PDF =====================

        public void OnUnlockPdf(Office.IRibbonControl c)
        {
            try
            {
                string input = PickFile("Select password-protected PDF", "PDF Files (*.pdf)|*.pdf", false, null);
                if (input == null) return;

                string password = RequestInput("Unlock PDF", "Enter the PDF password:", "");
                if (string.IsNullOrWhiteSpace(password)) return;

                string output = PickFile("Save unlocked PDF as", "PDF Files (*.pdf)|*.pdf", true,
                    Path.GetFileNameWithoutExtension(input) + "_unlocked.pdf");
                if (output == null) return;

                _pdf.UnlockPdf(input, output, password);
                MessageBox.Show("PDF unlocked successfully.", "AKSA PDF Tools");
            }
            catch (Exception ex) { MessageBox.Show("Error: " + ex.Message, "Unlock PDF Error"); }
        }

        // ===================== PDF INFO =====================

        public void OnPdfInfo(Office.IRibbonControl c)
        {
            try
            {
                string input = PickFile("Select PDF file", "PDF Files (*.pdf)|*.pdf", false, null);
                if (input == null) return;

                var info = _pdf.GetPdfInfo(input);
                var msg = $"File: {Path.GetFileName(input)}\n" +
                          $"Pages: {info.PageCount}\n" +
                          $"Size: {info.FileSize}\n" +
                          $"Encrypted: {(info.Encrypted ? "Yes" : "No")}" +
                          (!string.IsNullOrEmpty(info.EncryptionInfo) ? $"\nEncryption: {info.EncryptionInfo}" : "");

                MessageBox.Show(msg, "PDF Info");
            }
            catch (Exception ex) { MessageBox.Show("Error: " + ex.Message, "PDF Info Error"); }
        }

        // ===================== ABOUT =====================

        public void OnAbout(Office.IRibbonControl c)
        {
            MessageBox.Show(
                "AKSA PDF Tools v1.0\n\n" +
                "Word add-in for:\n" +
                "\u2022 Word \u2192 PDF conversion\n" +
                "\u2022 PDF \u2192 Word conversion\n" +
                "\u2022 Merge / Split / Extract PDF\n" +
                "\u2022 Protect / Unlock PDF\n" +
                "\u2022 Batch conversion\n\n" +
                "\u00a9 2026 AKSA DIGITAL POINT",
                "About AKSA PDF Tools");
        }

        // ===================== HELPERS =====================

        private static string RequestInput(string title, string prompt, string defaultValue)
        {
            var form = new Form
            {
                Text = title,
                Size = new System.Drawing.Size(400, 150),
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                StartPosition = FormStartPosition.CenterScreen
            };

            var lbl = new Label { Text = prompt, Left = 12, Top = 12, Width = 360 };
            var txt = new TextBox { Text = defaultValue, Left = 12, Top = 40, Width = 360 };
            var btnOk = new Button { Text = "OK", Left = 100, Top = 75, Width = 80, DialogResult = DialogResult.OK };
            var btnCancel = new Button { Text = "Cancel", Left = 200, Top = 75, Width = 80, DialogResult = DialogResult.Cancel };
            form.AcceptButton = btnOk;
            form.CancelButton = btnCancel;
            form.Controls.AddRange(new Control[] { lbl, txt, btnOk, btnCancel });

            return form.ShowDialog() == DialogResult.OK ? txt.Text.Trim() : null;
        }
    }
}
