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

        public string GetImage(Office.IRibbonControl c)
        {
            return "";
        }

        private Word.Document Doc
        {
            get
            {
                if (_app == null) { MessageBox.Show("Word application not connected.", "AKSA PDF Tools"); return null; }
                try { return _app.ActiveDocument; }
                catch { MessageBox.Show("No document open.", "AKSA PDF Tools"); return null; }
            }
        }

        // ===================== WORD → PDF =====================

        public void OnWordToPdf(Office.IRibbonControl c)
        {
            try
            {
                var doc = Doc;
                if (doc == null) return;

                var dlg = new SaveFileDialog
                {
                    Title = "Save as PDF",
                    Filter = "PDF Files (*.pdf)|*.pdf",
                    FileName = Path.GetFileNameWithoutExtension(doc.FullName) + ".pdf"
                };

                if (dlg.ShowDialog() != DialogResult.OK) return;

                _pdf.WordToPdf(doc, dlg.FileName);
                var result = MessageBox.Show(
                    "PDF saved successfully.\n\nOpen the PDF file now?",
                    "AKSA PDF Tools", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                    Process.Start(dlg.FileName);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Word → PDF Error");
            }
        }

        // ===================== PDF → WORD =====================

        public void OnPdfToWord(Office.IRibbonControl c)
        {
            try
            {
                var ofd = new OpenFileDialog
                {
                    Title = "Select PDF file",
                    Filter = "PDF Files (*.pdf)|*.pdf"
                };
                if (ofd.ShowDialog() != DialogResult.OK) return;

                var sfd = new SaveFileDialog
                {
                    Title = "Save Word document as",
                    Filter = "Word Documents (*.docx)|*.docx",
                    FileName = Path.GetFileNameWithoutExtension(ofd.FileName) + ".docx"
                };
                if (sfd.ShowDialog() != DialogResult.OK) return;

                _pdf.PdfToWord(_app, ofd.FileName, sfd.FileName);

                var result = MessageBox.Show(
                    "Word document saved successfully.\n\nOpen the document now?",
                    "AKSA PDF Tools", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                    Process.Start(sfd.FileName);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "PDF → Word Error");
            }
        }

        // ===================== BATCH =====================

        public void OnBatchWordToPdf(Office.IRibbonControl c)
        {
            try
            {
                var ofd = new OpenFileDialog
                {
                    Title = "Select Word files to convert",
                    Filter = "Word Documents (*.docx;*.doc)|*.docx;*.doc",
                    Multiselect = true
                };
                if (ofd.ShowDialog() != DialogResult.OK || ofd.FileNames.Length == 0) return;

                using (var fb = new FolderBrowserDialog { Description = "Select destination folder for PDFs" })
                {
                    if (fb.ShowDialog() != DialogResult.OK) return;
                    string dest = fb.SelectedPath;

                    int total = ofd.FileNames.Length;
                    int done = 0;
                    var errors = new List<string>();

                    foreach (var file in ofd.FileNames)
                    {
                        try
                        {
                            string pdfPath = Path.Combine(dest, Path.GetFileNameWithoutExtension(file) + ".pdf");
                            _pdf.BatchWordToPdf(_app, file, pdfPath);
                            done++;
                        }
                        catch (Exception ex)
                        {
                            errors.Add(Path.GetFileName(file) + ": " + ex.Message);
                        }
                    }

                    string msg = $"Converted {done} of {total} files successfully.";
                    if (errors.Count > 0)
                        msg += "\n\nErrors:\n" + string.Join("\n", errors);

                    MessageBox.Show(msg, "Batch Convert Complete");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Batch Convert Error");
            }
        }

        public void OnBatchPdfToWord(Office.IRibbonControl c)
        {
            try
            {
                var ofd = new OpenFileDialog
                {
                    Title = "Select PDF files to convert",
                    Filter = "PDF Files (*.pdf)|*.pdf",
                    Multiselect = true
                };
                if (ofd.ShowDialog() != DialogResult.OK || ofd.FileNames.Length == 0) return;

                using (var fb = new FolderBrowserDialog { Description = "Select destination folder for Word documents" })
                {
                    if (fb.ShowDialog() != DialogResult.OK) return;
                    string dest = fb.SelectedPath;

                    int total = ofd.FileNames.Length;
                    int done = 0;
                    var errors = new List<string>();

                    foreach (var file in ofd.FileNames)
                    {
                        try
                        {
                            string docxPath = Path.Combine(dest, Path.GetFileNameWithoutExtension(file) + ".docx");
                            _pdf.BatchPdfToWord(_app, file, docxPath);
                            done++;
                        }
                        catch (Exception ex)
                        {
                            errors.Add(Path.GetFileName(file) + ": " + ex.Message);
                        }
                    }

                    string msg = $"Converted {done} of {total} files successfully.";
                    if (errors.Count > 0)
                        msg += "\n\nErrors:\n" + string.Join("\n", errors);

                    MessageBox.Show(msg, "Batch Convert Complete");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Batch Convert Error");
            }
        }

        // ===================== MERGE PDF =====================

        public void OnMergePdf(Office.IRibbonControl c)
        {
            try
            {
                var ofd = new OpenFileDialog
                {
                    Title = "Select PDF files to merge",
                    Filter = "PDF Files (*.pdf)|*.pdf",
                    Multiselect = true
                };
                if (ofd.ShowDialog() != DialogResult.OK || ofd.FileNames.Length < 2)
                {
                    MessageBox.Show("Please select at least 2 PDF files.", "Merge PDF");
                    return;
                }

                var sfd = new SaveFileDialog
                {
                    Title = "Save merged PDF as",
                    Filter = "PDF Files (*.pdf)|*.pdf",
                    FileName = "Merged.pdf"
                };
                if (sfd.ShowDialog() != DialogResult.OK) return;

                _pdf.MergePdf(ofd.FileNames.ToList(), sfd.FileName);

                var result = MessageBox.Show(
                    "PDFs merged successfully.\n\nOpen the merged file?",
                    "AKSA PDF Tools", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                    Process.Start(sfd.FileName);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Merge PDF Error");
            }
        }

        // ===================== SPLIT PDF =====================

        public void OnSplitPdf(Office.IRibbonControl c)
        {
            try
            {
                var ofd = new OpenFileDialog
                {
                    Title = "Select PDF file to split",
                    Filter = "PDF Files (*.pdf)|*.pdf"
                };
                if (ofd.ShowDialog() != DialogResult.OK) return;

                using (var fb = new FolderBrowserDialog { Description = "Select destination folder for split pages" })
                {
                    if (fb.ShowDialog() != DialogResult.OK) return;

                    _pdf.SplitPdf(ofd.FileName, fb.SelectedPath);

                    MessageBox.Show(
                        "PDF split successfully.\nPages saved to: " + fb.SelectedPath,
                        "AKSA PDF Tools");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Split PDF Error");
            }
        }

        // ===================== EXTRACT PAGES =====================

        public void OnExtractPages(Office.IRibbonControl c)
        {
            try
            {
                var ofd = new OpenFileDialog
                {
                    Title = "Select PDF file",
                    Filter = "PDF Files (*.pdf)|*.pdf"
                };
                if (ofd.ShowDialog() != DialogResult.OK) return;

                string input = RequestInput(
                    "Extract Pages",
                    "Enter page numbers to extract (e.g. 1,3,5-8):",
                    "1");

                if (string.IsNullOrWhiteSpace(input)) return;

                var sfd = new SaveFileDialog
                {
                    Title = "Save extracted pages as",
                    Filter = "PDF Files (*.pdf)|*.pdf",
                    FileName = "Extracted.pdf"
                };
                if (sfd.ShowDialog() != DialogResult.OK) return;

                var pages = ParsePageRange(input);
                if (pages.Count == 0)
                {
                    MessageBox.Show("No valid page numbers entered.", "Extract Pages");
                    return;
                }

                _pdf.ExtractPages(ofd.FileName, sfd.FileName, pages);

                var result = MessageBox.Show(
                    "Pages extracted successfully.\n\nOpen the file?",
                    "AKSA PDF Tools", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                    Process.Start(sfd.FileName);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Extract Pages Error");
            }
        }

        private List<int> ParsePageRange(string input)
        {
            var pages = new List<int>();
            foreach (var part in input.Split(','))
            {
                var trimmed = part.Trim();
                if (trimmed.Contains('-'))
                {
                    var range = trimmed.Split('-');
                    if (int.TryParse(range[0], out int start) && int.TryParse(range[1], out int end))
                    {
                        for (int i = start; i <= end; i++)
                            pages.Add(i);
                    }
                }
                else
                {
                    if (int.TryParse(trimmed, out int num))
                        pages.Add(num);
                }
            }
            return pages.Distinct().OrderBy(p => p).ToList();
        }

        // ===================== PROTECT PDF =====================

        public void OnProtectPdf(Office.IRibbonControl c)
        {
            try
            {
                var ofd = new OpenFileDialog
                {
                    Title = "Select PDF file to protect",
                    Filter = "PDF Files (*.pdf)|*.pdf"
                };
                if (ofd.ShowDialog() != DialogResult.OK) return;

                string password = RequestInput(
                    "Protect PDF",
                    "Enter password to protect the PDF:",
                    "");

                if (string.IsNullOrWhiteSpace(password)) return;

                var sfd = new SaveFileDialog
                {
                    Title = "Save protected PDF as",
                    Filter = "PDF Files (*.pdf)|*.pdf",
                    FileName = Path.GetFileNameWithoutExtension(ofd.FileName) + "_protected.pdf"
                };
                if (sfd.ShowDialog() != DialogResult.OK) return;

                _pdf.ProtectPdf(ofd.FileName, sfd.FileName, password);

                MessageBox.Show("PDF protected successfully with password.", "AKSA PDF Tools");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Protect PDF Error");
            }
        }

        // ===================== UNLOCK PDF =====================

        public void OnUnlockPdf(Office.IRibbonControl c)
        {
            try
            {
                var ofd = new OpenFileDialog
                {
                    Title = "Select password-protected PDF",
                    Filter = "PDF Files (*.pdf)|*.pdf"
                };
                if (ofd.ShowDialog() != DialogResult.OK) return;

                string password = RequestInput(
                    "Unlock PDF",
                    "Enter the PDF password:",
                    "");

                if (string.IsNullOrWhiteSpace(password)) return;

                var sfd = new SaveFileDialog
                {
                    Title = "Save unlocked PDF as",
                    Filter = "PDF Files (*.pdf)|*.pdf",
                    FileName = Path.GetFileNameWithoutExtension(ofd.FileName) + "_unlocked.pdf"
                };
                if (sfd.ShowDialog() != DialogResult.OK) return;

                _pdf.UnlockPdf(ofd.FileName, sfd.FileName, password);

                MessageBox.Show("PDF unlocked successfully.", "AKSA PDF Tools");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Unlock PDF Error");
            }
        }

        // ===================== PDF INFO =====================

        public void OnPdfInfo(Office.IRibbonControl c)
        {
            try
            {
                var ofd = new OpenFileDialog
                {
                    Title = "Select PDF file",
                    Filter = "PDF Files (*.pdf)|*.pdf"
                };
                if (ofd.ShowDialog() != DialogResult.OK) return;

                var info = _pdf.GetPdfInfo(ofd.FileName);

                var msg = $"File: {Path.GetFileName(ofd.FileName)}\n" +
                          $"Pages: {info.PageCount}\n" +
                          $"Size: {info.FileSize}\n" +
                          $"Encrypted: {(info.Encrypted ? "Yes" : "No")}" +
                          (!string.IsNullOrEmpty(info.EncryptionInfo) ? $"\nEncryption: {info.EncryptionInfo}" : "");

                MessageBox.Show(msg, "PDF Info");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "PDF Info Error");
            }
        }

        // ===================== ABOUT =====================

        public void OnAbout(Office.IRibbonControl c)
        {
            MessageBox.Show(
                "AKSA PDF Tools v1.0\n\n" +
                "Word add-in for:\n" +
                "• Word → PDF conversion\n" +
                "• PDF → Word conversion\n" +
                "• Merge / Split / Extract PDF\n" +
                "• Protect / Unlock PDF\n" +
                "• Batch conversion\n\n" +
                "© 2026 AKSA DIGITAL POINT",
                "About AKSA PDF Tools");
        }

        // ===================== HELPERS =====================

        private static string RequestInput(string title, string prompt, string defaultValue)
        {
            var form = new Form
            {
                Text = title,
                Size = new System.Drawing.Size(400, 160),
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                StartPosition = FormStartPosition.CenterScreen
            };

            var lbl = new Label
            {
                Text = prompt,
                Left = 12,
                Top = 12,
                Width = 360,
                Height = 30
            };

            var txt = new TextBox
            {
                Left = 12,
                Top = 48,
                Width = 360,
                Height = 24,
                Text = defaultValue
            };

            var btnOk = new Button { Text = "OK", Left = 100, Top = 85, Width = 80, DialogResult = DialogResult.OK };
            var btnCancel = new Button { Text = "Cancel", Left = 200, Top = 85, Width = 80, DialogResult = DialogResult.Cancel };
            form.AcceptButton = btnOk;
            form.CancelButton = btnCancel;

            form.Controls.Add(lbl);
            form.Controls.Add(txt);
            form.Controls.Add(btnOk);
            form.Controls.Add(btnCancel);

            return form.ShowDialog() == DialogResult.OK ? txt.Text.Trim() : null;
        }
    }
}
