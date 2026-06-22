using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace AksaPdfAddin
{
    public class PdfInfo
    {
        public int PageCount { get; set; }
        public string FileSize { get; set; }
        public bool Encrypted { get; set; }
        public string EncryptionInfo { get; set; }
    }

    public class PdfService
    {
        // ===================== WORD → PDF =====================

        public void WordToPdf(Word.Document doc, string pdfPath)
        {
            object missing = System.Reflection.Missing.Value;
            object outputFile = pdfPath;
            object fileFormat = Word.WdSaveFormat.wdFormatPDF;

            doc.SaveAs2(ref outputFile,
                ref fileFormat, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing);
        }

        // ===================== PDF → WORD =====================

        public void PdfToWord(Word.Application app, string pdfPath, string docxPath)
        {
            object missing = System.Reflection.Missing.Value;
            object confirmConversions = false;
            object readOnly = false;
            object addToRecentFiles = false;
            object format = Word.WdOpenFormat.wdOpenFormatAuto;

            var doc = app.Documents.Open(pdfPath,
                ref confirmConversions, ref readOnly, ref addToRecentFiles,
                ref missing, ref missing, ref missing, ref missing,
                ref missing, ref format, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing);

            doc.SaveAs2(docxPath, Word.WdSaveFormat.wdFormatDocumentDefault);
            doc.Close();
        }

        public void BatchWordToPdf(Word.Application app, string wordPath, string pdfPath)
        {
            object missing = System.Reflection.Missing.Value;
            object confirmConversions = false;
            object readOnly = true;
            object addToRecentFiles = false;

            var doc = app.Documents.Open(wordPath,
                ref confirmConversions, ref readOnly, ref addToRecentFiles,
                ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing);

            WordToPdf(doc, pdfPath);
            doc.Close();
        }

        public void BatchPdfToWord(Word.Application app, string pdfPath, string docxPath)
        {
            PdfToWord(app, pdfPath, docxPath);
        }

        // ===================== MERGE PDF =====================

        public void MergePdf(List<string> pdfFiles, string outputPath)
        {
            // Uses Word's native PDF merging via InsertFile
            // Alternatively uses a PDF library helper
            NativeMergeUsingWord(pdfFiles, outputPath);
        }

        private void NativeMergeUsingWord(List<string> pdfFiles, string outputPath)
        {
            // Create a temporary Word document and insert each PDF as content
            Word.Application app = null;
            Word.Document doc = null;

            try
            {
                app = new Word.Application { Visible = false };
                doc = app.Documents.Add();

                for (int i = 0; i < pdfFiles.Count; i++)
                {
                    if (i > 0)
                    {
                        object breakType = Word.WdBreakType.wdSectionBreakNextPage;
                        doc.Content.InsertBreak(ref breakType);
                    }

                    object missing = System.Reflection.Missing.Value;
                    object fileName = pdfFiles[i];
                    object confirmConversions = false;
                    object link = false;
                    object attachment = false;

                    doc.Application.Selection.InsertFile(ref fileName,
                        ref missing, ref confirmConversions, ref link, ref attachment);
                }

                WordToPdf(doc, outputPath);
            }
            finally
            {
                if (doc != null) { doc.Close(false); Marshal.ReleaseComObject(doc); }
                if (app != null) { app.Quit(); Marshal.ReleaseComObject(app); }
            }
        }

        // ===================== SPLIT PDF =====================

        public void SplitPdf(string pdfPath, string outputFolder)
        {
            NativeSplitUsingWord(pdfPath, outputFolder);
        }

        private void NativeSplitUsingWord(string pdfPath, string outputFolder)
        {
            Word.Application app = null;
            Word.Document doc = null;

            try
            {
                app = new Word.Application { Visible = false };

                object missing = System.Reflection.Missing.Value;
                object confirmConversions = false;
                object readOnly = false;
                object addToRecentFiles = false;
                object format = Word.WdOpenFormat.wdOpenFormatAuto;

                doc = app.Documents.Open(pdfPath,
                    ref confirmConversions, ref readOnly, ref addToRecentFiles,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref format, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing);

                int pageCount = doc.ComputeStatistics(Word.WdStatistic.wdStatisticPages);

                for (int i = 1; i <= pageCount; i++)
                {
                    doc.Application.Selection.GoTo(Word.WdGoToItem.wdGoToPage,
                        Word.WdGoToDirection.wdGoToFirst, missing, i);

                    var range = doc.Application.Selection.Range;
                    range.Collapse(Word.WdCollapseDirection.wdCollapseStart);

                    if (i < pageCount)
                    {
                        object what = Word.WdGoToItem.wdGoToPage;
                        object which = Word.WdGoToDirection.wdGoToAbsolute;
                        object count = i + 1;
                        doc.Application.Selection.GoTo(ref what, ref which, ref count, ref missing);

                        var endRange = doc.Application.Selection.Range;
                        endRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);

                        range.End = endRange.Start;
                    }
                    else
                    {
                        range.End = doc.Content.End;
                    }

                    range.Copy();

                    var pageDoc = app.Documents.Add();
                    pageDoc.Application.Selection.Paste();
                    string pagePdf = Path.Combine(outputFolder,
                        $"{Path.GetFileNameWithoutExtension(pdfPath)}_page{i:D3}.pdf");
                    WordToPdf(pageDoc, pagePdf);
                    pageDoc.Close(false);
                }
            }
            finally
            {
                if (doc != null) { doc.Close(false); Marshal.ReleaseComObject(doc); }
                if (app != null) { app.Quit(); Marshal.ReleaseComObject(app); }
            }
        }

        // ===================== EXTRACT PAGES =====================

        public void ExtractPages(string pdfPath, string outputPath, List<int> pages)
        {
            NativeExtractUsingWord(pdfPath, outputPath, pages);
        }

        private void NativeExtractUsingWord(string pdfPath, string outputPath, List<int> pages)
        {
            Word.Application app = null;
            Word.Document doc = null;
            Word.Document resultDoc = null;

            try
            {
                app = new Word.Application { Visible = false };

                object missing = System.Reflection.Missing.Value;
                object confirmConversions = false;
                object readOnly = false;
                object addToRecentFiles = false;
                object format = Word.WdOpenFormat.wdOpenFormatAuto;

                doc = app.Documents.Open(pdfPath,
                    ref confirmConversions, ref readOnly, ref addToRecentFiles,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref format, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing);

                resultDoc = app.Documents.Add();
                int pageCount = doc.ComputeStatistics(Word.WdStatistic.wdStatisticPages);

                for (int pi = 0; pi < pages.Count; pi++)
                {
                    int pageNum = pages[pi];
                    if (pageNum < 1 || pageNum > pageCount) continue;

                    if (pi > 0)
                    {
                        object breakType = Word.WdBreakType.wdSectionBreakNextPage;
                        resultDoc.Content.InsertBreak(ref breakType);
                    }

                    doc.Application.Selection.GoTo(Word.WdGoToItem.wdGoToPage,
                        Word.WdGoToDirection.wdGoToFirst, missing, pageNum);

                    var range = doc.Application.Selection.Range;
                    range.Collapse(Word.WdCollapseDirection.wdCollapseStart);

                    if (pageNum < pageCount)
                    {
                        object what = Word.WdGoToItem.wdGoToPage;
                        object which = Word.WdGoToDirection.wdGoToAbsolute;
                        object count = pageNum + 1;
                        doc.Application.Selection.GoTo(ref what, ref which, ref count, ref missing);

                        var endRange = doc.Application.Selection.Range;
                        endRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                        range.End = endRange.Start;
                    }
                    else
                    {
                        range.End = doc.Content.End;
                    }

                    range.Copy();
                    resultDoc.Application.Selection.Paste();
                }

                WordToPdf(resultDoc, outputPath);
            }
            finally
            {
                if (resultDoc != null) { resultDoc.Close(false); Marshal.ReleaseComObject(resultDoc); }
                if (doc != null) { doc.Close(false); Marshal.ReleaseComObject(doc); }
                if (app != null) { app.Quit(); Marshal.ReleaseComObject(app); }
            }
        }

        // ===================== PROTECT PDF =====================

        public void ProtectPdf(string inputPath, string outputPath, string password)
        {
            NativeProtectUsingWord(inputPath, outputPath, password);
        }

        private void NativeProtectUsingWord(string inputPath, string outputPath, string password)
        {
            Word.Application app = null;
            Word.Document doc = null;

            try
            {
                app = new Word.Application { Visible = false };

                object missing = System.Reflection.Missing.Value;
                object confirmConversions = false;
                object readOnly = false;
                object addToRecentFiles = false;
                object format = Word.WdOpenFormat.wdOpenFormatAuto;

                doc = app.Documents.Open(inputPath,
                    ref confirmConversions, ref readOnly, ref addToRecentFiles,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref format, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing);

                // Password protect the Word document before saving as PDF
                object passwordObj = password;
                doc.Password = passwordObj.ToString();

                WordToPdf(doc, outputPath);
            }
            finally
            {
                if (doc != null) { doc.Close(false); Marshal.ReleaseComObject(doc); }
                if (app != null) { app.Quit(); Marshal.ReleaseComObject(app); }
            }
        }

        // ===================== UNLOCK PDF =====================

        public void UnlockPdf(string inputPath, string outputPath, string password)
        {
            // Opens the password-protected PDF (Word will prompt for password)
            // and saves as unprotected PDF
            Word.Application app = null;
            Word.Document doc = null;

            try
            {
                app = new Word.Application { Visible = false };

                object missing = System.Reflection.Missing.Value;
                object confirmConversions = false;
                object readOnly = false;
                object addToRecentFiles = false;
                object format = Word.WdOpenFormat.wdOpenFormatAuto;
                object pwd = password;

                doc = app.Documents.Open(inputPath,
                    ref confirmConversions, ref readOnly, ref addToRecentFiles,
                    ref missing, ref missing, ref missing, ref missing,
                    ref pwd, ref format, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing);

                doc.Password = "";

                WordToPdf(doc, outputPath);
            }
            finally
            {
                if (doc != null) { doc.Close(false); Marshal.ReleaseComObject(doc); }
                if (app != null) { app.Quit(); Marshal.ReleaseComObject(app); }
            }
        }

        // ===================== PDF INFO =====================

        public PdfInfo GetPdfInfo(string pdfPath)
        {
            var info = new PdfInfo();
            var file = new FileInfo(pdfPath);
            info.FileSize = FormatFileSize(file.Length);

            Word.Application app = null;
            Word.Document doc = null;

            try
            {
                app = new Word.Application { Visible = false };

                object missing = System.Reflection.Missing.Value;
                object confirmConversions = false;
                object readOnly = true;
                object addToRecentFiles = false;
                object format = Word.WdOpenFormat.wdOpenFormatAuto;

                doc = app.Documents.Open(pdfPath,
                    ref confirmConversions, ref readOnly, ref addToRecentFiles,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref format, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing);

                info.PageCount = doc.ComputeStatistics(Word.WdStatistic.wdStatisticPages);
                info.Encrypted = false;

                // Check protection
                if (doc.ProtectionType != Word.WdProtectionType.wdNoProtection)
                {
                    info.Encrypted = true;
                    info.EncryptionInfo = doc.ProtectionType.ToString();
                }
            }
            catch (Exception)
            {
                info.Encrypted = true;
                info.EncryptionInfo = "Password protected";
            }
            finally
            {
                if (doc != null) { doc.Close(false); Marshal.ReleaseComObject(doc); }
                if (app != null) { app.Quit(); Marshal.ReleaseComObject(app); }
            }

            return info;
        }

        private static string FormatFileSize(long bytes)
        {
            string[] sizes = { "B", "KB", "MB", "GB" };
            int order = 0;
            double len = bytes;
            while (len >= 1024 && order < sizes.Length - 1)
            {
                order++;
                len = len / 1024;
            }
            return $"{len:0.##} {sizes[order]}";
        }
    }
}
