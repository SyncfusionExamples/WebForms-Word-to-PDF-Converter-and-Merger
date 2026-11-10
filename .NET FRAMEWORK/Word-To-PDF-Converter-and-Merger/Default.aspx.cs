using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using Syncfusion.Compression.Zip;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocToPDFConverter;
using Syncfusion.Pdf;
using Syncfusion.Pdf.Parsing;

namespace Word_To_PDF_Converter_and_Merger
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
        }

        // ============== FEATURE 1: WORD TO PDF ==============
        protected void btnConvertWordToPdf_Click(object sender, EventArgs e)
        {
            try
            {
                var posted = CollectPostedFiles(fuWordFiles);
                if (posted.Count == 0)
                {
                    ShowError("word", "Please add at least one Word document before converting.");
                    return;
                }

                var wordFiles = posted
                    .Where(f => IsWordExtension(Path.GetExtension(f.FileName)))
                    .ToList();

                if (wordFiles.Count == 0)
                {
                    ShowError("word", "No valid Word documents found. Please upload .doc or .docx files.");
                    return;
                }

                if (wordFiles.Count == 1)
                {
                    StreamSinglePdf(wordFiles[0]);
                }
                else
                {
                    StreamAsZip(wordFiles);
                }
            }
            catch (Exception ex)
            {
                ShowError("word", "Conversion failed: " + ex.Message);
            }
        }

        private void StreamSinglePdf(HttpPostedFile file)
        {
            byte[] pdfBytes;
            using (var inputStream = file.InputStream)
            {
                pdfBytes = ConvertWordToPdfBytes(inputStream, file.FileName);
            }
            var safeName = Path.GetFileNameWithoutExtension(file.FileName) + ".pdf";
            StreamPdfToResponse(pdfBytes, safeName);
        }

        private void StreamAsZip(List<HttpPostedFile> wordFiles)
        {
            using (var archive = new ZipArchive())
            {
                foreach (var f in wordFiles)
                {
                    var entryName = Path.GetFileNameWithoutExtension(f.FileName) + ".pdf";
                    using (var entryStream = new MemoryStream())
                    {
                        f.InputStream.CopyTo(entryStream);
                        entryStream.Position = 0;
                        byte[] pdfBytes = ConvertWordToPdfBytes(entryStream, f.FileName);
                        var finalStream = new MemoryStream(pdfBytes);
                        archive.AddItem(entryName, finalStream, true, FileAttributes.Normal);
                    }
                }
                using (var outStream = new MemoryStream())
                {
                    archive.Save(outStream, true);
                    var bytes = outStream.ToArray();
                    Response.Clear();
                    Response.ContentType = "application/zip";
                    Response.AddHeader("Content-Disposition", "attachment; filename=\"Converted_Documents.zip\"");
                    Response.AddHeader("Content-Length", bytes.Length.ToString());
                    SetDownloadCompleteCookie();
                    Response.OutputStream.Write(bytes, 0, bytes.Length);
                    Response.Flush();
                    Response.End();
                }
            }
        }

        private static byte[] ConvertWordToPdfBytes(Stream inputStream, string fileName)
        {
            string ext = Path.GetExtension(fileName);
            FormatType ft = string.Equals(ext, ".doc", StringComparison.OrdinalIgnoreCase)
                ? FormatType.Doc
                : FormatType.Docx;
            using (var wordDoc = new WordDocument(inputStream, ft))
            using (var converter = new DocToPDFConverter())
            using (var pdfDoc = converter.ConvertToPDF(wordDoc))
            {
                var outStream = new MemoryStream();
                pdfDoc.Save(outStream);
                pdfDoc.Close(true);
                return outStream.ToArray();
            }
        }

        // ============== FEATURE 2: MERGE TO PDF ==============
        protected void btnMergeToPdf_Click(object sender, EventArgs e)
        {
            try
            {
                var posted = CollectPostedFiles(fuMergeFiles);
                if (posted.Count < 2)
                {
                    ShowError("merge", "Please add at least two files to merge.");
                    return;
                }

                var orderedFiles = OrderByClient(posted, hfMergeOrder.Value);

                var validFiles = orderedFiles
                    .Where(f => IsWordExtension(Path.GetExtension(f.FileName)) ||
                                string.Equals(Path.GetExtension(f.FileName), ".pdf", StringComparison.OrdinalIgnoreCase))
                    .ToList();

                if (validFiles.Count < 2)
                {
                    ShowError("merge", "Please upload at least two valid Word or PDF files.");
                    return;
                }

                byte[] mergedBytes;
                var loadedDocs = new List<PdfLoadedDocument>();
                var tempStreams = new List<Stream>();
                try
                {
                    using (var finalDoc = new PdfDocument())
                    {
                        foreach (var file in validFiles)
                        {
                            Stream pdfStream = OpenFileAsPdfStream(file);
                            tempStreams.Add(pdfStream);
                            var loaded = new PdfLoadedDocument(pdfStream);
                            loadedDocs.Add(loaded);
                            if (loaded.Pages.Count > 0)
                            {
                                finalDoc.ImportPageRange(loaded, 0, loaded.Pages.Count - 1);
                            }
                        }
                        using (var outStream = new MemoryStream())
                        {
                            finalDoc.Save(outStream);
                            finalDoc.Close(true);
                            mergedBytes = outStream.ToArray();
                        }
                    }
                }
                finally
                {
                    foreach (var doc in loadedDocs)
                    {
                        try { doc.Close(true); } catch { /* ignore */ }
                    }
                    foreach (var stm in tempStreams)
                    {
                        try { stm.Dispose(); } catch { /* ignore */ }
                    }
                }

                StreamPdfToResponse(mergedBytes, "Merged_Document.pdf");
            }
            catch (Exception ex)
            {
                ShowError("merge", "Merge failed: " + ex.Message);
            }
        }

        private Stream OpenFileAsPdfStream(HttpPostedFile file)
        {
            string ext = Path.GetExtension(file.FileName);
            if (string.Equals(ext, ".pdf", StringComparison.OrdinalIgnoreCase))
            {
                var mem = new MemoryStream();
                file.InputStream.CopyTo(mem);
                mem.Position = 0;
                return mem;
            }

            string fileName = file.FileName;
            var ms = new MemoryStream();
            file.InputStream.CopyTo(ms);
            ms.Position = 0;
            FormatType ft = string.Equals(ext, ".doc", StringComparison.OrdinalIgnoreCase)
                ? FormatType.Doc
                : FormatType.Docx;
            using (var wordDoc = new WordDocument(ms, ft))
            using (var converter = new DocToPDFConverter())
            using (var pdfDoc = converter.ConvertToPDF(wordDoc))
            {
                var outMs = new MemoryStream();
                pdfDoc.Save(outMs);
                pdfDoc.Close(true);
                outMs.Position = 0;
                return outMs;
            }
        }

        // ============== UTILITIES ==============
        private List<HttpPostedFile> CollectPostedFiles(System.Web.UI.WebControls.FileUpload fu)
        {
            var list = new List<HttpPostedFile>();
            if (fu == null) return list;
            for (int i = 0; i < fu.PostedFiles.Count; i++)
            {
                var pf = fu.PostedFiles[i];
                if (pf != null && pf.ContentLength > 0)
                {
                    list.Add(pf);
                }
            }
            return list;
        }

        private static bool IsWordExtension(string ext)
        {
            if (string.IsNullOrEmpty(ext)) return false;
            ext = ext.ToLowerInvariant();
            return ext == ".doc" || ext == ".docx";
        }

        private static List<HttpPostedFile> OrderByClient(List<HttpPostedFile> files, string orderList)
        {
            if (string.IsNullOrEmpty(orderList))
                return files;

            var byName = files
                .GroupBy(f => f.FileName, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

            var ordered = new List<HttpPostedFile>();
            var usedKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var name in orderList.Split('|'))
            {
                if (string.IsNullOrWhiteSpace(name)) continue;
                if (byName.TryGetValue(name, out var pf) && usedKeys.Add(name))
                {
                    ordered.Add(pf);
                }
            }
            foreach (var kvp in byName)
            {
                if (!usedKeys.Contains(kvp.Key))
                {
                    ordered.Add(kvp.Value);
                }
            }
            return ordered;
        }

        private void StreamPdfToResponse(byte[] pdfBytes, string downloadFileName)
        {
            Response.Clear();
            Response.ContentType = "application/pdf";
            Response.AddHeader("Content-Disposition", "attachment; filename=\"" + downloadFileName + "\"");
            Response.AddHeader("Content-Length", pdfBytes.Length.ToString());
            SetDownloadCompleteCookie();
            Response.OutputStream.Write(pdfBytes, 0, pdfBytes.Length);
            Response.Flush();
            Response.End();
        }

        private static void SetDownloadCompleteCookie()
        {
            HttpContext.Current.Response.Cookies["wp_download_done"].Value = "1";
            HttpContext.Current.Response.Cookies["wp_download_done"].Path = "/";
            HttpContext.Current.Response.Cookies["wp_download_done"].Expires = DateTime.Now.AddMinutes(1);
        }

        private void ShowError(string panelKey, string message)
        {
            var alertHost = panelKey == "word" ? wpAlertHost : mpAlertHost;
            if (alertHost == null) return;
            alertHost.InnerHtml =
                "<div class='custom-alert custom-alert-danger'>" +
                "<div><strong>" + System.Web.HttpUtility.HtmlEncode(message) + "</strong></div>" +
                "<button type='button' class='btn-close' onclick=\"this.parentNode.remove()\"></button>" +
                "</div>";
        }
    }
}
