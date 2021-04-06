using System;
using System.IO;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Core;

namespace TemplateEngine.Docx.Runner
{
    /// <summary>
    /// Works only on Windows with installed Ms Word.
    /// Took the code from https://stackoverflow.com/questions/607669/how-do-i-convert-word-files-to-pdf-programmatically
    /// </summary>
    public static class DocumentFormatConvertor
    {
        public static void Convert(string docPath, string outputFolder, string strFormat)
        {
            var wordApp = CreateWordApp();
            var timer = Stopwatch.StartNew();

            var doc = wordApp.Documents.Open(docPath, MsoTriState.msoTrue, MsoTriState.msoTrue, MsoTriState.msoFalse);
            doc.Activate();

            try
            {
                var targetFormat = Enum.Parse<WdSaveFormat>(strFormat);

                // See https://social.msdn.microsoft.com/Forums/vstudio/en-US/b28e41a5-e476-470d-9b8f-e9f5ab0f33ca/default-file-extension-from-wdsaveformat?forum=vsto
                var outputPath = Path.Combine(outputFolder, ChangeFileExtension(docPath, targetFormat));

#if true
                doc.SaveAs2(outputPath, targetFormat, MisValue, MisValue, MsoTriState.msoFalse,
                    MisValue, MsoTriState.msoTrue, MsoTriState.msoTrue, MisValue, MisValue, MisValue, MisValue);

#else
                doc.ExportAsFixedFormat(
                    outputPath,
                    WdExportFormat.wdExportFormatPDF,
                    false,
                    WdExportOptimizeFor.wdExportOptimizeForPrint,
                    WdExportRange.wdExportAllDocument
                );
#endif

                timer.Stop();
                Console.WriteLine($"Converted document path: '{outputPath}'");
                Console.WriteLine($"Time spent (without Word app start) to convert document: '{timer.Elapsed}'");
            }
            finally
            {
                doc.Close(MsoTriState.msoFalse, MisValue, MisValue);
                wordApp.Quit();

                releaseObject(doc);
                releaseObject(wordApp);
            }
        }

        private static string ChangeFileExtension(string path, WdSaveFormat fmt)
        {
#if true
            return Path.GetFileNameWithoutExtension(path);
#else
            var ext = GetTargetExtension(fmt);
            path = Path.ChangeExtension(path, ext);

            return path;
#endif
        }

        private static string GetTargetExtension(WdSaveFormat fmt)
        {
            if (fmt == WdSaveFormat.wdFormatPDF)
                return "pdf";
            else if (fmt == WdSaveFormat.wdFormatRTF)
                return "rtf";
            else if (fmt == WdSaveFormat.wdFormatHTML)
                return "html";
            else if (fmt == WdSaveFormat.wdFormatOpenDocumentText)
                return "odt";
            else
                throw new NotSupportedException($"Target convertion format '{fmt}' is not supported yet.");
        }

        private static Application CreateWordApp()
        {
            var wordApp = new Application();
            wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;
            wordApp.Visible = false;
            wordApp.Options.SavePropertiesPrompt = false;
            wordApp.Options.SaveNormalPrompt = false;

            return wordApp;
        }

        private static void releaseObject(object obj)
        {
            try
            {
                Marshal.ReleaseComObject(obj);
            }
            finally
            {
                GC.Collect();
            }
        }

        private static readonly object MisValue = System.Reflection.Missing.Value;
    }
}
