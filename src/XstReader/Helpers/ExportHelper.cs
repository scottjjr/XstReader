// Project site: https://github.com/iluvadev/XstReader
//
// Based on the great work of Dijji. 
// Original project: https://github.com/dijji/XstReader
//
// Issues: https://github.com/iluvadev/XstReader/issues
// License (Ms-PL): https://github.com/iluvadev/XstReader/blob/master/license.md
//
// Copyright (c) 2021, iluvadev, and released under Ms-PL License.

using System.Diagnostics;
using XstReader.Exporter;

namespace XstReader.App.Helpers
{
    public static class ExportHelper
    {
        private static readonly SaveFileDialog SaveFileDialog = new();
        private static readonly FolderBrowserDialog FolderBrowserDialog = new() { ShowNewFolderButton = true };

        private static bool ConfigureExport()
            => !ExportOptionsForm.IsFirstTime || new ExportOptionsForm().ShowDialog() == DialogResult.OK;

        private static bool AskDirectoryPath(ref string path)
        {
            if (!string.IsNullOrWhiteSpace(path))
                FolderBrowserDialog.SelectedPath = path;
            if (FolderBrowserDialog.ShowDialog() != DialogResult.OK)
                return false;
            path = FolderBrowserDialog.SelectedPath;
            return true;
        }

        private static bool AskFileName(ref string fileName)
        {
            if (!string.IsNullOrWhiteSpace(fileName))
                SaveFileDialog.FileName = fileName;
            if (SaveFileDialog.ShowDialog() != DialogResult.OK)
                return false;
            fileName = SaveFileDialog.FileName;
            return true;
        }

        private static string GetFilePrefixFromFile(XstFile file)
        {
            try
            {
                var type = file.GetType();
                var prop = type.GetProperty("FileName") ?? type.GetProperty("Path") ?? type.GetProperty("Name");
                var full = prop?.GetValue(file) as string;
                if (!string.IsNullOrWhiteSpace(full))
                {
                    var name = System.IO.Path.GetFileNameWithoutExtension(full);
                    return string.IsNullOrWhiteSpace(name) ? string.Empty : name + "_";
                }
            }
            catch { }
            return string.Empty;
        }

        private static string GetFilePrefix(XstElement elem, XstFile? owningFile)
        {
            if (elem is XstFile file) return GetFilePrefixFromFile(file);
            if (owningFile != null) return GetFilePrefixFromFile(owningFile);
            return string.Empty;
        }

        private static void RunExport(XstElement elem, string basePath, XstFile? owningFile, Action<XstExporter,string> exporterAction)
        {
            using (var frm = new WaitingForm($"Exporting from {elem.DisplayName}"))
            {
                var prefix = GetFilePrefix(elem, owningFile);
                var path = basePath;
                if (!string.IsNullOrEmpty(prefix))
                {
                    path = System.IO.Path.Combine(basePath, prefix.TrimEnd('_'));
                    try { System.IO.Directory.CreateDirectory(path); } catch { }
                }
                var exporter = new XstExporter(XstReaderEnvironment.Options.ExportOptions, frm.ReportExportProgress);
                frm.Start(() => exporterAction(exporter, path));
                frm.ShowDialog();
                try { Process.Start(new ProcessStartInfo() { FileName = path, UseShellExecute = true }); } catch { }
            }
        }

        // New: explicit message export with owning file prefix
        public static bool ExportMessage(XstMessage? message, XstFile? owningFile)
        {
            if (message == null) return false;
            string path = "";
            if (!ConfigureExport() || !AskDirectoryPath(ref path)) return false;
            RunExport(message, path, owningFile, (ex,p) => ex.SaveMessage(message, p));
            return true;
        }

        // UPDATED: now accepts owningFile for folder/message prefixing
        public static bool ExportMessages<T>(T? elem, XstFile? owningFile = null) where T : XstElement
        {
            string path = "";
            if (elem != null && ConfigureExport() && AskDirectoryPath(ref path))
            {
                if (elem is XstFile file)
                    RunExport(file, path, file, (ex,p) => ex.SaveMessages(file, p));
                else if (elem is XstFolder folder)
                    RunExport(folder, path, owningFile, (ex,p) => ex.SaveMessages(folder, p));
                else if (elem is XstMessage message)
                    RunExport(message, path, owningFile, (ex,p) => ex.SaveMessage(message, p));
            }
            return false;
        }

        // UPDATED: now accepts owningFile for folder/message prefixing
        public static bool ExportAttachments<T>(T? elem, XstFile? owningFile = null) where T : XstElement
        {
            string path = "";
            if (elem != null && ConfigureExport() && AskDirectoryPath(ref path))
            {
                if (elem is XstFile file)
                    RunExport(file, path, file, (ex,p) => ex.SaveAttachments(file, p));
                else if (elem is XstFolder folder)
                    RunExport(folder, path, owningFile, (ex,p) => ex.SaveAttachments(folder, p));
                else if (elem is XstMessage message)
                    RunExport(message, path, owningFile, (ex,p) => ex.SaveAttachments(message, p));
            }
            return false;
        }

        public static bool ExportMessageAttachments(XstMessage? message, XstFile? owningFile)
        {
            if (message == null) return false;
            string path = "";
            if (!ConfigureExport() || !AskDirectoryPath(ref path)) return false;
            RunExport(message, path, owningFile, (ex,p) => ex.SaveAttachments(message, p));
            return true;
        }
    }
}
