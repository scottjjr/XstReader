// Project site: https://github.com/iluvadev/XstReader
//
// Based on the great work of Dijji. 
// Original project: https://github.com/dijji/XstReader
//
// Issues: https://github.com/iluvadev/XstReader/issues
// License (Ms-PL): https://github.com/iluvadev/XstReader/blob/master/license.md
//
// Copyright (c) 2021, iluvadev, and released under Ms-PL License.

using Krypton.Docking;
using Krypton.Toolkit;
using System.Data;
using XstReader.App.Controls;
using XstReader.App.Helpers;
using System.Collections.Generic;

namespace XstReader.App
{
    public partial class MainForm : KryptonForm
    {
        private XstFolderTreeControl FolderTreeControl { get; } = new XstFolderTreeControl() { Name = "Folders Tree" };
        private XstMessageListControl MessageListControl { get; } = new XstMessageListControl() { Name = "Message List" };
        private XstMessageViewControl MessageViewControl { get; } = new XstMessageViewControl() { Name = "Message View" };
        private XstPropertiesControl PropertiesControl { get; } = new XstPropertiesControl() { Name = "Properties" };
        private XstPropertiesInfoControl InfoControl { get; } = new XstPropertiesInfoControl() { Name = "Information" };

        private readonly List<XstFile> _XstFiles = new();
        private XstFile? _CurrentFile = null; // for Close file operation
        private XstFile? CurrentFile { get => _CurrentFile; set => _CurrentFile = value; }

        private XstElement? _CurrentXstElement = null;
        private XstElement? CurrentXstElement { get => _CurrentXstElement; set => SetCurrentXstElement(value); }
        private void SetCurrentXstElement(XstElement? value)
        {
            if (_CurrentXstElement == value) return;
            _CurrentXstElement?.ClearContents();

            if (value is XstFolder folder)
            {
                MessageListControl.SetDataSource(folder?.Messages?.OrderByDescending(m => m.Date));
                MessageListControl.SetError(folder?.HasErrorInMessages ?? false, folder?.ErrorInMessages ?? "");
                var owningFile = FolderTreeControl.GetSelectedFile();
                if (owningFile != null) CurrentFile = owningFile;
            }
            else if (value is XstMessage message)
            {
                MessageViewControl.SetDataSource(message);
                var owningFile = FolderTreeControl.GetSelectedFile();
                if (owningFile != null) CurrentFile = owningFile;
            }

            MessageToolStripMenuItem.Enabled = value is XstMessage;
            InfoControl.SetDataSource(value);
            PropertiesControl.SetDataSource(value);
            UpdateMenu();
            _CurrentXstElement = value;
        }

        private XstFile? GetCurrentXstFile() => CurrentFile;
        private XstFolder? GetCurrentXstFolder() => FolderTreeControl.GetSelectedItem();
        private XstMessage? GetCurrentXstMessage() => MessageViewControl.GetDataSource();

        private void UpdateMenu()
        {
            // File exports now operate on ALL files; enable if any loaded
            FileExportFoldersToolStripMenuItem.Enabled = FileExportAttachmentsToolStripMenuItem.Enabled = _XstFiles.Count > 0;
            FolderToolStripMenuItem.Enabled = GetCurrentXstFolder() != null;
            MessageToolStripMenuItem.Enabled = GetCurrentXstMessage() != null;
            MessageExportAttachmentsToolStripMenuItem.Enabled = GetCurrentXstMessage()?.Attachments?.Any(a => a.IsFile) ?? false;
            CloseFileToolStripMenuItem.Enabled = CurrentFile != null;
        }

        public MainForm() { InitializeComponent(); Initialize(); }

        private void Initialize()
        {
            OpenXstFileDialog.Multiselect = true;
            OpenToolStripMenuItem.Click += OpenXstFile;
            CloseFileToolStripMenuItem.Click += (s, e) => CloseCurrentXstFile();

            FolderTreeControl.SelectedItemChanged += (s, e) => CurrentXstElement = e.Element;
            FolderTreeControl.FileActionRequested += FolderTreeControl_FileActionRequested;
            FolderTreeControl.GotFocus += (s, e) => CurrentXstElement = FolderTreeControl.GetSelectedItem();
            MessageListControl.SelectedItemChanged += (s, e) => CurrentXstElement = e.Element;
            MessageListControl.GotFocus += (s, e) => CurrentXstElement = MessageListControl.GetSelectedItem();
            MessageViewControl.SelectedItemChanged += (s, e) => CurrentXstElement = e.Element;
            MessageViewControl.GotFocus += (s, e) => CurrentXstElement = MessageViewControl.GetSelectedItem();
            MessageViewControl.ExportMessageRequested += (s, msg) =>
            {
                var owningFile = FolderTreeControl.GetSelectedFile();
                ExportHelper.ExportMessage(msg, owningFile);
            }; // NEW handler using owning file

            AboutToolStripMenuItem.Click += (s, e) => { using var f = new AboutForm(); f.ShowDialog(); };

            FileExportFoldersToolStripMenuItem.Click += (s, e) => { foreach (var f in _XstFiles) ExportHelper.ExportMessages(f, f); };
            FileExportAttachmentsToolStripMenuItem.Click += (s, e) => { foreach (var f in _XstFiles) ExportHelper.ExportAttachments(f, f); };

            FolderExportMessagesToolStripMenuItem.Click += (s, e) =>
            {
                var folder = GetCurrentXstFolder();
                var owningFile = FolderTreeControl.GetSelectedFile();
                ExportHelper.ExportMessages(folder, owningFile);
            };
            FolderExportAttachmentsToolStripMenuItem.Click += (s, e) =>
            {
                var folder = GetCurrentXstFolder();
                var owningFile = FolderTreeControl.GetSelectedFile();
                ExportHelper.ExportAttachments(folder, owningFile);
            };
            MessagePrintToolStripMenuItem.Click += (s, e) => MessageViewControl.Print();
            MessageExportToolStripMenuItem.Click += (s, e) =>
            {
                var msg = GetCurrentXstMessage();
                var owningFile = FolderTreeControl.GetSelectedFile();
                ExportHelper.ExportMessages(msg, owningFile);
            };
            MessageExportAttachmentsToolStripMenuItem.Click += (s, e) =>
            {
                var msg = GetCurrentXstMessage();
                var owningFile = FolderTreeControl.GetSelectedFile();
                ExportHelper.ExportMessageAttachments(msg, owningFile);
            };
            SettingsToolStripMenuItem.Click += (s, e) => new ExportOptionsForm().ShowDialog();

            LayoutDefaultToolStripMenuItem.Click += (s, e) => { try { var tmp = Path.GetTempFileName() + ".xml"; File.WriteAllBytes(tmp, Properties.Resources.layout_default); KryptonDockingManager.LoadConfigFromFile(tmp); File.Delete(tmp); } catch { } };
            LayoutClassic3PanelToolStripMenuItem.Click += (s, e) => { try { var tmp = Path.GetTempFileName() + ".xml"; File.WriteAllBytes(tmp, Properties.Resources.layout_3panels); KryptonDockingManager.LoadConfigFromFile(tmp); File.Delete(tmp); } catch { } };

            ResetAllFiles();
            UpdateMenu();
        }

        private void FolderTreeControl_FileActionRequested(object? sender, XstFileActionEventArgs e)
        {
            switch (e.Action)
            {
                case XstFileAction.ExportMessages:
                    ExportHelper.ExportMessages(e.File);
                    break;
                case XstFileAction.ExportAttachments:
                    ExportHelper.ExportAttachments(e.File);
                    break;
                case XstFileAction.CloseFile:
                    CurrentFile = e.File; // set current then close
                    CloseCurrentXstFile();
                    break;
            }
            UpdateMenu();
        }

        private void OpenXstFile(object? sender, EventArgs e)
        {
            if (OpenXstFileDialog.ShowDialog(this) != DialogResult.OK) return;
            foreach (var fileName in OpenXstFileDialog.FileNames) AddXstFile(fileName);
            FolderTreeControl.SetDataSources(_XstFiles);
            if (CurrentFile == null) CurrentFile = _XstFiles.FirstOrDefault();
            UpdateMenu();
        }

        private void ResetAllFiles()
        {
            foreach (var f in _XstFiles) { try { f.ClearContents(); f.Dispose(); } catch { } }
            _XstFiles.Clear();
            CurrentFile = null;
            FolderTreeControl.ClearContents();
            FolderTreeControl.SetDataSources(_XstFiles);
            MessageListControl.ClearContents();
            MessageViewControl.ClearContents();
            PropertiesControl.ClearContents();
            InfoControl.ClearContents();
        }

        private void AddXstFile(string filename)
        {
            var xst = new XstFile(filename);
            _XstFiles.Add(xst);
            if (CurrentFile == null) CurrentFile = xst;
        }

        private void CloseCurrentXstFile()
        {
            if (CurrentFile == null) return;
            try { CurrentFile.ClearContents(); CurrentFile.Dispose(); } catch { }
            _XstFiles.Remove(CurrentFile);
            CurrentFile = _XstFiles.FirstOrDefault();
            FolderTreeControl.SetDataSources(_XstFiles);
            if (CurrentFile == null)
            {
                FolderTreeControl.ClearContents();
                MessageListControl.ClearContents();
                MessageViewControl.ClearContents();
                PropertiesControl.ClearContents();
                InfoControl.ClearContents();
            }
            UpdateMenu();
        }

        protected override void OnClosed(EventArgs e)
        {
            try { KryptonDockingManager.SaveConfigToFile(Path.Combine(Application.StartupPath, "Layout.xml")); } catch { }
            try { XstReaderOptions.SaveToFile(Path.Combine(Application.StartupPath, "Options.xml"), XstReaderEnvironment.Options); } catch { }
            ResetAllFiles();
            base.OnClosed(e);
        }

        protected override void OnLoad(EventArgs e) { base.OnLoad(e); LoadForm(); }

        private void LoadForm()
        {
            KryptonMessagePanel.BeginInit();
            KryptonMessagePanel.Controls.Add(MessageViewControl);
            MessageViewControl.Dock = DockStyle.Fill;
            KryptonDockingManager.ManageControl(KryptonMainPanel);
            KryptonDockingManager.AddXstDockSpaceInTabs(DockingEdge.Top, MessageListControl);
            KryptonDockingManager.AddXstDockSpaceInTabs(DockingEdge.Right, InfoControl, PropertiesControl);
            KryptonDockingManager.AddXstDockSpaceInTabs(DockingEdge.Left, FolderTreeControl);
            try { KryptonDockingManager.LoadConfigFromFile(Path.Combine(Application.StartupPath, "Layout.xml")); }
            catch { LayoutDefaultToolStripMenuItem.PerformClick(); }
            KryptonMessagePanel.EndInit();
            try { XstReaderEnvironment.Options = XstReaderOptions.LoadFromFile(Path.Combine(Application.StartupPath, "Options.xml")); } catch { }
        }
    }
}
