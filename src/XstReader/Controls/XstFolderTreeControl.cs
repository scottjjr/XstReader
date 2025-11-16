// Project site: https://github.com/iluvadev/XstReader
//
// Based on the great work of Dijji. 
// Original project: https://github.com/dijji/XstReader
//
// Issues: https://github.com/iluvadev/XstReader/issues
// License (Ms-PL): https://github.com/iluvadev/XstReader/blob/master/license.md
//
// Copyright (c) 2021, iluvadev, and released under Ms-PL License.

using Krypton.Toolkit;
using XstReader.App.Common;
using System.IO;
using System.Linq;

namespace XstReader.App.Controls
{
    public enum XstFileAction { ExportMessages, ExportAttachments, CloseFile }
    public class XstFileActionEventArgs : EventArgs
    {
        public XstFile File { get; }
        public XstFileAction Action { get; }
        public XstFileActionEventArgs(XstFile file, XstFileAction action) { File = file; Action = action; }
    }

    public partial class XstFolderTreeControl : UserControl,
                                                IXstDataSourcedControl<XstFile>,
                                                IXstElementSelectable<XstFolder>
    {
        public event EventHandler<XstFileActionEventArgs>? FileActionRequested; // NEW

        public XstFolderTreeControl()
        {
            InitializeComponent();
            Initialize();
        }

        private ContextMenuStrip? _fileContextMenu; // NEW

        private void Initialize()
        {
            if (DesignMode) return;

            MainTreeView.AfterSelect += (s, e) => RaiseSelectedItemChanged();
            MainTreeView.GotFocus += (s, e) => OnGotFocus(e);
            MainTreeView.MouseUp += MainTreeView_MouseUp; // NEW
            CreateFileContextMenu(); // NEW
            SetDataSource(null);
        }
        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            MainTreeView.BackColor = this.BackColor;
            MainTreeView.Font = this.Font;
            MainTreeView.ForeColor = this.ForeColor;
        }

        private void CreateFileContextMenu()
        {
            _fileContextMenu = new ContextMenuStrip();
            _fileContextMenu.Items.Add("Export all Messages", null, (s, e) => RaiseFileAction(XstFileAction.ExportMessages));
            _fileContextMenu.Items.Add("Export all Attachments", null, (s, e) => RaiseFileAction(XstFileAction.ExportAttachments));
            _fileContextMenu.Items.Add(new ToolStripSeparator());
            _fileContextMenu.Items.Add("Close File", null, (s, e) => RaiseFileAction(XstFileAction.CloseFile));
        }

        private KryptonTreeNode? _contextNode; // track node for which menu opened
        private void MainTreeView_MouseUp(object? sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right) return;
            var node = MainTreeView.GetNodeAt(e.Location) as KryptonTreeNode;
            if (node == null) return;
            _contextNode = node;
            // Show context only for root nodes (file roots)
            if (node.Parent == null)
            {
                MainTreeView.SelectedNode = node;
                _fileContextMenu?.Show(MainTreeView, e.Location);
            }
        }

        private void RaiseFileAction(XstFileAction action)
        {
            if (_contextNode == null) return;
            var folder = _contextNode.Tag as XstFolder;
            if (folder == null) return;
            // Get owning file (root folder node Tag == root folder of file)
            var file = _DataSources.FirstOrDefault(f => f.RootFolder.GetId() == folder.GetId());
            if (file == null) return;
            FileActionRequested?.Invoke(this, new XstFileActionEventArgs(file, action));
        }

        public event EventHandler<XstElementEventArgs>? SelectedItemChanged;
        private void RaiseSelectedItemChanged() => SelectedItemChanged?.Invoke(this, new XstElementEventArgs(GetSelectedItem()));

        // ORIGINAL single file data source (kept for backward compatibility)
        private XstFile? _DataSource;
        public XstFile? GetDataSource() => _DataSource;

        // NEW: multiple file support
        private readonly List<XstFile> _DataSources = new();

        // Backwards compatibility: set single file (clears list and loads that one)
        public void SetDataSource(XstFile? dataSource)
        {
            _DataSource = dataSource;
            _DataSources.Clear();
            if (dataSource != null)
                _DataSources.Add(dataSource);
            LoadFolders();
            if (MainTreeView.Nodes.Count == 0)
                RaiseSelectedItemChanged();
        }

        // NEW API: set multiple files
        public void SetDataSources(IEnumerable<XstFile>? dataSources)
        {
            _DataSource = null; // disable single-file tracking when using multi
            _DataSources.Clear();
            if (dataSources != null)
                _DataSources.AddRange(dataSources.Where(f => f != null));
            LoadFolders();
            if (MainTreeView.Nodes.Count == 0)
                RaiseSelectedItemChanged();
        }

        // Returns current selected folder (folder stored in Tag of folder nodes)
        public XstFolder? GetSelectedItem() => MainTreeView.SelectedNode?.Tag as XstFolder;

        // Helper: get the XstFile owning the selected folder
        public XstFile? GetSelectedFile()
        {
            var folder = GetSelectedItem();
            if (folder == null) return null;
            // root folder comparison
            foreach (var f in _DataSources)
            {
                if (ReferenceEquals(f.RootFolder, folder) || folder.GetId() == f.RootFolder.GetId())
                    return f;
                // If folder is a descendant, we can attempt to walk up via Parent if available
                // (Assuming XstFolder may have Parent property; if not, fallback to search by ID contained in subtree.)
            }
            // Fallback: search by traversing each file's folders for matching id
            foreach (var f in _DataSources)
            {
                if (ContainsFolder(f.RootFolder, folder)) return f;
            }
            return null;
        }

        private bool ContainsFolder(XstFolder root, XstFolder target)
        {
            if (root.GetId() == target.GetId()) return true;
            foreach (var child in root.Folders)
                if (ContainsFolder(child, target)) return true;
            return false;
        }

        public void SetSelectedItem(XstFolder? item)
            => MainTreeView.SelectedNode = (item != null && _DicMapFoldersNodes.ContainsKey(item.GetId())) ?
                                           _DicMapFoldersNodes[item.GetId()]
                                           : null;

        private Dictionary<string, KryptonTreeNode> _DicMapFoldersNodes = new Dictionary<string, KryptonTreeNode>();

        private void LoadFolders()
        {
            MainTreeView.Nodes.Clear();
            _DicMapFoldersNodes = new Dictionary<string, KryptonTreeNode>();
            if (_DataSources.Count == 0) return;

            foreach (var file in _DataSources)
            {
                AddFileRootWithFolders(file);
            }
            // Select first file root if nothing selected
            if (MainTreeView.SelectedNode == null && MainTreeView.Nodes.Count > 0)
                MainTreeView.SelectedNode = MainTreeView.Nodes[0];
        }

        private void AddFileRootWithFolders(XstFile file)
        {
            var rootFolder = file.RootFolder;
            // Root node text: file name (fallback to folder original ToString if file name not accessible)
            string fileName = TryGetFileName(file, rootFolder);
            var fileNode = new KryptonTreeNode(fileName) { Tag = rootFolder };
            MainTreeView.Nodes.Add(fileNode);
            _DicMapFoldersNodes[rootFolder.GetId()] = fileNode;

            // Add child folders of root
            foreach (var childFolder in rootFolder.Folders)
                AddFolderToTree(childFolder, fileNode);
            fileNode.Expand();
        }

        private string TryGetFileName(XstFile file, XstFolder rootFolder)
        {
            // Try common property names via reflection to avoid tight coupling
            try
            {
                var type = file.GetType();
                var prop = type.GetProperty("FileName") ?? type.GetProperty("Path") ?? type.GetProperty("Name");
                if (prop != null)
                {
                    var val = prop.GetValue(file) as string;
                    if (!string.IsNullOrWhiteSpace(val))
                        return Path.GetFileName(val);
                }
            }
            catch { }
            // Fallback: use root folder string
            return rootFolder.ToString() ?? "<file>";
        }

        private KryptonTreeNode AddFolderToTree(XstFolder folder, KryptonTreeNode parentNode)
        {
            string name = $"{folder.ToString() ?? "<no_name>"} ({folder.ContentCount}|{folder.ContentUnreadCount})";
            var node = new KryptonTreeNode(name) { Tag = folder };
            _DicMapFoldersNodes[folder.GetId()] = node;

            if (folder.HasErrorInFolders || folder.HasErrorInMessages)
                node.ForeColor = Color.Red;

            parentNode.Nodes.Add(node);

            foreach (var childFolder in folder.Folders)
                AddFolderToTree(childFolder, node);

            node.Expand();
            return node;
        }

        public void ClearContents()
        {
            GetSelectedItem()?.ClearContents();
            SetDataSource(null); // clears list as well
            _DataSources.Clear();
        }

    }
}
