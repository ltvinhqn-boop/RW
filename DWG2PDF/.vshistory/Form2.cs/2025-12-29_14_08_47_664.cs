using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using iText.StyledXmlParser.Jsoup.Parser;
using RWTools;
using System.IO;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Autodesk.AutoCAD.Interop.Common;
using acadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using Autodesk.AutoCAD.EditorInput;
using System.Reflection;
using Autodesk.AutoCAD.Geometry;

namespace DWG2PDF
{
    public partial class Form2 : Form
    {
        private Form1 _form1;
        private LoadFolder loadFolder;
        private ImageList previewImageList;
        PromptPointResult pointResult;
        IntPtr acadHwnd1 = new IntPtr(acadApp.MainWindow.Handle.ToInt64());

        public Form2(Form1 form1)
        {
            InitializeComponent();
            _form1 = form1;
            this.FolderTreeView.AfterSelect += new TreeViewEventHandler(this.FolderTreeView_AfterSelect);
            this.FolderTreeView.BeforeExpand += new TreeViewCancelEventHandler(this.FolderTreeView_BeforeExpand);
            this.FolderContentsListView.DoubleClick += new EventHandler(this.FolderContentsListView_DoubleClick);
            comboBox1.DropDownStyle = ComboBoxStyle.DropDown;
            comboBox1.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            comboBox1.AutoCompleteSource = AutoCompleteSource.ListItems;

            // Đăng ký sự kiện cho ComboBox
            comboBox1.SelectedIndexChanged += ComboBox1_SelectedIndexChanged;
        }

        public ObservableCollection<FolderItem> FolderItems { get; set; } = new ObservableCollection<FolderItem>();
        private FolderTreeBuilder treeBuilder = new FolderTreeBuilder();

        private void LoadAllDwgToComboBox(string rootFolder)
        {
            if (!Directory.Exists(rootFolder))
                return;

            comboBox1.DisplayMember = "Name";
            comboBox1.ValueMember = "FullPath";
            comboBox1.Items.Clear();

            try
            {
                foreach (var file in Directory.GetFiles(
                    rootFolder,
                    "*.dwg",
                    SearchOption.AllDirectories))
                {
                    comboBox1.Items.Add(new
                    {
                        Name = Path.GetFileNameWithoutExtension(file),
                        FullPath = file
                    });
                }
            }
            catch (UnauthorizedAccessException)
            {
                // bỏ qua thư mục không có quyền
            }
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            // Khởi tạo ImageList cho preview
            previewImageList = new ImageList();
            previewImageList.ImageSize = new Size(96, 96);
            previewImageList.ColorDepth = ColorDepth.Depth32Bit;

            // Khởi tạo LoadFolder với đường dẫn preview
            string previewPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\Libraries";
            loadFolder = new LoadFolder(previewImageList, previewPath);

            // Gán ImageList cho ListView
            FolderContentsListView.LargeImageList = previewImageList;
            FolderContentsListView.SmallImageList = previewImageList;
            FolderContentsListView.View = View.Details;
            FolderContentsListView.OwnerDraw = true;

            // Đăng ký sự kiện vẽ
            FolderContentsListView.DrawColumnHeader += FolderContentsListView_DrawColumnHeader;
            FolderContentsListView.DrawSubItem += FolderContentsListView_DrawSubItem;

            // Chỉ thêm cột nếu chưa có
            if (FolderContentsListView.Columns.Count == 0)
            {
                FolderContentsListView.Columns.Add("Tên", 200);
                FolderContentsListView.Columns.Add("Preview", 150);
                FolderContentsListView.Columns.Add("Đường dẫn", 300);
            }

            // Load TreeView
            FolderTreeView.Nodes.Clear();
            string path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\Libraries";
            LoadAllDwgToComboBox(path);
            TreeNode root = treeBuilder.BuildTreeFromRoot(path);
            if (root != null)
            {
                FolderTreeView.Nodes.Add(root);
            }
        }

        // Sự kiện khi chọn item trong ComboBox
        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedItem != null)
            {
                try
                {
                    dynamic selectedItem = comboBox1.SelectedItem;
                    string filePath = selectedItem.FullPath;

                    FindAndSelectFile(filePath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Lỗi: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void FindAndSelectFile(string filePath)
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    MessageBox.Show($"File không tồn tại: {filePath}", "Thông báo",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string directoryPath = Path.GetDirectoryName(filePath);

                // Tìm và expand node
                TreeNode foundNode = FindAndExpandNodeByPath(FolderTreeView.Nodes, directoryPath);

                if (foundNode != null)
                {
                    // Select node trong TreeView
                    FolderTreeView.SelectedNode = foundNode;
                    foundNode.EnsureVisible();

                    // Load ListView (sẽ được trigger bởi AfterSelect event)
                    // Nhưng để chắc chắn, gọi trực tiếp
                    LoadFolderContents(directoryPath);

                    Application.DoEvents();
                    System.Threading.Thread.Sleep(50);

                    // Select file trong ListView
                    SelectFileInListView(filePath);
                }
                else
                {
                    MessageBox.Show($"Không tìm thấy thư mục: {directoryPath}", "Thông báo",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private TreeNode FindAndExpandNodeByPath(TreeNodeCollection nodes, string path)
        {
            foreach (TreeNode node in nodes)
            {
                if (node.Tag != null && node.Tag.ToString().Equals(path, StringComparison.OrdinalIgnoreCase))
                {
                    // Expand tất cả node cha từ dưới lên
                    TreeNode parent = node.Parent;
                    while (parent != null)
                    {
                        if (!parent.IsExpanded)
                        {
                            parent.Expand();
                            Application.DoEvents();
                        }
                        parent = parent.Parent;
                    }

                    // Expand node này nếu có children
                    if (!node.IsExpanded && (node.Nodes.Count > 0 || Directory.Exists(path)))
                    {
                        node.Expand();
                        Application.DoEvents();
                    }

                    return node;
                }

                // Kiểm tra nếu path nằm trong subtree của node này
                if (node.Tag != null && path.StartsWith(node.Tag.ToString() + "\\", StringComparison.OrdinalIgnoreCase))
                {
                    // Load children nếu cần
                    if (node.Nodes.Count == 0 ||
                        (node.Nodes.Count == 1 && node.Nodes[0].Text == "Loading..."))
                    {
                        treeBuilder.LoadChildren(node);
                        Application.DoEvents();
                    }

                    // Expand node này
                    if (!node.IsExpanded)
                    {
                        node.Expand();
                        Application.DoEvents();
                    }

                    // Tìm tiếp trong children
                    TreeNode foundNode = FindAndExpandNodeByPath(node.Nodes, path);
                    if (foundNode != null)
                    {
                        return foundNode;
                    }
                }
            }
            return null;
        }

        private void SelectFileInListView(string filePath)
        {
            try
            {
                foreach (ListViewItem lvItem in FolderContentsListView.Items)
                {
                    FolderItem item = lvItem.Tag as FolderItem;
                    if (item != null && item.FullPath.Equals(filePath, StringComparison.OrdinalIgnoreCase))
                    {
                        FolderContentsListView.SelectedItems.Clear();
                        lvItem.Selected = true;
                        lvItem.Focused = true;
                        lvItem.EnsureVisible();
                        FolderContentsListView.Focus();
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi SelectFileInListView: {ex.Message}");
            }
        }

        private void FolderContentsListView_DrawColumnHeader(object sender, DrawListViewColumnHeaderEventArgs e)
        {
            e.DrawDefault = true;
        }

        private void FolderContentsListView_DrawSubItem(object sender, DrawListViewSubItemEventArgs e)
        {
            // Vẽ background với màu phù hợp
            if (e.Item.Selected)
            {
                e.Graphics.FillRectangle(SystemBrushes.Highlight, e.Bounds);
            }
            else
            {
                e.Graphics.FillRectangle(SystemBrushes.Window, e.Bounds);
            }

            // Nếu là cột Preview (index 1)
            if (e.ColumnIndex == 1)
            {
                FolderItem item = e.Item.Tag as FolderItem;
                if (item != null && !string.IsNullOrEmpty(item.ImageKey))
                {
                    Image previewImg = previewImageList.Images[item.ImageKey];

                    if (previewImg != null)
                    {
                        int padding = 2;
                        int maxWidth = e.Bounds.Width - (padding * 2);
                        int maxHeight = e.Bounds.Height - (padding * 2);

                        float scale = Math.Min((float)maxWidth / previewImg.Width, (float)maxHeight / previewImg.Height);
                        int newWidth = (int)(previewImg.Width * scale);
                        int newHeight = (int)(previewImg.Height * scale);

                        int x = e.Bounds.X + (e.Bounds.Width - newWidth) / 2;
                        int y = e.Bounds.Y + (e.Bounds.Height - newHeight) / 2;

                        e.Graphics.DrawImage(previewImg, x, y, newWidth, newHeight);
                    }
                }
            }
            else
            {
                if (!string.IsNullOrEmpty(e.SubItem.Text))
                {
                    StringFormat sf = new StringFormat();
                    sf.Alignment = StringAlignment.Near;
                    sf.LineAlignment = StringAlignment.Center;
                    sf.Trimming = StringTrimming.EllipsisCharacter;
                    sf.FormatFlags = StringFormatFlags.NoWrap;

                    int leftPadding = 5;
                    Rectangle textBounds = new Rectangle(
                        e.Bounds.X + leftPadding,
                        e.Bounds.Y,
                        e.Bounds.Width - leftPadding,
                        e.Bounds.Height
                    );

                    Brush textBrush = e.Item.Selected ? SystemBrushes.HighlightText : SystemBrushes.WindowText;

                    e.Graphics.DrawString(
                        e.SubItem.Text,
                        e.SubItem.Font ?? FolderContentsListView.Font,
                        textBrush,
                        textBounds,
                        sf
                    );
                }
            }

            using (Pen gridPen = new Pen(SystemColors.ControlLight))
            {
                e.Graphics.DrawRectangle(gridPen, e.Bounds);
            }

            if (e.Item.Selected && e.Item.ListView.Focused)
            {
                ControlPaint.DrawFocusRectangle(e.Graphics, e.Bounds);
            }
        }

        private void FolderTreeView_BeforeExpand(object sender, TreeViewCancelEventArgs e)
        {
            treeBuilder.LoadChildren(e.Node);
        }

        private void FolderTreeView_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (e.Node != null && e.Node.Tag != null)
            {
                string selectedPath = e.Node.Tag.ToString();
                LoadFolderContents(selectedPath);
            }
        }

        private void LoadFolderContents(string path)
        {
            try
            {
                FolderItems.Clear();
                loadFolder.LoadFolderContents(FolderItems, path);
                FolderContentsListView.Items.Clear();

                FolderContentsListView.OwnerDraw = true;

                foreach (var item in FolderItems)
                {
                    ListViewItem lvItem = new ListViewItem(item.Name);
                    lvItem.ImageKey = item.ImageKey;
                    lvItem.SubItems.Add("");
                    lvItem.SubItems.Add(item.FullPath);
                    lvItem.Tag = item;
                    FolderContentsListView.Items.Add(lvItem);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi load thư mục: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void FolderContentsListView_DoubleClick(object sender, EventArgs e)
        {
            if (FolderContentsListView.SelectedItems.Count > 0)
            {
                FolderItem item = FolderContentsListView.SelectedItems[0].Tag as FolderItem;
                if (item != null && item.IsDirectory)
                {
                    LoadFolderContents(item.FullPath);
                    SelectTreeNodeByPath(item.FullPath);
                }
                else if (item != null && !item.IsDirectory)
                {
                    string filePath = item.FullPath;
                    string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(filePath);

                    NativeMethods.SetForegroundWindow(acadHwnd1);
                    var doc = acadApp.DocumentManager.MdiActiveDocument;
                    var db = doc.Database;
                    var ed = doc.Editor;

                    pointResult = ed.GetPoint("Chọn vị trí để chèn " + fileNameWithoutExtension);

                    if (pointResult.Status == PromptStatus.OK)
                    {
                        Insertdynamicblock.InsertBlock2(filePath, fileNameWithoutExtension,
                            new Point3d(pointResult.Value.X, pointResult.Value.Y, 0), null, 1.0, 0);
                    }
                }
            }
        }

        private void SelectTreeNodeByPath(string path)
        {
            TreeNode foundNode = FindNodeByPath(FolderTreeView.Nodes, path);
            if (foundNode != null)
            {
                FolderTreeView.SelectedNode = foundNode;
                foundNode.Expand();
            }
        }

        private TreeNode FindNodeByPath(TreeNodeCollection nodes, string path)
        {
            foreach (TreeNode node in nodes)
            {
                if (node.Tag != null && node.Tag.ToString().Equals(path, StringComparison.OrdinalIgnoreCase))
                {
                    return node;
                }

                TreeNode foundNode = FindNodeByPath(node.Nodes, path);
                if (foundNode != null)
                {
                    return foundNode;
                }
            }
            return null;
        }
    }
}