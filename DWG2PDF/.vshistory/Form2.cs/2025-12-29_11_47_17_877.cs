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

namespace DWG2PDF
{
    public partial class Form2 : Form
    {
        private Form1 _form1;
        private LoadFolder loadFolder;
        private ImageList previewImageList;

        public Form2(Form1 form1)
        {
            InitializeComponent();
            _form1 = form1;
            this.FolderTreeView.AfterSelect += new TreeViewEventHandler(this.FolderTreeView_AfterSelect);
            this.FolderContentsListView.DoubleClick += new EventHandler(this.FolderContentsListView_DoubleClick);
            comboBox1.DropDownStyle = ComboBoxStyle.DropDown;
            comboBox1.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            comboBox1.AutoCompleteSource = AutoCompleteSource.ListItems;
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
            previewImageList.ImageSize = new Size(96, 96); // Kích thước preview
            previewImageList.ColorDepth = ColorDepth.Depth32Bit;

            // Khởi tạo LoadFolder với đường dẫn preview
            string previewPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\Libraries";
            loadFolder = new LoadFolder(previewImageList, previewPath);

            // Gán ImageList cho ListView
            FolderContentsListView.LargeImageList = previewImageList;
            FolderContentsListView.SmallImageList = previewImageList;
            FolderContentsListView.View = View.Details;
            FolderContentsListView.OwnerDraw = true; // Bật chế độ vẽ tùy chỉnh

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

        private void FolderContentsListView_DrawColumnHeader(object sender, DrawListViewColumnHeaderEventArgs e)
        {
            e.DrawDefault = true;
        }

        private void FolderContentsListView_DrawSubItem(object sender, DrawListViewSubItemEventArgs e)
        {
            // Vẽ background
            e.DrawBackground();

            // Nếu là cột Preview (index 1)
            if (e.ColumnIndex == 1)
            {
                FolderItem item = e.Item.Tag as FolderItem;
                if (item != null && !string.IsNullOrEmpty(item.ImageKey))
                {
                    // Lấy preview image từ ImageList
                    Image previewImg = previewImageList.Images[item.ImageKey];

                    if (previewImg != null)
                    {
                        // Tính toán kích thước để fit vào cell
                        int padding = 2;
                        int maxWidth = e.Bounds.Width - (padding * 2);
                        int maxHeight = e.Bounds.Height - (padding * 2);

                        float scale = Math.Min((float)maxWidth / previewImg.Width, (float)maxHeight / previewImg.Height);
                        int newWidth = (int)(previewImg.Width * scale);
                        int newHeight = (int)(previewImg.Height * scale);

                        // Căn giữa hình ảnh
                        int x = e.Bounds.X + (e.Bounds.Width - newWidth) / 2;
                        int y = e.Bounds.Y + (e.Bounds.Height - newHeight) / 2;

                        e.Graphics.DrawImage(previewImg, x, y, newWidth, newHeight);
                    }
                }
            }
            else
            {
                // Vẽ text căn trái giữa cho các cột text
                if (!string.IsNullOrEmpty(e.SubItem.Text))
                {
                    // Format text căn trái và căn giữa theo chiều dọc
                    StringFormat sf = new StringFormat();
                    sf.Alignment = StringAlignment.Near;          // Căn trái
                    sf.LineAlignment = StringAlignment.Center;     // Căn giữa dọc
                    sf.Trimming = StringTrimming.EllipsisCharacter; // Cắt chữ nếu quá dài
                    sf.FormatFlags = StringFormatFlags.NoWrap;     // Không xuống dòng

                    // Thêm padding bên trái
                    int leftPadding = 5;
                    Rectangle textBounds = new Rectangle(
                        e.Bounds.X + leftPadding,
                        e.Bounds.Y,
                        e.Bounds.Width - leftPadding,
                        e.Bounds.Height
                    );

                    // Sử dụng màu text mặc định
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

            // Vẽ focus rectangle nếu item được chọn
            if (e.Item.Selected)
            {
                e.DrawFocusRectangle(e.Bounds);
            }
        }

        private void FolderTreeView_BeforeExpand(object sender, TreeViewCancelEventArgs e)
        {
            treeBuilder.LoadChildren(e.Node);
        }

        // Sự kiện khi click vào TreeView
        private void FolderTreeView_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (e.Node != null && e.Node.Tag != null)
            {
                string selectedPath = e.Node.Tag.ToString();
                LoadFolderContents(selectedPath);
            }
        }

        // Load nội dung thư mục vào ListView
        private void LoadFolderContents(string path)
        {
            try
            {
                FolderItems.Clear();
                loadFolder.LoadFolderContents(FolderItems, path);
                FolderContentsListView.Items.Clear();

                // Tăng chiều cao của item để hiển thị preview tốt hơn
                FolderContentsListView.OwnerDraw = true;

                foreach (var item in FolderItems)
                {
                    ListViewItem lvItem = new ListViewItem(item.Name);
                    lvItem.ImageKey = item.ImageKey; // Sử dụng ImageKey
                    lvItem.SubItems.Add(""); // Cột preview (sẽ vẽ hình ảnh qua DrawSubItem)
                    lvItem.SubItems.Add(item.FullPath);
                    lvItem.Tag = item; // Lưu FolderItem để dùng khi vẽ
                    FolderContentsListView.Items.Add(lvItem);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi load thư mục: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Sự kiện double click vào ListView (nếu muốn mở thư mục con)
        private void FolderContentsListView_DoubleClick(object sender, EventArgs e)
        {
            if (FolderContentsListView.SelectedItems.Count > 0)
            {
                FolderItem item = FolderContentsListView.SelectedItems[0].Tag as FolderItem;
                if (item != null && item.IsDirectory)
                {
                    // Nếu là thư mục, load nội dung thư mục đó
                    LoadFolderContents(item.FullPath);

                    // Tìm và chọn node tương ứng trong TreeView (optional)
                    SelectTreeNodeByPath(item.FullPath);
                }
                else if (item != null && !item.IsDirectory)
                {
                    // Nếu là file DWG, có thể xử lý thêm (ví dụ: mở file, chọn file...)
                    MessageBox.Show($"File đã chọn: {item.Name}", "Thông tin", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        // Tìm và chọn node trong TreeView theo đường dẫn (optional)
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