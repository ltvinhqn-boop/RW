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
//using System.Windows.Controls;
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
        private ImageList imageList;

        public Form2(Form1 form1)
        {
            InitializeComponent();
            _form1 = form1;
            this.FolderTreeView.AfterSelect += new TreeViewEventHandler(this.FolderTreeView_AfterSelect);
            this.FolderContentsListView.DoubleClick += new EventHandler(this.FolderContentsListView_DoubleClick);
        }

        public ObservableCollection<FolderItem> FolderItems { get; set; } = new ObservableCollection<FolderItem>();
        private FolderTreeBuilder treeBuilder = new FolderTreeBuilder();

        private void Form2_Load(object sender, EventArgs e)
        {
            // Khởi tạo ImageList
            imageList = new ImageList();
            imageList.ImageSize = new Size(24, 24);
            imageList.ColorDepth = ColorDepth.Depth32Bit;

            // Khởi tạo LoadFolder với đường dẫn icon
            string iconPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Resources", "Icons");
            loadFolder = new LoadFolder(imageList, iconPath);

            // Gán ImageList cho ListView
            FolderContentsListView.SmallImageList = imageList;
            FolderContentsListView.LargeImageList = imageList;
            FolderContentsListView.View = View.Details;

            // Xóa các cột cũ nếu có
            FolderContentsListView.Columns.Clear();

            // Thêm cột cho ListView
            FolderContentsListView.Columns.Add("Tên", 300);
            FolderContentsListView.Columns.Add("Đường dẫn", 400);

            // Load TreeView
            FolderTreeView.Nodes.Clear();
            string path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\Libraries";
            TreeNode root = treeBuilder.BuildTreeFromRoot(path);
            if (root != null)
            {
                FolderTreeView.Nodes.Add(root);
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
                // Xóa ObservableCollection cũ
                FolderItems.Clear();

                // Load dữ liệu mới
                loadFolder.LoadFolderContents(FolderItems, path);

                // Xóa items cũ trong ListView
                FolderContentsListView.Items.Clear();

                // Thêm items mới vào ListView
                foreach (var item in FolderItems)
                {
                    ListViewItem lvItem = new ListViewItem(item.Name);
                    lvItem.ImageIndex = item.ImageIndex;
                    lvItem.SubItems.Add(item.FullPath);
                    lvItem.Tag = item; // Lưu object để dùng sau
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
