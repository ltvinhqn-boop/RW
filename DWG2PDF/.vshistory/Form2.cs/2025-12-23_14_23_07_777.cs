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
using System.Windows.Controls;
using System.Windows.Forms;
using iText.StyledXmlParser.Jsoup.Parser;
using RWTools;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace DWG2PDF
{
    public partial class Form2 : Form
    {
        private Form1 _form1;

        public Form2(Form1 form1)
        {
            InitializeComponent();
            _form1 = form1;
        }
        public ObservableCollection<FolderItem> FolderItems { get; set; } = new ObservableCollection<FolderItem>();
        LoadFolder loaditems = new LoadFolder();
        //string path = @"C:\Users\User-PC10\Desktop\_Lib_WKV";
        private FolderTreeBuilder treeBuilder = new FolderTreeBuilder();
        private void Form2_Load(object sender, EventArgs e)
        {
            FolderTreeView.Nodes.Clear();

            string path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\Libraries"; // thư mục gốc

            TreeNode root = treeBuilder.BuildTreeFromRoot(path);
            if (root != null)
            {
                FolderTreeView.Nodes.Add(root);
            }
            //----------------------------------------

            //----------------------------------------
        }
    }
}
