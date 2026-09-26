using System.IO;
using System.Windows.Forms;

namespace DWG2PDF
{
    internal class FolderTreeBuilder
    {
        // Build node gốc
        public TreeNode BuildTreeFromRoot(string rootPath)
        {
            if (!Directory.Exists(rootPath))
                return null;

            string rootName = Path.GetFileName(rootPath.TrimEnd(Path.DirectorySeparatorChar));
            if (string.IsNullOrEmpty(rootName))
                rootName = rootPath;

            TreeNode rootNode = CreateNode(rootPath, rootName, isFolder: true);
            return rootNode;
        }

        // Tạo TreeNode
        private TreeNode CreateNode(string fullPath, string name, bool isFolder)
        {
            string icon = isFolder ? "📁" : "📄";

            TreeNode node = new TreeNode($"{icon} {name}");
            node.Tag = fullPath;

            if (isFolder)
            {
                // Placeholder để lazy load
                node.Nodes.Add("DUMMY");
            }

            return node;
        }

        // Gọi trong sự kiện TreeView.BeforeExpand
        public void LoadChildren(TreeNode node)
        {
            // Nếu đã load rồi thì thôi
            if (node.Nodes.Count != 1 || node.Nodes[0].Text != "DUMMY")
                return;

            node.Nodes.Clear();

            string path = node.Tag as string;
            if (!Directory.Exists(path))
                return;

            // Thư mục con
            try
            {
                foreach (var dir in Directory.GetDirectories(path))
                {
                    string name = Path.GetFileName(dir);
                    node.Nodes.Add(CreateNode(dir, name, true));
                }
            }
            catch { }

            // File DWG
            try
            {
                foreach (var file in Directory.GetFiles(path, "*.dwg"))
                {
                    string name = Path.GetFileNameWithoutExtension(file);
                    node.Nodes.Add(CreateNode(file, name, false));
                }
            }
            catch { }
        }
    }
}
