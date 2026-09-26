using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows;

namespace RWTools
{
    internal class FolderTreeBuilder
    {

        // Hàm khởi tạo node gốc và gán sự kiện mở nhánh
        public TreeViewItem BuildTreeFromRoot(string rootPath)
        {
            if (!Directory.Exists(rootPath))
                return null;

            var rootName = Path.GetFileName(rootPath.TrimEnd(Path.DirectorySeparatorChar));
            var rootItem = CreateTreeItem(rootPath, string.IsNullOrEmpty(rootName) ? rootPath : rootName, isFolder: true);
            return rootItem;
        }

        // Tạo TreeViewItem cho thư mục hoặc file
        private TreeViewItem CreateTreeItem(string fullPath, string name, bool isFolder)
        {
            var icon = isFolder ? "📁" : "📄";
            var item = new TreeViewItem
            {
                Header = $"{icon} {name}",
                Tag = fullPath
            };

            if (isFolder)
            {
                // Add placeholder để hỗ trợ lazy loading
                item.Items.Add(null);
                item.Expanded += Folder_Expanded;
            }

            return item;
        }

        // Xử lý khi người dùng mở một nhánh
        private void Folder_Expanded(object sender, RoutedEventArgs e)
        {
            var item = (TreeViewItem)sender;

            if (item.Items.Count == 1 && item.Items[0] == null)
            {
                item.Items.Clear();

                string path = item.Tag as string;

                // Add thư mục con
                try
                {
                    foreach (var dir in Directory.GetDirectories(path))
                    {
                        string name = Path.GetFileName(dir);
                        var subItem = CreateTreeItem(dir, name, isFolder: true);
                        item.Items.Add(subItem);
                    }
                }
                catch { }

                // Add file con
                try
                {
                    foreach (var file in Directory.GetFiles(path, "*.dwg"))
                    {
                        string name = Path.GetFileName(file);
                        var fileItem = CreateTreeItem(file, name, isFolder: false);
                        item.Items.Add(fileItem);
                    }
                }
                catch { }
            }
        }
    
}
}
