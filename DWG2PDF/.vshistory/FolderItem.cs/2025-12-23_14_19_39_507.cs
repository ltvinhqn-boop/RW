using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DWG2PDF
{
    public class FolderItem
    {
        public string Name { get; set; }
        public string Icon { get; set; }
        public string FullPath { get; set; }
        public bool IsDirectory => Directory.Exists(FullPath);
    }

    public class LoadFolder
    {

        public void LoadFolderContents(ObservableCollection<FolderItem> targetCollection, string path)
        {
            targetCollection.Clear();

            try
            {
                if (Directory.Exists(path))
                {
                    //if (Directory.GetParent(path) != null)
                    //{
                    //    targetCollection.Add(new FolderItem
                    //    {
                    //        Name = "Back",
                    //        Icon = "◀️",
                    //        FullPath = Directory.GetParent(path).FullName
                    //    });
                    //}
                    foreach (var dir in Directory.GetDirectories(path))
                    {
                        targetCollection.Add(new FolderItem
                        {
                            Name = Path.GetFileName(dir),
                            Icon = "📁",
                            FullPath = dir
                        });
                    }

                    foreach (var file in Directory.GetFiles(path, "*.dwg"))
                    {
                        targetCollection.Add(new FolderItem
                        {
                            Name = Path.GetFileName(file),
                            Icon = "📄",
                            FullPath = file
                        });
                    }
                }
                else if (File.Exists(path) && Path.GetExtension(path).Equals(".dwg", StringComparison.OrdinalIgnoreCase))
                {
                    targetCollection.Add(new FolderItem
                    {
                        Name = Path.GetFileName(path),
                        Icon = "📄",
                        FullPath = path
                    });
                }
            }
            catch (Exception ex)
            {
                // Ghi log hoặc thông báo nhẹ nếu cần
                Console.WriteLine($"Lỗi đọc thư mục: {ex.Message}");
            }
        }





    }
}
