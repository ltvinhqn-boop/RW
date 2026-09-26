using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Windows.Forms;

namespace DWG2PDF
{
    public class FolderItem
    {
        public string Name { get; set; }
        public Image Icon { get; set; }
        public int ImageIndex { get; set; } // Index trong ImageList
        public string FullPath { get; set; }
        public bool IsDirectory => Directory.Exists(FullPath);
    }

    public class LoadFolder
    {
        private ImageList _imageList;
        private string _iconBasePath;
        private const int FOLDER_INDEX = 0;
        private const int FILE_INDEX = 1;

        public LoadFolder(ImageList imageList, string iconBasePath = "")
        {
            _imageList = imageList;
            _iconBasePath = iconBasePath;
            InitializeDefaultIcons();
        }

        private void InitializeDefaultIcons()
        {
            // Khởi tạo 2 icon mặc định: Folder và File
            _imageList.Images.Clear();
            _imageList.ImageSize = new Size(24, 24);
            _imageList.ColorDepth = ColorDepth.Depth32Bit;

            // Icon mặc định cho Folder (vẽ một icon đơn giản)
            Bitmap folderIcon = new Bitmap(24, 24);
            using (Graphics g = Graphics.FromImage(folderIcon))
            {
                g.Clear(Color.Transparent);
                g.FillRectangle(Brushes.Orange, 4, 8, 16, 12);
                g.DrawRectangle(Pens.DarkOrange, 4, 8, 16, 12);
            }
            _imageList.Images.Add("folder", folderIcon);

            // Icon mặc định cho File (vẽ một icon đơn giản)
            Bitmap fileIcon = new Bitmap(24, 24);
            using (Graphics g = Graphics.FromImage(fileIcon))
            {
                g.Clear(Color.Transparent);
                g.FillRectangle(Brushes.White, 6, 4, 12, 16);
                g.DrawRectangle(Pens.Black, 6, 4, 12, 16);
                g.DrawLine(Pens.Black, 8, 8, 16, 8);
                g.DrawLine(Pens.Black, 8, 11, 16, 11);
                g.DrawLine(Pens.Black, 8, 14, 16, 14);
            }
            _imageList.Images.Add("file", fileIcon);
        }

        private int LoadDwgIcon(string fileNameWithoutExtension)
        {
            try
            {
                // Thử tìm file icon với tên trùng với file DWG
                string[] extensions = { ".png", ".jpg", ".bmp", ".ico" };

                foreach (string ext in extensions)
                {
                    string iconPath = Path.Combine(_iconBasePath, fileNameWithoutExtension + ext);

                    if (File.Exists(iconPath))
                    {
                        // Kiểm tra xem icon đã tồn tại trong ImageList chưa
                        if (!_imageList.Images.ContainsKey(fileNameWithoutExtension))
                        {
                            Image icon = Image.FromFile(iconPath);
                            _imageList.Images.Add(fileNameWithoutExtension, icon);
                        }
                        return _imageList.Images.IndexOfKey(fileNameWithoutExtension);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi load icon: {ex.Message}");
            }

            // Nếu không tìm thấy, trả về icon file mặc định
            return FILE_INDEX;
        }

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
                    //        ImageIndex = FOLDER_INDEX,
                    //        FullPath = Directory.GetParent(path).FullName
                    //    });
                    //}

                    // Thêm các thư mục
                    foreach (var dir in Directory.GetDirectories(path))
                    {
                        targetCollection.Add(new FolderItem
                        {
                            Name = Path.GetFileName(dir),
                            ImageIndex = FOLDER_INDEX,
                            FullPath = dir
                        });
                    }

                    // Thêm các file DWG
                    foreach (var file in Directory.GetFiles(path, "*.dwg"))
                    {
                        string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(file);
                        int iconIndex = LoadDwgIcon(fileNameWithoutExtension);

                        targetCollection.Add(new FolderItem
                        {
                            Name = Path.GetFileName(file),
                            ImageIndex = iconIndex,
                            FullPath = file
                        });
                    }
                }
                else if (File.Exists(path) && Path.GetExtension(path).Equals(".dwg", StringComparison.OrdinalIgnoreCase))
                {
                    string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(path);
                    int iconIndex = LoadDwgIcon(fileNameWithoutExtension);

                    targetCollection.Add(new FolderItem
                    {
                        Name = Path.GetFileName(path),
                        ImageIndex = iconIndex,
                        FullPath = path
                    });
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi đọc thư mục: {ex.Message}");
            }
        }

        public void SetIconBasePath(string path)
        {
            _iconBasePath = path;
        }

        public ImageList GetImageList()
        {
            return _imageList;
        }
    }
}