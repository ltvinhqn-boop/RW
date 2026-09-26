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
        public int ImageIndex { get; set; }
        public string FullPath { get; set; }
        public Image PreviewImage { get; set; } // Hình ảnh preview
        public bool IsDirectory => Directory.Exists(FullPath);
    }

    public class LoadFolder
    {
        private ImageList _imageList;
        private string _iconBasePath;
        private string _previewBasePath; // Đường dẫn thư mục chứa preview images
        private const int FOLDER_INDEX = 0;
        private const int FILE_INDEX = 1;

        public LoadFolder(ImageList imageList, string iconBasePath = "", string previewBasePath = "")
        {
            _imageList = imageList;
            _iconBasePath = iconBasePath;
            _previewBasePath = previewBasePath;
            InitializeDefaultIcons();
        }

        private void InitializeDefaultIcons()
        {
            _imageList.Images.Clear();
            _imageList.ImageSize = new Size(24, 24);
            _imageList.ColorDepth = ColorDepth.Depth32Bit;

            // Icon mặc định cho Folder
            Bitmap folderIcon = new Bitmap(24, 24);
            using (Graphics g = Graphics.FromImage(folderIcon))
            {
                g.Clear(Color.Transparent);
                g.FillRectangle(Brushes.Orange, 4, 8, 16, 12);
                g.DrawRectangle(Pens.DarkOrange, 4, 8, 16, 12);
            }
            _imageList.Images.Add("folder", folderIcon);

            // Icon mặc định cho File
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
                string[] extensions = { ".png", ".jpg", ".bmp", ".ico" };

                foreach (string ext in extensions)
                {
                    string iconPath = Path.Combine(_iconBasePath, fileNameWithoutExtension + ext);

                    if (File.Exists(iconPath))
                    {
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

            return FILE_INDEX;
        }

        private Image LoadPreviewImage(string fileNameWithoutExtension)
        {
            try
            {
                if (string.IsNullOrEmpty(_previewBasePath))
                    return null;

                string[] extensions = { ".png", ".jpg", ".jpeg", ".bmp", ".gif" };

                foreach (string ext in extensions)
                {
                    string previewPath = Path.Combine(_previewBasePath, fileNameWithoutExtension + ext);

                    if (File.Exists(previewPath))
                    {
                        // Load image và resize để hiển thị
                        Image originalImage = Image.FromFile(previewPath);
                        return originalImage;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi load preview image: {ex.Message}");
            }

            return null;
        }

        public void LoadFolderContents(ObservableCollection<FolderItem> targetCollection, string path)
        {
            targetCollection.Clear();
            try
            {
                if (Directory.Exists(path))
                {
                    // Thêm các thư mục
                    foreach (var dir in Directory.GetDirectories(path))
                    {
                        targetCollection.Add(new FolderItem
                        {
                            Name = Path.GetFileName(dir),
                            ImageIndex = FOLDER_INDEX,
                            FullPath = dir,
                            PreviewImage = null
                        });
                    }

                    // Thêm các file DWG
                    foreach (var file in Directory.GetFiles(path, "*.dwg"))
                    {
                        string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(file);
                        int iconIndex = LoadDwgIcon(fileNameWithoutExtension);
                        Image previewImage = LoadPreviewImage(fileNameWithoutExtension);

                        targetCollection.Add(new FolderItem
                        {
                            Name = Path.GetFileName(file),
                            ImageIndex = iconIndex,
                            FullPath = file,
                            PreviewImage = previewImage
                        });
                    }
                }
                else if (File.Exists(path) && Path.GetExtension(path).Equals(".dwg", StringComparison.OrdinalIgnoreCase))
                {
                    string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(path);
                    int iconIndex = LoadDwgIcon(fileNameWithoutExtension);
                    Image previewImage = LoadPreviewImage(fileNameWithoutExtension);

                    targetCollection.Add(new FolderItem
                    {
                        Name = Path.GetFileName(path),
                        ImageIndex = iconIndex,
                        FullPath = path,
                        PreviewImage = previewImage
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

        public void SetPreviewBasePath(string path)
        {
            _previewBasePath = path;
        }

        public ImageList GetImageList()
        {
            return _imageList;
        }
    }
}