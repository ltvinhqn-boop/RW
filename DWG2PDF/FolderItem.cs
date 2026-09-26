using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Drawing;
using System.Windows.Forms;

namespace DWG2PDF
{
    public class FolderItem
    {
        public string Name { get; set; }
        public string ImageKey { get; set; } // Key cho preview image trong ImageList
        public string FullPath { get; set; }
        public bool IsDirectory => Directory.Exists(FullPath);
    }

    public class LoadFolder
    {
        private ImageList _previewImageList; // ImageList cho preview
        private string _previewBasePath;
        private Dictionary<string, string> _previewImagePaths;
        private const string FOLDER_KEY = "folder";
        private const string FILE_KEY = "file";

        public LoadFolder(ImageList previewImageList, string previewBasePath = "")
        {
            _previewImageList = previewImageList;
            _previewBasePath = previewBasePath;
            _previewImagePaths = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            InitializeDefaultImages();

            if (!string.IsNullOrEmpty(_previewBasePath))
            {
                LoadAllPreviewImages();
            }
        }

        private void InitializeDefaultImages()
        {
            _previewImageList.Images.Clear();
            _previewImageList.ImageSize = new Size(96, 96);
            _previewImageList.ColorDepth = ColorDepth.Depth32Bit;

            // Hình mặc định cho Folder
            Bitmap folderIcon = new Bitmap(96, 96);
            using (Graphics g = Graphics.FromImage(folderIcon))
            {
                g.Clear(Color.Transparent);
                g.FillRectangle(Brushes.Orange, 16, 32, 64, 48);
                g.DrawRectangle(Pens.DarkOrange, 16, 32, 64, 48);
            }
            _previewImageList.Images.Add(FOLDER_KEY, folderIcon);

            // Hình mặc định cho File
            Bitmap fileIcon = new Bitmap(96, 96);
            using (Graphics g = Graphics.FromImage(fileIcon))
            {
                g.Clear(Color.Transparent);
                g.FillRectangle(Brushes.White, 24, 16, 48, 64);
                g.DrawRectangle(Pens.Black, 24, 16, 48, 64);
                g.DrawLine(Pens.Black, 32, 32, 64, 32);
                g.DrawLine(Pens.Black, 32, 44, 64, 44);
                g.DrawLine(Pens.Black, 32, 56, 64, 56);
            }
            _previewImageList.Images.Add(FILE_KEY, fileIcon);
        }

        /// <summary>
        /// Quét tất cả các file ảnh trong thư mục preview bao gồm cả thư mục con
        /// </summary>
        private void LoadAllPreviewImages()
        {
            try
            {
                if (!Directory.Exists(_previewBasePath))
                    return;

                string[] imageExtensions = { "*.png", "*.jpg", "*.jpeg", "*.bmp", "*.gif" };

                foreach (string extension in imageExtensions)
                {
                    var imageFiles = Directory.GetFiles(_previewBasePath, extension, SearchOption.AllDirectories);

                    foreach (string imagePath in imageFiles)
                    {
                        string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(imagePath);

                        if (!_previewImagePaths.ContainsKey(fileNameWithoutExtension))
                        {
                            _previewImagePaths[fileNameWithoutExtension] = imagePath;

                            try
                            {
                                Image img = Image.FromFile(imagePath);
                                _previewImageList.Images.Add(fileNameWithoutExtension, img);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Lỗi load preview image {imagePath}: {ex.Message}");
                            }
                        }
                    }
                }

                Console.WriteLine($"Đã load {_previewImagePaths.Count} preview images");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi quét thư mục preview: {ex.Message}");
            }
        }

        /// <summary>
        /// Lấy key của preview image từ ImageList
        /// </summary>
        private string GetPreviewImageKey(string fileNameWithoutExtension)
        {
            if (_previewImageList.Images.ContainsKey(fileNameWithoutExtension))
            {
                return fileNameWithoutExtension;
            }
            return FILE_KEY; // Trả về key mặc định nếu không tìm thấy
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
                            ImageKey = FOLDER_KEY,
                            FullPath = dir
                        });
                    }

                    // Thêm các file DWG
                    foreach (var file in Directory.GetFiles(path, "*.dwg"))
                    {
                        string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(file);
                        string previewKey = GetPreviewImageKey(fileNameWithoutExtension);

                        targetCollection.Add(new FolderItem
                        {
                            Name = fileNameWithoutExtension,
                            ImageKey = previewKey,
                            FullPath = file
                        });
                    }
                }
                else if (File.Exists(path) && Path.GetExtension(path).Equals(".dwg", StringComparison.OrdinalIgnoreCase))
                {
                    string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(path);
                    string previewKey = GetPreviewImageKey(fileNameWithoutExtension);

                    targetCollection.Add(new FolderItem
                    {
                        Name = fileNameWithoutExtension,
                        ImageKey = previewKey,
                        FullPath = path
                    });
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi đọc thư mục: {ex.Message}");
            }
        }

        public void SetPreviewBasePath(string path)
        {
            _previewBasePath = path;
            _previewImagePaths.Clear();
            _previewImageList.Images.Clear();
            InitializeDefaultImages();

            if (!string.IsNullOrEmpty(path))
            {
                LoadAllPreviewImages();
            }
        }

        public ImageList GetPreviewImageList()
        {
            return _previewImageList;
        }

        public void RefreshPreviewImages()
        {
            _previewImagePaths.Clear();
            _previewImageList.Images.Clear();
            InitializeDefaultImages();

            if (!string.IsNullOrEmpty(_previewBasePath))
            {
                LoadAllPreviewImages();
            }
        }
    }
}