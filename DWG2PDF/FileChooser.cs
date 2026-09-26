using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace DWG2PDF
{
    public class FileChooser
    {
        public List<string> OpenAndSelectFiles(string fileExtension)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Multiselect = true;
            fileDialog.Filter = $"Files (*.{fileExtension})|*.{fileExtension}";

            DialogResult result = fileDialog.ShowDialog();

            if (result == DialogResult.OK)
            {
                return new List<string>(fileDialog.FileNames);
            }
            else
            {
                return new List<string>(); // Trả về danh sách rỗng nếu người dùng không chọn file
            }
        }
    }
}
