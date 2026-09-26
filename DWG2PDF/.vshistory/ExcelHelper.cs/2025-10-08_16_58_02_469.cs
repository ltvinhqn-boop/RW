using System;
using System.Collections.Generic;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;

namespace RWTools
{
    

    public class ExcelHelper
    {
        // Đọc tất cả checkbox Form Control và trả về Dictionary
        public class CheckBoxInfo
        {
            public string Name { get; set; }      // Tên shape (CheckBox 1 ...)
            public string Caption { get; set; }   // Text hiển thị cạnh checkbox
            public bool IsChecked { get; set; }   // Trạng thái
        }

        public static List<CheckBoxInfo> GetFormCheckBoxes(string filePath, int sheetIndex = 1)
        {
            var result = new List<CheckBoxInfo>();

            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            try
            {
                excelApp = new Excel.Application();
                workbook = excelApp.Workbooks.Open(filePath);
                Excel._Worksheet sheet = workbook.Sheets[sheetIndex];

                foreach (Excel.Shape shape in sheet.Shapes)
                {
                    if (shape.Type == MsoShapeType.msoFormControl &&
                        shape.FormControlType == Excel.XlFormControl.xlCheckBox)
                    {
                        int val = shape.ControlFormat.Value;
                        bool isChecked = (val == 1);

                        // Lấy text hiển thị cạnh checkbox
                        string caption = shape.TextFrame.Characters().Text;

                        result.Add(new CheckBoxInfo
                        {
                            Name = shape.Name,
                            Caption = caption,
                            IsChecked = isChecked
                        });
                    }
                }

                return result;
            }
            finally
            {
                if (workbook != null)
                {
                    workbook.Close(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                }
            }
        }

        public static void ExportCheckBoxesToTxt(string excelPath, string txtPath, int sheetIndex = 1)
        {
            var checkBoxes = GetFormCheckBoxes(excelPath, sheetIndex);

            using (StreamWriter sw = new StreamWriter(txtPath, false))
            {
                foreach (var cb in checkBoxes)
                {
                    sw.WriteLine($"{cb.Name} | {cb.Caption} : {(cb.IsChecked ? "Checked" : "Unchecked")}");
                }
            }
        }
    }

}
