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
        public static Dictionary<string, bool> GetFormCheckBoxes(string filePath, int sheetIndex = 1)
        {
            var result = new Dictionary<string, bool>();

            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            try
            {
                excelApp = new Excel.Application();
                workbook = excelApp.Workbooks.Open(filePath);
                Excel._Worksheet sheet = workbook.Sheets[sheetIndex];

                foreach (Excel.Shape shape in sheet.Shapes)
                {
                    if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoFormControl &&
                        shape.FormControlType == Excel.XlFormControl.xlCheckBox)
                    {
                        int val = shape.ControlFormat.Value;
                        bool isChecked = (val == 1);
                        result[shape.Name] = isChecked;
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

        // Xuất kết quả ra file txt
        public static void ExportCheckBoxesToTxt(string excelPath, string txtPath, int sheetIndex = 1)
        {
            var checkBoxes = GetFormCheckBoxes(excelPath, sheetIndex);

            using (StreamWriter sw = new StreamWriter(txtPath, false))
            {
                foreach (var cb in checkBoxes)
                {
                    sw.WriteLine($"{cb.Key} : {(cb.Value ? "Checked" : "Unchecked")}");
                }
            }
        }
    }

}
