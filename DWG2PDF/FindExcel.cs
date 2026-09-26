using System;
using Microsoft.Office.Interop.Excel;

public class ExcelDataProcessor
{
    public void ProcessData(Application excelApp)
    {
        try
        {
            // Khai báo biến
            object val1, val2;
            Worksheet source, inVal;
            int xCol, yRow;
            int i, j, pitch1, pitch2;
            object aVal = null, bVal = null, cVal = null, dVal;

            // Đặt pitch cho giá trị trên hàng 5
            pitch1 = 500;
            pitch2 = 300;

            // Đặt biến cho sheet chứa dữ liệu
            source = excelApp.Worksheets["M4Žd—l•\\"];
            inVal = excelApp.Worksheets["Tra TT"];

            // Lấy giá trị 1 và giá trị 2 từ ô C1 và C2 của sheet "Tra TT"
            val1 = inVal.Range["C1"].Value2;
            val2 = inVal.Range["C2"].Value2;

            // Chuyển đổi về số để tính toán
            double val1Double = Convert.ToDouble(val1);
            double val2Double = Convert.ToDouble(val2);

            // Làm tròn giá trị 1,2 lên pitch gần nhất (ví dụ: 5230 -> 5500)
            val1Double = Math.Ceiling(val1Double / pitch1) * pitch1;
            val2Double = Math.Ceiling(val2Double / pitch2) * pitch2;

            // Tìm vị trí cột chứa giá trị1 tại hàng 5
            xCol = 0;
            for (i = 4; i <= 200; i++)
            {
                var cellValue = source.Cells[5, i].Value2;
                if (cellValue != null && Convert.ToDouble(cellValue) == val1Double)
                {
                    xCol = i;
                    break;
                }
            }

            // Tìm vị trí dòng chứa giá trị2 tại cột B (cột 2)
            yRow = 0;
            for (j = 26; j <= 200; j++)
            {
                var cellValue = source.Cells[j, 2].Value2;
                if (cellValue != null && Convert.ToDouble(cellValue) == val2Double)
                {
                    yRow = j;
                    break;
                }
            }

            // Xóa nội dung các ô từ C3 đến C11
            inVal.Range["C3:C11"].ClearContents();

            // Kiểm tra nếu tìm thấy cả hàng và cột thì lấy giá trị tại điểm giao
            if (xCol > 0 && yRow > 0)
            {
                dVal = source.Cells[yRow, xCol].Value2;
                inVal.Range["C7"].Value2 = dVal;
            }
            else
            {
                inVal.Range["C7"].Value2 = "Invalid";
            }

            // Xử lý chuỗi từ ô phía trên (YRow - 1)
            string tempStr = source.Cells[yRow - 1, xCol].Value2?.ToString() ?? "";

            if (tempStr.Contains("*") && tempStr.Contains(","))
            {
                // Tách a*b,c thành 3 phần tử AVal, BVal, CVal
                string[] parts1 = tempStr.Split('*');
                if (parts1.Length >= 2)
                {
                    cVal = parts1[0].Trim();

                    string[] parts2 = parts1[1].Split(',');
                    if (parts2.Length >= 2)
                    {
                        aVal = parts2[0].Trim();
                        bVal = parts2[1].Trim();
                    }
                }
            }

            // Gán giá trị với nhân 10
            if (aVal != null && double.TryParse(aVal.ToString(), out double aValDouble))
                inVal.Range["C4"].Value2 = aValDouble * 10;

            if (bVal != null && double.TryParse(bVal.ToString(), out double bValDouble))
                inVal.Range["C5"].Value2 = bValDouble * 10;

            if (cVal != null && double.TryParse(cVal.ToString(), out double cValDouble))
                inVal.Range["C6"].Value2 = cValDouble * 10;

            // Xử lý chuỗi từ ô phía dưới (YRow + 1)
            string tempStr2 = source.Cells[yRow + 1, xCol].Value2?.ToString() ?? "";

            if (!string.IsNullOrEmpty(tempStr2))
            {
                string[] parts = tempStr2.Split(' ');
                if (parts.Length >= 1)
                    inVal.Range["C8"].Value2 = parts[0].Trim();

                if (parts.Length >= 2)
                    inVal.Range["C9"].Value2 = parts[1].Trim();
            }

            // Xử lý chuỗi từ cột 3 của hàng YRow
            string tempStr3 = source.Cells[yRow, 3].Value2?.ToString() ?? "";

            if (!string.IsNullOrEmpty(tempStr3) && tempStr3.Contains("-"))
            {
                string[] parts = tempStr3.Split('-');
                if (parts.Length >= 2)
                    inVal.Range["C11"].Value2 = parts[1].Trim();
            }
        }
        catch (Exception ex)
        {
            throw new Exception($"Lỗi khi xử lý dữ liệu Excel: {ex.Message}");
        }
    }

    // Phương thức helper để gọi hàm với file Excel cụ thể
    public void ProcessDataFromFile(string filePath)
    {
        Application excelApp = null;
        Workbook workbook = null;

        try
        {
            excelApp = new Application();
            excelApp.Visible = false;
            excelApp.DisplayAlerts = false;

            workbook = excelApp.Workbooks.Open(filePath);

            ProcessData(excelApp);

            workbook.Save();
        }
        catch (Exception ex)
        {
            throw new Exception($"Lỗi khi mở file Excel: {ex.Message}");
        }
        finally
        {
            // Giải phóng tài nguyên
            if (workbook != null)
            {
                workbook.Close();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            }

            if (excelApp != null)
            {
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            }
        }
    }
}

// Cách sử dụng:
// var processor = new ExcelDataProcessor();
// processor.ProcessDataFromFile(@"C:\path\to\your\excel\file.xlsx");