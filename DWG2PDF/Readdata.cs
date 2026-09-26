using System;
using System.Linq;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

public class ExcelLookupResult
{
    public double AVal { get; set; }
    public double BVal { get; set; }
    public double CVal { get; set; }
    public object DVal { get; set; }
    public string Value1 { get; set; }
    public string Value2 { get; set; }
    public string Spacing { get; set; }
    public string AdditionalInfo { get; set; }
    public bool IsValid { get; set; }
    public string ErrorMessage { get; set; }
}

public class ExcelLookupService
{
    private Excel.Application _excelApp;
    private Excel.Workbook _workbook;

    public ExcelLookupService(string filePath)
    {
        _excelApp = new Excel.Application();
        _workbook = _excelApp.Workbooks.Open(filePath);
    }

    /// <summary>
    /// Thực hiện tra cứu dữ liệu tương tự hàm VBA Val()
    /// </summary>
    /// <param name="val1">Giá trị đầu vào 1</param>
    /// <param name="val2">Giá trị đầu vào 2</param>
    /// <param name="sourceSheetName">Tên sheet chứa dữ liệu tra cứu (mặc định: "M4Žd—l•\")</param>
    /// <param name="inputSheetName">Tên sheet nhập liệu (mặc định: "Tra TT")</param>
    /// <param name="pitch1">Bước nhảy cho giá trị 1 (mặc định: 500)</param>
    /// <param name="pitch2">Bước nhảy cho giá trị 2 (mặc định: 300)</param>
    /// <returns>Kết quả tra cứu</returns>
    public ExcelLookupResult PerformLookup(
        double val1,
        double val2,
        string sourceSheetName = "M4-1",
        string inputSheetName = "Tra TT",
        int pitch1 = 500,
        int pitch2 = 300)
    {
        var result = new ExcelLookupResult();
       // MessageBox.Show("ddã vào đây1");
        try
        {
            // Lấy worksheet
            Excel.Worksheet sourceSheet = _workbook.Sheets[sourceSheetName];
            Excel.Worksheet inputSheet = _workbook.Sheets[inputSheetName];

            // Làm tròn giá trị lên theo pitch
            double roundedVal1 = Math.Ceiling(val1 / pitch1) * pitch1;
            double roundedVal2 = Math.Ceiling(val2 / pitch2) * pitch2;

            // Tìm vị trí cột chứa val1 trong hàng 5 (từ cột 4 đến 200)
            int xCol = 0;
            for (int i = 5; i <= 200; i++)
            {
                var cellValue = sourceSheet.Cells[5, i].Value;
                if (cellValue != null && Convert.ToDouble(cellValue) == roundedVal1)
                {
                    xCol = i;
                    break;
                }
            }

            // Tìm vị trí hàng chứa val2 trong cột B (từ hàng 26 đến 200)
            int yRow = 0;
            for (int j = 27; j <= 200; j=j+4)
            {
                var cellValue = sourceSheet.Cells[j, 2].Value;
                if (cellValue != null && Convert.ToDouble(cellValue) == roundedVal2)
                {
                    yRow = j;
                    break;
                }
            }

            // Kiểm tra nếu tìm thấy cả hàng và cột
            if (xCol > 0 && yRow > 0)
            {
                result.IsValid = true;

                // Lấy giá trị chính tại giao điểm
                result.DVal = sourceSheet.Cells[yRow, xCol].Value;

                // Xử lý chuỗi từ hàng trên (YRow - 1)
                string tempStr = sourceSheet.Cells[yRow - 1, xCol].Value?.ToString();
                if (!string.IsNullOrEmpty(tempStr) && tempStr.Contains("*") && tempStr.Contains(","))
                {
                    var parts = tempStr.Split('*');
                    if (parts.Length == 2)
                    {
                        result.CVal = Convert.ToDouble(parts[0].Trim()) * 10;

                        var subParts = parts[1].Split(',');
                        if (subParts.Length >= 2)
                        {
                            result.AVal = Convert.ToDouble(subParts[0].Trim()) * 10;
                            result.BVal = Convert.ToDouble(subParts[1].Trim()) * 10;
                        }
                    }
                }

                // Xử lý chuỗi từ hàng dưới (YRow + 1)
                string tempStr2 = sourceSheet.Cells[yRow + 1, xCol].Value?.ToString();
                if (!string.IsNullOrEmpty(tempStr2))
                {
                    var spaceParts = tempStr2.Split(' ');
                    if (spaceParts.Length >= 2)
                    {
                        result.Value1 = spaceParts[0].Trim();
                        result.Value2 = spaceParts[1].Trim();
                    }
                }

                // Lấy thông tin bổ sung từ cột C
                string additionalStr = sourceSheet.Cells[yRow, 3].Value?.ToString();
                if (!string.IsNullOrEmpty(additionalStr) && additionalStr.Contains("-"))
                {
                    var dashParts = additionalStr.Split('-');
                    if (dashParts.Length >= 2)
                    {
                        result.AdditionalInfo = dashParts[1].Trim();
                        result.Spacing = double.Parse(dashParts[1].Trim())< 7 ? "600" : "900"; // Giả sử nếu khoảng cách < 7 thì là 600, ngược lại là 900   
                    }
                }
                
            }
            else
            {
                result.IsValid = false;
                result.ErrorMessage = "Không tìm thấy giá trị phù hợp trong bảng tra cứu";
            }
        }
        catch (Exception ex)
        {
            result.IsValid = false;
            result.ErrorMessage = $"Lỗi khi thực hiện tra cứu: {ex.Message}";
        }
      //  MessageBox.Show("ddã vào đây2");
        return result;
    }

    /// <summary>
    /// Ghi kết quả tra cứu vào sheet (tương tự VBA gốc)
    /// </summary>
    /// <param name="result">Kết quả tra cứu</param>
    /// <param name="inputSheetName">Tên sheet để ghi kết quả</param>
    public void WriteResultToSheet(ExcelLookupResult result, string inputSheetName = "Tra TT")
    {
        try
        {
            Excel.Worksheet inputSheet = _workbook.Sheets[inputSheetName];

            // Xóa nội dung cũ (C3:C11)
            inputSheet.Range["C3:C11"].ClearContents();

            if (result.IsValid)
            {
                // Ghi các giá trị vào các ô tương ứng
                inputSheet.Cells[4, 3].Value = result.AVal;      // C4
                inputSheet.Cells[5, 3].Value = result.BVal;      // C5
                inputSheet.Cells[6, 3].Value = result.CVal;      // C6
                inputSheet.Cells[7, 3].Value = result.DVal;      // C7
                inputSheet.Cells[8, 3].Value = result.Value1;    // C8
                inputSheet.Cells[9, 3].Value = result.Value2;    // C9
                inputSheet.Cells[11, 3].Value = result.AdditionalInfo; // C11
               
            }
            else
            {
                inputSheet.Cells[7, 3].Value = "Invalid";
            }
        }
        catch (Exception ex)
        {
            throw new Exception($"Lỗi khi ghi kết quả: {ex.Message}");
        }
    }

    /// <summary>
    /// Hàm kết hợp: thực hiện tra cứu và ghi kết quả (giống hàm VBA gốc)
    /// </summary>
    /// <param name="val1">Giá trị 1 từ ô C1</param>
    /// <param name="val2">Giá trị 2 từ ô C2</param>
    /// <returns>Kết quả tra cứu</returns>
    public ExcelLookupResult Val(double val1, double val2)
    {
        var result = PerformLookup(val1, val2);
        WriteResultToSheet(result);
        return result;
    }

    /// <summary>
  
    /// </summary>
    /// <param name="inputSheetName">Tên sheet chứa dữ liệu đầu vào</param>
    /// <returns>Kết quả tra cứu</returns>
    public ExcelLookupResult Val(string inputSheetName = "Tra TT")
    {
        try
        {
            Excel.Worksheet inputSheet = _workbook.Sheets[inputSheetName];

            // Lấy giá trị từ C1 và C2
            double val1 = Convert.ToDouble(inputSheet.Cells[1, 3].Value);
            double val2 = Convert.ToDouble(inputSheet.Cells[2, 3].Value);

            return Val(val1, val2);
        }
        catch (Exception ex)
        {
            return new ExcelLookupResult
            {
                IsValid = false,
                ErrorMessage = $"Lỗi khi đọc dữ liệu đầu vào: {ex.Message}"
            };
        }
    }

    public void Dispose()
    {
        _workbook?.Close(false);
        _excelApp?.Quit();

        // Giải phóng COM objects
        if (_workbook != null)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(_workbook);
        if (_excelApp != null)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(_excelApp);
    }
}
