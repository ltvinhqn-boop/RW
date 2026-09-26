using System;
using System.Data;
using System.IO;
using OfficeOpenXml;

public class ExcelToDataTableReader
{
    /// <summary>
    /// Đọc Excel và chuyển thành DataTable.
    /// </summary>
    /// <param name="filePath">Đường dẫn file Excel.</param>
    /// <param name="sheetIndex">Chỉ số sheet (mặc định là 0 – sheet đầu tiên).</param>
    /// <param name="hasHeader">Có dòng tiêu đề không (mặc định true).</param>
    /// <returns>DataTable chứa dữ liệu từ Excel.</returns>
    public DataTable ReadExcelToDataTable(string filePath, int sheetIndex , bool hasHeader = true)
    {
        if (!File.Exists(filePath))
            throw new FileNotFoundException("File Excel không tồn tại.", filePath);

        // EPPlus license (chỉ cho bản miễn phí, không dùng thương mại)
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets[sheetIndex];
            if (worksheet == null)
                throw new Exception($"Không tìm thấy sheet tại index {sheetIndex}.");

            var dt = new DataTable();
            int startRow = hasHeader ? 2 : 1;
            int colCount = worksheet.Dimension.Columns;
            int rowCount = worksheet.Dimension.Rows;

            // Tạo cột DataTable
            for (int col = 1; col <= colCount; col++)
            {
                string columnName = hasHeader
                    ? worksheet.Cells[1, col].Text
                    : $"Column{col}";
                dt.Columns.Add(columnName);
            }

            // Thêm dữ liệu vào DataTable
            for (int row = startRow; row <= rowCount; row++)
            {
                var dataRow = dt.NewRow();
                for (int col = 1; col <= colCount; col++)
                {
                    dataRow[col - 1] = worksheet.Cells[row, col].Text;
                }
                dt.Rows.Add(dataRow);
            }

            return dt;
        }
    }
}
