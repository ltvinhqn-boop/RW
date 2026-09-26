using iText.Kernel.Pdf;
using iText.Kernel.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace DWG2PDF
{
    class JOINPDF
    {
        public static void MergePDFFiles(List<string> filePaths, string outputFileName)
        {
            try
            {
                using (var pdfWriter = new PdfWriter(outputFileName))
                using (var pdfDoc = new PdfDocument(pdfWriter))
                {
                    var merger = new PdfMerger(pdfDoc);

                    foreach (var file in filePaths)
                    {
                        if (File.Exists(file))
                        {
                            try
                            {
                                using (var src = new PdfDocument(new PdfReader(file)))
                                {
                                    merger.Merge(src, 1, src.GetNumberOfPages());
                                }
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"Lỗi khi đọc file PDF:\n{file}\n{ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                        }
                        else
                        {
                            MessageBox.Show($"Không tìm thấy file PDF:\n{file}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }

                    pdfDoc.Close();

                    foreach (var file in filePaths)
                    {
                        try { File.Delete(file); } catch { }
                    }
                }

                //MessageBox.Show("Gộp file PDF thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi gộp PDF:\n{ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
