using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using acadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using Autodesk.AutoCAD.EditorInput;
using System.Reflection;
using Autodesk.AutoCAD.Geometry;
using OfficeOpenXml;
using LicenseContext = OfficeOpenXml.LicenseContext;
using static OfficeOpenXml.ExcelErrorValue;
using Autodesk.AutoCAD.DatabaseServices;
using System.Globalization;
using _3D2CAD;
using Autodesk.AutoCAD.ApplicationServices;
using System.Security.Cryptography.Xml;
using System.Security.Cryptography;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using HBTools;
using System.Runtime.InteropServices;
using Autodesk.AutoCAD.PlottingServices;
using iText.Layout.Element;
using RWTools;
using static DWG2PDF.Form1;
using Google.Protobuf.Compiler;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Rebar;
using Autodesk.AutoCAD.DatabaseServices.Filters;
using Org.BouncyCastle.Asn1.X509;
//using Autodesk.AutoCAD.GraphicsInterface;
//using iText.Forms.Form.Element;
//using System.Windows;
namespace DWG2PDF
{
    public partial class Form1 : Form
    {
        IntPtr acadHwnd = new IntPtr(acadApp.MainWindow.Handle.ToInt64());
        //------------------RW
        Double OW = new Double();
        Double OH = new Double();
        Double A = new Double();
        Double B = new Double();
        Double C = new Double();
        Double D = new Double();
        string motor = "";
        string gearbox = "";
        string qlypp;
        string spacing = "";
        List<string> liststylename = new List<string> { };
        //----------------------List Blockname

        List<string> listblockname1 = new List<string>() { "A3-KT-RW", "BangTSRW" };
        List<string> listblockname2 = new List<string>() { "A3-KT-RW", "BangTSRW", "BangKLRW" };
        List<string> listblockname3 = new List<string>() { "A3-KT-RW", "BangTSRW-2" };
        List<string> listblockname4 = new List<string>() { "A3-KT-RW", "BangTSRW-2", "BangKLRW" };
        List<string> listdimstyle = new List<string>() { "WKV_10", "WKV_15", "WKV_20", "WKV_25", "WKV_30", "WKV_35", "WKV_40", "WKV_50", "WKV_60", "WKV_70", "WKV_80" };
        List<string> A3_KT_RW = new List<string>() { };
        List<string> Bang_TSRW = new List<string>() { };
        List<string> Bang_KLRW = new List<string>() { };

        List<string>[] WKV_Array = new List<string>[] { };
        PromptPointResult pointResult;
        Point3d p0;
        //--------------------  

        List<double> values = new List<double>();
        List<double> values2 = new List<double>();
        public List<string> list914 = new List<string>() { };
        public List<string> list1219 = new List<string>() { };
        public List<ZamPara> ListZam = new List<ZamPara>();
        List<Item> items1 = new List<Item> { };
        List<Item> items2 = new List<Item> { };
        // Tạo mảng các TextBox để duyệt nhanh
        TextBox[] textBoxes;
        CheckBox[] RD1;
        // Tạo mảng các TextBox để duyệt nhanh
        Button[] buttons;
        double scale = 1.0;

        List<SizeView1> sizeview = new List<SizeView1>();
        public List<Extents2d> liextent2d = new List<Extents2d>();
        // int[,] array_hole = new int[2, 3];
        // double gapx = 0.0;
        //double gapy = 0.0;
        ObjectId dimstyle_text = ObjectId.Null;
        ObjectId dimstyle;
        // string filePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\HB特注原価計算.xlsx";
        double Length = 0.0;
        double Height = 0.0;
        string t_thep = "";
        string t_ong1 = "";
        string t_ong2 = "";
        string t_ong3 = "";
        Hatch hatch = new Hatch();
        // ExcelPackage package;// = new ExcelPackage(new FileInfo(filePath));
        // Listx
        string tempFilePath;
        string version = "RWtool_v1.0.1";
        public Form1()
        {
            InitializeComponent();
            this.AllowDrop = true;

        }

        //-----------------------------------xuất excel1
        public static void ExportPDFListToExcel()
        {
            try
            {
                // Chọn thư mục
                string folderPath = SelectFolder();
                if (string.IsNullOrEmpty(folderPath))
                    return;

                // Lấy danh sách file PDF
                string[] pdfFiles = GetPDFFiles(folderPath);
                if (pdfFiles.Length == 0)
                {
                    MessageBox.Show("Không tìm thấy file PDF nào trong thư mục!");
                    return;
                }

                // Tạo Excel và xuất dữ liệu
                CreateExcelFile(pdfFiles);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Hiển thị hộp thoại chọn thư mục
        /// </summary>
        private static string SelectFolder()
        {
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Chọn thư mục chứa file PDF";

                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    return folderDialog.SelectedPath;
                }
            }
            return null;
        }

        /// <summary>
        /// Lấy tất cả file PDF trong thư mục
        /// </summary>
        private static string[] GetPDFFiles(string folderPath)
        {
            return Directory.GetFiles(folderPath, "*.pdf");
        }

        /// <summary>
        /// Tách tên file thành 2 phần: mã và tên
        /// </summary>
        private static (string code, string name) SplitFileName(string fileName)
        {
            // Lấy tên file không có extension
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(fileName);

            // Lấy 6 ký tự đầu tiên
            string code = fileNameWithoutExt.Length >= 7
                ? fileNameWithoutExt.Substring(0, 7)
                : fileNameWithoutExt;

            // Lấy phần còn lại và trim dấu cách ở đầu
            string name = fileNameWithoutExt.Length > 7
                ? fileNameWithoutExt.Substring(7).TrimStart()
                : "";

            return (code, name);
        }

        /// <summary>
        /// Tạo file Excel và xuất dữ liệu
        /// </summary>
        private static void CreateExcelFile(string[] pdfFiles)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                // Tạo Excel application
                excelApp = new Excel.Application();
                excelApp.Visible = true;
                workbook = excelApp.Workbooks.Add();
                worksheet = (Excel.Worksheet)workbook.Sheets[1];

                // Thêm header
                AddHeader(worksheet);

                // Thêm dữ liệu
                AddDataToWorksheet(worksheet, pdfFiles);

                // Tự động điều chỉnh độ rộng cột
                worksheet.Columns.AutoFit();

                // Lưu file
                SaveExcelFile(workbook, pdfFiles.Length);
            }
            finally
            {
                // Giải phóng COM objects
                ReleaseComObjects(worksheet, workbook, excelApp);
            }
        }

        /// <summary>
        /// Thêm tiêu đề vào worksheet
        /// </summary>
        private static void AddHeader(Excel.Worksheet worksheet)
        {
            worksheet.Cells[1, 1] = "Mã";
            worksheet.Cells[1, 2] = "Tên";

            // Định dạng tiêu đề
            Excel.Range headerRange = worksheet.Range["A1", "B1"];
            headerRange.Font.Bold = true;
            headerRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
        }

        /// <summary>
        /// Thêm dữ liệu PDF vào worksheet
        /// </summary>
        private static void AddDataToWorksheet(Excel.Worksheet worksheet, string[] pdfFiles)
        {
            int row = 2;
            foreach (string pdfFile in pdfFiles)
            {
                var (code, name) = SplitFileName(pdfFile);

                worksheet.Cells[row, 1] = code;
                worksheet.Cells[row, 2] = name;
                row++;
            }
        }

        /// <summary>
        /// Lưu file Excel
        /// </summary>
        private static void SaveExcelFile(Excel.Workbook workbook, int fileCount)
        {
            using (SaveFileDialog saveDialog = new SaveFileDialog())
            {
                saveDialog.Filter = "Excel Files|*.xlsx";
                saveDialog.Title = "Lưu file Excel";
                saveDialog.FileName = "DanhSachPDF.xlsx";

                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    workbook.SaveAs(saveDialog.FileName);
                    MessageBox.Show($"Đã xuất thành công {fileCount} file PDF ra Excel!",
                        "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        /// <summary>
        /// Giải phóng COM objects
        /// </summary>
        private static void ReleaseComObjects(Excel.Worksheet worksheet, Excel.Workbook workbook, Excel.Application excelApp)
        {
            if (worksheet != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);

            if (workbook != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);

            if (excelApp != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
        }
        //---------------------------------



        private bool IsNumeric(string text)
        {
            return double.TryParse(text, out _);
        }
        private void Form1_DragDrop(object sender, DragEventArgs e)
        {

        }

        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            //if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
        }

        double RoundToNearest(double number)
        {
            // Làm tròn về bội số gần nhất của 5
            double round5 = Math.Round(number / 5.0) * 5.0;

            // Làm tròn về bội số gần nhất của 10
            double round10 = Math.Round(number / 10.0) * 10.0;

            // Tính khoảng cách đến số gốc
            double diff5 = Math.Abs(number - round5);
            double diff10 = Math.Abs(number - round10);

            // Trả về bội số gần hơn
            return (diff5 <= diff10) ? round5 : round10;
        }

        private void button1_Click(object sender, EventArgs e)
        {


        }
        public void AutoHoles(double x, double y)
        {

        }
        public static void AddLeaderNote2(Point3d startPoint, Point3d textPoint, string note, ObjectId dimStyleId)
        {
            Document doc = acadApp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            using (DocumentLock m_DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument())
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                    // Lấy thông tin từ DimStyle
                    DimStyleTableRecord dimStyle = tr.GetObject(dimStyleId, OpenMode.ForRead) as DimStyleTableRecord;
                    double textHeight = dimStyle.Dimtxt;
                    double dimscale = dimStyle.Dimscale;
                    double arrowSize = dimStyle.Dimasz;
                    ObjectId arrowhead = dimStyle.Dimblk1;
                    ObjectId textstyle = dimStyle.Dimtxsty;
                    //MessageBox.Show($"textHeight: {textHeight}, dimscale: {dimscale}");
                    ObjectId textStyleId = dimStyleId;

                    // Tạo MText
                    MText mtext = new MText
                    {
                        Contents = note,
                        Height = textHeight * dimscale,
                        TextHeight = textHeight * dimscale,

                        ColorIndex = 2, // Màu sắc của text
                        Location = textPoint,
                        // TextStyleId = textstyle

                    };
                    mtext.LineWeight = LineWeight.LineWeight009; // Đặt độ dày của văn bản
                    mtext.Attachment = AttachmentPoint.MiddleCenter;                                             // mtext.EdgeStyleId\
                                                                                                                 // Tạo MLeader
                    MLeader mleader = new MLeader
                    {
                        ContentType = ContentType.MTextContent,
                        MText = mtext,

                        Layer = "0", // hoặc tuỳ chọn layer
                        ColorIndex = 2, // Màu sắc của leader
                                        //   TextStyleId = textStyleId

                    };
                    mleader.TextLocation = new Point3d(textPoint.X - 200, textPoint.Y, 0);
                    mleader.LineWeight = LineWeight.LineWeight009; // Đặt độ dày của đường dẫn
                    mleader.TextStyleId = textstyle;

                    mleader.SetDatabaseDefaults();
                    // mleader.add
                    int leaderIndex = mleader.AddLeader();
                    int lineIndex = mleader.AddLeaderLine(leaderIndex);
                    mleader.AddFirstVertex(leaderIndex, startPoint);
                    mleader.AddLastVertex(leaderIndex, textPoint);
                    // mleader.Annotative = AnnotativeStates.True;
                    //  mleader.TextAttachmentDirection = TextAttachmentDirection.AttachmentVertical;
                    // Căn trái nếu cần
                    mleader.SetArrowSize(leaderIndex, arrowSize * dimscale);
                    mleader.SetArrowSymbolId(leaderIndex, arrowhead);
                    //mleader.SetArrowSymbolId
                    //mleader.TextAttachmentType = TextAttachmentType.AttachmentBottomLine;
                    // Đặt căn chỉnh văn bản sang phải
                    mleader.TextAttachmentType = TextAttachmentType.AttachmentBottomOfTopLine;
                    mleader.SetTextAttachmentType(TextAttachmentType.AttachmentBottomOfTopLine, LeaderDirectionType.RightLeader);
                    mleader.SetTextAttachmentType(TextAttachmentType.AttachmentBottomOfTopLine, LeaderDirectionType.LeftLeader);
                    // mleader.TextAlignmentType = TextAlignmentType.RightAlignment;
                    // mleader.TextHeight = textHeight;

                    // Thêm vào bản vẽ
                    btr.AppendEntity(mleader);
                    tr.AddNewlyCreatedDBObject(mleader, true);
                    // === DI CHUYỂN ĐIỂM TEXT SAU KHI TẠO ===
                    // Lấy vị trí hiện tại

                    Point3d oldPos = mtext.Location;

                    // Dịch xuống 3 lần chiều cao text (ví dụ)
                    Point3d newPos = new Point3d(oldPos.X - 1000, oldPos.Y, oldPos.Z);

                    // Cập nhật vị trí text
                    mtext.Location = newPos;
                    tr.Commit();
                }
            }
        }
        public static void AddLeaderNote3(Point3d startPoint, Point3d textLocation, string note, ObjectId dimStyleId)

        {
            Document doc = acadApp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            using (var lockDoc = doc.LockDocument())
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                DimStyleTableRecord dim = tr.GetObject(dimStyleId, OpenMode.ForRead) as DimStyleTableRecord;

                double txtH = dim.Dimtxt * dim.Dimscale;
                double arrow = dim.Dimasz * dim.Dimscale;

                // 1) MText đặt đúng vị trí
                MText mt = new MText();
                mt.Contents = note;
                mt.TextStyleId = dim.Dimtxsty;
                mt.TextHeight = txtH;
                mt.Location = textLocation;
                mt.Attachment = AttachmentPoint.MiddleCenter;
                mt.ColorIndex = 2;

                btr.AppendEntity(mt);
                tr.AddNewlyCreatedDBObject(mt, true);

                // 2) Leader vẽ bằng Polyline
                Polyline pl = new Polyline();
                pl.ColorIndex = 2;

                // Điểm 1 = mũi tên
                pl.AddVertexAt(0, new Point2d(startPoint.X, startPoint.Y), 0, 0, 0);

                // Điểm 2 = điểm gãy (ngang hàng với text)
                Point3d breakPoint = new Point3d(textLocation.X - 20 * dim.Dimscale, startPoint.Y, 0);
                pl.AddVertexAt(1, new Point2d(breakPoint.X, breakPoint.Y), 0, 0, 0);

                // Điểm 3 = điểm cuối leader (ngay dưới chữ)
                Point3d endPoint = new Point3d(textLocation.X, textLocation.Y - txtH * 0.7, 0);
                pl.AddVertexAt(2, new Point2d(endPoint.X, endPoint.Y), 0, 0, 0);

                btr.AppendEntity(pl);
                tr.AddNewlyCreatedDBObject(pl, true);

                // 3) Chèn mũi tên theo DimStyle
                if (!dim.Dimblk.IsNull && arrow > 0)
                {
                    BlockTableRecord arrowBlk =
                        tr.GetObject(dim.Dimblk, OpenMode.ForRead) as BlockTableRecord;

                    using (BlockReference arrowRef = new BlockReference(startPoint, dim.Dimblk))
                    {
                        arrowRef.ScaleFactors = new Scale3d(dim.Dimscale);
                        arrowRef.Rotation = Math.Atan2(
                            breakPoint.Y - startPoint.Y,
                            breakPoint.X - startPoint.X);

                        arrowRef.ColorIndex = 2;
                        btr.AppendEntity(arrowRef);
                        tr.AddNewlyCreatedDBObject(arrowRef, true);
                    }
                }

                tr.Commit();
            }
        }




        public void CreateLeaderAndMText(Point3d p1, Point3d p2, Point3d p3, string note, ObjectId dimStyleId)
        {
            Document doc = acadApp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            using (DocumentLock m_DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument())
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord btr = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                    DimStyleTableRecord dimStyle = tr.GetObject(dimStyleId, OpenMode.ForRead) as DimStyleTableRecord;
                    double textHeight = dimStyle.Dimtxt * dimStyle.Dimscale;

                    double dimscale = dimStyle.Dimscale;
                    double arrowSize = dimStyle.Dimasz * dimStyle.Dimscale;
                    ObjectId arrowhead = dimStyle.Dimblk1;
                    ObjectId textstyle = dimStyle.Dimtxsty;
                    //MessageBox.Show($"textHeight: {textHeight}, dimscale: {dimscale}");
                    ObjectId textStyleId = dimStyleId;


                    // ====== Tạo MText ======
                    MText mtext = new MText();
                    mtext.Contents = note; // \\P = xuống dòng trong MText
                    mtext.TextHeight = textHeight;
                    mtext.Location = p3;
                    mtext.Attachment = AttachmentPoint.MiddleCenter;
                    mtext.ColorIndex = 2; // Màu vàng (giống hình)
                    ObjectId mtextId = textstyle;
                    mtextId = btr.AppendEntity(mtext);
                    tr.AddNewlyCreatedDBObject(mtext, true);

                    // ====== Tạo Leader ======
                    Leader leader = new Leader();
                    leader.SetDatabaseDefaults();
                    leader.TextStyleId = textstyle;
                    leader.HasArrowHead = true;
                  //  leader.seta
                    leader.ColorIndex = 2;
                    leader.AppendVertex(p1);
                    leader.AppendVertex(p2);
                    leader.AppendVertex(p3);


                    // Liên kết với MText
                    //  leader.Annotation = mtextId;
                    //// leader.AnnotType = LeaderAnnotType.MText;
                    //leader.Dimasz = 2.5; // kích thước mũi tên
                    //leader.Dimgap = 2.0; // khoảng cách text với đường leader
                    //leader.EvaluateLeader();

                    btr.AppendEntity(leader);
                    tr.AddNewlyCreatedDBObject(leader, true);

                    tr.Commit();
                }

                ed.WriteMessage("\nLeader + MText created successfully!");
            }
        }
        public static void AddLeaderAndText(Point3d startPoint, Point3d textPoint, string note, ObjectId dimStyleId)
        {
            Document doc = acadApp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            using (var docLock = doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                BlockTableRecord btr =
                    (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                // Lấy Dimstyle
                DimStyleTableRecord dimStyle =
                    (DimStyleTableRecord)tr.GetObject(dimStyleId, OpenMode.ForRead);

                double textHeight = dimStyle.Dimtxt * dimStyle.Dimscale;
                double arrowSize = dimStyle.Dimasz * dimStyle.Dimscale;
                ObjectId arrowBlockId = dimStyle.Dimblk;      // block mũi tên
                ObjectId textStyleId = dimStyle.Dimtxsty;     // text style

                //----------------------------------
                // 1) LEADER (cổ điển)
                //----------------------------------
                Leader leader = new Leader();
                leader.SetDatabaseDefaults();
                leader.ColorIndex = 2;
                leader.HasArrowHead = true;        // bật mũi tên

                

                leader.AppendVertex(startPoint);
                leader.AppendVertex(textPoint);

                btr.AppendEntity(leader);
                tr.AddNewlyCreatedDBObject(leader, true);

                //----------------------------------
                // 2) MTEXT
                //----------------------------------
                MText mtext = new MText();
                mtext.Contents = @"\W0.8;" + note;
                mtext.TextHeight = textHeight;
                mtext.TextStyleId = textStyleId;
                mtext.ColorIndex = 2;

                // Căn chữ theo vị trí
                if (textPoint.X > startPoint.X)
                    mtext.Attachment = AttachmentPoint.BottomLeft;
                else
                    mtext.Attachment = AttachmentPoint.BottomRight;

                mtext.Location = textPoint;

                btr.AppendEntity(mtext);
                tr.AddNewlyCreatedDBObject(mtext, true);

                tr.Commit();
            }
        }


        /*   public static void AddLeaderNote2(Point3d startPoint, Point3d textPoint, string note, ObjectId dimStyleId)
           {
               Document doc = acadApp.DocumentManager.MdiActiveDocument;
               Database db = doc.Database;

               using (DocumentLock docLock = doc.LockDocument())
               using (Transaction tr = db.TransactionManager.StartTransaction())
               {
                   BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                   DimStyleTableRecord dimStyle = (DimStyleTableRecord)tr.GetObject(dimStyleId, OpenMode.ForRead);

                   double textHeight = dimStyle.Dimtxt;
                   double dimscale = dimStyle.Dimscale;
                   double arrowSize = dimStyle.Dimasz;
                   ObjectId arrowhead = dimStyle.Dimblk1;
                   ObjectId textstyle = dimStyle.Dimtxsty;

                   // --- MText ---
                   MText mtext = new MText
                   {
                       Contents = note,
                       TextHeight = textHeight * dimscale,
                       ColorIndex = 2
                   };

                   // --- MLeader ---
                   MLeader mleader = new MLeader();
                   mleader.SetDatabaseDefaults();
                   mleader.Layer = "0";
                   mleader.ColorIndex = 2;
                   mleader.TextStyleId = textstyle;
                   mleader.ArrowSize = arrowSize * dimscale;
                 //  mleader.SetArrowSymbolId(arrowhead);
                   mleader.EnableDogleg = true;
                   mleader.ContentType = ContentType.MTextContent;
                   mleader.MText = mtext;

                   int leaderIndex = mleader.AddLeader();
                   int lineIndex = mleader.AddLeaderLine(leaderIndex);
                   mleader.AddFirstVertex(leaderIndex, startPoint);
                   mleader.AddLastVertex(leaderIndex, textPoint);

                   // --- Thêm vào bản vẽ ---
                   btr.AppendEntity(mleader);
                   tr.AddNewlyCreatedDBObject(mleader, true);

                   // === DI CHUYỂN ĐIỂM TEXT SAU KHI TẠO ===
                   // Lấy vị trí hiện tại
                   Point3d oldPos = mleader.TextLocation;

                   // Dịch xuống 3 lần chiều cao text (ví dụ)
                //   Point3d newPos = new Point3d(oldPos.X-200, oldPos.Y , oldPos.Z);

                   // Cập nhật vị trí text
                //   mleader.TextLocation = newPos;

                   tr.Commit();
               }
           }*/
        public static void AddLeaderNotecmd(Point3d startPoint, Point3d textPoint, string note, ObjectId dimStyleId)
        {
            Document doc = acadApp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            double textHeight = 2.5;
            double dimscale = 1.0;
            double arrowSize = 2.5;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                DimStyleTableRecord dimStyle = tr.GetObject(dimStyleId, OpenMode.ForRead) as DimStyleTableRecord;
                dimscale = dimStyle.Dimscale;
                textHeight = dimStyle.Dimtxt* dimscale;
               
                arrowSize = dimStyle.Dimasz* dimscale;
                tr.Commit();
            }

            // 1) Set các biến hệ thống trước khi chạy MLEADER
            string cmd =
                "_DIMASZ " + arrowSize + " " +
                "_DIMSCALE " + dimscale + " " +
                "_MLEADERTEXTHEIGHT " + textHeight + " " +

                // 2) Gọi MLEADER như bình thường
                "_MLEADER " +
                startPoint.X + "," + startPoint.Y + " " +
                textPoint.X + "," + textPoint.Y + " " +
                note + " " +
                " ";

            doc.SendStringToExecute(cmd, false, false, false);
        }

        public static ObjectId AddLeaderNote1(Point3d startPoint, Point3d textPoint, string note, ObjectId dimStyleId)
        {
            Document doc = acadApp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            using (DocumentLock m_DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument())
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                    // Lấy thông tin từ DimStyle
                    DimStyleTableRecord dimStyle = tr.GetObject(dimStyleId, OpenMode.ForRead) as DimStyleTableRecord;
                    double textHeight = dimStyle.Dimtxt;
                    double dimscale = dimStyle.Dimscale;
                    double arrowSize = dimStyle.Dimasz;
                    ObjectId arrowhead = dimStyle.Dimblk1;
                    ObjectId textstyle = dimStyle.Dimtxsty;
                    //MessageBox.Show($"textHeight: {textHeight}, dimscale: {dimscale}");
                    ObjectId textStyleId = dimStyleId;

                    // Tạo MText
                    MText mtext = new MText
                    {
                        Contents = @"\W0.8;" + note,
                        Height = textHeight * dimscale,
                        TextHeight = textHeight * dimscale,

                        ColorIndex = 2, // Màu sắc của text
                        Location = textPoint,
                        // TextStyleId = textstyle

                    };
                    mtext.LineWeight = LineWeight.LineWeight009; // Đặt độ dày của văn bản
                    mtext.Attachment = AttachmentPoint.MiddleCenter;                                             // mtext.EdgeStyleId\
                                                                                                                 // Tạo MLeader
                    MLeader mleader = new MLeader
                    {
                        ContentType = ContentType.MTextContent,
                        MText = mtext,

                        Layer = "0", // hoặc tuỳ chọn layer
                        ColorIndex = 2, // Màu sắc của leader
                                        //   TextStyleId = textStyleId

                    };
                    mleader.LineWeight = LineWeight.LineWeight009; // Đặt độ dày của đường dẫn
                    mleader.TextStyleId = textstyle;
                    mleader.SetDatabaseDefaults();
                    // mleader.add
                    int leaderIndex = mleader.AddLeader();
                    int lineIndex = mleader.AddLeaderLine(leaderIndex);
                    mleader.AddFirstVertex(leaderIndex, startPoint);
                    mleader.AddLastVertex(leaderIndex, textPoint);
                    // mleader.Annotative = AnnotativeStates.True;
                    //  mleader.TextAttachmentDirection = TextAttachmentDirection.AttachmentVertical;
                    // Căn trái nếu cần
                    mleader.SetArrowSize(leaderIndex, arrowSize * dimscale);
                    mleader.SetArrowSymbolId(leaderIndex, arrowhead);
                    //mleader.SetArrowSymbolId
                    //mleader.TextAttachmentType = TextAttachmentType.AttachmentBottomLine;
                   // Đặt căn chỉnh văn bản sang phải
                      mleader.TextAttachmentType = TextAttachmentType.AttachmentBottomLine;
                    mleader.SetTextAttachmentType(TextAttachmentType.AttachmentBottomOfTopLine, LeaderDirectionType.LeftLeader);
                    mleader.SetTextAttachmentType(TextAttachmentType.AttachmentBottomOfTopLine, LeaderDirectionType.RightLeader);
                    // Xác định leader đi sang trái hay sang phải
                    //LeaderDirectionType dir =
                    //    (textPoint.X >= startPoint.X)
                    //        ? LeaderDirectionType.LeftLeader   // text nằm bên phải mũi tên
                    //        : LeaderDirectionType.RightLeader; // text nằm bên trái mũi tên

                    //mleader.TextAttachmentType = TextAttachmentType.AttachmentBottomOfTopLine;
                    //mleader.SetTextAttachmentType(
                    //    TextAttachmentType.AttachmentBottomOfTopLine, dir);

                    // mleader.TextAlignmentType = TextAlignmentType.RightAlignment;
                    // mleader.TextHeight = textHeight;

                    // Thêm vào bản vẽ
                    btr.AppendEntity(mleader);
                    tr.AddNewlyCreatedDBObject(mleader, true);
                    ObjectId id = mleader.ObjectId;

                    tr.Commit();
                    return id;
                }
            }
        }
        //********************************************
        public static Point3d MirrorPointAcrossVerticalLine(Point3d pt, double mirrorX)
        {
            double dx = mirrorX - pt.X;
            return new Point3d(mirrorX + dx, pt.Y, pt.Z);
        }
        public static void AddLeaderNoteMirror(
    Point3d startPoint,
    Point3d textPoint,
    string note,
    ObjectId dimStyleId,
    double mirrorX)    // trục mirror đứng: X = mirrorX
        {
            // Tạo 2 điểm sau khi mirror
            Point3d startMirror = MirrorPointAcrossVerticalLine(startPoint, mirrorX);
            Point3d textMirror = MirrorPointAcrossVerticalLine(textPoint, mirrorX);

            // Đảo hướng: theo lý thuyết textRightOfArrow sẽ đảo ngược
            bool originalRight = textPoint.X > startPoint.X;
            bool mirroredRight = !originalRight;

            // Gọi lại hàm cũ nhưng đảo hướng của textPoint so với startPoint
            AddLeaderNote1(startMirror, textMirror, note, dimStyleId);
        }
        //**********************************************************************
        public static void MirrorByCommand(ObjectId objId, Point3d p1, Point3d p2, int era)
        {
            Document doc = acadApp.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            // Tạo selection set tạm thời
            ObjectId[] ids = new ObjectId[] { objId };
            ed.SetImpliedSelection(ids);
            string yn = era==1 ? "Y " : "N ";
            // Gửi lệnh MIRROR + đường mirror
            string cmd =
                "_MIRROR " +
                "_SI " +                // dùng selection hiện tại
                p1.X + "," + p1.Y + " " +
                p2.X + "," + p2.Y + " " +
                yn;                   // Không xóa đối tượng gốc (Yes/No)

            doc.SendStringToExecute(cmd, false, false, false);
            Thread.Sleep(100);
            // Clear selection
          //  ed.SetImpliedSelection(new ObjectId[0]);
        }
        public static ObjectId AddLeaderNotemir(
    Point3d startPoint,   // mũi tên
    Point3d textPoint,    // điểm gốc của text (góc dưới trái / phải)
    string note,
    ObjectId dimStyleId)
        {
            Document doc = acadApp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            using (DocumentLock m_DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument())
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                    // Lấy thông tin từ DimStyle
                    DimStyleTableRecord dimStyle = tr.GetObject(dimStyleId, OpenMode.ForRead) as DimStyleTableRecord;
                    double textHeight = dimStyle.Dimtxt;
                    double dimscale = dimStyle.Dimscale;
                    double arrowSize = dimStyle.Dimasz;
                    ObjectId arrowhead = dimStyle.Dimblk1;
                    ObjectId textstyle = dimStyle.Dimtxsty;
                    //MessageBox.Show($"textHeight: {textHeight}, dimscale: {dimscale}");
                    ObjectId textStyleId = dimStyleId;

                    // Tạo MText
                    MText mtext = new MText
                    {
                        Contents = @"\W0.8;" + note,
                        Height = textHeight * dimscale,
                        TextHeight = textHeight * dimscale,

                        ColorIndex = 2, // Màu sắc của text
                        Location = textPoint,
                        // TextStyleId = textstyle

                    };
                    mtext.LineWeight = LineWeight.LineWeight009; // Đặt độ dày của văn bản
                    mtext.Attachment = AttachmentPoint.MiddleCenter;                                             // mtext.EdgeStyleId\
                                                                                                                 // Tạo MLeader
                    MLeader mleader = new MLeader
                    {
                        ContentType = ContentType.MTextContent,
                        MText = mtext,

                        Layer = "0", // hoặc tuỳ chọn layer
                        ColorIndex = 2, // Màu sắc của leader
                                        //   TextStyleId = textStyleId

                    };
                    mleader.LineWeight = LineWeight.LineWeight009; // Đặt độ dày của đường dẫn
                    mleader.TextStyleId = textstyle;
                    mleader.SetDatabaseDefaults();
                    // mleader.add
                    int leaderIndex = mleader.AddLeader();
                    int lineIndex = mleader.AddLeaderLine(leaderIndex);
                    mleader.AddFirstVertex(leaderIndex, startPoint);
                    mleader.AddLastVertex(leaderIndex, textPoint);
                    // mleader.Annotative = AnnotativeStates.True;
                    //  mleader.TextAttachmentDirection = TextAttachmentDirection.AttachmentVertical;
                    // Căn trái nếu cần
                    mleader.SetArrowSize(leaderIndex, arrowSize * dimscale);
                    mleader.SetArrowSymbolId(leaderIndex, arrowhead);
                    //mleader.SetArrowSymbolId
                    //mleader.TextAttachmentType = TextAttachmentType.AttachmentBottomLine;
                    // Đặt căn chỉnh văn bản sang phải
                    //  mleader.TextAttachmentType = TextAttachmentType.AttachmentBottomLine;
                    //  mleader.SetTextAttachmentType(TextAttachmentType.AttachmentBottomOfTopLine, LeaderDirectionType.LeftLeader);
                    //   mleader.SetTextAttachmentType(TextAttachmentType.AttachmentBottomOfTopLine, LeaderDirectionType.RightLeader);
                    // Xác định leader đi sang trái hay sang phải
                    LeaderDirectionType dir =
                        (textPoint.X >= startPoint.X)
                            ? LeaderDirectionType.LeftLeader   // text nằm bên phải mũi tên
                            : LeaderDirectionType.RightLeader; // text nằm bên trái mũi tên

                    mleader.TextAttachmentType = TextAttachmentType.AttachmentBottomOfTopLine;
                    mleader.SetTextAttachmentType(
                        TextAttachmentType.AttachmentBottomOfTopLine, dir);

                    // mleader.TextAlignmentType = TextAlignmentType.RightAlignment;
                    // mleader.TextHeight = textHeight;

                    // Thêm vào bản vẽ
                    btr.AppendEntity(mleader);
                    tr.AddNewlyCreatedDBObject(mleader, true);

                    ObjectId id = mleader.ObjectId;

                    tr.Commit();
                    return id;
                }
            }
        }

        //*****************************************************************
        public static void AddLeaderNotetest(
    Point3d startPoint,   // mũi tên
    Point3d textPoint,    // điểm gốc của text (góc dưới trái / phải)
    string note,
    ObjectId dimStyleId)
        {
            Document doc = acadApp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            using (doc.LockDocument())
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTableRecord btr =
                    (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                // Lấy thông tin từ DimStyle
                DimStyleTableRecord dimStyle =
                    (DimStyleTableRecord)tr.GetObject(dimStyleId, OpenMode.ForRead);

                double dimscale = dimStyle.Dimscale;
                double textHeight = dimStyle.Dimtxt * dimscale;
                double arrowSize = dimStyle.Dimasz * dimscale;
                ObjectId textstyle = dimStyle.Dimtxsty;
                ObjectId arrowhead = dimStyle.Dimblk1;

                // text nằm bên phải hay bên trái mũi tên?
                bool textRightOfArrow = textPoint.X > startPoint.X;

                // ----- MText -----
                MText mtext = new MText
                {
                    Contents = @"\W0.8;" + note,
                    
                    TextHeight = textHeight,
                    Height = textHeight,
                    ColorIndex = 2,
                    Location = textPoint,
                    TextStyleId = textstyle
                };
                mtext.LineWeight = LineWeight.LineWeight009;
                
                // Gắn gốc text ở GÓC DƯỚI TRÁI / PHẢI
                mtext.Attachment = textRightOfArrow
                    ? AttachmentPoint.BottomLeft   // text nằm bên phải mũi tên
                    : AttachmentPoint.BottomRight; // text nằm bên trái mũi tên

                // ----- MLeader -----
                MLeader mleader = new MLeader();
                mleader.SetDatabaseDefaults();
                mleader.Layer = "0";
                mleader.ColorIndex = 2;
                mleader.LineWeight = LineWeight.LineWeight009;
                
                mleader.TextStyleId = textstyle;

                int leaderIndex = mleader.AddLeader();
                int lineIndex = mleader.AddLeaderLine(leaderIndex);
                mleader.AddFirstVertex(lineIndex, startPoint);
                mleader.AddLastVertex(lineIndex, textPoint);

                // kiểu gắn text vào đường gạch dưới (giống dim)
                mleader.SetTextAttachmentType(
                    TextAttachmentType.AttachmentBottomOfTopLine,
                    LeaderDirectionType.LeftLeader);
                mleader.SetTextAttachmentType(
                    TextAttachmentType.AttachmentBottomOfTopLine,
                    LeaderDirectionType.RightLeader);
                mleader.TextAttachmentType = TextAttachmentType.AttachmentBottomOfTopLine;
                mleader.TextAttachmentDirection = TextAttachmentDirection.AttachmentHorizontal;

                // gán nội dung
                mleader.ContentType = ContentType.MTextContent;
                mleader.MText = mtext;

                // căn lề text: text nằm bên phải mũi tên thì căn trái, ngược lại
                mleader.TextAlignmentType = textRightOfArrow
                    ? TextAlignmentType.LeftAlignment
                    : TextAlignmentType.RightAlignment;

                // mũi tên
                mleader.SetArrowSymbolId(leaderIndex, arrowhead);
                mleader.SetArrowSize(leaderIndex, arrowSize);

                btr.AppendEntity(mleader);
                tr.AddNewlyCreatedDBObject(mleader, true);
                tr.Commit();
            }
        }
        //-------------------------------------------------------------------------------------------------------
        public static ObjectId MirrorMLeaderByTransform(ObjectId mleaderId, Point3d p1, Point3d p2)
        {
            Document doc = acadApp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            using (doc.LockDocument())
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // MLeader gốc
                MLeader oldML = (MLeader)tr.GetObject(mleaderId, OpenMode.ForWrite);

                // Tạo ma trận mirror
                Matrix3d mirrorMatrix = Matrix3d.Mirroring(new Line3d(p1, p2));

                // Clone
                MLeader newML = (MLeader)oldML.Clone();

                // Mirror bằng TransformBy
                newML.TransformBy(mirrorMatrix);

                // Sửa lại hướng text vì MText bị lật sau mirror
                if (newML.MText != null)
                {
                    MText mt = newML.MText;

                    // AutoCAD sẽ lật rotation, phải đảo lại
                    
                    mt.Rotation = -mt.Rotation;

                    // Đảo alignment
                    switch (newML.TextAlignmentType)
                    {
                        case TextAlignmentType.LeftAlignment:
                            newML.TextAlignmentType = TextAlignmentType.RightAlignment;
                            break;

                        case TextAlignmentType.RightAlignment:
                            newML.TextAlignmentType = TextAlignmentType.LeftAlignment;
                            break;
                    }
                   // mt.Contents = "";
                    newML.MText = mt;
                }

                // ArrowSize là property
                double oldArrow = newML.ArrowSize;
                newML.ArrowSize = oldArrow;  // ép AutoCAD cập nhật lại hướng arrow

                // Ghi MLeader mới vào bản vẽ
                BlockTableRecord btr =
                    (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                ObjectId newId = btr.AppendEntity(newML);
                tr.AddNewlyCreatedDBObject(newML, true);
              //  ObjectId id = newML.ObjectId;
                // XÓA MLEADER CŨ
                oldML.Erase(true);

                tr.Commit();
                return newId;
            }
        }


        //----------------------------------------------------------------------------------------------
        public static void AddLeaderNote(Point3d startPoint, Point3d textPoint, string note, ObjectId dimStyleId)
        {
            Document doc = acadApp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            using (DocumentLock m_DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument())
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                    // Lấy thông tin từ DimStyle
                    DimStyleTableRecord dimStyle = tr.GetObject(dimStyleId, OpenMode.ForRead) as DimStyleTableRecord;
                    double textHeight = dimStyle.Dimtxt;
                    double dimscale = dimStyle.Dimscale;
                    double arrowSize = dimStyle.Dimasz;
                    ObjectId arrowhead = dimStyle.Dimblk1;
                    ObjectId textstyle = dimStyle.Dimtxsty;
                    //MessageBox.Show($"textHeight: {textHeight}, dimscale: {dimscale}");
                    ObjectId textStyleId = dimStyleId;

                    // Tạo MText
                    MText mtext = new MText
                    {
                        Contents = @"\W0.8;" + note,
                        Height = textHeight * dimscale,
                        TextHeight = textHeight * dimscale,

                        ColorIndex = 2, // Màu sắc của text
                        Location = textPoint,
                        // TextStyleId = textstyle

                    };
                    mtext.LineWeight = LineWeight.LineWeight009; // Đặt độ dày của văn bản
                                                                 //  mtext.w = 0.8;                                          // mtext.EdgeStyleId\
                                                                 // Tạo MLeader
                    MLeader mleader = new MLeader
                    {
                        ContentType = ContentType.MTextContent,
                        MText = mtext,

                        Layer = "0", // hoặc tuỳ chọn layer
                        ColorIndex = 2, // Màu sắc của leader
                                        //   TextStyleId = textStyleId

                    };
                    mleader.LineWeight = LineWeight.LineWeight009; // Đặt độ dày của đường dẫn
                    mleader.TextStyleId = textstyle;
                    mleader.SetDatabaseDefaults();
                    // mleader.add
                    int leaderIndex = mleader.AddLeader();
                    int lineIndex = mleader.AddLeaderLine(leaderIndex);
                    mleader.AddFirstVertex(leaderIndex, startPoint);
                    mleader.AddLastVertex(leaderIndex, textPoint);
                    // mleader.Annotative = AnnotativeStates.True;
                    //  mleader.TextAttachmentDirection = TextAttachmentDirection.AttachmentVertical;
                    // Căn trái nếu cần
                    mleader.SetArrowSize(leaderIndex, arrowSize * dimscale);
                    mleader.SetArrowSymbolId(leaderIndex, arrowhead);
                    //mleader.SetArrowSymbolId
                    //mleader.TextAttachmentType = TextAttachmentType.AttachmentBottomLine;
                    // Đặt căn chỉnh văn bản sang phải
                    mleader.TextAttachmentType = TextAttachmentType.AttachmentAllLine;
                    mleader.SetTextAttachmentType(TextAttachmentType.AttachmentBottomOfTopLine, LeaderDirectionType.LeftLeader);
                    mleader.SetTextAttachmentType(TextAttachmentType.AttachmentBottomOfTopLine, LeaderDirectionType.RightLeader);
                    // mleader.TextAlignmentType = TextAlignmentType.RightAlignment;
                    // mleader.TextHeight = textHeight;

                    // Thêm vào bản vẽ
                    btr.AppendEntity(mleader);
                    tr.AddNewlyCreatedDBObject(mleader, true);

                    tr.Commit();
                }
            }
        }
        public List<int> listholes(List<Point3d> points)
        {
            List<int> list = new List<int>();

            return list;
        }

        public static void Dimentionholes1(List<Point3d> points, double xstart, double ystart, double fixedStartOffset, out List<Point3d> Lx, out List<Point3d> Ly)
        {
            Lx = new List<Point3d>();
            Ly = new List<Point3d>();

            foreach (Point3d point in points)
            {
                if (point.X == xstart + fixedStartOffset)
                {
                    Ly.Add(point);
                }
                if (point.Y == ystart + fixedStartOffset)
                {
                    Lx.Add(point);
                }
            }
            Lx.Insert(0, new Point3d(xstart, ystart, 0));

            Lx.Sort((p1, p2) => p1.X.CompareTo(p2.X));
            Point3d lastPoint1 = Lx[Lx.Count - 1];
            Lx.Add(new Point3d(lastPoint1.X + fixedStartOffset, ystart, 0));
            //----------------------------------------------------
            Ly.Insert(0, new Point3d(xstart, ystart, 0));
            Ly.Sort((p1, p2) => p1.Y.CompareTo(p2.Y));
            Point3d lastPoint2 = Ly[Ly.Count - 1];
            Ly.Add(new Point3d(xstart, lastPoint2.Y + fixedStartOffset, 0));



        }
        public static void Dimentionholes2(List<Point3d> points, double xstart, double ystart, double fixedStartOffset, out List<Point3d> Lx, out List<Point3d> Ly)
        {
            Lx = new List<Point3d>();
            Ly = new List<Point3d>();

            foreach (Point3d point in points)
            {
                if (point.X <= xstart + fixedStartOffset)
                {
                    Ly.Add(point);
                }
                if (point.Y <= ystart + fixedStartOffset)
                {
                    Lx.Add(point);
                }
            }
        }




        //-------------------------------------------------------


        public static void createDimagu(Point3d[] points, Point3d dimLinePos, ObjectId dimstyle)
        {
            Document doc = acadApp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;


            using (DocumentLock m_DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument())
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord btr = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                    // Danh sách điểm
                    // Point3d[] points = new Point3d[] { p1, p2, p3, p4 };

                    // Tạo từng đoạn dim liên tiếp
                    LineAngularDimension2 dim = new LineAngularDimension2(
                        points[0],         // Điểm trên đường thứ nhất (cạnh 3)
           points[1],         // Giao điểm của hai cạnh
           points[0],
           points[2],
           // Điểm trên đường thứ hai (cạnh 4)
           dimLinePos,   // Điểm định vị text dim (nằm gần cung góc cần đo)
           "",         // Ghi chú (để trống = tự sinh text)
           dimstyle // Dùng dimstyle mặc định
       );
                    dim.LineWeight = LineWeight.LineWeight009; // Chọn độ dày tùy ý
                    btr.AppendEntity(dim);
                    tr.AddNewlyCreatedDBObject(dim, true);


                    tr.Commit();
                }
            }

        }
        public static void createDim(Point3d[] points, Point3d dimLinePos, ObjectId dimstyle, double angle, Point3d p1, Point3d p2, bool checkpos, string Text_left, string Text_right, int posaddtext)
        {
            Document doc = acadApp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            using (DocumentLock m_DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument())
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord btr = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                    // Danh sách điểm
                    // Point3d[] points = new Point3d[] { p1, p2, p3, p4 };

                    // Tạo từng đoạn dim liên tiếp
                    for (int i = 0; i < points.Length - 1; i++)
                    {
                        //dim
                        RotatedDimension dim = new RotatedDimension(angle,
                            points[i],
                            points[i + 1],
                            dimLinePos,
                            "",             // Auto-generated text
                            dimstyle
                        );
                        dim.LineWeight = LineWeight.LineWeight009; // Chọn độ dày tùy ý
                        if (i == posaddtext) { dim.DimensionText = Text_left + "<>" + Text_right; }
                        if (checkpos)
                        {
                            if (i == 0)
                            {
                                // dim.DimensionText = ""; // tắt text mặc định
                                dim.TextPosition = p1; // tự đặt text ra ngoài
                            }
                            if (i == points.Length - 2)
                            {
                                // dim.DimensionText = ""; // tắt text mặc định
                                dim.TextPosition = p2; // tự đặt text ra ngoài
                            }
                        }
                        // dim.Dimtmove = 0;
                        btr.AppendEntity(dim);
                        tr.AddNewlyCreatedDBObject(dim, true);
                    }

                    tr.Commit();
                }
            }

        }

        public static ObjectId CreatePolyline2D(List<Point2d> points, string layerName, LineWeight lineWeight)
        {
            if (points == null || points.Count < 2)
                throw new ArgumentException("Cần ít nhất hai điểm để tạo polyline.");

            Document doc = acadApp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            ObjectId polyId;
            using (DocumentLock m_DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument())
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // Đảm bảo layer tồn tại
                    LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    if (!lt.Has(layerName))
                    {
                        lt.UpgradeOpen();
                        LayerTableRecord newLayer = new LayerTableRecord
                        {
                            Name = layerName
                        };
                        lt.Add(newLayer);
                        tr.AddNewlyCreatedDBObject(newLayer, true);
                    }

                    // Vẽ polyline
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                    Polyline poly = new Polyline();

                    for (int i = 0; i < points.Count; i++)
                    {
                        poly.AddVertexAt(i, points[i], 0, 0, 0);
                    }

                    poly.Closed = points.First() == points.Last(); // Đóng nếu điểm đầu = điểm cuối
                    poly.Layer = layerName;
                    poly.LineWeight = lineWeight;

                    polyId = btr.AppendEntity(poly);
                    tr.AddNewlyCreatedDBObject(poly, true);

                    tr.Commit();
                }
            }

            return polyId;
        }
        private void exportExcel(string outputFolder)
        {

        }



        void RecalculateExcel(string filePath)
        {
            var app = new Excel.Application();
            app.Visible = false;
            var workbook = app.Workbooks.Open(filePath);

            workbook.RefreshAll(); // Refresh external data
            app.Calculate(); // Tính toán lại tất cả công thức

            workbook.Save();
            workbook.Close();
            app.Quit();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
        }
        private void button3_Click(object sender, EventArgs e)
        {

        }
        private void Checkedcb(object sender, EventArgs e)
        {
            CheckBox[] RD1 = new CheckBox[] { CheckCB1, CheckCB2, CheckCB3, CheckCB4 };
            int count = 0;
            foreach (var cb in RD1)
            {
                if (cb.Checked)
                {
                    count++;
                }
            }
            TB_SLBUTTON.Text = count.ToString();
            //  writedata();


        }
        private void Form1_Load(object sender, EventArgs e)
        {
            this.AllowDrop = true;
            this.Text = version;


            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            RegistryHelper.LoadAllSettings(version, this);
            // _ = ExampleAsync(500);
            tb_date.Text = DateTime.Now.ToString("yyyy/MM/dd");
            // tb_ken.Text = "";
            // rdb_4.Checked = true;

            DimStyleImporter.ImportDimStyle(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\RW.dwg", listdimstyle);
            btn_ex.ContextMenuStrip = contextMenuStrip1;
            cb_pdf.ContextMenuStrip = contextMenuStrip2;
            PrinterUtility.ListPlotDevices(cbb_printer);
            PrinterUtility.additemcombobox(cbb_plotstyle, PrinterUtility.PlotStyleList());
            string printer = (string)RegistryHelper.GetSetting(version, "Printer", "");
            var item = cbb_printer.Items.Cast<object>().FirstOrDefault(i => i.ToString().Equals(printer, StringComparison.OrdinalIgnoreCase));
            if (item != null)
            {
                cbb_printer.SelectedItem = item;
            }
            string plotstyle = (string)RegistryHelper.GetSetting(version, "Plotstyle", "");
            var item2 = cbb_plotstyle.Items.Cast<object>().FirstOrDefault(i => i.ToString().Equals(plotstyle, StringComparison.OrdinalIgnoreCase));
            if (item2 != null)
            {
                cbb_plotstyle.SelectedItem = item2;
            }
            // AutoCAD PDF(High Quality Print).pc3
            //ISO_full_bleed_A3_(420.00_x_297.00_MM)
            //monochrome.ctb
            bool found = cbb_printer.Items
    .Cast<object>()
    .Any(item3 => item3.ToString().Equals(cbb_printer.Text, StringComparison.OrdinalIgnoreCase));

            if (found)
            {

                PrinterUtility.additemcombobox(cbb_papersize, PrinterUtility.GetPaperSizes(cbb_printer.Text));
                string papersize = (string)RegistryHelper.GetSetting(version, "Papersize", "");
                var item4 = cbb_papersize.Items.Cast<object>().FirstOrDefault(i => i.ToString().Equals(papersize, StringComparison.OrdinalIgnoreCase));
                if (item4 != null)
                {
                    cbb_papersize.SelectedItem = item4;
                }

            }
            string blockname = (string)RegistryHelper.GetSetting(version, "Block", "");
            // TB_Block.Text = blockname;

            //if (RDB_auto.Checked)
            //{
            //    cbb_Scale.Enabled = false;
            //}
            //else
            //{
            //    cbb_Scale.Enabled = true;
            //}
            textBoxes = new TextBox[] { tb_RSow, tb_RSoh, tb_RSa, tb_RSb, tb_RSc, tb_RSd, tb_VtP, tb_SlB, tb_qlypp, tb_vtb };
            RD1 = new CheckBox[] { CheckCB1, CheckCB2, CheckCB3, CheckCB4 };
            LoadingForm loadingForm = new LoadingForm();
            Task showFormTask = Task.Run(() => loadingForm.ShowDialog());

            button1.PerformClick();
            loadingForm.Invoke(new Action(() => loadingForm.Close()));
            foreach (var cb in RD1)
            {
                cb.CheckedChanged += Checkedcb;
            }
            lb_A.DataBindings.Add("Text", tb_RSa, "Text");
            lb_B1.DataBindings.Add("Text", tb_RSb, "Text");
            lb_B2.DataBindings.Add("Text", tb_RSb, "Text");
            lb_C.DataBindings.Add("Text", tb_RSc, "Text");
            lb_D.DataBindings.Add("Text", tb_RSd, "Text");
            lb_ow.DataBindings.Add("Text", tb_RSow, "Text");
            lb_oh.DataBindings.Add("Text", tb_RSoh, "Text");
            lb_Vtb1.DataBindings.Add("Text", tb_vtb, "Text");
            lb_Vtb2.DataBindings.Add("Text", tb_vtb, "Text");
            lb_sq_a.DataBindings.Add("Text", tb_sq_a, "Text");
            lb_sq_b.DataBindings.Add("Text", tb_sq_b, "Text");
            switch (CBB_TUDIEN.SelectedIndex)
            {
                case 0:
                    tb_tudien_d.Text = "270";
                    tb_tudien_r.Text = "170";
                    tb_tudien_c.Text = "450";
                    break;
                case 1:
                    tb_tudien_d.Text = "300";
                    tb_tudien_r.Text = "160";
                    tb_tudien_c.Text = "500";
                    break;
                case 2:
                    tb_tudien_d.Text = "210";
                    tb_tudien_r.Text = "170";
                    tb_tudien_c.Text = "650";
                    break;
                case 3:
                    tb_tudien_d.Text = "600";
                    tb_tudien_r.Text = "200";
                    tb_tudien_c.Text = "600";
                    break;
            }
        }

        private void label26_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            TextBox tb = sender as TextBox;

            // Cho phép: số, phím điều khiển (Backspace...), dấu trừ đầu tiên, dấu chấm thập phân nếu chưa có
            if (!char.IsControl(e.KeyChar) &&
                !char.IsDigit(e.KeyChar) &&
                (e.KeyChar != '.') &&
                (e.KeyChar != '-'))
            {
                e.Handled = true;
            }

            // Chỉ cho phép một dấu chấm (thập phân)
            if (e.KeyChar == '.' && tb.Text.Contains("."))
            {
                e.Handled = true;
            }

            // Dấu trừ chỉ được đứng đầu
            if (e.KeyChar == '-' && (tb.SelectionStart != 0 || tb.Text.Contains("-")))
            {
                e.Handled = true;
            }
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            TextBox tb = sender as TextBox;

            // Cho phép: số, phím điều khiển (Backspace...), dấu trừ đầu tiên, dấu chấm thập phân nếu chưa có
            if (!char.IsControl(e.KeyChar) &&
                !char.IsDigit(e.KeyChar) &&
                (e.KeyChar != '.') &&
                (e.KeyChar != '-'))
            {
                e.Handled = true;
            }

            // Chỉ cho phép một dấu chấm (thập phân)
            if (e.KeyChar == '.' && tb.Text.Contains("."))
            {
                e.Handled = true;
            }

            // Dấu trừ chỉ được đứng đầu
            if (e.KeyChar == '-' && (tb.SelectionStart != 0 || tb.Text.Contains("-")))
            {
                e.Handled = true;
            }
        }

        private void tb_c_KeyPress(object sender, KeyPressEventArgs e)
        {
            TextBox tb = sender as TextBox;

            // Cho phép: số, phím điều khiển (Backspace...), dấu trừ đầu tiên, dấu chấm thập phân nếu chưa có
            if (!char.IsControl(e.KeyChar) &&
                !char.IsDigit(e.KeyChar) &&
                (e.KeyChar != '.') &&
                (e.KeyChar != '-'))
            {
                e.Handled = true;
            }

            // Chỉ cho phép một dấu chấm (thập phân)
            if (e.KeyChar == '.' && tb.Text.Contains("."))
            {
                e.Handled = true;
            }

            // Dấu trừ chỉ được đứng đầu
            if (e.KeyChar == '-' && (tb.SelectionStart != 0 || tb.Text.Contains("-")))
            {
                e.Handled = true;
            }
        }

        private void tb_j_KeyPress(object sender, KeyPressEventArgs e)
        {
            TextBox tb = sender as TextBox;

            // Cho phép: số, phím điều khiển (Backspace...), dấu trừ đầu tiên, dấu chấm thập phân nếu chưa có
            if (!char.IsControl(e.KeyChar) &&
                !char.IsDigit(e.KeyChar) &&
                (e.KeyChar != '.') &&
                (e.KeyChar != '-'))
            {
                e.Handled = true;
            }

            // Chỉ cho phép một dấu chấm (thập phân)
            if (e.KeyChar == '.' && tb.Text.Contains("."))
            {
                e.Handled = true;
            }

            // Dấu trừ chỉ được đứng đầu
            if (e.KeyChar == '-' && (tb.SelectionStart != 0 || tb.Text.Contains("-")))
            {
                e.Handled = true;
            }
        }

        private void tb_a_KeyPress(object sender, KeyPressEventArgs e)
        {
            TextBox tb = sender as TextBox;

            // Cho phép: số, phím điều khiển (Backspace...), dấu trừ đầu tiên, dấu chấm thập phân nếu chưa có
            if (!char.IsControl(e.KeyChar) &&
                !char.IsDigit(e.KeyChar) &&
                (e.KeyChar != '.') &&
                (e.KeyChar != '-'))
            {
                e.Handled = true;
            }

            // Chỉ cho phép một dấu chấm (thập phân)
            if (e.KeyChar == '.' && tb.Text.Contains("."))
            {
                e.Handled = true;
            }

            // Dấu trừ chỉ được đứng đầu
            if (e.KeyChar == '-' && (tb.SelectionStart != 0 || tb.Text.Contains("-")))
            {
                e.Handled = true;
            }
        }

        private void tb_b_KeyPress(object sender, KeyPressEventArgs e)
        {
            TextBox tb = sender as TextBox;

            // Cho phép: số, phím điều khiển (Backspace...), dấu trừ đầu tiên, dấu chấm thập phân nếu chưa có
            if (!char.IsControl(e.KeyChar) &&
                !char.IsDigit(e.KeyChar) &&
                (e.KeyChar != '.') &&
                (e.KeyChar != '-'))
            {
                e.Handled = true;
            }

            // Chỉ cho phép một dấu chấm (thập phân)
            if (e.KeyChar == '.' && tb.Text.Contains("."))
            {
                e.Handled = true;
            }

            // Dấu trừ chỉ được đứng đầu
            if (e.KeyChar == '-' && (tb.SelectionStart != 0 || tb.Text.Contains("-")))
            {
                e.Handled = true;
            }
        }

        private void tb_d_TextChanged(object sender, EventArgs e)
        {

        }

        private void tb_d_KeyPress(object sender, KeyPressEventArgs e)
        {
            TextBox tb = sender as TextBox;

            // Cho phép: số, phím điều khiển (Backspace...), dấu trừ đầu tiên, dấu chấm thập phân nếu chưa có
            if (!char.IsControl(e.KeyChar) &&
                !char.IsDigit(e.KeyChar) &&
                (e.KeyChar != '.') &&
                (e.KeyChar != '-'))
            {
                e.Handled = true;
            }

            // Chỉ cho phép một dấu chấm (thập phân)
            if (e.KeyChar == '.' && tb.Text.Contains("."))
            {
                e.Handled = true;
            }

            // Dấu trừ chỉ được đứng đầu
            if (e.KeyChar == '-' && (tb.SelectionStart != 0 || tb.Text.Contains("-")))
            {
                e.Handled = true;
            }
        }

        private void tb_e_KeyPress(object sender, KeyPressEventArgs e)
        {
            TextBox tb = sender as TextBox;

            // Cho phép: số, phím điều khiển (Backspace...), dấu trừ đầu tiên, dấu chấm thập phân nếu chưa có
            if (!char.IsControl(e.KeyChar) &&
                !char.IsDigit(e.KeyChar) &&
                (e.KeyChar != '.') &&
                (e.KeyChar != '-'))
            {
                e.Handled = true;
            }

            // Chỉ cho phép một dấu chấm (thập phân)
            if (e.KeyChar == '.' && tb.Text.Contains("."))
            {
                e.Handled = true;
            }

            // Dấu trừ chỉ được đứng đầu
            if (e.KeyChar == '-' && (tb.SelectionStart != 0 || tb.Text.Contains("-")))
            {
                e.Handled = true;
            }
        }

        private void tb_f_KeyPress(object sender, KeyPressEventArgs e)
        {
            TextBox tb = sender as TextBox;

            // Cho phép: số, phím điều khiển (Backspace...), dấu trừ đầu tiên, dấu chấm thập phân nếu chưa có
            if (!char.IsControl(e.KeyChar) &&
                !char.IsDigit(e.KeyChar) &&
                (e.KeyChar != '.') &&
                (e.KeyChar != '-'))
            {
                e.Handled = true;
            }

            // Chỉ cho phép một dấu chấm (thập phân)
            if (e.KeyChar == '.' && tb.Text.Contains("."))
            {
                e.Handled = true;
            }

            // Dấu trừ chỉ được đứng đầu
            if (e.KeyChar == '-' && (tb.SelectionStart != 0 || tb.Text.Contains("-")))
            {
                e.Handled = true;
            }
        }

        private void tb_g_TextChanged(object sender, EventArgs e)
        {

        }

        private void tb_g_KeyPress(object sender, KeyPressEventArgs e)
        {
            TextBox tb = sender as TextBox;

            // Cho phép: số, phím điều khiển (Backspace...), dấu trừ đầu tiên, dấu chấm thập phân nếu chưa có
            if (!char.IsControl(e.KeyChar) &&
                !char.IsDigit(e.KeyChar) &&
                (e.KeyChar != '.') &&
                (e.KeyChar != '-'))
            {
                e.Handled = true;
            }

            // Chỉ cho phép một dấu chấm (thập phân)
            if (e.KeyChar == '.' && tb.Text.Contains("."))
            {
                e.Handled = true;
            }

            // Dấu trừ chỉ được đứng đầu
            if (e.KeyChar == '-' && (tb.SelectionStart != 0 || tb.Text.Contains("-")))
            {
                e.Handled = true;
            }
        }

        private void tb_h_KeyPress(object sender, KeyPressEventArgs e)
        {
            TextBox tb = sender as TextBox;

            // Cho phép: số, phím điều khiển (Backspace...), dấu trừ đầu tiên, dấu chấm thập phân nếu chưa có
            if (!char.IsControl(e.KeyChar) &&
                !char.IsDigit(e.KeyChar) &&
                (e.KeyChar != '.') &&
                (e.KeyChar != '-'))
            {
                e.Handled = true;
            }

            // Chỉ cho phép một dấu chấm (thập phân)
            if (e.KeyChar == '.' && tb.Text.Contains("."))
            {
                e.Handled = true;
            }

            // Dấu trừ chỉ được đứng đầu
            if (e.KeyChar == '-' && (tb.SelectionStart != 0 || tb.Text.Contains("-")))
            {
                e.Handled = true;
            }
        }

        private void tb_i_KeyPress(object sender, KeyPressEventArgs e)
        {
            TextBox tb = sender as TextBox;

            // Cho phép: số, phím điều khiển (Backspace...), dấu trừ đầu tiên, dấu chấm thập phân nếu chưa có
            if (!char.IsControl(e.KeyChar) &&
                !char.IsDigit(e.KeyChar) &&
                (e.KeyChar != '.') &&
                (e.KeyChar != '-'))
            {
                e.Handled = true;
            }

            // Chỉ cho phép một dấu chấm (thập phân)
            if (e.KeyChar == '.' && tb.Text.Contains("."))
            {
                e.Handled = true;
            }

            // Dấu trừ chỉ được đứng đầu
            if (e.KeyChar == '-' && (tb.SelectionStart != 0 || tb.Text.Contains("-")))
            {
                e.Handled = true;
            }
        }

        private void tb_k_KeyPress(object sender, KeyPressEventArgs e)
        {
            TextBox tb = sender as TextBox;

            // Cho phép: số, phím điều khiển (Backspace...), dấu trừ đầu tiên, dấu chấm thập phân nếu chưa có
            if (!char.IsControl(e.KeyChar) &&
                !char.IsDigit(e.KeyChar) &&
                (e.KeyChar != '.') &&
                (e.KeyChar != '-'))
            {
                e.Handled = true;
            }

            // Chỉ cho phép một dấu chấm (thập phân)
            if (e.KeyChar == '.' && tb.Text.Contains("."))
            {
                e.Handled = true;
            }

            // Dấu trừ chỉ được đứng đầu
            if (e.KeyChar == '-' && (tb.SelectionStart != 0 || tb.Text.Contains("-")))
            {
                e.Handled = true;
            }
        }

        private void tb_deg_KeyPress(object sender, KeyPressEventArgs e)
        {
            TextBox tb = sender as TextBox;

            // Cho phép: số, phím điều khiển (Backspace...), dấu trừ đầu tiên, dấu chấm thập phân nếu chưa có
            if (!char.IsControl(e.KeyChar) &&
                !char.IsDigit(e.KeyChar) &&
                (e.KeyChar != '.') &&
                (e.KeyChar != '-'))
            {
                e.Handled = true;
            }

            // Chỉ cho phép một dấu chấm (thập phân)
            if (e.KeyChar == '.' && tb.Text.Contains("."))
            {
                e.Handled = true;
            }

            // Dấu trừ chỉ được đứng đầu
            if (e.KeyChar == '-' && (tb.SelectionStart != 0 || tb.Text.Contains("-")))
            {
                e.Handled = true;
            }
        }

        private void tb_caotong_KeyPress(object sender, KeyPressEventArgs e)
        {


        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            RegistryHelper.SaveAllSettings(version, this);
            RegistryHelper.SaveSetting(version, "Printer", cbb_printer.Text);
            RegistryHelper.SaveSetting(version, "Plotstyle", cbb_plotstyle.Text);
            RegistryHelper.SaveSetting(version, "Papersize", cbb_papersize.Text);

        }

        private void btn_parent_Click(object sender, EventArgs e)
        {

        }

        private void btn_left_Click(object sender, EventArgs e)
        {
        }

        private void btn_right_Click(object sender, EventArgs e)
        {

        }

        private void btn_top_Click(object sender, EventArgs e)
        {
        }

        private void btn_parent_MouseHover(object sender, EventArgs e)
        {

        }

        private void btn_top_MouseHover(object sender, EventArgs e)
        {

        }

        private void btn_right_MouseHover(object sender, EventArgs e)
        {

        }

        private void btn_left_MouseHover(object sender, EventArgs e)
        {

        }
        private void btn_back_MouseHover(object sender, EventArgs e)
        {

        }
        private void button4_Click(object sender, EventArgs e)
        {

        }



        private void btn_autoholes_Click(object sender, EventArgs e)
        {
            var doc = acadApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;
            // int PP = dtb_gr1.RowCount;
            //int p = 1;
            //  dtb_gr1.Rows.RemoveAt(PP-1);
            PromptPointResult pointResult = ed.GetPoint("Chọn tâm hình chiếu bằng: ");
            AutoHoles(pointResult.Value.X, pointResult.Value.Y);
        }

        private void btn_back_Click(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {

        }

        private void btn_update_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click_1(object sender, EventArgs e)
        {

        }

        private void rbt_lapsau_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void bt_updateinput_Click(object sender, EventArgs e)
        {


        }

        private void ExportToDwg(string outputPath, List<bool> exportOptions)
        {
            var doc = acadApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;

            using (DocumentLock docLock = doc.LockDocument())
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // Save the DWG file  
                    if (exportOptions[0]) // DWG export  
                    {
                        db.SaveAs(outputPath, DwgVersion.Current);
                    }

                    // Additional export logic for DXF, PDF, etc., can be added here  
                    // Example:  
                    if (exportOptions[1]) // DXF export  
                    {
                        string dxfPath = Path.ChangeExtension(outputPath, ".dxf");
                        db.DxfOut(dxfPath, 16, true);
                    }

                    tr.Commit();
                }
            }
        }

        private void openFolderToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void btn_folder_Click(object sender, EventArgs e)
        {

        }

        private void tb_per_MouseHover(object sender, EventArgs e)
        {

        }

        private void button3_Click_2(object sender, EventArgs e)
        {
            Form2 form2 = new Form2(this); // Truyền chính instance này
            form2.Show();
        }

        //private void tb_todayrw_Click(object sender, EventArgs e)
        //{
        //    tb_daterw.Text = DateTime.Now.ToString("yyyy/MM/dd");
        //}

        private void label36_Click(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {
            //try
            //{
            //    double.TryParse(tb_OW.Text, out OW);
            //    double.TryParse(tb_OH.Text, out OH);
            //    string filePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\TRATHONGSO.xlsM";
            //    // Khởi tạo
            //    var service = new ExcelLookupService(filePath);
            // //   MessageBox.Show("ddã vào đây1-1");
            //    // Sử dụng và lấy kết quả
            //    var result = service.Val(OW, OH);
            //   // MessageBox.Show("ddã vào đây1-2");
            //    // Kiểm tra kết quả
            //    if (result.IsValid)
            //    {
            //        tb_RSow.Text = OW.ToString();
            //        tb_RSoh.Text = OH.ToString();
            //        A= result.AVal;
            //        tb_RSa.Text = A.ToString();
            //        B = result.BVal;
            //        tb_RSb.Text=B.ToString();
            //        C = result.CVal;
            //        tb_RSc.Text = C.ToString();
            //        D = result.DVal.ToString();
            //        tb_RSd.Text = D;
            //        motor = result.Value1;
            //        tb_VtP.Text = motor;
            //        gearbox = result.Value2;
            //        tb_SlB.Text = gearbox;
            //        qlypp = result.AdditionalInfo;
            //        tb_qlypp.Text = qlypp;
            //        spacing = result.Spacing;
            //        tb_scs.Text = spacing.ToString();


            //    }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("Lỗi: " + ex.Message);
            //}
        }
        public class LoadingForm : Form
        {
            private Label label;
            private ProgressBar progressBar;

            public LoadingForm()
            {
                this.Text = "Đang xử lý...";
                this.Size = new Size(250, 140);
                this.StartPosition = FormStartPosition.CenterScreen;
                this.FormBorderStyle = FormBorderStyle.FixedDialog;
                this.ControlBox = false;
                this.TopMost = true;

                label = new Label()
                {
                    Text = "Vui lòng chờ...",
                    Dock = DockStyle.Top,
                    TextAlign = ContentAlignment.MiddleCenter,
                    Height = 50
                };

                progressBar = new ProgressBar()
                {
                    Style = ProgressBarStyle.Marquee,
                    Dock = DockStyle.Fill,
                    MarqueeAnimationSpeed = 50
                };

                this.Controls.Add(progressBar);
                this.Controls.Add(label);
            }
        }

        private void button8_Click(object sender, EventArgs e)
        { var result = MessageBox.Show("Update dữ liệu theo OW-OH?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                LoadingForm loadingForm = new LoadingForm();
                Task showFormTask = Task.Run(() => loadingForm.ShowDialog());             
               
                button1.PerformClick();
                loadingForm.Invoke(new Action(() => loadingForm.Close()));
            }
            NativeMethods.SetForegroundWindow(acadHwnd);
            //  WriteData(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\Bang tra.xlsx");
            var doc = acadApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;
            // loadingForm.Invoke(new Action(() => loadingForm.Close()));
             pointResult = ed.GetPoint("Chọn vị trí để chèn Block Reference: ");

            //----------------------------------
            string targetName = "";
            //if (RDB_manual.Checked)
            //{
            targetName = listdimstyle[cbb_Scale.SelectedIndex];
            scale = double.Parse(cbb_Scale.Text.Split('/')[1]) / 50;
            //    //  MessageBox.Show(scale.ToString());
            //}
            //else
            //{
            //    if (Length < 7800 && Height < 4000)
            //    {
            //        if (true)
            //        {
            //            scale = 1.0; // Nếu chiều cao trung bình nhỏ hơn 1780, không cần scale
            //            targetName = "WKV_20";
            //        }
            //        else
            //        {
            //            scale = 1.5; // Nếu chiều cao trung bình lớn hơn hoặc bằng 1780, áp dụng scale
            //            targetName = "WKV_30";
            //        }
            //        // scale = 1.0; // Nếu tổng chiều dài và chiều cao nhỏ hơn 7800 và 4000, không cần scale
            //    }
            //    else
            //    {
            //        scale = 1.5; // Nếu tổng chiều dài hoặc chiều cao lớn hơn, áp dụng scale
            //             targetName = "WKV_30";
            //    }
            //}
            using (DocumentLock m_DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument())
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    DimStyleTable dimTable = (DimStyleTable)tr.GetObject(db.DimStyleTableId, OpenMode.ForRead);

                    if (dimTable.Has(targetName))
                    {

                        dimstyle = dimTable[targetName];
                        // ed.WriteMessage($"\ntest '{targetName}'.");
                        // DimStyleTableRecord dimStyle = (DimStyleTableRecord)tr.GetObject(dimstyle, OpenMode.ForRead);

                        //  ed.WriteMessage($"\nTìm thấy DimStyle: {dimstyle.}, ID: {dimstyle1}");
                    }
                    else
                    {
                        ed.WriteMessage($"\nKhông tìm thấy DimStyle có tên '{targetName}'.");
                    }

                    tr.Commit();
                }
            }
            //List<string> A3_KT_RW = new List<string>() { };
            //List<string> Bang_TSRW = new List<string>() { };
            //List<string> Bang_KLRW = new List<string>() { };
            //-----------------------------------------
            TextBox[] textBoxes = new TextBox[] { tb_ong1, tb_ong2, tb_ong3 };
            string t = "";
            foreach (var tb in textBoxes)
            {
                //switch (tb.Text)
                //{
                //    case "Ø48.6":
                //        t = "2.4";
                //        break;
                //    case "Ø42.7":
                //        t = "2.3";
                //        break;
                //    case "Ø34.0":
                //        t = "3.2";
                //        break;
                //    case "Ø27.2":
                //        t = "1.9";
                //        break;
                //    case "Ø21.7":
                //        t = "1.9";
                //        break;


                //}
            }
            // MessageBox.Show(t_ong1 + "-" + t_ong2 + "-" + t_ong3);
            A3_KT_RW = new List<string>() { tb_tencongtrinh.Text, tb_diadiem.Text, "ロールウエイ外観図", tb_maso.Text, cbb_nguoitao.Text, tb_date.Text, $"1/{Math.Round(scale * 50, 0)}" };
            Bang_TSRW = tb_ong2.Text == "Ø0.0" ?

            new List<string>()
            {
                cbb_loairw.Text,
                tb_vlkaten.Text.Split(' ')[0],
                "ポリ塩化ビニール",
                TB_maukaten.Text.Split(' ')[1],
                (tb_vlkaten.SelectedIndex==0? "不燃認定番号" : "防炎登録番号"),
                (tb_vlkaten.SelectedIndex==0? "NM-5361" : "B1130062"),
                (cb_cst2.Checked||cb_cst3.Checked? "有り" : "無し"),
                tb_ong3.Text.Replace("Ø","%%C")+"-"+"t="+t_ong3,
                tb_ong1.Text.Replace("Ø","%%C")+"-"+"t="+t_ong1,
                t_thep,
                tb_mmotor.Text,
                tb_tsmotor.Text,
                tb_slmotor.Text+"台",
                "3相 AC200/50Hz AC200/220/60Hz",
                tb_v.Text+ " m/s"
            } :
            new List<string>(){
                cbb_loairw.Text,
                tb_vlkaten.Text.Split(' ')[0],
                "ポリ塩化ビニール",
                TB_maukaten.Text.Split(' ')[1],
                (tb_vlkaten.SelectedIndex == 0 ? "不燃認定番号" : "防炎登録番号"),
                (tb_vlkaten.SelectedIndex == 0 ? "NM-5361" : "B1130062"),
                (cb_cst2.Checked || cb_cst3.Checked ? "有り" : "無し"),
                tb_ong3.Text.Replace("Ø", "%%C") + "-" + "t=" + t_ong3,
                tb_ong2.Text.Replace("Ø", "%%C") + "-" + "t=" + t_ong2,
                tb_ong1.Text.Replace("Ø", "%%C") + "-" + "t=" + t_ong1,
                t_thep,
                tb_mmotor.Text,
                tb_tsmotor.Text,
                tb_slmotor.Text + "台",
                "3相 AC200/50Hz AC200/220/60Hz",
                tb_v.Text+ " m/s"
            };
            // MessageBox.Show(Bang_TSRW.Count.ToString());
            WKV_Array = new List<string>[] { A3_KT_RW, Bang_TSRW, Bang_KLRW };
            int ind = 0;
            List<string> listblockname_ins = new List<string>();
            if (rdb_kl1.Checked == true)
            {
                listblockname_ins = tb_ong2.Text == "Ø0.0" ? listblockname2 : listblockname4;
                // listblockname_ins = listblockname2;
            }
            else
            {
                listblockname_ins = tb_ong2.Text == "Ø0.0" ? listblockname1 : listblockname3;
                // listblockname_ins = listblockname1 ;
            }
            double X_add = 0.0;
            double y_add = 0.0;
            // MessageBox.Show(listblockname_ins.Count.ToString());
            foreach (string blockname in listblockname_ins)
            {
                switch (ind)
                {
                    case 0:
                        X_add = 0.0;
                        y_add = 0.0;
                        break;
                    case 1:
                        X_add = -101;
                        y_add = 0;
                        break;
                    case 2:
                        X_add = 0.0;
                        y_add = 0.0;
                        break;

                }
                Insertdynamicblock.InsertBlock(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\RW.dwg", blockname, new Point3d(pointResult.Value.X + X_add, pointResult.Value.Y + y_add, 0), WKV_Array[ind], scale, 0);
                ind++;
            }

            //-------------------------------- tạm thời để test xóa bỏ để mở code dưới sau này
            double x = pointResult.Value.X - 12000 * scale;
            double y = pointResult.Value.Y + 6300 * scale;
             p0 = new Point3d(x, y, 0);
            //Point3d p1 = new Point3d(x1, y, 0);
            //Point3d p2 = new Point3d(x2, y, 0);
            mattruoc(p0, ed);


            double.TryParse(tb_SlB.Text, out double slb);
            double.TryParse(tb_RSa.Text, out double a);
            double.TryParse(tb_RSb.Text, out double b);
            double.TryParse(tb_RSow.Text, out double ow);
            double.TryParse(tb_RSoh.Text, out double oh);
            double.TryParse(tb_RSd.Text, out double d);
            double.TryParse(tb_vtb.Text, out double vtb);
            double.TryParse(tb_VtP.Text, out double vtp);
            double.TryParse(tb_qlypp.Text, out double P);
            double.TryParse(tb_scs.Text, out double scs);
            double.TryParse(tb_slmotor.Text, out double slmt);
            double.TryParse(tb_tudien_d.Text, out double DT);
            double.TryParse(tb_tudien_r.Text, out double DT2);
            double.TryParse(tb_tudien_c.Text, out double HT);
            double.TryParse(tb_y_tudien.Text, out double y_tudien);

            CHIEUBANG(new Point3d(p0.X - ow / 2 - b, p0.Y - 3000, 0), ed);
            double scale2 = scale * 50;
            double kc = (ow - 2.0 * vtb) / (slb - 1.0);
            List<Point3d> p1 = new List<Point3d> { };
            for (int i = 0; i < (slb); i++)
            {

                p1.Add(new Point3d(x - (ow - 2 * vtb) / 2 + (i) * kc, y, 0));

            }
            Belt(p1, ed, slb);

            List<Point3d> p2_1 = new List<Point3d> { };
            List<Point3d> p2_2 = new List<Point3d> { };
            List<Point3d> p2_3 = new List<Point3d> { };
            double y_pipe = y + 40;
            for (int i = 0; i < (P); i++)
            {
                //ống ngoài
                if (i > 6)
                {
                    y_pipe = y + 40 + 6 * 600 + (i - 6) * 900;
                }
                else
                {
                    y_pipe = y + 40 + i * 600;
                }
                // int pipe_distance = i > 6 ? 900 : 600;


                p2_1.Add(new Point3d(x - ow / 2.0, y_pipe, 0));
                p2_2.Add(new Point3d(x + ow / 2.0 - vtb + 25, y_pipe, 0));


                //---------------------------------------------------------------------------
                //-----------------------------------------------------ống trong
                for (int j = 0; j < slb - 1; j++)
                {
                    p2_3.Add(new Point3d(x - ow / 2.0 + vtb + 25 + j * (kc), y_pipe, 0));


                }
                //---------------------------------------------------------------------------cửa sổ


            }
            Tube(p2_1, ed, p2_1.Count, vtb - 25);
            Tube(p2_2, ed, p2_2.Count, vtb - 25);
            Tube(p2_3, ed, p2_3.Count, kc - 50);


            List<Point3d> p3 = new List<Point3d> { };
            List<Point3d> p3_1 = new List<Point3d> { };
            List<Point3d> p3_2 = new List<Point3d> { };
            List<Point3d> p3_4 = new List<Point3d> { };
            List<Point3d> p3_5 = new List<Point3d> { };
            List<Point3d> p3_6 = new List<Point3d> { };
            if (cb_cst2.Checked == true)
            {

                for (int j = 0; j < slb - 1; j++)
                {
                    if (scs == 1)
                    {

                        p3.Add(new Point3d(x - ow / 2.0 + vtb + 110 + j * (kc) + (kc - 2 * 110) / 2.0, y + 800, 0));


                    }
                    else if (scs == 2)
                    {
                        p3_1.Add(new Point3d(x - ow / 2.0 + vtb + 110 + j * (kc) + (kc / 2 - 110 - 125) / 2.0, y + 800, 0));

                        p3_2.Add(new Point3d(x - ow / 2.0 + vtb + 110 + j * (kc) + (kc / 2 - 110 - 125) + 250 + (kc / 2 - 110 - 125) / 2.0, y + 800, 0));

                    }
                }

            }
            Window(p3, ed, p3.Count, (kc - 2 * 110) / 2.0);
            Window(p3_1, ed, p3_1.Count, (kc / 2 - 110 - 125) / 2.0);
            Window(p3_2, ed, p3_2.Count, (kc / 2 - 110 - 125) / 2.0);


            if (cb_cst3.Checked == true)
            {
                for (int j = 0; j < slb - 1; j++)
                {
                    if (scs == 1)
                    {

                        p3_4.Add(new Point3d(x - ow / 2.0 + vtb + 110 + j * (kc) + (kc - 2 * 110) / 2.0, y + 800 + 600, 0));



                    }
                    else if (scs == 2)
                    {
                        p3_5.Add(new Point3d(x - ow / 2.0 + vtb + 110 + j * (kc) + (kc / 2 - 110 - 125) / 2.0, y + 800 + 600, 0));

                        p3_6.Add(new Point3d(x - ow / 2.0 + vtb + 110 + j * (kc) + (kc / 2 - 110 - 125) + 250 + (kc / 2 - 110 - 125) / 2.0, y + 800 + 600, 0));

                    }
                }
            }
            Window(p3_4, ed, p3_4.Count, (kc - 2 * 110) / 2.0);
            Window(p3_5, ed, p3_5.Count, (kc / 2 - 110 - 125) / 2.0);
            Window(p3_6, ed, p3_6.Count, (kc / 2 - 110 - 125) / 2.0);
            if (rad_B.Checked || slmt == 2)
            {
                Insertdynamicblock.InsertDynamicBlock(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\RW.dwg", "Motor2", new Point3d(x + ow / 2 - 2 * vtb / 3, y + oh + d, 0), new Dictionary<string, double> { { "A-vtp", a - vtp } }/*, cbb_test*/, dimstyle, 1.0);
                // Insertdynamicblock.InsertBlock(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\RW.dwg", "NT",
                //      new Point3d(x-ow/2+1200, y + oh+d-a+vtp, 0), null, 1.0, 0);
                AddLeaderNote(new Point3d(x + ow / 2 - 2 * vtb / 3 + 120, y + oh + d - 20, 0), new Point3d(x + ow / 2 - 2 * vtb / 3 + 6 * scale2 + 120, y + oh + d + 6 * scale2, 0), "モーター", dimstyle_text);
            }
            if (rab_A.Checked || slmt == 2)
            {
                Insertdynamicblock.InsertDynamicBlock(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\RW.dwg", "Motor", new Point3d(x - ow / 2 + 2 * vtb / 3, y + oh + d, 0), new Dictionary<string, double> { { "A-vtp", a - vtp } }/*, cbb_test*/, dimstyle, 1.0);
                AddLeaderNote(new Point3d(x - ow / 2 + 2 * vtb / 3, y + oh + d - 20, 0), new Point3d(x - ow / 2 + 2 * vtb / 3 + 6 * scale2, y + oh + d + 6 * scale2, 0), "モーター", dimstyle_text);
                // Insertdynamicblock.InsertBlock(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\RW.dwg", "Motor2",
                //     new Point3d(x+ow/2 - 670, y + oh + d, 0), null, 1.0, 0);

            }
            if (/*slmt == 2 ||*/ rad_B.Checked)
            {
                Insertdynamicblock.InsertBlock(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\RW.dwg", "NT", slmt == 2 ? new Point3d(x - ow / 2 + vtb + 350+50, y + oh + d - a + vtp, 0) : new Point3d(x - ow / 2 + vtb / 2 + 100, y + oh + d - a + vtp, 0), null, 1.0, 0);
                AddLeaderNote(slmt == 2 ? new Point3d(x - ow / 2 + vtb + 350 + 100 - 30+50, y + oh + d - a + vtp + 29, 0) : new Point3d(x - ow / 2 + vtb / 2 + 100 - 100 + 30, y + oh + d - a + vtp + 29, 0), slmt == 2 ? new Point3d(x - ow / 2 + vtb + 350 - 100 + 7 * scale2 + 30+50, y + oh + d + 6 * scale2, 0) : new Point3d(x - ow / 2 + vtb / 2 - 100 + 100 + 7 * scale2 + 30, y + oh + d + 6 * scale2, 0), "リミッター", dimstyle_text);
            }

            //if (rad_B.Checked /*&& slmt != 2*/)
            //{
            //    //Insertdynamicblock.InsertDynamicBlock(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\RW.dwg", "Motor", new Point3d(x - ow / 2 + 670, y + oh + d, 0), new Dictionary<string, double> { { "A-vtp", a - vtp } }/*, cbb_test*/, dimstyle, 1.0);
            //    Insertdynamicblock.InsertBlock(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\RW.dwg", "NT", new Point3d(x - ow / 2 + vtb / 2 - 100, y + oh + d - a + vtp, 0), null, 1.0, 0);
            //    AddLeaderNote(new Point3d(x - ow / 2 + vtb / 2 - 100 + 100, y + oh + d - a + vtp + 29, 0), new Point3d(x - ow / 2 + vtb / 2 - 100 + 100 + 6 * scale2, y + oh + d + 6 * scale2, 0), "リミッター", dimstyle_text);
            //}
            if (/*slmt == 2 ||*/ rab_A.Checked)
            {
                Insertdynamicblock.InsertBlock(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\RW.dwg", "NT", slmt == 2 ? new Point3d(x + ow / 2 - vtb - 350-50, y + oh + d - a + vtp, 0) : new Point3d(x + ow / 2 - vtb / 2 - 100, y + oh + d - a + vtp, 0), null, 1.0, 0);
                AddLeaderNote(slmt == 2 ? new Point3d(x + ow / 2 - vtb - 350 + 100 - 30-50, y + oh + d - a + vtp + 29, 0) : new Point3d(x + ow / 2 - vtb / 2 - 100 + 100 - 30, y + oh + d - a + vtp + 29, 0), slmt == 2 ? new Point3d(x + ow / 2 - vtb - 350 + 100 + 6 * scale2 - 30-50, y + oh + d + 6 * scale2, 0) : new Point3d(x + ow / 2 - vtb / 2 - 100 + 100 + 6 * scale2 - 30, y + oh + d + 6 * scale2, 0), "リミッター", dimstyle_text);
            }
            ///----------------------------------Tạo dim hình chiếu chính
            List<Point3d> Pb1 = new List<Point3d> { new Point3d(x - ow / 2, y, 0), new Point3d(x - ow / 2 + vtb, y - 20, 0) };
            List<Point3d> Pb2 = new List<Point3d> { new Point3d(x + ow / 2, y, 0), new Point3d(x + ow / 2 - vtb, y - 20, 0) };

            //MessageBox.Show(scale2.ToString());
            createDim(Pb1.ToArray(), new Point3d(x - ow / 2, y - 12 * scale2, 0), dimstyle_text, 0, new Point3d(x - ow / 2 - 200, y - 12 * scale2, 0), new Point3d(x - ow / 2 + vtb + 200, y - 12 * scale2, 0), false, "", "", 0);
            createDim(Pb2.ToArray(), new Point3d(x + ow / 2, y - 12 * scale2, 0), dimstyle_text, 0, new Point3d(x + ow / 2 - vtb - 200, y - 12 * scale2, 0), new Point3d(x + ow / 2 + vtb + 200, y - 12 * scale2, 0), false, "", "", 0);
            List<Point3d> Pb3 = new List<Point3d> { new Point3d(x - ow / 2 - b, y, 0), new Point3d(x - ow / 2, y, 0), new Point3d(x + ow / 2, y, 0), new Point3d(x + ow / 2 + b, y, 0) };
            createDim(Pb3.ToArray(), new Point3d(x - ow / 2, y - 18 * scale2, 0), dimstyle_text, 0, new Point3d(x - ow / 2 - b - 200, y - 18 * scale2, 0), new Point3d(x + ow / 2 + b + 200, y - 18 * scale2, 0), true, "OW:", "", 1);
            ///----------------------------------
            AddLeaderNote(new Point3d(x + 50, y + oh + d, 0), new Point3d(x + 50 + 6 * scale2, y + oh + d + 6 * scale2, 0), "上部フレーム", dimstyle_text);
            AddLeaderNote(new Point3d(x + ow / 2 + b, y + oh + d - 1300, 0), new Point3d(x + ow / 2 + b + 6 * scale2, y + oh + d - 1300 + 6 * scale2, 0), "サイドフレーム", dimstyle_text);
            List<Point3d> p_cabinet = new List<Point3d>() { rab_A.Checked ? new Point3d(x + ow / 2 + b + 935, y + y_tudien, 0) : new Point3d(x - ow / 2 - b - 935, y + y_tudien, 0) };
            if (CBB_TUDIEN.SelectedIndex != 2 && CBB_TUDIEN.SelectedIndex != 3)

            {
                ElectricalCabinet(p_cabinet, ed, 1.0, new List<double> { DT / 2, DT / 2, HT }, "Tutieuchuan_CD");

            }
            else
            {
                string blocknametudien = CBB_TUDIEN.SelectedIndex == 2 ? "Tu_RL1_CD" : "Tu_RL2_CD";

                Insertdynamicblock.InsertBlock(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\RW.dwg", blocknametudien, p_cabinet[0], null, 1.0, 0);
            }
            if (CBB_TUDIEN.SelectedIndex == 2) { DT = 210; HT = 650; }
            if (CBB_TUDIEN.SelectedIndex == 3) { DT = 600; HT = 600; }
            List<Point3d> PD1;
            List<Point3d> PD3;
            Point3d PD2 = new Point3d();
            Point3d PD4 = new Point3d();
            if (rab_A.Checked)
            {
                PD1 = new List<Point3d>() { new Point3d(x + ow / 2 + 1040, y, 0), new Point3d(x + ow / 2 + b + 935 + DT / 2, y + y_tudien, 0), new Point3d(x + ow / 2 + b + 935 + DT / 2, y + y_tudien + HT, 0) };
                PD2 = new Point3d(x + ow / 2 + b + 935 + DT / 2 + 9 * scale2, y + y_tudien, 0);
                PD3 = new List<Point3d>() { new Point3d(x + ow / 2 + b + 935 + DT / 2, y + y_tudien, 0), new Point3d(x + ow / 2 + b + 935 - DT / 2, y + y_tudien, 0) };
                PD4 = new Point3d(x + ow / 2 + b + 935, y + y_tudien + HT, 0);
            }
            else
            {
                PD1 = new List<Point3d>() { new Point3d(x - ow / 2 - 1040, y, 0), new Point3d(x - ow / 2 - b - 935 - DT / 2, y + y_tudien, 0), new Point3d(x - ow / 2 - b - 935 - DT / 2, y + y_tudien + HT, 0) };
                PD2 = new Point3d(x - ow / 2 - b - 935 - DT / 2 - 9 * scale2, y + y_tudien, 0);
                PD3 = new List<Point3d>() { new Point3d(x - ow / 2 - b - 935 - DT / 2, y + y_tudien, 0), new Point3d(x - ow / 2 - b - 935 + DT / 2, y + y_tudien, 0) };
                PD4 = new Point3d(x - ow / 2 - b - 935, y + y_tudien + HT, 0);
            }
            createDim(PD1.ToArray(), PD2, dimstyle_text, 0.5 * Math.PI, PD2, PD2, false, "", "", 0);
            createDim(PD3.ToArray(), new Point3d(x - ow / 2 - b - 935 + DT / 2, y + y_tudien - 6 * scale2, 0), dimstyle_text, 0, PD2, PD2, false, "", "", 0);
            AddLeaderNote(PD4, new Point3d(PD4.X + scale2 * 6, PD4.Y + scale2 * 6, 0), "制御箱", dimstyle_text);
            //---------------------------------------------------------
            CHIEUCANH(new Point3d(p0.X + ow / 2 + b + 4000 * scale, p0.Y, 0), ed);

        }
        //------------------------------------------------HÀM CHIẾU BẰNG
        private void CHIEUBANG(Point3d point1, Editor ed)
        {
            double.TryParse(tb_sq_a.Text, out double sq_a);
            double.TryParse(tb_sq_b.Text, out double sq_b);
            double.TryParse(tb_sq_t.Text, out double sq_t);
            double.TryParse(tb_kc_t.Text, out double kc_t);
            double.TryParse(tb_kc_p.Text, out double kc_p);

            //--------------------------------------------

            double.TryParse(tb_RSa.Text, out double a);
            double.TryParse(tb_RSb.Text, out double b);
            double.TryParse(tb_RSc.Text, out double c);
            double.TryParse(tb_RSow.Text, out double ow);
            double.TryParse(tb_RSoh.Text, out double oh);
            double.TryParse(tb_RSd.Text, out double d);
            double.TryParse(tb_vtb.Text, out double vtb);
            double.TryParse(tb_VtP.Text, out double vtp);
            double.TryParse(tb_qlypp.Text, out double P);
            double.TryParse(tb_scs.Text, out double scs);
            double.TryParse(tb_slmotor.Text, out double slmt);
            double.TryParse(tb_tudien_d.Text, out double DT);
            double.TryParse(tb_tudien_r.Text, out double DT2);
            double.TryParse(tb_tudien_c.Text, out double HT);
            double.TryParse(tb_y_tudien.Text, out double y_tudien);

            //---------------------------------------------
            foreach (var tb in textBoxes)
            {
                if (double.TryParse(tb.Text, out double result))
                {
                    values.Add(result);
                    //  MessageBox.Show(values[1].ToString());
                    // MessageBox.Show("vô");
                }
                else
                {
                    // MessageBox.Show("sai");
                    // Nếu muốn bỏ qua lỗi parse, có thể ghi log ở đây
                    values.Add(0); // hoặc throw / log warning
                }
            }
            //foreach (var tb1 in values)
            //{
            //    MessageBox.Show(tb1.ToString());
            //}
            //    textBoxes = new TextBox[]
            //{tb_OW,tb_OH, tb_RSa, tb_RSb, tb_RSc, tb_RSd,tb_VtP, tb_SlB, tb_qlypp,tb_vtb   };

            var dynProps = new Dictionary<string, double>
            {
                { "B1", values[3] },
                { "B2", values[3] },
                { "C1", values[4]/2 },
                { "C", values[4] },
                { "OW", values[0] }










            };
            dimstyle_text = Insertdynamicblock.InsertDynamicBlock(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\RW.dwg", "RW_CB", point1, dynProps/*, cbb_test*/, dimstyle, 1.0);
            values.Clear();

            Point3d p2_t;
            Point3d p2_p;
            Point3d p_button1;
            Point3d p_button2;
            Point3d p_button3;
            Point3d p_button4;
            Point3d point_tuong_t;
            Point3d point_tuong_p;
            Point3d p_leader_td;
            List<Point3d> p_cabinet2 = new List<Point3d>();
            Point3d p2_leader_t;
            Point3d p2_leader_p;
            List<Point3d> P_point_td_cb;
            if (rad_lapngoai.Checked)
            {
                p2_t = new Point3d(point1.X + sq_a / 2.0 + kc_t, point1.Y - sq_b / 2.0, 0);
                p2_p = new Point3d(point1.X + ow + 2 * b + sq_a / 2.0 - kc_p - sq_a, point1.Y - sq_b / 2.0, 0);
                point_tuong_p = new Point3d(p2_p.X - sq_a / 2.0, p2_p.Y - sq_b / 2.0 - 200, 0);
                point_tuong_t = new Point3d(p2_t.X + sq_a / 2.0, p2_t.Y - sq_b / 2.0 - 200, 0);
                //  p_cabinet2 = new List<Point3d>() { rab_A.Checked ? new Point3d(point1.X + ow + 2*b + 935,  point_tuong_p.Y, 0) : new Point3d(point1.X  - 935 , point_tuong_p.Y , 0) };
                p2_leader_t = new Point3d(p2_t.X + scale * 5 * 50 + sq_a / 2.0, p2_t.Y - scale * 8 * 50, 0);
                p2_leader_p = new Point3d(p2_p.X - scale * 5 * 50 - sq_a / 2.0, p2_p.Y - scale * 8 * 50, 0);

                p_button1 = new Point3d(point_tuong_p.X + 20 + 86 / 2, point_tuong_p.Y, 0);
                p_button2 = new Point3d(point_tuong_t.X - 20 - 86 / 2, point_tuong_t.Y, 0); ;
                p_button3 = new Point3d(point1.X + ow + 2 * b + 86 - 93+50, point1.Y + c - 82, 0);
                p_button4 = new Point3d(point1.X - 86 / 2 - 50+50, point1.Y + c - 82, 0);
                // MessageBox.Show(point1.X)
            }
            else
            {
                p2_t = new Point3d(point1.X + sq_a / 2.0 + kc_t, point1.Y - sq_b / 2.0 + c + sq_b, 0);
                p2_p = new Point3d(point1.X + ow + 2 * b + sq_a / 2.0 - kc_p - sq_a, point1.Y - sq_b / 2.0 + c + sq_b, 0);
                point_tuong_p = new Point3d(p2_p.X - sq_a / 2.0, p2_p.Y + sq_b / 2.0, 0);
                point_tuong_t = new Point3d(p2_t.X + sq_a / 2.0, p2_t.Y + sq_b / 2.0, 0);
                p2_leader_t = new Point3d(p2_t.X + scale * 5 * 50 + sq_a / 2.0, p2_t.Y + scale * 8 * 50, 0);
                p2_leader_p = new Point3d(p2_p.X - scale * 5 * 50 - sq_a / 2.0, p2_p.Y + scale * 8 * 50, 0);
                p_button3 = new Point3d(point_tuong_p.X + 20 + 86 / 2, point_tuong_p.Y + 200, 0);
                p_button4 = new Point3d(point_tuong_t.X - 20 - 86 / 2, point_tuong_t.Y + 200, 0); ;
                p_button1 = new Point3d(point1.X + ow + 2 * b + 86 / 2 - 50 + 50, point1.Y + 82, 0);
                p_button2 = new Point3d(point1.X - 86 / 2 - 50 + 50, point1.Y + 82, 0);
            }
            p_cabinet2 = new List<Point3d>() { rab_A.Checked ? new Point3d(point1.X + ow + 2 * b + 935, point_tuong_p.Y, 0) : new Point3d(point1.X - 935, point_tuong_p.Y, 0) };
            p_leader_td = new Point3d(p_cabinet2[0].X + 20, p_cabinet2[0].Y - DT2, 0);

            CreateBoxTube(a: sq_a, b: sq_b, t: sq_t, R: 10, center: p2_t, 3);
            CreateBoxTube(a: sq_a, b: sq_b, t: sq_t, R: 10, center: p2_p, 3);
            string text_ks = rdb_y.Checked ? "口-" + sq_a + "x" + sq_b + "x" + sq_t : "口-" + sq_a + "x" + sq_b;
            ObjectId abv1 = AddLeaderNote1(new Point3d(p2_t.X + sq_a / 2.0, p2_t.Y, 0), p2_leader_t, text_ks + "\\P" + "（当社工事外）", dimstyle_text);
           // AddLeaderNotecmd(new Point3d(p2_p.X - sq_a / 2.0, p2_p.Y, 0), p2_leader_p, text_ks + "\\P" + "（当社工事外）", dimstyle_text);
            // AddLeaderNoteMirror(new Point3d(p2_t.X + sq_a / 2.0, p2_t.Y, 0), p2_leader_t, text_ks + "\\P" + "（当社工事外）", dimstyle_text, point1.X+ow/2);

            // CreateLeaderAndMText(new Point3d(p2_p.X - sq_a / 2.0, p2_p.Y, 0), p2_leader_p, p2_leader_p, text_ks + "\\P" + "（当社工事外）", dimstyle_text);
             // ObjectId abv =    AddLeaderNotemir(new Point3d(p2_p.X - sq_a / 2.0, p2_p.Y, 0), p2_leader_p, text_ks + "\\P" + "（当社工事外）", dimstyle_text);
            MirrorByCommand(abv1, p0, new Point3d(p0.X, p0.Y+100, 0),0);
            // MirrorMLeaderByTransform(abv, p0, new Point3d(p0.X, p0.Y + 100, 0));
            // Insertdynamicblock.InsertBlockMleader1(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\RW.dwg", "mleaderinsert", text_ks + "\\P" + "（当社工事外）", new Point3d(p2_p.X - sq_a / 2.0, p2_p.Y, 0), scale, p2_leader_p);
             
            //----------------------------------tạo tường


            double L = 1500.0;
            var dynProps1 = new Dictionary<string, double>
            {
                { "L", L }
            };
            Insertdynamicblock.InsertDynamicBlock(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\RW.dwg", "wall", point_tuong_p, dynProps1/*, cbb_test*/, dimstyle, 1.0);
            var p = new List<Point2d>()
             {
                new Point2d(point_tuong_p.X, point_tuong_p.Y),
                new Point2d(point_tuong_p.X+L, point_tuong_p.Y),
                new Point2d(point_tuong_p.X+L, point_tuong_p.Y+200),
                new Point2d(point_tuong_p.X, point_tuong_p.Y+200),
            };
            CreateHatchFrom4Points(p, "ANSI31", LineWeight.LineWeight009);
            Insertdynamicblock.InsertDynamicBlock(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\RW.dwg", "wall_t", point_tuong_t, dynProps1/*, cbb_test*/, dimstyle, 1.0);
            var p2 = new List<Point2d>()
             {
                new Point2d(point_tuong_t.X, point_tuong_t.Y),
                new Point2d(point_tuong_t.X-L, point_tuong_t.Y),
                new Point2d(point_tuong_t.X-L, point_tuong_t.Y+200),
                new Point2d(point_tuong_t.X, point_tuong_t.Y+200),
            };
            CreateHatchFrom4Points(p2, "ANSI31", LineWeight.LineWeight009);
            //------------------------------------------tạo tường
            Insertdynamicblock.InsertBlock(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\RW.dwg", "trong", new Point3d(point1.X + b + ow / 2.0, point1.Y - 400, 0), null, scale, 0);
            Insertdynamicblock.InsertBlock(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\RW.dwg", "ngoai", new Point3d(point1.X + b + ow / 2.0, point1.Y + c + 400, 0), null, scale, 0);

            //----------------------------------tủ điện mặt bằng


            if (CBB_TUDIEN.SelectedIndex != 2 && CBB_TUDIEN.SelectedIndex != 3)

            {
                ElectricalCabinet(p_cabinet2, ed, 1.0, new List<double> { DT / 2, DT / 2, DT2 }, "Tutieuchuan_CB");

            }
            else
            {
                string blocknametudien = CBB_TUDIEN.SelectedIndex == 2 ? "Tu_RL1_CB" : "Tu_RL2_CB";

                Insertdynamicblock.InsertBlock(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\RW.dwg", blocknametudien, p_cabinet2[0], null, 1.0, 0);
            }


            AddLeaderNote(p_leader_td, new Point3d(p_leader_td.X + scale * 5 * 50, p_leader_td.Y - scale * 50 * 8, 0), "制御箱", dimstyle_text);
            List<Point3d> PD3;
            if (rab_A.Checked)
            {
                PD3 = new List<Point3d>() { new Point3d(p_cabinet2[0].X + DT / 2, p_cabinet2[0].Y - DT2, 0), new Point3d(point_tuong_p.X + L, p_cabinet2[0].Y, 0) };
                createDim(PD3.ToArray(), new Point3d(PD3[1].X + 300 * scale, PD3[1].Y, 0), dimstyle_text, 0.5 * Math.PI, new Point3d(PD3[1].X + 300 * scale, PD3[1].Y + 200, 0), new Point3d(PD3[1].X + 300 * scale, PD3[1].Y + 200, 0), true, "", "", 0);
            }
            else
            {

                PD3 = new List<Point3d>() { new Point3d(p_cabinet2[0].X - DT / 2, p_cabinet2[0].Y - DT2, 0), new Point3d(point_tuong_t.X - L, p_cabinet2[0].Y, 0) };
                createDim(PD3.ToArray(), new Point3d(PD3[1].X - 300 * scale, PD3[1].Y, 0), dimstyle_text, 0.5 * Math.PI, new Point3d(PD3[1].X - 300 * scale, PD3[1].Y + 200, 0), new Point3d(PD3[1].X - 300 * scale, PD3[1].Y + 200, 0), true, "", "", 0);
            }
            
            //---------------------------------------------------------------------
            List<bool> buttonyn = new List<bool> { CheckCB1.Checked, CheckCB2.Checked, CheckCB3.Checked, CheckCB4.Checked };
            List<Point3d> point3Ds = new List<Point3d>() { p_button1, p_button2, p_button3, p_button4 };

            var dynProps2 = new Dictionary<string, double>
            {
            { "Angle", 0 }
            };
            var dynProps3 = new Dictionary<string, double>
            {
            { "Angle", Math.PI }
            };
            for (int i = 0; i < 4; i++)
            {
                if (buttonyn[i] == true)
                {
                    if (i < 2)
                    {
                        Insertdynamicblock.InsertDynamicBlock(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\RW.dwg", "Button", point3Ds[i], dynProps3/*, cbb_test*/, dimstyle, 1.0);
                    }
                    else
                    {
                        Insertdynamicblock.InsertDynamicBlock(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\RW.dwg", "Button", point3Ds[i], dynProps2/*, cbb_test*/, dimstyle, 1.0);
                    }
                }
            }

            //---------------------------------------------------------------


        }
        public static void MirrorMTextInPlace_Y(ObjectId mleaderId)
        {
            Document doc = acadApp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            using (doc.LockDocument())
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                MLeader ml = tr.GetObject(mleaderId, OpenMode.ForWrite) as MLeader;
                if (ml == null || ml.ContentType != ContentType.MTextContent)
                {
                    tr.Commit();
                    return;
                }

                MText mt = ml.MText;

                // Vị trí gốc text
                Point3d basePoint = mt.Location;

                // Đường mirror: đường thẳng dọc (song song trục Y) đi qua điểm gốc
                Line3d mirrorLine = new Line3d(
                    basePoint,
                    basePoint + Vector3d.YAxis
                );

                // Ma trận mirror
                Matrix3d mx = Matrix3d.Mirroring(mirrorLine);

                // Mirror MText ngay tại vị trí đó
                mt.TransformBy(mx);

                // Gán lại MText vào MLeader
                ml.MText = mt;

                // Điều chỉnh attachment và alignment cho đúng hướng sau khi flip
                FixAttachment(mt);
                FixAlignment(ml);

                tr.Commit();
            }
        }
        private static void FixAttachment(MText mt)
        {
            switch (mt.Attachment)
            {
                case AttachmentPoint.TopLeft:
                    mt.Attachment = AttachmentPoint.TopRight;
                    break;
                case AttachmentPoint.TopRight:
                    mt.Attachment = AttachmentPoint.TopLeft;
                    break;
                case AttachmentPoint.BottomLeft:
                    mt.Attachment = AttachmentPoint.BottomRight;
                    break;
                case AttachmentPoint.BottomRight:
                    mt.Attachment = AttachmentPoint.BottomLeft;
                    break;
                case AttachmentPoint.MiddleLeft:
                    mt.Attachment = AttachmentPoint.MiddleRight;
                    break;
                case AttachmentPoint.MiddleRight:
                    mt.Attachment = AttachmentPoint.MiddleLeft;
                    break;
            }
        }
        private static void FixAlignment(MLeader ml)
        {
            if (ml.TextAlignmentType == TextAlignmentType.LeftAlignment)
                ml.TextAlignmentType = TextAlignmentType.RightAlignment;
            else if (ml.TextAlignmentType == TextAlignmentType.RightAlignment)
                ml.TextAlignmentType = TextAlignmentType.LeftAlignment;
        }

        //----------------------------------------------Hàm chiếu cạnh

        //------------------------------------------------HÀM CHIẾU BẰNG
        private void CHIEUCANH(Point3d point1, Editor ed)
        {
            double.TryParse(tb_sq_a.Text, out double sq_a);
            double.TryParse(tb_sq_b.Text, out double sq_b);
            double.TryParse(tb_sq_t.Text, out double sq_t);
            double.TryParse(tb_kc_t.Text, out double kc_t);
            double.TryParse(tb_kc_p.Text, out double kc_p);

            //--------------------------------------------

            double.TryParse(tb_RSa.Text, out double a);
            double.TryParse(tb_RSb.Text, out double b);
            double.TryParse(tb_RSc.Text, out double c);
            double.TryParse(tb_RSow.Text, out double ow);
            double.TryParse(tb_RSoh.Text, out double oh);
            double.TryParse(tb_RSd.Text, out double d);
            double.TryParse(tb_kcbl.Text, out double kc_bl);
            double.TryParse(tb_vtb.Text, out double vtb);
            double.TryParse(tb_VtP.Text, out double vtp);
            double.TryParse(tb_qlypp.Text, out double P);
            double.TryParse(tb_scs.Text, out double scs);
            double.TryParse(tb_slmotor.Text, out double slmt);
            double.TryParse(tb_tudien_d.Text, out double DT);
            double.TryParse(tb_tudien_r.Text, out double DT2);
            double.TryParse(tb_tudien_c.Text, out double HT);
            double.TryParse(tb_y_tudien.Text, out double y_tudien);
            double.TryParse(tb_dk2.Text, out double dkong);
            double.TryParse(tb_dk1.Text, out double dkpuly);
            double.TryParse(tb_caoks_1.Text, out double caoks_1);
            double.TryParse(tb_caoks_2.Text, out double caoks_2);
            double.TryParse(tb_caotuong.Text, out double caotuong);
            //---------------------------------------------

            //foreach (var tb1 in values)
            //{
            //    MessageBox.Show(tb1.ToString());
            //}
            //    textBoxes = new TextBox[]
            //{tb_OW,tb_OH, tb_RSa, tb_RSb, tb_RSc, tb_RSd,tb_VtP, tb_SlB, tb_qlypp,tb_vtb   };

            var dynProps = new Dictionary<string, double>
            {
                { "OHD", oh + d},
                { "A", a },
                { "C", c },
                { "K", 50.0 + kc_bl }

            };
            dimstyle_text = Insertdynamicblock.InsertDynamicBlock(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\RW.dwg", "RW_CC", point1, dynProps/*, cbb_test*/, dimstyle, 1.0);
            values.Clear();

            Insertdynamicblock.InsertBlock(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\RW.dwg", "Tam1", new Point3d(point1.X + c / 2, point1.Y + oh + d - a + vtp, 0), null, 1.0, 0);
            Insertdynamicblock.InsertBlock(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\RW.dwg", "Tam2", new Point3d(point1.X + c - 17 - dkong / 2, point1.Y + oh + dkong / 2, 0), null, 1.0, 0);
            Point3d p1 = new Point3d(point1.X + c / 2, point1.Y + oh + d - a + vtp, 0);
            Point3d p2 = new Point3d(point1.X + c - 17 - dkong / 2, point1.Y + oh + dkong / 2, 0);

            CreateCircle(p1, dkpuly / 2);
            CreateCircle(p2, dkong / 2);
            CreateLine(new Point3d(point1.X + c - 17, point1.Y + oh + dkong / 2, 0), new Point3d(point1.X + c - 17, point1.Y + oh + d - a, 0), 3);
            CreateLine(new Point3d(p1.X + dkpuly / 2 * Math.Cos(Math.PI / 8), p1.Y - dkpuly / 2 * Math.Sin(Math.PI / 8), 0), new Point3d(p2.X - dkong / 2 * Math.Cos(Math.PI / 8), p2.Y + dkong / 2 * Math.Sin(Math.PI / 8), 0), 3);
            Point3d p_kscd = new Point3d();
            double tuongcc_1 = new double();
            double tuongcc_2 = new double();
            if (rad_lapngoai.Checked)
            {
                p_kscd = new Point3d(point1.X + c + sq_a / 2, point1.Y, 0);
                tuongcc_1 = 200;
                tuongcc_2 = sq_a/2;
            }
            else
            {
                p_kscd = new Point3d(point1.X - sq_a / 2, point1.Y, 0);
                tuongcc_1 = -200;
                tuongcc_2 = -sq_a / 2;
            }
            Point3d p_kscc_1 = new Point3d(p_kscd.X, p_kscd.Y + oh + d + caoks_1 - sq_b / 2, 0);
            Point3d p_kscc_2 = new Point3d(p_kscd.X, p_kscd.Y + oh + d + caoks_1 - sq_b / 2 - caoks_2, 0);
            Khungsat_cd(p_kscd, p_kscc_1, p_kscc_2, sq_a, sq_b, sq_t);

            Point2d p0_tuong = new Point2d(p_kscd.X+tuongcc_2, point1.Y + oh+caotuong);
            Point2d p1_tuong = new Point2d(p0_tuong.X + tuongcc_1, p0_tuong.Y);
            Point2d p2_tuong = new Point2d(p1_tuong.X , point1.Y+oh+d+200);
            Point2d p3_tuong = new Point2d(p0_tuong.X, p2_tuong.Y);
            tuong_cc(new List<Point2d> {
                p0_tuong,p1_tuong,p2_tuong,p3_tuong
            });
            List<Point3d> p_dim = new List<Point3d>();
            List<Point3d> p_dim2 = new List<Point3d>();
            List<Point3d> p_dim3 = new List<Point3d>();
            List<Point3d> p_dim4 = new List<Point3d>();
            List<Point3d> p_dim5 = new List<Point3d>();
            //----------------------------------------dim kích thước

            if (rad_lapngoai.Checked)
            {
                p_dim = new List<Point3d>() { new Point3d(point1.X-563, point1.Y, 0), new Point3d(p2.X , p2.Y-dkong/2, 0), new Point3d(point1.X, point1.Y+oh+d, 0) };
                createDim(p_dim.ToArray(), new Point3d(p_dim[0].X - 630 * scale, p_dim[0].Y, 0), dimstyle_text, 0.5 * Math.PI, new Point3d(p_dim[0].X - 630 * scale, p_dim[0].Y + 200, 0), p_dim[0], false, "OH:", "", 0);
                p_dim2 = new List<Point3d>() { new Point3d(point1.X, point1.Y + oh + d-a, 0), new Point3d(point1.X, point1.Y + oh + d, 0) };
                createDim(p_dim2.ToArray(), new Point3d(p_dim2[0].X - 630/1.5 * scale, p_dim2[0].Y, 0), dimstyle_text, 0.5 * Math.PI, new Point3d(p_dim2[0].X - 630/1.5 * scale, p_dim2[1].Y + 200, 0), new Point3d(p_dim2[0].X - 630 / 1.5 * scale, p_dim2[1].Y + 200, 0), true, "", "", 0);
                p_dim3 = new List<Point3d>() { new Point3d(point1.X - 563, point1.Y, 0), new Point3d(point1.X, point1.Y + oh + d, 0) };
                createDim(p_dim3.ToArray(), new Point3d(p_dim3[0].X - 900 * scale, p_dim3[0].Y, 0), dimstyle_text, 0.5 * Math.PI, new Point3d(p_dim3[0].X - 900 * scale, p_dim3[0].Y + 200, 0), p_dim3[0], false, "", "", 0);
                p_dim4 = new List<Point3d>() {  new Point3d(point1.X, point1.Y + oh + d, 0) , new Point3d(point1.X+c , point1.Y + oh + d + caoks_1, 0) };
                createDim(p_dim4.ToArray(), new Point3d(p_dim3[0].X - 900 * scale, p_dim4[0].Y, 0), dimstyle_text, 0.5 * Math.PI, new Point3d(p_dim3[0].X - 900 * scale, p_dim4[0].Y + 200, 0), new Point3d(p_dim3[0].X - 900 * scale, p_dim4[0].Y + 200, 0), true, "", "", 0);
                p_dim5 = new List<Point3d>() { new Point3d(point1.X - 563, point1.Y, 0), new Point3d(point1.X + c, point1.Y + oh + d + caoks_1, 0) };
                createDim(p_dim5.ToArray(), new Point3d(p_dim5[0].X - 1170 * scale, p_dim5[0].Y, 0), dimstyle_text, 0.5 * Math.PI, new Point3d(p_dim5[0].X - 1170 * scale, p_dim5[0].Y + 200, 0), p_dim5[0], false, "", "", 0);
            }
            else
            {
                p_dim = new List<Point3d>() { new Point3d(point1.X + 739, point1.Y, 0), new Point3d(p2.X, p2.Y - dkong / 2, 0), new Point3d(point1.X+c, point1.Y + oh + d, 0) };
                createDim(p_dim.ToArray(), new Point3d(p_dim[0].X + 630 * scale, p_dim[0].Y, 0), dimstyle_text, 0.5 * Math.PI, new Point3d(p_dim[0].X + 630 * scale, p_dim[0].Y + 200, 0), p_dim[0], false, "OH:", "", 0);
                p_dim2 = new List<Point3d>() { new Point3d(point1.X+c, point1.Y + oh + d - a, 0), new Point3d(point1.X+c, point1.Y + oh + d, 0) };
                createDim(p_dim2.ToArray(), new Point3d(p_dim2[0].X + 630 / 1.5 * scale, p_dim2[0].Y, 0), dimstyle_text, 0.5 * Math.PI, new Point3d(p_dim2[0].X + 630 / 1.5 * scale, p_dim2[1].Y + 200, 0), new Point3d(p_dim2[0].X + 630 / 1.5 * scale, p_dim2[1].Y + 200, 0), true, "", "", 0);
                p_dim3 = new List<Point3d>() { new Point3d(point1.X + 739, point1.Y, 0), new Point3d(point1.X+c, point1.Y + oh + d, 0) };
                createDim(p_dim3.ToArray(), new Point3d(p_dim3[0].X + 900 * scale, p_dim3[0].Y, 0), dimstyle_text, 0.5 * Math.PI, new Point3d(p_dim3[0].X + 900 * scale, p_dim3[0].Y + 200, 0), p_dim3[0], false, "", "", 0);
                p_dim4 = new List<Point3d>() { new Point3d(point1.X+c, point1.Y + oh + d, 0), new Point3d(point1.X , point1.Y + oh + d + caoks_1, 0) };
                createDim(p_dim4.ToArray(), new Point3d(p_dim3[0].X + 900 * scale, p_dim4[0].Y, 0), dimstyle_text, 0.5 * Math.PI, new Point3d(p_dim3[0].X + 900 * scale, p_dim4[0].Y + 200, 0), new Point3d(p_dim3[0].X + 900 * scale, p_dim4[0].Y + 200, 0), true, "", "", 0);
                p_dim5 = new List<Point3d>() { new Point3d(point1.X + 739, point1.Y, 0), new Point3d(point1.X , point1.Y + oh + d + caoks_1, 0) };
                createDim(p_dim5.ToArray(), new Point3d(p_dim5[0].X + 1170 * scale, p_dim5[0].Y, 0), dimstyle_text, 0.5 * Math.PI, new Point3d(p_dim5[0].X + 1170 * scale, p_dim5[0].Y + 200, 0), p_dim5[0], false, "", "", 0);

                //PD3 = new List<Point3d>() { new Point3d(p_cabinet2[0].X - DT / 2, p_cabinet2[0].Y - DT2, 0), new Point3d(point_tuong_t.X - L, p_cabinet2[0].Y, 0) };
                //createDim(PD3.ToArray(), new Point3d(PD3[1].X - 300 * scale, PD3[1].Y, 0), dimstyle_text, 0.5 * Math.PI, new Point3d(PD3[1].X - 300 * scale, PD3[1].Y + 200, 0), PD3[0], false, "", "", 0);
            }
            Insertdynamicblock.InsertBlockMleader(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\RW.dwg", "ld1", "", new Point3d(point1.X+45,point1.Y-9,0), scale);
            Insertdynamicblock.InsertBlockMleader(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\RW.dwg", "ld3", "", new Point3d(point1.X + 30, point1.Y +815, 0), scale);
            Insertdynamicblock.InsertBlockMleader(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\RW.dwg", "ld2", "", new Point3d(point1.X + 30, point1.Y +293, 0), scale);
            //----------------------------------------
            Point3d p_leader = new Point3d(p_kscc_1.X, p_kscc_1.Y + sq_b / 2, 0);
            Point3d p2_leader_t  =   new Point3d(p_leader.X + scale * 5 * 50 , p_leader.Y + scale * 10 * 50, 0);
            // Point3d p2_leader_p = new Point3d(p_leader.X - scale * 5 * 50, p_leader.Y + scale * 8 * 50, 0);
            string text_ks = rdb_y.Checked ? "口-" + sq_a + "x" + sq_b + "x" + sq_t : "口-" + sq_a + "x" + sq_b;
           
            // AddLeaderNotecmd(new Point3d(p2_p.X - sq_a / 2.0, p2_p.Y, 0), p2_leader_p, text_ks + "\\P" + "（当社工事外）", dimstyle_text);
            // AddLeaderNoteMirror(new Point3d(p2_t.X + sq_a / 2.0, p2_t.Y, 0), p2_leader_t, text_ks + "\\P" + "（当社工事外）", dimstyle_text, point1.X+ow/2);

            // CreateLeaderAndMText(new Point3d(p2_p.X - sq_a / 2.0, p2_p.Y, 0), p2_leader_p, p2_leader_p, text_ks + "\\P" + "（当社工事外）", dimstyle_text);
            // ObjectId abv =    AddLeaderNotemir(new Point3d(p2_p.X - sq_a / 2.0, p2_p.Y, 0), p2_leader_p, text_ks + "\\P" + "（当社工事外）", dimstyle_text);
            if(rad_lapngoai.Checked)
            {
                Insertdynamicblock.InsertBlockMleader5(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\RW.dwg", "ld4", @"\W0.8;" + text_ks + "\\P" + "（当社工事外）", p_leader, scale);
                //MirrorByCommand(abv, p0, new Point3d(p0.X, p0.Y + 100, 0),1);
            }  
            else
            {
                ObjectId abv = AddLeaderNote1(p_leader, p2_leader_t, text_ks + "\\P" + "（当社工事外）", dimstyle_text);
            }    
            
        }

        public static void Khungsat_cd(Point3d center, Point3d center1, Point3d center2, double sq_a, double sq_b, double sq_t)
        {

            CreateBoxTube(a: sq_a, b: sq_b, t: sq_t, R: 10, center: center1, 4);
            CreateBoxTube(a: sq_a, b: sq_b, t: sq_t, R: 10, center: center2, 4);
            CreateLine(new Point3d(center.X + sq_a / 2, center.Y, 0), new Point3d(center1.X + sq_a / 2, center1.Y, 0), 4);
            CreateLine(new Point3d(center.X - sq_a / 2, center.Y, 0), new Point3d(center1.X - sq_a / 2, center1.Y, 0), 4);
            CreateLine2(new Point3d(center.X + sq_a / 2 - sq_t, center.Y, 0), new Point3d(center1.X + sq_a / 2 - sq_t, center1.Y, 0), "Dashed", 2.0);
            CreateLine2(new Point3d(center.X - sq_a / 2 + sq_t, center.Y, 0), new Point3d(center1.X - sq_a / 2 + sq_t, center1.Y, 0), "Dashed", 2.0);


        }
        public static void tuong_cc(List<Point2d> p)
        {
            Point3d p0 = new Point3d(p[0].X, p[0].Y, 0);
            Point3d p1 = new Point3d(p[1].X, p[1].Y, 0);
            Point3d p2 = new Point3d(p[2].X, p[2].Y, 0);
            Point3d p3 = new Point3d(p[3].X, p[3].Y, 0);
            CreateLine(p3, p0, 7, LineWeight.LineWeight009);
            CreateLine(p0, p1, 7, LineWeight.LineWeight009);
            CreateLine(p1, p2, 7, LineWeight.LineWeight009);
            CreateHatchFrom4Points(p, "ANSI31", LineWeight.LineWeight009);
        }
        public static void CreateCircle(Point3d center, double radius)
        {
            Document doc = acadApp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            using (DocumentLock m_DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument())
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    // Tạo đường tròn
                    Circle circle = new Circle();
                    circle.Center = center;
                    circle.Radius = radius;
                    circle.LineWeight = LineWeight.LineWeight013;
                    circle.ColorIndex = 3;

                    // Thêm vào bản vẽ
                    btr.AppendEntity(circle);
                    tr.AddNewlyCreatedDBObject(circle, true);

                    tr.Commit();
                }
            }
        }
        public static void CreateLine(Point3d startPoint, Point3d endPoint, int color, LineWeight lineWeight = LineWeight.LineWeight013)
        {
            Document doc = acadApp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            using (DocumentLock m_DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument())
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    // Tạo đường line
                    Line line = new Line(startPoint, endPoint);
                    line.LineWeight = lineWeight;
                    line.ColorIndex = color;
                    // Thêm vào bản vẽ
                    btr.AppendEntity(line);
                    tr.AddNewlyCreatedDBObject(line, true);

                    tr.Commit();
                }
            }
        }
        public static void CreateLine2(Point3d startPoint, Point3d endPoint, string linetype, double linescale, LineWeight lineWeight = LineWeight.LineWeight009)
        {
            Document doc = acadApp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            using (DocumentLock m_DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument())
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                    // Load linetype nếu cần
                    if (linetype != "Continuous" && linetype != "ByLayer")
                    {
                        try
                        {
                            LoadLinetype(db, tr, linetype);
                        }
                        catch
                        {
                            // Nếu không load được, dùng Continuous
                            linetype = "Continuous";
                        }
                    }
                    // Tạo đường line
                    Line line = new Line(startPoint, endPoint);
                    line.LineWeight = lineWeight;
                    line.Linetype = linetype;
                    line.LinetypeScale = linescale;
                    line.ColorIndex = 4;
                    // Thêm vào bản vẽ
                    btr.AppendEntity(line);
                    tr.AddNewlyCreatedDBObject(line, true);

                    tr.Commit();
                }
            }
        }
        private static void LoadLinetype(Database db, Transaction tr, string linetypeName)
        {
            LinetypeTable lt = tr.GetObject(db.LinetypeTableId, OpenMode.ForRead) as LinetypeTable;

            if (!lt.Has(linetypeName))
            {
                // Load linetype từ file acad.lin hoặc acadiso.lin
                db.LoadLineTypeFile(linetypeName, "acad.lin");
            }
        }
        //------------------------------------------------------------
        public static ObjectId CreateHatchFrom4Points(
    List<Point2d> pts,
    string hatchName,
    LineWeight lw)
        {
            if (pts == null || pts.Count != 4)
                throw new ArgumentException("Danh sách phải gồm đúng 4 điểm.");

            Document doc = acadApp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            using (DocumentLock m_DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument())
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    // ---- Tạo polyline rectangle từ 4 điểm ----
                    Polyline pl = new Polyline();
                    for (int i = 0; i < 4; i++)
                        pl.AddVertexAt(i, pts[i], 0, 0, 0);

                    pl.Closed = true;
                    pl.LineWeight = lw;

                    ObjectId plId = btr.AppendEntity(pl);
                    tr.AddNewlyCreatedDBObject(pl, true);
                    //--------------------


                    // ---- Tạo Hatch ----
                    Hatch hatch = new Hatch();
                    hatch.PatternScale = 400;
                    hatch.SetHatchPattern(HatchPatternType.PreDefined, hatchName);
                    hatch.Associative = false;
                    hatch.LineWeight = lw;
                    ObjectId hatchId = btr.AppendEntity(hatch);
                    tr.AddNewlyCreatedDBObject(hatch, true);// KHÔNG associative để cho phép xóa đường bao






                    // ---- Gán boundary ----
                    hatch.AppendLoop(HatchLoopTypes.Outermost, new ObjectIdCollection { plId });
                    hatch.EvaluateHatch(true);

                    // ---- XÓA polyline boundary ----
                    Entity ent = tr.GetObject(plId, OpenMode.ForWrite) as Entity;
                    ent.Erase();

                    tr.Commit();

                    return hatchId;
                }
            }
        }
        //---------------------------------------------------------vẽ hộp 
        public static void CreateBoxTube(double a, double b, double t, double R, Point3d center, int color)
        {
            Document doc = acadApp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            using (DocumentLock m_DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument())
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    // ===== 1) Tạo polyline ngoài (0,0) → (a,b) tại góc trái dưới =====
                    Polyline outer = CreateRoundedRect(a, b, R);
                    outer.LineWeight = LineWeight.LineWeight013;
                    outer.ColorIndex = color;

                    // ===== 2) Tạo polyline trong =====
                    double ai = a - 2 * t;
                    double bi = b - 2 * t;
                    double Ri = R - t;

                    Polyline inner = CreateRoundedRect(ai, bi, Ri);
                    inner.TransformBy(Matrix3d.Displacement(new Vector3d(t, t, 0)));
                    inner.LineWeight = LineWeight.LineWeight009;
                    inner.ColorIndex = color;
                    double a_dich = (R * (Math.Sqrt(2) - 1) + t) * Math.Cos(Math.PI / 4);
                    // ===== 3) Tạo 2 đường chéo =====
                    Line d1 = new Line(new Point3d(0 + a_dich, 0 + a_dich, 0), new Point3d(a - a_dich, b - a_dich, 0));
                    Line d2 = new Line(new Point3d(0 + a_dich, b - a_dich, 0), new Point3d(a - a_dich, 0 + a_dich, 0));
                    d1.LineWeight = LineWeight.LineWeight009;
                    d2.LineWeight = LineWeight.LineWeight009;
                    d1.ColorIndex = color;
                    d2.ColorIndex = color;
                    // ===== 4) Tính vector dịch để đưa tâm về vị trí cần vẽ =====
                    Point3d currentCenter = new Point3d(a / 2.0, b / 2.0, 0);
                    Vector3d move = center - currentCenter;

                    Matrix3d mat = Matrix3d.Displacement(move);

                    // Dịch tất cả geometry
                    outer.TransformBy(mat);
                    inner.TransformBy(mat);
                    d1.TransformBy(mat);
                    d2.TransformBy(mat);

                    // ===== 5) Ghi vào bản vẽ =====
                    btr.AppendEntity(outer); tr.AddNewlyCreatedDBObject(outer, true);
                    btr.AppendEntity(inner); tr.AddNewlyCreatedDBObject(inner, true);
                    btr.AppendEntity(d1); tr.AddNewlyCreatedDBObject(d1, true);
                    btr.AppendEntity(d2); tr.AddNewlyCreatedDBObject(d2, true);

                    tr.Commit();
                }
            }
        }


        // ---------------------------------
        // Hình chữ nhật bo góc
        // ---------------------------------
        private static Polyline CreateRoundedRect(double width, double height, double radius)
        {
            Polyline pl = new Polyline();

            double k = 0.414213562;   // bulge 90 degrees

            pl.AddVertexAt(0, new Point2d(radius, 0), 0, 0, 0);              // Bottom edge
            pl.AddVertexAt(1, new Point2d(width - radius, 0), k, 0, 0);
            pl.AddVertexAt(2, new Point2d(width, radius), 0, 0, 0);          // Right edge
            pl.AddVertexAt(3, new Point2d(width, height - radius), k, 0, 0);
            pl.AddVertexAt(4, new Point2d(width - radius, height), 0, 0, 0);      // Top edge
            pl.AddVertexAt(5, new Point2d(radius, height), k, 0, 0);
            pl.AddVertexAt(6, new Point2d(0, height - radius), 0, 0, 0);          // Left edge
            pl.AddVertexAt(7, new Point2d(0, radius), k, 0, 0);

            pl.Closed = true;
            return pl;
        }
        //---------------------------------------------Hàm hình chiếu đứng
        private void mattruoc(Point3d point1, Editor ed)
        {

            foreach (var tb in textBoxes)
            {
                if (double.TryParse(tb.Text, out double result))
                {
                    values.Add(result);
                    //  MessageBox.Show(values[1].ToString());
                    // MessageBox.Show("vô");
                }
                else
                {
                    // MessageBox.Show("sai");
                    // Nếu muốn bỏ qua lỗi parse, có thể ghi log ở đây
                    values.Add(0); // hoặc throw / log warning
                }
            }
            //foreach (var tb1 in values)
            //{
            //    MessageBox.Show(tb1.ToString());
            //}
            //    textBoxes = new TextBox[]
            //{tb_OW,tb_OH, tb_RSa, tb_RSb, tb_RSc, tb_RSd,tb_VtP, tb_SlB, tb_qlypp,tb_vtb   };

            var dynProps = new Dictionary<string, double>
            {
            { "OW1", values[0]/2.0 },
            { "OW2", values[0]/2.0 },
            { "H", values[1]+values[5] },
            { "B1", values[3] },
            { "B2", values[3] },
            { "A", values[2] },
            { "Vtp", values[6] }





            };
            dimstyle_text = Insertdynamicblock.InsertDynamicBlock(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\RW.dwg", "RW-011", point1, dynProps/*, cbb_test*/, dimstyle, 1.0);
            values.Clear();
        }
        private void Belt(List<Point3d> point1, Editor ed, double cout)
        {

            foreach (var tb in textBoxes)
            {
                if (double.TryParse(tb.Text, out double result))
                {
                    values.Add(result);
                    //  MessageBox.Show(values[1].ToString());
                    // MessageBox.Show("vô");
                }
                else
                {
                    // MessageBox.Show("sai");
                    // Nếu muốn bỏ qua lỗi parse, có thể ghi log ở đây
                    values.Add(0); // hoặc throw / log warning
                }
            }
            //foreach (var tb1 in values)
            //{
            //    MessageBox.Show(tb1.ToString());
            //}
            //    textBoxes = new TextBox[]
            //{tb_OW,tb_OH, tb_RSa, tb_RSb, tb_RSc, tb_RSd,tb_VtP, tb_SlB, tb_qlypp,tb_vtb   };

            var dynProps = new Dictionary<string, double>
            {

            { "H1", values[1]+values[5]-values[2] },

            { "Vtp", values[6] }




            };
            dimstyle_text = Insertdynamicblock.Insertbelt(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\RW.dwg", "Belt", point1, cout, dynProps/*, cbb_test*/, dimstyle, 1.0);
            values.Clear();
        }
        private void Tube(List<Point3d> point1, Editor ed, double cout, double value)
        {



            var dynProps = new Dictionary<string, double>
            {

            { "T1", value }




            };
            dimstyle_text = Insertdynamicblock.Insertbelt(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\RW.dwg", "Tube", point1, cout, dynProps/*, cbb_test*/, dimstyle, 1.0);
            values.Clear();
        }
        private void Window(List<Point3d> point1, Editor ed, double cout, double value)
        {

            //foreach (var tb in textBoxes)
            //{
            //    if (double.TryParse(tb.Text, out double result))
            //    {
            //        values.Add(result);
            //        //  MessageBox.Show(values[1].ToString());
            //        // MessageBox.Show("vô");
            //    }
            //    else
            //    {
            //        // MessageBox.Show("sai");
            //        // Nếu muốn bỏ qua lỗi parse, có thể ghi log ở đây
            //        values.Add(0); // hoặc throw / log warning
            //    }
            //}
            //foreach (var tb1 in values)
            //{
            //    MessageBox.Show(tb1.ToString());
            //}
            //    textBoxes = new TextBox[]
            //{tb_OW,tb_OH, tb_RSa, tb_RSb, tb_RSc, tb_RSd,tb_VtP, tb_SlB, tb_qlypp,tb_vtb   };

            var dynProps = new Dictionary<string, double>
            {

            { "WD1", value },
            { "WD2", value }




            };
            dimstyle_text = Insertdynamicblock.Insertbelt(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\RW.dwg", "Window", point1, cout, dynProps/*, cbb_test*/, dimstyle, 1.0);
            values.Clear();
        }

        private void ElectricalCabinet(List<Point3d> point1, Editor ed, double cout, List<double> value, string blockname)
        {



            var dynProps = new Dictionary<string, double>
            {

            { "DT1", value[0] },
            { "DT2", value[1] },
            { "HT1", value[2] }




        };
            dimstyle_text = Insertdynamicblock.Insertbelt(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\RW.dwg", blockname, point1, cout, dynProps/*, cbb_test*/, dimstyle, 1.0);
            values.Clear();
        }
        //---------------------------------------------nhập xuất dữ liệu excel
        private void WriteData(string filepathdata)
        {
            if (!File.Exists(filepathdata))
            {
                MessageBox.Show("Đường dẫn file không tồn tại.");
                return;
            }

            // Kiểm tra file có bị khóa không
            try
            {
                using (FileStream fs = File.Open(filepathdata, FileMode.Open, FileAccess.Read, FileShare.None))
                {
                    fs.Close();
                }
            }
            catch
            {
                MessageBox.Show("File đang được mở bởi chương trình khác. Vui lòng đóng file Excel trước.");
                return;
            }

            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet1 = null;
            Excel.Worksheet worksheet2 = null;

            try
            {


                // Khởi tạo Excel - Cách an toàn hơn
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null)
                {
                    MessageBox.Show("Không tìm thấy Excel trên máy tính. Vui lòng cài đặt Microsoft Excel.");
                    return;
                }

                excelApp = (Excel.Application)Activator.CreateInstance(excelType);

                if (excelApp == null)
                {
                    MessageBox.Show("Không thể khởi tạo Excel Application.");
                    return;
                }

                excelApp.Visible = false;
                excelApp.DisplayAlerts = false;
                excelApp.ScreenUpdating = false;

                // Mở workbook
                workbook = excelApp.Workbooks.Open(
                    filepathdata,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing);

                worksheet1 = (Excel.Worksheet)workbook.Sheets[1];
                worksheet2 = (Excel.Worksheet)workbook.Sheets[2];

                // Gán giá trị vào các ô (đã cập nhật 2025/09/09)
                worksheet1.Cells[3, 3] = tb_RSow.Text;
                worksheet1.Cells[4, 3] = tb_RSoh.Text;




                // Tính toán lại
                workbook.RefreshAll();
                excelApp.CalculateFullRebuild();

                // Đợi Excel tính toán xong
                System.Threading.Thread.Sleep(100);
                // MessageBox.Show(worksheet1.Cells[15, 3].Text?.ToString() ?? "");
                // Lấy dữ liệu kết quả sau tính toán
                tb_RSa.Text = worksheet1.Cells[15, 3].Text?.ToString() ?? "";
                tb_RSb.Text = worksheet1.Cells[16, 3].Text?.ToString() ?? "";
                tb_RSc.Text = worksheet1.Cells[17, 3].Text?.ToString() ?? "";
                tb_RSd.Text = worksheet1.Cells[18, 3].Text?.ToString() ?? "";
                tb_qlypp.Text = worksheet1.Cells[23, 3].Text?.ToString() ?? "";
                tb_SlB.Text = worksheet1.Cells[30, 3].Text?.ToString() ?? "";
                tb_vtb.Text = worksheet1.Cells[24, 3].Text?.ToString() ?? "";
                tb_VtP.Text = worksheet1.Cells[29, 3].Text?.ToString() ?? "";
                tb_scs.Text = worksheet1.Cells[31, 3].Text?.ToString() ?? "";
                tb_slmotor.Text = worksheet1.Cells[19, 3].Text?.ToString() ?? "";
                tb_tsmotor.Text = worksheet1.Cells[20, 3].Text?.ToString() ?? "";
                tb_mmotor.Text = worksheet1.Cells[32, 3].Text?.ToString() ?? "";
                tb_speedwind.Text = worksheet1.Cells[33, 3].Text?.ToString() ?? "";
                tb_ong1.Text = worksheet1.Cells[34, 3].Text?.ToString() ?? "";
                tb_ong2.Text = worksheet1.Cells[36, 3].Text?.ToString() ?? "";
                tb_ong3.Text = worksheet1.Cells[35, 3].Text?.ToString() ?? "";
                tb_v.Text = worksheet1.Cells[37, 3].Text?.ToString() ?? "";
                tb_dk2.Text = worksheet1.Cells[27, 3].Text?.ToString() ?? "";
                tb_dk1.Text = worksheet1.Cells[28, 3].Text?.ToString() ?? "";
                groupBox2.Text = worksheet1.Cells[13, 3].Text?.ToString() ?? "";
                t_thep = worksheet1.Cells[38, 3].Text?.ToString() ?? "";
                t_ong1 = worksheet1.Cells[34, 5].Text?.ToString() ?? "";
                t_ong3 = worksheet1.Cells[35, 5].Text?.ToString() ?? "";
                t_ong2 = worksheet1.Cells[36, 5].Text?.ToString() ?? "";
                tb_kcbl.Text = worksheet1.Cells[25, 3].Text?.ToString() ?? "";
                if (RDB_auto.Checked)
                {
                    cbb_Scale.SelectedItem = "1/" + worksheet1.Cells[39, 3].Text?.ToString() ?? "1/50";
                }

                // Bật lại screen updating
                excelApp.ScreenUpdating = true;
            }
            catch (COMException comEx)
            {
                MessageBox.Show($"Lỗi COM khi xử lý Excel:\n{comEx.Message}\n\nMã lỗi: {comEx.ErrorCode:X}\n\nVui lòng thử:\n" +
                               "1. Đóng tất cả Excel đang mở\n" +
                               "2. Chạy chương trình với quyền Administrator\n" +
                               "3. Repair Microsoft Office",
                               "Lỗi COM", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi xử lý Excel: {ex.Message}\n\nChi tiết: {ex.StackTrace}",
                               "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                //   Giải phóng tài nguyên COM theo thứ tự: Worksheet → Workbook → Application
                try
                {
                    if (worksheet1 != null)
                    {
                        Marshal.ReleaseComObject(worksheet1);
                        worksheet1 = null;
                    }

                    if (worksheet2 != null)
                    {
                        Marshal.ReleaseComObject(worksheet2);
                        worksheet2 = null;
                    }

                    if (workbook != null)
                    {
                        workbook.Close(false, Type.Missing, Type.Missing);
                        Marshal.ReleaseComObject(workbook);
                        workbook = null;
                    }

                    if (excelApp != null)
                    {
                        excelApp.Quit();
                        Marshal.ReleaseComObject(excelApp);
                        excelApp = null;
                    }

                    // Buộc garbage collection
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                }
                catch (Exception cleanupEx)
                {
                    MessageBox.Show($"Lỗi khi dọn dẹp tài nguyên: {cleanupEx.Message}");
                }


            }
        }


        //---------------------------------------------
        private void WriteData1(string filepathdata)
        {
            if (!File.Exists(filepathdata))
            {
                MessageBox.Show("Đường dẫn file không tồn tại.");
                return;
            }

            // Kiểm tra file có bị khóa không
            try
            {
                using (FileStream fs = File.Open(filepathdata, FileMode.Open, FileAccess.Read, FileShare.None))
                {
                    fs.Close();
                }
            }
            catch
            {
                MessageBox.Show("File đang được mở bởi chương trình khác. Vui lòng đóng file Excel trước.");
                return;
            }

            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet1 = null;
            Excel.Worksheet worksheet2 = null;

            try
            {


                // Khởi tạo Excel - Cách an toàn hơn
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null)
                {
                    MessageBox.Show("Không tìm thấy Excel trên máy tính. Vui lòng cài đặt Microsoft Excel.");
                    return;
                }

                excelApp = (Excel.Application)Activator.CreateInstance(excelType);

                if (excelApp == null)
                {
                    MessageBox.Show("Không thể khởi tạo Excel Application.");
                    return;
                }

                excelApp.Visible = false;
                excelApp.DisplayAlerts = false;
                excelApp.ScreenUpdating = false;

                // Mở workbook
                workbook = excelApp.Workbooks.Open(
                    filepathdata,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing);

                List<string> sheetNames = new List<string>();

                var doc = acadApp.DocumentManager.MdiActiveDocument;
                var db = doc.Database;
                var ed = doc.Editor;
                // loadingForm.Invoke(new Action(() => loadingForm.Close()));
                PromptPointResult pointResult = ed.GetPoint("Chọn vị trí để chèn Block Reference: ");
                MessageBox.Show(workbook.Sheets.Count.ToString());
                int abc = 0;
                foreach (Excel.Worksheet ws in workbook.Sheets)
                {
                    if (ws.Cells[25, 2].Text.ToString() == "No")
                    {
                        Insertdynamicblock.InsertBlock(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\RW.dwg", "List_HB", new Point3d(pointResult.Value.X, pointResult.Value.Y - abc * 1000, 0), new List<string>() { ws.Name, "-", "-" }, scale, 0);
                        // sheetNames.Add(ws.Name);
                        for (int i = 26; i <= 50; i++)
                        {
                            if (ws.Cells[i, 14].Text.ToString() != "")
                            {
                                Insertdynamicblock.InsertBlock(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\RW.dwg", "List_HB", new Point3d(pointResult.Value.X, pointResult.Value.Y - 14 * (i - 25) - abc * 1000, 0), new List<string>() { ws.Cells[i, 3].Text?.ToString() ?? "-", ws.Cells[i, 2].Text?.ToString() ?? "-", ws.Cells[i, 14].Text?.ToString() ?? "-" }, scale, 0);
                            }
                            if (ws.Cells[i, 15].Text.ToString() != "")
                            {
                                Insertdynamicblock.InsertBlock(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\RW.dwg", "List_HB", new Point3d(pointResult.Value.X + 500, pointResult.Value.Y - 14 * (i - 25) - abc * 1000, 0), new List<string>() { ws.Cells[i, 3].Text?.ToString() ?? "-", ws.Cells[i, 2].Text?.ToString() ?? "-", ws.Cells[i, 15].Text?.ToString() ?? "-" }, scale, 0);
                            }

                        }

                        abc++;
                    }
                }








                //// MessageBox.Show(worksheet1.Cells[15, 3].Text?.ToString() ?? "");
                //// Lấy dữ liệu kết quả sau tính toán
                //tb_RSa.Text = worksheet1.Cells[15, 3].Text?.ToString() ?? "";
                //tb_RSb.Text = worksheet1.Cells[16, 3].Text?.ToString() ?? "";
                //tb_RSc.Text = worksheet1.Cells[17, 3].Text?.ToString() ?? "";
                //tb_RSd.Text = worksheet1.Cells[18, 3].Text?.ToString() ?? "";
                //tb_qlypp.Text = worksheet1.Cells[23, 3].Text?.ToString() ?? "";
                //tb_SlB.Text = worksheet1.Cells[30, 3].Text?.ToString() ?? "";
                //tb_vtb.Text = worksheet1.Cells[24, 3].Text?.ToString() ?? "";
                //tb_VtP.Text = worksheet1.Cells[29, 3].Text?.ToString() ?? "";
                //tb_scs.Text = worksheet1.Cells[31, 3].Text?.ToString() ?? "";
                //tb_slmotor.Text = worksheet1.Cells[19, 3].Text?.ToString() ?? "";
                //tb_tsmotor.Text = worksheet1.Cells[20, 3].Text?.ToString() ?? "";
                //tb_mmotor.Text = worksheet1.Cells[32, 3].Text?.ToString() ?? "";
                //tb_speedwind.Text = worksheet1.Cells[33, 3].Text?.ToString() ?? "";
                //tb_ong1.Text = worksheet1.Cells[34, 3].Text?.ToString() ?? "";
                //tb_ong2.Text = worksheet1.Cells[36, 3].Text?.ToString() ?? "";
                //tb_ong3.Text = worksheet1.Cells[35, 3].Text?.ToString() ?? "";
                //tb_v.Text = worksheet1.Cells[37, 3].Text?.ToString() ?? "";
                //tb_dk2.Text = worksheet1.Cells[27, 3].Text?.ToString() ?? "";
                //tb_dk1.Text = worksheet1.Cells[28, 3].Text?.ToString() ?? "";
                //groupBox2.Text = worksheet1.Cells[13, 3].Text?.ToString() ?? "";
                //// Bật lại screen updating
                //excelApp.ScreenUpdating = true;
            }
            catch (COMException comEx)
            {
                MessageBox.Show($"Lỗi COM khi xử lý Excel:\n{comEx.Message}\n\nMã lỗi: {comEx.ErrorCode:X}\n\nVui lòng thử:\n" +
                               "1. Đóng tất cả Excel đang mở\n" +
                               "2. Chạy chương trình với quyền Administrator\n" +
                               "3. Repair Microsoft Office",
                               "Lỗi COM", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi xử lý Excel: {ex.Message}\n\nChi tiết: {ex.StackTrace}",
                               "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                //   Giải phóng tài nguyên COM theo thứ tự: Worksheet → Workbook → Application
                try
                {
                    if (worksheet1 != null)
                    {
                        Marshal.ReleaseComObject(worksheet1);
                        worksheet1 = null;
                    }

                    if (worksheet2 != null)
                    {
                        Marshal.ReleaseComObject(worksheet2);
                        worksheet2 = null;
                    }

                    if (workbook != null)
                    {
                        workbook.Close(false, Type.Missing, Type.Missing);
                        Marshal.ReleaseComObject(workbook);
                        workbook = null;
                    }

                    if (excelApp != null)
                    {
                        excelApp.Quit();
                        Marshal.ReleaseComObject(excelApp);
                        excelApp = null;
                    }

                    // Buộc garbage collection
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                }
                catch (Exception cleanupEx)
                {
                    MessageBox.Show($"Lỗi khi dọn dẹp tài nguyên: {cleanupEx.Message}");
                }


            }
        }
        //------------
        private void button5_Click_2(object sender, EventArgs e)
        {
            string filePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\TRATHONGSO.xlsl";
            var processor = new ExcelDataProcessor();
            processor.ProcessDataFromFile(filePath);
        }

        private void button11_Click(object sender, EventArgs e)
        {
            string excelFile = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\RW 依頼書.xlsx";
            string txtFile = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\test.txt";

            ExcelHelper.ExportCheckBoxesToTxt(excelFile, txtFile);
        }

        private void tb_RSow_TextChanged(object sender, EventArgs e)
        {

        }

        private void tb_RSow_Leave(object sender, EventArgs e)
        {

        }

        private void tb_RSoh_Leave(object sender, EventArgs e)
        {

        }

        private void label66_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {

        }

        private void tb_RSow_Enter(object sender, EventArgs e)
        {
          //  WriteData(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\Bang tra.xlsx");
        }

        private void tb_RSoh_Enter(object sender, EventArgs e)
        {
           // WriteData(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\Bang tra.xlsx");
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            tb_date.Text = DateTime.Now.ToString("yyyy/MM/dd");
        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void tb_RSow_KeyPress(object sender, KeyPressEventArgs e)
        {
            TextBox tb = sender as TextBox;

            // Cho phép: số, phím điều khiển (Backspace...), dấu trừ đầu tiên, dấu chấm thập phân nếu chưa có
            if (!char.IsControl(e.KeyChar) &&
                !char.IsDigit(e.KeyChar) &&
                (e.KeyChar != '.') &&
                (e.KeyChar != '-'))
            {
                e.Handled = true;
            }

            // Chỉ cho phép một dấu chấm (thập phân)
            if (e.KeyChar == '.' && tb.Text.Contains("."))
            {
                e.Handled = true;
            }

            // Dấu trừ chỉ được đứng đầu
            if (e.KeyChar == '-' && (tb.SelectionStart != 0 || tb.Text.Contains("-")))
            {
                e.Handled = true;
            }
        }

        private void tb_RSoh_KeyPress(object sender, KeyPressEventArgs e)
        {
            TextBox tb = sender as TextBox;

            // Cho phép: số, phím điều khiển (Backspace...), dấu trừ đầu tiên, dấu chấm thập phân nếu chưa có
            if (!char.IsControl(e.KeyChar) &&
                !char.IsDigit(e.KeyChar) &&
                (e.KeyChar != '.') &&
                (e.KeyChar != '-'))
            {
                e.Handled = true;
            }

            // Chỉ cho phép một dấu chấm (thập phân)
            if (e.KeyChar == '.' && tb.Text.Contains("."))
            {
                e.Handled = true;
            }

            // Dấu trừ chỉ được đứng đầu
            if (e.KeyChar == '-' && (tb.SelectionStart != 0 || tb.Text.Contains("-")))
            {
                e.Handled = true;
            }
        }

        private void RDB_auto_CheckedChanged(object sender, EventArgs e)
        {
            WriteData(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\Bang tra.xlsx");
            // cbb_Scale.Enabled = false;

        }

        private void RDB_manual_CheckedChanged(object sender, EventArgs e)
        {
            cbb_Scale.Enabled = true;
            cbb_Scale.SelectedIndex = 5;
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void tableLayoutPanel10_Paint(object sender, PaintEventArgs e)
        {

        }

        private void checkBox54_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox58_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox55_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void CBB_TUDIEN_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (CBB_TUDIEN.SelectedIndex)
            {
                case 0:
                    tb_tudien_d.Text = "270";
                    tb_tudien_r.Text = "170";
                    tb_tudien_c.Text = "450";
                    break;
                case 1:
                    tb_tudien_d.Text = "300";
                    tb_tudien_r.Text = "160";
                    tb_tudien_c.Text = "500";
                    break;
                case 2:
                    tb_tudien_d.Text = "210";
                    tb_tudien_r.Text = "170";
                    tb_tudien_c.Text = "650";
                    break;
                case 3:
                    tb_tudien_d.Text = "600";
                    tb_tudien_r.Text = "200";
                    tb_tudien_c.Text = "600";
                    break;
            }
        }

        private void btn_ex_Click(object sender, EventArgs e)
        {

        }

        private void button11_Click_1(object sender, EventArgs e)
        {
            WriteData1(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\HBリスト 251023.xlsx");
        }

        private void button10_Click(object sender, EventArgs e)
        {
            WriteData(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\Bang tra.xlsx");
        }

        private void button9_Click(object sender, EventArgs e)
        {
            ExportPDFListToExcel();
        }

        private void TB_maukaten_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (TB_maukaten.SelectedIndex)
            {
                case 0:
                    pn_color.BackColor = Color.Orange;
                    break;
                case 1:
                    pn_color.BackColor = Color.Green;

                    break;
                case 2:
                    pn_color.BackColor = Color.Blue;

                    break;
                case 3:
                    pn_color.BackColor = Color.Ivory;

                    break;
                case 4:
                    pn_color.BackColor = Color.Gray;

                    break;
                case 5:
                    pn_color.BackColor = Color.WhiteSmoke;

                    break;
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            LoadingForm loadingForm = new LoadingForm();
            Task showFormTask = Task.Run(() => loadingForm.ShowDialog());

            WriteData(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\Bang tra.xlsx");
            loadingForm.Invoke(new Action(() => loadingForm.Close()));
           
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            folderBrowserDialog.Description = "Select the file output folder";

            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {

                cbb_path.Text = folderBrowserDialog.SelectedPath;
                // readdata();

            }
        }
    }
}
