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
        List<string> listblockname1 = new List<string>() { "A3-KT-RW", "Bang TSRW" };
        List<string> listblockname2 = new List<string>() { "A3-KT-RW",  "Bang TSRW","Bang KLRW"};
        List<string> listdimstyle = new List<string>() { "WKV_15", "WKV_20", "WKV_25", "WKV_30", "WKV_35", "WKV_40", "WKV_50", "WKV_60" };
        List<string> A3_KT_RW = new List<string>() { };
        List<string> Bang_TSRW = new List<string>() { };
        List<string> Bang_KLRW= new List<string>() { };
        
        List<string>[] WKV_Array = new List<string>[] { };
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
                        Contents = note,
                        Height = textHeight * dimscale,
                        TextHeight = textHeight * dimscale,

                        ColorIndex = 2, // Màu sắc của text
                        Location = textPoint,
                        // TextStyleId = textstyle

                    };
                    mtext.LineWeight = LineWeight.LineWeight009; // Đặt độ dày của văn bản
                                                                 // mtext.EdgeStyleId\
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
                    mleader.SetTextAttachmentType(TextAttachmentType.AttachmentBottomLine, LeaderDirectionType.LeftLeader);
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
        public static void createDim(Point3d[] points, Point3d dimLinePos, ObjectId dimstyle, double angle, Point3d p1, Point3d p2, bool checkpos)
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
            
            DimStyleImporter.ImportDimStyle(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\HB-CD.dwg", "WKV_15", "WKV_20", "WKV_25", "WKV_30", "WKV_35", "WKV_40", "WKV_50", "WKV_60");
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
           
            if (RDB_auto.Checked)
            {
                cbb_Scale.Enabled = false;
            }
            else
            {
                cbb_Scale.Enabled = true;
            }
            textBoxes = new TextBox[] {tb_RSow,tb_RSoh, tb_RSa, tb_RSb, tb_RSc, tb_RSd,tb_VtP, tb_SlB, tb_qlypp,tb_vtb   };
           
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
            RegistryHelper.SaveAllSettings("RWtool-01", this);
            RegistryHelper.SaveSetting("RWtool-01", "Printer", cbb_printer.Text);
            RegistryHelper.SaveSetting("RWtool-01", "Plotstyle", cbb_plotstyle.Text);
            RegistryHelper.SaveSetting("RWtool-01", "Papersize", cbb_papersize.Text);

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

        private void tb_todayrw_Click(object sender, EventArgs e)
        {
            tb_daterw.Text = DateTime.Now.ToString("yyyy/MM/dd");
        }

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
        {
            NativeMethods.SetForegroundWindow(acadHwnd);
            WriteData(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\Bang tra.xlsx");
            var doc = acadApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;
           // loadingForm.Invoke(new Action(() => loadingForm.Close()));
            PromptPointResult pointResult = ed.GetPoint("Chọn vị trí để chèn Block Reference: ");

            //----------------------------------
            string targetName = "";
            if (RDB_manual.Checked)
            {
                targetName = listdimstyle[cbb_Scale.SelectedIndex];
                scale = double.Parse(cbb_Scale.Text.Split('/')[1]) / 20;
                //  MessageBox.Show(scale.ToString());
            }
            else
            {
                if (Length < 7800 && Height < 4000)
                {
                    if (true)
                    {
                        scale = 1.0; // Nếu chiều cao trung bình nhỏ hơn 1780, không cần scale
                        targetName = "WKV_20";
                    }
                    else
                    {
                        scale = 1.5; // Nếu chiều cao trung bình lớn hơn hoặc bằng 1780, áp dụng scale
                        targetName = "WKV_30";
                    }
                    // scale = 1.0; // Nếu tổng chiều dài và chiều cao nhỏ hơn 7800 và 4000, không cần scale
                }
                else
                {
                    scale = 1.5; // Nếu tổng chiều dài hoặc chiều cao lớn hơn, áp dụng scale
                    targetName = "WKV_30";
                }
            }
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
            A3_KT_RW = new List<string>() { tb_tencongtrinh.Text, tb_diadiem.Text, "ロールウエイ外観図",tb_maso.Text,cbb_nguoitao.Text,tb_date.Text, $"1/{Math.Round(scale * 20, 0)}" };
            Bang_TSRW = new List<string>() { 
                cbb_loairw.Text,
                tb_vlkaten.Text.Split(' ')[0],
                TB_maukaten.Text.Split(' ')[0],
                "ポリ塩化ビニール",
                (tb_vlkaten.SelectedIndex==0? "不燃認定番号" : "防炎登録番号"),
                (tb_vlkaten.SelectedIndex==0? "NM-5361" : "B1130062"),
                (cb_cst2.Checked||cb_cst3.Checked? "有り" : "無し"),




            };
            WKV_Array = new List<string>[] { A3_KT_RW};
            int ind = 0;
            List<string> listblockname_ins = new List<string>();
            if (rdb_kl1.Checked == true)
            {
                listblockname_ins = listblockname2;
            }
            else
            {
                listblockname_ins = listblockname1;
            }
            foreach (string blockname in listblockname_ins)
            {
                Insertdynamicblock.InsertBlock(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\RW.dwg", blockname, new Point3d(pointResult.Value.X, pointResult.Value.Y, 0), WKV_Array[ind], scale, 0);
                ind++;
            }

            //--------------------------------
            double x = pointResult.Value.X + 6500 * scale ;
            double y = pointResult.Value.Y + 6300 * scale ;
            Point3d p0 = new Point3d(x, y, 0);
            //Point3d p1 = new Point3d(x1, y, 0);
            //Point3d p2 = new Point3d(x2, y, 0);
            mattruoc(p0, ed);
            double.TryParse(tb_SlB.Text, out double slb);
            double.TryParse(tb_RSow.Text, out double ow);
            double.TryParse(tb_vtb.Text, out double vtb);
            double.TryParse(tb_qlypp.Text, out double P);
            double.TryParse(tb_scs.Text, out double scs);
            double kc = (ow - 2.0 * vtb) / (slb - 1.0);
            for (int i = 0; i < (slb); i++)
            {
                
                Point3d p1 = new Point3d(x- (ow-2*vtb) / 2+(i)*kc, y, 0);
                Belt(p1, ed);
            }
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


                Point3d p2_1 = new Point3d(x - ow / 2.0, y_pipe, 0);
                Point3d p2_2 = new Point3d(x + ow / 2.0 - vtb + 25, y_pipe, 0);

                Tube(p2_1, ed, vtb - 25);
                Tube(p2_2, ed, vtb - 25);
                //---------------------------------------------------------------------------
                //-----------------------------------------------------ống trong
                for (int j = 0; j < slb - 1; j++)
                {
                    Point3d p2_3 = new Point3d(x - ow / 2.0 + vtb + 25 + j * (kc), y_pipe, 0);
                    Tube(p2_3, ed, kc - 50);

                }
                //---------------------------------------------------------------------------cửa sổ
                

            }
            if (cb_cst2.Checked == true)
            {

                for (int j = 0; j < slb - 1; j++)
                {
                    if (scs == 1)
                    {

                        Point3d p3 = new Point3d(x - ow / 2.0 + vtb + 110 + j * (kc)+ (kc - 2 * 110) / 2.0, y + 800, 0);
                        Window(p3, ed, (kc - 2 * 110) / 2.0);


                    }
                    else if (scs == 2)
                    {
                        Point3d p3_1 = new Point3d(x - ow / 2.0 + vtb + 110 + j * (kc)+ (kc / 2 - 110 - 125) / 2.0, y + 800, 0);
                        Window(p3_1, ed, (kc / 2 - 110 - 125) / 2.0);
                        Point3d p3_2 = new Point3d(x - ow / 2.0 + vtb + 110 + j * (kc) + (kc / 2 - 110 - 125) + 250+ (kc / 2 - 110 - 125) / 2.0, y + 800, 0);
                        Window(p3_2, ed, (kc / 2 - 110 - 125) / 2.0);
                    }
                }

            }
            if (cb_cst3.Checked == true)
            {
                for (int j = 0; j < slb - 1; j++)
                {
                    if (scs == 1)
                    {

                        Point3d p3 = new Point3d(x - ow / 2.0 + vtb + 110 + j * (kc)+(kc - 2 * 110) / 2.0, y + 800 + 600, 0);
                        Window(p3, ed, (kc - 2 * 110) / 2.0);


                    }
                    else if (scs == 2)
                    {
                        Point3d p3_1 = new Point3d(x - ow / 2.0 + vtb + 110 + j * (kc) + (kc / 2 - 110 - 125) / 2.0, y + 800 + 600, 0);
                        Window(p3_1, ed, (kc / 2 - 110 - 125) / 2.0);
                        Point3d p3_2 = new Point3d(x - ow / 2.0 + vtb + 110 + j * (kc) + (kc / 2 - 110 - 125) + 250 + (kc / 2 - 110 - 125) / 2.0, y + 800 + 600, 0);
                        Window(p3_2, ed, (kc / 2 - 110 - 125) / 2.0);
                    }
                }
            }


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
        private void Belt(Point3d point1, Editor ed)
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
            dimstyle_text = Insertdynamicblock.InsertDynamicBlock(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\RW.dwg", "Belt", point1, dynProps/*, cbb_test*/, dimstyle, 1.0);
            values.Clear();
        }
        private void Tube(Point3d point1, Editor ed, double value)
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

            { "T1", value }




            };
            dimstyle_text = Insertdynamicblock.InsertDynamicBlock(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\RW.dwg", "Tube", point1, dynProps/*, cbb_test*/, dimstyle, 1.0);
            values.Clear();
        }
        private void Window(Point3d point1, Editor ed, double value)
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
            dimstyle_text = Insertdynamicblock.InsertDynamicBlock(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\RW.dwg", "Window", point1, dynProps/*, cbb_test*/, dimstyle, 1.0);
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
                tb_ong2.Text = worksheet1.Cells[35, 3].Text?.ToString() ?? "";
                tb_v.Text = worksheet1.Cells[36, 3].Text?.ToString() ?? "";
                tb_dk2.Text = worksheet1.Cells[27, 3].Text?.ToString() ?? "";
                tb_dk1.Text = worksheet1.Cells[28, 3].Text?.ToString() ?? "";

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
            WriteData(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\Bang tra.xlsx");
        }

        private void tb_RSoh_Enter(object sender, EventArgs e)
        {
           WriteData(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\Bang tra.xlsx");
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
            cbb_Scale.Enabled = false;
        }

        private void RDB_manual_CheckedChanged(object sender, EventArgs e)
        {
            cbb_Scale.Enabled = true;
            cbb_Scale.SelectedIndex = 5;
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }
    }
}
