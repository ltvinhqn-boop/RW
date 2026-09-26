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
        string D = "";
        string motor = "";
        string gearbox = "";
        string qlypp;
        string spacing = "";
        List<string> liststylename = new List<string> { };
        //----------------------List Blockname
        List<string> listblockname1 = new List<string>() { "A3-KT-RW", "Bang TSRW" };
        List<string> listblockname2 = new List<string>() { "A3-KT-RW", "Bang KLRW", "Bang TSRW" };
        List<string> A3_KT_RW = new List<string>() { };
        List<string> Bang_TSRW = new List<string>() { };
        List<string> Bang_KLRW= new List<string>() { };
        //List<string> WKV_KT2 = new List<string>() { };
        //List<string> WKV_KT3 = new List<string>() { };
        //List<string> WKV_KT4 = new List<string>() { };
        //List<string> WKV_KT5 = new List<string>() { };
        //List<string> WKV_KT6 = new List<string>() { };
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
        //List<string> listblockname = new List<string>() { "HB-A3-CL", "WKV-KT1", "WKV-KT2", "WKV-KT3", "WKV-KT4", "WKV-KT5", "WKV-KT6" };
        //List<string> listblockname2 = new List<string>() { "HB-A3-LS", "WKV-KT1", "WKV-KT2", "WKV-KT3", "WKV-KT4", "WKV-KT5", "WKV-KT6" };
        //List<string> HB_A3_CL = new List<string>() { };
        //List<string> WKV_KT1 = new List<string>() { };
        //List<string> WKV_KT2 = new List<string>() { };
        //List<string> WKV_KT3 = new List<string>() { };
        //List<string> WKV_KT4 = new List<string>() { };
        //List<string> WKV_KT5 = new List<string>() { };
        //List<string> WKV_KT6 = new List<string>() { };
        //List<string>[] WKV_Array = new List<string>[] { };
        List<SizeView1> sizeview = new List<SizeView1>();
        public List<Extents2d> liextent2d = new List<Extents2d>();
        int[,] array_hole = new int[2, 3];
        double gapx = 0.0;
        double gapy = 0.0;
        ObjectId dimstyle_text = ObjectId.Null;
        ObjectId dimstyle;
        string filePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\HB特注原価計算.xlsx";
        double Length = 0.0;
        double Height = 0.0;
        int total914 = 0;
        int total1219 = 0;
        Hatch hatch = new Hatch();
        // ExcelPackage package;// = new ExcelPackage(new FileInfo(filePath));
        // Listx
        string tempFilePath;
        public Form1()
        {
            InitializeComponent();
            this.AllowDrop = true;
            textBoxes = new TextBox[]
           {
    tb_a, tb_b, tb_c, tb_j,
    tb_d, tb_e, tb_f, tb_g,
    tb_h, tb_i, tb_k, tb_deg
           };
            buttons = new Button[]
            {
                btn_front, btn_right,btn_top,btn_left,btn_back
            };
            tb_w.TextChanged += ValidateInput;
            tb_dh.TextChanged += ValidateInput;
            tb_c.KeyDown += ValidateInput;
            tb_j.TextChanged += ValidateInput;
            tb_caotong.KeyDown += ValidateInput1;
            rdb_3.CheckedChanged += SelectView;
            rdb_4.CheckedChanged += SelectView;
            // rdb_5.CheckedChanged += SelectView;


        }
        private void SelectView(object sender, EventArgs e)
        {
            foreach (var btn in buttons)
            {
                btn.BackColor = Color.White;
            }
            int slindex = 3;
            if (rdb_3.Checked)
            {
                slindex = 3;
            }
            else if (rdb_4.Checked)
            {
                slindex = 4;
            }
            else if (rdb_5.Checked)
            {
                slindex = 5;
            }

            for (int i = 0; i < slindex; i++)
            {
                Button button = buttons[i];
                button.BackColor = Color.FromArgb(255, 224, 192);


            }
        }
        private void SelectView1()
        {
            foreach (var btn in buttons)
            {
                btn.BackColor = Color.White;
            }
            int slindex = 3;
            if (rdb_3.Checked == true)
            {
                slindex = 3;
            }
            else if (rdb_4.Checked == true)
            {
                slindex = 4;
            }
            else if (rdb_5.Checked == true)
            {
                slindex = 5;
            }

            for (int i = 0; i < slindex; i++)
            {
                Button button = buttons[i];
                button.BackColor = Color.FromArgb(255, 224, 192);


            }
        }
        private void ValidateInput(object sender, EventArgs e)
        {
            bool allValid = IsNumeric(tb_w.Text)
                         && IsNumeric(tb_dh.Text)
                         && IsNumeric(tb_c.Text)
                         && IsNumeric(tb_j.Text);
            if (allValid) { outputdata(); }
            //  writedata();


        }
        private void ValidateInput1(object sender, EventArgs e)
        {
            bool allValid = IsNumeric(tb_w.Text)
                         && IsNumeric(tb_dh.Text)
                         && IsNumeric(tb_caotong.Text)
                         && IsNumeric(tb_j.Text);
            if (allValid) { outputdata2(); }
            //  writedata();

        }
        public void outputdata()
        {
            // Lấy giá trị từ các TextBox
            Double.TryParse(tb_i.Text, out Double i);
            Double.TryParse(tb_w.Text, out Double w);
            Double.TryParse(tb_dh.Text, out Double dh);
            Double.TryParse(tb_c.Text, out Double c);
            Double.TryParse(tb_j.Text, out Double j);
            Double.TryParse(tb_deg.Text, out Double deg);
            double.TryParse(tb_m.Text, out Double m);
            tb_a.Text = (w + m * 2).ToString();
            tb_b.Text = (dh + m * 2).ToString();
            tb_caotong.Text = Math.Round(j + c + ((dh + m * 2 + i * 2) / 2.0) * Math.Tan(deg * Math.PI / 180.0), 0, MidpointRounding.AwayFromZero).ToString();
        }
        public void outputdata2()
        {
            Double.TryParse(tb_i.Text, out Double i);
            Double.TryParse(tb_w.Text, out Double w);
            Double.TryParse(tb_dh.Text, out Double dh);
            Double.TryParse(tb_c.Text, out Double c);
            Double.TryParse(tb_j.Text, out Double j);
            Double.TryParse(tb_deg.Text, out Double deg);
            Double.TryParse(tb_caotong.Text, out Double caotong);
            double.TryParse(tb_m.Text, out Double m);
            tb_a.Text = (w + m * 2).ToString();
            tb_b.Text = (dh + m * 2).ToString();
            tb_c.Text = Math.Round(caotong - j - ((dh + m * 2 + i * 2) / 2.0) * Math.Tan(deg * Math.PI / 180.0), 0, MidpointRounding.AwayFromZero).ToString();
            // tb_c.Text = (caotong - m * 2).ToString();
        }

        private bool IsNumeric(string text)
        {
            return double.TryParse(text, out _);
        }
        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            //string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            //string filepath = files[0];
            //// string[] ex = filepath.Split('.');

            //// In kết quả
            //if (Path.GetExtension(filepath).ToLower() == ".xlsx")
            //{
            //    comboBox1.Text = filepath;
            //    comboBox1.ForeColor = Color.Black;
            //    readdata();
            //}
            //else
            //{
            //    MessageBox.Show("Vui lòng kéo thả File có đuôi mở rộng .xlsx");
            //}
        }

        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            //if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
        }
        private void mattruoc(Point3d point1, Editor ed)
        {

            foreach (var tb in textBoxes)
            {
                if (double.TryParse(tb.Text, out double result))
                {
                    values.Add(result);
                }
                else
                {
                    // Nếu muốn bỏ qua lỗi parse, có thể ghi log ở đây
                    values.Add(0); // hoặc throw / log warning
                }
            }


            //var doc = acadApp.DocumentManager.MdiActiveDocument;
            //var db = doc.Database;
            //var ed = doc.Editor;
            // int PP = dtb_gr1.RowCount;
            //int p = 1;
            ////  dtb_gr1.Rows.RemoveAt(PP-1);
            ////PromptPointResult pointResult = ed.GetPoint("Chọn vị trí để chèn Block Reference: ");
            //double x = pointResult.Value.X;
            //double y = pointResult.Value.Y;
            Double.TryParse(tb_deg.Text, out Double deg);
            Double.TryParse(tb_c.Text, out Double c);
            // Double.TryParse(tb_test1.Text, out Double th);
            Double.TryParse(tb_w.Text, out Double w);
            Double.TryParse(tb_dh.Text, out Double dh);
            Double.TryParse(tb_i.Text, out Double i);
            Double.TryParse(tb_j.Text, out Double j);

            Double.TryParse(tb_b.Text, out Double b);
            Double.TryParse(tb_f.Text, out Double f);
            Double.TryParse(tb_g.Text, out Double g);
            double.TryParse(tb_k.Text, out Double k);
            double.TryParse(tb_a.Text, out Double a);
            double.TryParse(tb_i.Text, out Double i1);
            double.TryParse(tb_d.Text, out Double d);
            double.TryParse(tb_e.Text, out Double e1);
            double.TryParse(tb_m.Text, out Double m);
            Double ct = j + c + ((dh + m * 2 + i * 2) / 2.0) * Math.Tan(deg * Math.PI / 180.0);
            var dynProps = new Dictionary<string, double>
            {
            { "I1", values[9] },
            { "I2", values[9] },
            { "E", values[5] },
            { "D", values[4] },
            { "J", values[3] },
            //{ "P1", values[12] },
            //{ "P2", values[12] },
            { "H", values[8] },
            { "C", values[2] },
            { "K1", values[10] },
            { "K2", values[10] },
            { "A1", values[0]/2 },
            { "A2", values[0]/2 },
            { "T", 30 * Math.Cos(deg * Math.PI / 180.0)},
                {"W2",RoundToNearest( c-90-60)  },
                {"CT", ct },
                {"M1",m },
                {"M2",m }



            };
            dimstyle_text = Insertdynamicblock.InsertDynamicBlock(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\HB-CD.dwg", "HB-MT", point1, dynProps/*, cbb_test*/, dimstyle, out hatch, 1.0);
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
        private void matsau(Point3d point1, Editor ed)
        {

            foreach (var tb in textBoxes)
            {
                if (double.TryParse(tb.Text, out double result))
                {
                    values.Add(result);
                }
                else
                {
                    // Nếu muốn bỏ qua lỗi parse, có thể ghi log ở đây
                    values.Add(0); // hoặc throw / log warning
                }
            }


            //var doc = acadApp.DocumentManager.MdiActiveDocument;
            //var db = doc.Database;
            //var ed = doc.Editor;
            // int PP = dtb_gr1.RowCount;
            //int p = 1;
            ////  dtb_gr1.Rows.RemoveAt(PP-1);
            ////PromptPointResult pointResult = ed.GetPoint("Chọn vị trí để chèn Block Reference: ");
            //double x = pointResult.Value.X;
            //double y = pointResult.Value.Y;
            Double.TryParse(tb_deg.Text, out Double deg);
            Double.TryParse(tb_c.Text, out Double c);
            //  Double.TryParse(tb_test1.Text, out Double th);
            Double.TryParse(tb_w.Text, out Double w);
            Double.TryParse(tb_dh.Text, out Double dh);
            Double.TryParse(tb_i.Text, out Double i);
            Double.TryParse(tb_j.Text, out Double j);

            Double.TryParse(tb_b.Text, out Double b);
            Double.TryParse(tb_f.Text, out Double f);
            Double.TryParse(tb_g.Text, out Double g);
            double.TryParse(tb_k.Text, out Double k);
            double.TryParse(tb_a.Text, out Double a);
            double.TryParse(tb_i.Text, out Double i1);
            double.TryParse(tb_d.Text, out Double d);
            double.TryParse(tb_e.Text, out Double e1);
            double.TryParse(tb_m.Text, out Double m);
            Double ct = j + c + ((dh + m * 2 + i * 2) / 2.0) * Math.Tan(deg * Math.PI / 180.0);
            var dynProps = new Dictionary<string, double>
            {
            { "I1", values[9] },
            { "I2", values[9] },
            { "E", values[5] },
            { "D", values[4] },
            { "J", values[3] },
            //{ "P1", values[12] },
            //{ "P2", values[12] },
            { "H", values[8] },
            { "C", values[2] },
            { "K1", values[10] },
            { "K2", values[10] },
            { "A1", values[0]/2 },
            { "A2", values[0]/2 },
            { "T", 30 * Math.Cos(deg * Math.PI / 180.0)},
                {"W2", RoundToNearest( c-90-60) },
                {"CT", ct },
                {"M1",m },
                {"M2",m }



            };
            dimstyle_text = Insertdynamicblock.InsertDynamicBlock(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\HB-CD.dwg", "HB-MS", point1, dynProps, dimstyle/*, cbb_test*/, out hatch, 1.0);
        }
        private void mattren(Point3d point1, Editor ed)
        {
            values.Clear();
            foreach (var tb in textBoxes)
            {
                if (double.TryParse(tb.Text, out double result))
                {
                    values.Add(result);
                }
                else
                {

                    values.Add(0); // hoặc throw / log warning
                }
            }
            // Double.TryParse(tb_test1.Text, out Double th);
            double W = double.Parse(tb_w.Text);
            double DH = double.Parse(tb_dh.Text);
            double.TryParse(tb_m.Text, out Double m);
            var dynProps = new Dictionary<string, double>
            {
            { "W1", W/2},
            { "W2", W/2},
            { "DH1", DH/2 },
            { "DH2", DH/2 },
                { "N1", m },
                { "N2", m },
                { "N3", m },
                { "N4", m }

            };
            Insertdynamicblock.InsertDynamicBlock(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\HB-CD.dwg", "HB-MTR", point1, dynProps/*, cbb_test*/, dimstyle, out hatch, 1.0);
            AutoHoles(point1.X, point1.Y);

        }
        private void mattren2(Point3d point1, Editor ed)
        {
            values.Clear();
            foreach (var tb in textBoxes)
            {
                if (double.TryParse(tb.Text, out double result))
                {
                    values.Add(result);
                }
                else
                {

                    values.Add(0); // hoặc throw / log warning
                }
            }
            // Double.TryParse(tb_test1.Text, out Double th);
            double W = double.Parse(tb_w.Text);
            double DH = double.Parse(tb_dh.Text);
            double.TryParse(tb_m.Text, out Double m);
            var dynProps = new Dictionary<string, double>
            {
            { "W1", W/2},
            { "W2", W/2},
            { "DH1", DH/2 },
            { "DH2", DH/2 },
                { "N1", m },
                { "N2", m },
                { "N3", m },
                { "N4", m }

            };
            Insertdynamicblock.InsertDynamicBlock(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\HB-CD.dwg", "HB-MTR-2", point1, dynProps/*, cbb_test*/, dimstyle, out hatch, 1.0);
            AutoHoles(point1.X, point1.Y);

        }
        private void matphai(Point3d point1, Editor ed)
        {
            values.Clear();
            foreach (var tb in textBoxes)
            {
                if (double.TryParse(tb.Text, out double result))
                {
                    values.Add(result);
                }
                else
                {
                    // Nếu muốn bỏ qua lỗi parse, có thể ghi log ở đây
                    values.Add(0); // hoặc throw / log warning
                }
            }
            //  Double.TryParse(tb_test1.Text, out Double th);
            //Double.TryParse(tb_c.Text, out Double c);
            double W = double.Parse(tb_w.Text);
            double DH = double.Parse(tb_dh.Text);
            // Lấy giá trị từ các TextBox
            Double.TryParse(tb_i.Text, out Double i);
            Double.TryParse(tb_w.Text, out Double w);
            Double.TryParse(tb_dh.Text, out Double dh);
            Double.TryParse(tb_c.Text, out Double c);
            Double.TryParse(tb_j.Text, out Double j);
            Double.TryParse(tb_deg.Text, out Double deg);
            double.TryParse(tb_m.Text, out Double m);
            double x2 = dh / 2 + 50 + m + w / 2 + m + 50 + 600;
            Double ct = j + c + ((dh + m * 2 + i * 2) / 2.0) * Math.Tan(deg * Math.PI / 180.0);
            double cp1 = ct - (dh / 2 + m) * Math.Tan(deg * Math.PI / 180.0) - 30 / Math.Cos(deg * Math.PI / 180.0) - j;
            var dynProps = new Dictionary<string, double>
            {
            { "B1", DH/2+m},
            { "B2", DH/2+m},
            { "E", values[5] },
            { "D", values[4] },
            { "K1", values[10] },
            { "K2", values[10] },
            { "CT", ct },
            { "J", values[3] },
            { "CP1", cp1 },
            {"W2", RoundToNearest( c-90-60) },
                {"M1",m },
                {"M2",m }
            };
            Insertdynamicblock.InsertDynamicBlock(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\HB-CD.dwg", "HB-CC", point1, dynProps/*, cbb_test*/, dimstyle, out hatch, 1.0);
            DrawLid(point1.X, point1.Y);
            //  MessageBox.Show($"CT: {ct}, CP1: {cp1}");
        }
        private void mattrai(Point3d point1, Editor ed)
        {
            values.Clear();
            foreach (var tb in textBoxes)
            {
                if (double.TryParse(tb.Text, out double result))
                {
                    values.Add(result);
                }
                else
                {
                    // Nếu muốn bỏ qua lỗi parse, có thể ghi log ở đây
                    values.Add(0); // hoặc throw / log warning
                }
            }
            //Double.TryParse(tb_test1.Text, out Double th);
            //Double.TryParse(tb_c.Text, out Double c);
            double W = double.Parse(tb_w.Text);
            double DH = double.Parse(tb_dh.Text);
            // Lấy giá trị từ các TextBox
            Double.TryParse(tb_i.Text, out Double i);
            Double.TryParse(tb_w.Text, out Double w);
            Double.TryParse(tb_dh.Text, out Double dh);
            Double.TryParse(tb_c.Text, out Double c);
            Double.TryParse(tb_j.Text, out Double j);
            Double.TryParse(tb_deg.Text, out Double deg);
            double.TryParse(tb_m.Text, out Double m);
            double x2 = dh / 2 + 50 + m + w / 2 + m + 50 + 600;
            Double ct = j + c + ((dh + m * 2 + i * 2) / 2.0) * Math.Tan(deg * Math.PI / 180.0);
            double cp1 = ct - (dh / 2 + m) * Math.Tan(deg * Math.PI / 180.0) - 30 / Math.Cos(deg * Math.PI / 180.0) - j;
            var dynProps = new Dictionary<string, double>
            {
            { "B1", DH/2+m},
            { "B2", DH/2+m},
            { "E", values[5] },
            { "D", values[4] },
            { "K1", values[10] },
            { "K2", values[10] },
            { "CT", ct },
            { "J", values[3] },
            { "CP1", cp1 },
            {"W2", RoundToNearest( c-90-60) },
            {"M1",m },
                {"M2",m }
            };
            Insertdynamicblock.InsertDynamicBlock(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\HB-CD.dwg", "HB-MTT", point1, dynProps/*, cbb_test*/, dimstyle, out hatch, 1.0);
            DrawLid(point1.X, point1.Y);
            //  MessageBox.Show($"CT: {ct}, CP1: {cp1}");
        }
        private void button1_Click(object sender, EventArgs e)
        {
            NativeMethods.SetForegroundWindow(acadHwnd);
            values.Clear();
            if (tb_ken.Text == "")
            {
                MessageBox.Show("Vui lòng nhập tên công trình");
                return;
            }
            LoadingForm loadingForm = new LoadingForm();
            Task showFormTask = Task.Run(() => loadingForm.ShowDialog());
            btn_update.PerformClick();
            Length = 0.0;
            Height = 0.0;



            var doc = acadApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;
            // int PP = dtb_gr1.RowCount;
            //int p = 1;
            //  dtb_gr1.Rows.RemoveAt(PP-1);
            loadingForm.Invoke(new Action(() => loadingForm.Close()));
            PromptPointResult pointResult = ed.GetPoint("Chọn vị trí để chèn Block Reference: ");
            Double.TryParse(tb_w.Text, out Double w);
            Double.TryParse(tb_dh.Text, out Double dh);
            Double.TryParse(tb_c.Text, out Double c);
            Double.TryParse(tb_j.Text, out Double j);
            Double.TryParse(tb_deg.Text, out Double deg);
            Double.TryParse(tb_b.Text, out Double b);
            Double.TryParse(tb_f.Text, out Double f);
            Double.TryParse(tb_g.Text, out Double g);
            double.TryParse(tb_k.Text, out Double k);
            double.TryParse(tb_a.Text, out Double a);
            // double.TryParse(tb_i.Text, out Double i1);
            double.TryParse(tb_d.Text, out Double d);
            double.TryParse(tb_e.Text, out Double e1);
            double.TryParse(tb_i.Text, out Double i);
            double.TryParse(tb_m.Text, out Double m);
            Double ct = j + c + ((dh + m * 2 + i * 2) / 2.0) * Math.Tan(deg * Math.PI / 180.0);
            //----------------------------------

            SizeView1 Frontview = new SizeView1()
            {
                Size_L = a / 2 + 225,
                Size_R = a / 2 + 440,
                Size_H = ct + 500 + 500
            };
            SizeView1 Backview = new SizeView1()
            {
                Size_L = a / 2 + 225,
                Size_R = a / 2 + 440,
                Size_H = ct + 500 + 500
            };
            SizeView1 Rightview = new SizeView1()
            {
                Size_L = b / 2 + 235,
                Size_R = b / 2 + 320,
                Size_H = ct + 500 + 500
            };
            SizeView1 Lefttview = new SizeView1()
            {
                Size_L = b / 2 + 235,
                Size_R = b / 2 + 320,
                Size_H = ct + 500 + 500
            };
            SizeView1 Topview = new SizeView1()
            {
                Size_L = a / 2 + 1100,
                Size_R = a / 2 + 800,
                Size_H = b + 370 + 370
            };
            sizeview = new List<SizeView1>()
            {
                Frontview,
                Rightview,
                Backview,
                Lefttview,
                Topview
            };
            // Màu cần kiểm tra
            Color targetColor = Color.FromArgb(255, 224, 192);

            // Tạo list bool theo điều kiện
            List<bool> colorMatches = new List<bool>
            {
             btn_front.BackColor == targetColor,
             btn_right.BackColor == targetColor,
             btn_back.BackColor == targetColor,
             btn_left.BackColor == targetColor,
             btn_top.BackColor == targetColor
            };
            gapx = double.Parse(tb_gapx.Text);
            gapy = double.Parse(tb_gapy.Text);
            Length = gapx; // Khởi tạo Length về 0 trước khi tính toán
            Height = 0.0; // Khởi tạo Height về 0 trước khi tính toán
            string targetName = "";
            double heightSum = 0.0; // Biến tạm để tính tổng chiều cao
            double lengthSum = 0.0; // Biến tạm để tính tổng chiều dài
                                    // Khoảng cách giữa các kích thước
                                    // Duyệt qua danh sách sizeview và cộng kích thước tương ứng vào Length và Height
            int index_view = 0; // Biến để theo dõi chỉ số view hiện tại
            for (int i_view = 0; i_view < sizeview.Count; i_view++)
            {
                if (colorMatches[i_view] && i_view != sizeview.Count - 1)
                {
                    // Nếu màu trùng, thêm kích thước tương ứng vào Length
                    Length += sizeview[i_view].Size_L + sizeview[i_view].Size_R + gapx; // Cộng thêm 600 cho khoảng cách
                    heightSum += sizeview[i_view].Size_H + gapy; // Cộng thêm kích thước chiều cao
                    index_view += 1; // Tăng chỉ số view hiện tại
                }
                if (colorMatches[i_view] && i_view == sizeview.Count - 1)
                {
                    // Nếu là view cuối cùng, không cộng thêm khoảng cách
                    Height += sizeview[i_view].Size_H + gapy; // Cộng thêm kích thước chiều cao
                }
                if (colorMatches[i_view] && i_view != sizeview.Count - 1 && i_view != 1)
                {
                    lengthSum += sizeview[i_view].Size_L + sizeview[i_view].Size_R + gapx; // Cộng thêm kích thước chiều dài
                }
            }
            if (index_view != 0)
            {
                Height = heightSum / index_view + Height; // Cập nhật tổng chiều cao chia cho số lượng view đã chọn 
            }
            if (Length < 7800 && Height < 4000)
            {
                if (index_view != 0 && (heightSum / index_view - gapy) < 1780 && lengthSum < 5700)
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
            //----------------------------------
            //if (btn_front.BackColor == Color.FromArgb(255, 224, 192))
            //{

            //    Length = Length + a + i * 2;
            //}
            //if (btn_top.BackColor == Color.FromArgb(255, 224, 192))
            //{
            //   // mattren(p3, ed);
            //}
            //if (btn_right.BackColor == Color.FromArgb(255, 224, 192))
            //{
            //   // matphai(p4, ed);
            //    Length = Length + b + i * 2 + 750;
            //}
            //if (btn_left.BackColor == Color.FromArgb(255, 224, 192) && btn_back.BackColor == Color.FromArgb(255, 224, 192))
            //{
            //    //mattrai(p2, ed);
            //    Length = Length + b + i * 2 + 750;
            //   // matsau(p5, ed);
            //    Length = Length + a + i * 2 + 750;
            //}
            //else if (btn_back.BackColor == Color.FromArgb(255, 224, 192))
            //{
            //  //  matsau(p2, ed);
            //    Length = Length + a + i * 2 + 750;
            //}
            //else if (btn_left.BackColor == Color.FromArgb(255, 224, 192))
            //{
            //    //mattrai(p2, ed);
            //    Length = Length + b + i * 2 + 750;
            //}
            //double giaso1 = 0.0;
            //double giaso2 = 0.0;
            //Length += 600 * 2;

            //    if (Length > 7800) {
            //    scale = 1.5;
            //    giaso1 = (dh + w + m * 2 + 50 + 600)/2-900; // Tính toán giá trị giaso1 dựa trên chiều dài
            //    giaso2 = 500;
            //    targetName = "WKV_30";
            //}          
            //else { scale = 1.0;
            //    targetName = "WKV_20";
            //}
            // -------------------------
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
            // Tính toán các giá trị cần thiết

            double x = (index_view == 4) ? (pointResult.Value.X + (7800 - 2125) * scale - sizeview[4].Size_R) : (pointResult.Value.X + (7800 / 2) * scale);
            double y = pointResult.Value.Y + 1640 * scale + 500;
            double x1 = x + (sizeview[0].Size_R + sizeview[1].Size_L + gapx);
            double x2 = x - (sizeview[0].Size_L + sizeview[3].Size_R + gapx);
            double x3 = x2 - (sizeview[2].Size_R + sizeview[3].Size_L + gapx);
            double x4 = x - (sizeview[0].Size_L + sizeview[2].Size_R + gapx);
            double y2 = y + sizeview[0].Size_H - 500 + gapy + sizeview[4].Size_H / 2;
            double L = w + m * 2 + k * 2 - d * 2;
            double H = dh + m * 2 + k * 2 - f * 2;
            Point3d p0 = new Point3d(x, y, 0);
            Point3d p1 = new Point3d(x1, y, 0);
            Point3d p2 = new Point3d(x2, y, 0);
            Point3d p3 = new Point3d(x3, y, 0);
            Point3d p4 = new Point3d(x, y2, 0);
            Point3d p5 = new Point3d(x4, y, 0);
            if (btn_front.BackColor == Color.FromArgb(255, 224, 192))
            {
                mattruoc(p0, ed);

            }
            if (btn_top.BackColor == Color.FromArgb(255, 224, 192))
            {
                if (rbt_lapsau.Checked == true)
                {
                    mattren2(p4, ed);
                }
                else
                {
                    mattren(p4, ed);
                }

            }
            if (btn_right.BackColor == Color.FromArgb(255, 224, 192))
            {
                matphai(p1, ed);
                //  Length = Length + b + i * 2+750;
            }
            if (btn_left.BackColor == Color.FromArgb(255, 224, 192) && btn_back.BackColor == Color.FromArgb(255, 224, 192))
            {
                mattrai(p2, ed);
                // Length = Length + b + i * 2 + 750;
                matsau(p3, ed);
                // Length = Length + a + i * 2+750;
            }
            else if (btn_back.BackColor == Color.FromArgb(255, 224, 192))
            {
                matsau(p5, ed);
                //Length = Length + a + i * 2 + 750;
            }
            else if (btn_left.BackColor == Color.FromArgb(255, 224, 192))
            {
                mattrai(p2, ed);
                // Length = Length + b + i * 2 + 750;
            }
            // Length += 800 * 2;
            // if (Length > 7800) { scale = 1.5; }


            //HB_A3_CL = new List<string>() { };
            //WKV_KT1 = new List<string>() { tb_date.Text, tb_ken.Text, tb_hbname.Text };

            double ss1 = a + i * 2;
            double ss2 = a - d - e1 + 2 * k;
            string selectedValue1 = (ss1 > ss2) ? ss1.ToString() : ss2.ToString();
            double ss3 = b + i * 2;
            double ss4 = b - f - g + 2 * k;
            string selectedValue2 = (ss3 > ss4) ? ss3.ToString() : ss4.ToString();

            //WKV_KT3 = new List<string> { selectedValue1, selectedValue2, tb_caotong.Text };
            //WKV_KT4 = new List<string>() { lb_tong.Text, lb_de.Text, lb_than.Text, lb_nap.Text };
            //WKV_KT5 = new List<string>() { $"屋根スラブ配管取出最大開口寸法 {w},{dh} x {RoundToNearest(c - 90 - 60)}", "配管箱開口部最大開口寸法 " + w + " x " + dh };
            //WKV_KT6 = new List<string>() { $"1/{Math.Round(scale * 20, 0)}" };
            //WKV_Array = new List<string>[] { HB_A3_CL, WKV_KT1, WKV_KT2, WKV_KT3, WKV_KT4, WKV_KT5, WKV_KT6 };
            int ind = 0;
            List<string> listblockname_ins = new List<string>();
            Point3d insertpoint = new Point3d(pointResult.Value.X, pointResult.Value.Y, 0);
            if (cb_bangKL.Checked == true)
            {
                listblockname_ins = listblockname2;
            }
            else
            {
                listblockname_ins = listblockname1;
            }
            foreach (string blockname in listblockname_ins)
            {
                insertpoint = blockname== "Bang KLRW" ? new Point3d(pointResult.Value.X, pointResult.Value.Y + 998.5*scale, 0) : new Point3d(pointResult.Value.X, pointResult.Value.Y, 0);
                Insertdynamicblock.InsertBlock(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\HB-CD.dwg", blockname, new Point3d(pointResult.Value.X, pointResult.Value.Y, 0), WKV_Array[ind], scale, 0);
                ind++;
            }

        }
        public void AutoHoles(double x, double y)
        {
            List<Point3d> pointholes = new List<Point3d>();

            Double.TryParse(tb_w.Text, out Double w);
            Double.TryParse(tb_dh.Text, out Double dh);
            Double.TryParse(tb_c.Text, out Double c);
            Double.TryParse(tb_j.Text, out Double j);
            Double.TryParse(tb_deg.Text, out Double deg);
            Double.TryParse(tb_b.Text, out Double b);
            Double.TryParse(tb_f.Text, out Double f);
            Double.TryParse(tb_g.Text, out Double g);
            double.TryParse(tb_k.Text, out Double k);
            double.TryParse(tb_a.Text, out Double a);
            double.TryParse(tb_i.Text, out Double i1);
            double.TryParse(tb_d.Text, out Double d);
            double.TryParse(tb_e.Text, out Double e1);
            double.TryParse(tb_i.Text, out Double i);
            double.TryParse(tb_m.Text, out Double m);
            Double ct = j + c + ((dh + m * 2 + i * 2) / 2.0) * Math.Tan(deg * Math.PI / 180.0);
            // x = pointResult.Value.X ;
            //  y = pointResult.Value.Y;
            double x1 = double.Parse(tb_a.Text) + 2 * double.Parse(tb_k.Text) + 600;
            double x2 = dh / 2 + 50 + m + w / 2 + m + 50 + 600;
            double y2 = ct + b / 2.0 - f - g + 2 * k + 600;
            double L = w + m * 2 + k * 2 - d * 2;
            double H = dh + m * 2 + k * 2 - f * 2;


            double fixedStartOffset;
            if (rbt_lapsau.Checked == true)
            {
                fixedStartOffset = 25;
            }
            else
            {
                fixedStartOffset = 40;
            }
            var holes = GetHoleBorderPoints(xStart: x - L / 2, xLength: L, yStart: y - H / 2, yLength: H, fixedStartOffset, fixedEndOffset: 110, maxSpacing: 350);
            foreach (Point3d point in holes)
            {
                Insertdynamicblock.InsertBlock(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\HB-CD.dwg", "fai7+", point, new List<string>(), scale, 0);

            }
            var holes2 = GetHoleBorderPoints2(xStart: x - L / 2, xLength: L, yStart: y - H / 2, yLength: H, fixedStartOffset: 109, fixedEndOffset: 100, maxSpacing: 350);
            foreach (Point3d point in holes2)
            {
                Insertdynamicblock.InsertBlock(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\HB-CD.dwg", "fai9", point, new List<string>(), scale, 0);

            }
            var holes3 = GetHoleBorderPoints3(xStart: x - L / 2, xLength: L, yStart: y - H / 2, yLength: H, fixedStartOffset: 36.6, fixedEndOffset: 100, maxSpacing: 350);
            foreach (Point3d point in holes3)
            {
                Insertdynamicblock.InsertBlock(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\HB-CD.dwg", "fai8", point, new List<string>(), scale, 0);

            }
            List<Point3d> ListPy = new List<Point3d>();
            List<Point3d> ListPx = new List<Point3d>();
            Dimentionholes1(holes, x - L / 2, y - H / 2, fixedStartOffset, out ListPx, out ListPy);//đế
            //// Fix for CS0022: Wrong number of indices inside []; expected 2
            //// The issue occurs because `array_hole` is declared as a 2D array, but the code is trying to access it as if it were a jagged array.
            //// Correcting the indexing to use two indices for the 2D array.
            //List<Point3d> ListPy2 = new List<Point3d>();
            //List<Point3d> ListPx2 = new List<Point3d>();
            //Dimentionholes2(holes2, x - L / 2, y - H / 2, 110, out ListPx2, out ListPy2);//thân
            //List<Point3d> ListPy3 = new List<Point3d>();
            //List<Point3d> ListPx3 = new List<Point3d>();
            //Dimentionholes2(holes3, x - L / 2, y - H / 2, 40, out ListPx3, out ListPy3);//nắp
            //array_hole[0, 0] = ListPx.Count-4;
            //array_hole[1, 0] = ListPy.Count-2;
            //array_hole[0, 1] = ListPx2.Count;
            //array_hole[1, 1] = ListPy2.Count;
            //array_hole[0, 2] = ListPx3.Count;
            //array_hole[1, 2] = ListPy3.Count;
            //ed.WriteMessage($"\nKhông tìm thấy DimStyle có tên '{targetName}'.");
            if (rbt_lapsau.Checked == true)
            {
                //WKV_KT2 = new List<string>() { "オールアンカー FT - 640(鉄)", holes.Count.ToString(), holes3.Count.ToString(), holes2.Count.ToString(), "1", "1", "1" };
                AddLeaderNote(new Point3d(ListPy[ListPy.Count - 2].X + L - 2 * fixedStartOffset - 110, ListPy[ListPy.Count - 2].Y, 0), new Point3d(ListPy[ListPy.Count - 2].X + L - 2 * fixedStartOffset - 110 + 300, ListPy[ListPy.Count - 2].Y + 300 - 40, 0), $"{holes.Count}-φ7 （アンカー用穴)", dimstyle_text);
            }
            else
            {
               // WKV_KT2 = new List<string>() { $"ﾚﾍﾞﾙｱﾝｶｰ6個＋埋込アンカー{holes.Count - 6}個", holes.Count.ToString(), holes3.Count.ToString(), holes2.Count.ToString(), "1", "1", "1" };
                AddLeaderNote(new Point3d(ListPy[ListPy.Count - 2].X + L - 2 * fixedStartOffset, ListPy[ListPy.Count - 2].Y, 0), new Point3d(ListPy[ListPy.Count - 2].X + L - 2 * fixedStartOffset - 110 + 300, ListPy[ListPy.Count - 2].Y + 150 - 40, 0), "6-φ10（レベルアンカー用穴", dimstyle_text);
                AddLeaderNote(new Point3d(ListPy[ListPy.Count - 2].X + L - 2 * fixedStartOffset - 110, ListPy[ListPy.Count - 2].Y, 0), new Point3d(ListPy[ListPy.Count - 2].X + L - 2 * fixedStartOffset - 110 + 300, ListPy[ListPy.Count - 2].Y + 300 - 40, 0), $"{holes.Count - 6}-φ7 （埋設アンカー用穴)", dimstyle_text);
            }

            AddLeaderNote(new Point3d(x - L / 2 + L - 109, y - H / 2 + 209, 0), new Point3d(x - L / 2 + L - 109 + 375, y - H / 2 + 209 - 305, 0), "（台座、配管箱 固定用穴）", dimstyle_text);
            var filteredPoints = holes3
    .GroupBy(p => p.Y)
    .Select(q => q.OrderByDescending(p => p.X).First())
    .ToList();
            Point3d PP = new Point3d();
            if (filteredPoints.Count > 2)
            {
                PP = filteredPoints[2];
            }
            else
            {
                PP = filteredPoints[1];
            }
            AddLeaderNote(PP, new Point3d(PP.X + 375, PP.Y - 305, 0), "（配管箱、蓋 固定用穴）", dimstyle_text);
            createDim(ListPy.ToArray(), new Point3d(ListPy[0].X - 130 - 50, 0, 0), dimstyle_text, 0.5 * Math.PI, new Point3d(ListPy[0].X - 130 - 50, ListPy[0].Y - 120, 0), new Point3d(ListPy[ListPy.Count - 1].X - 130 - 50, ListPy[ListPy.Count - 1].Y + 120, 0), true);
            createDim(ListPx.ToArray(), new Point3d(0, ListPy[0].Y - 200, 0), dimstyle_text, 0, new Point3d(ListPy[0].X - 90, ListPy[0].Y - 200, 0), new Point3d(ListPx[ListPx.Count - 1].X + 85, ListPx[ListPx.Count - 1].Y - 200, 0), true);


            //return pointholes=holes;
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

        private List<Point3d> ListHoles(double x, double y)
        {
            Double.TryParse(tb_w.Text, out Double w);
            Double.TryParse(tb_dh.Text, out Double dh);
            double.TryParse(tb_i.Text, out Double i1);
            Double.TryParse(tb_b.Text, out Double b);
            Double.TryParse(tb_f.Text, out Double f);
            Double.TryParse(tb_g.Text, out Double g);
            double.TryParse(tb_k.Text, out Double k);
            double.TryParse(tb_a.Text, out Double a);
            // double.TryParse(tb_i.Text, out Double i1);
            double.TryParse(tb_d.Text, out Double d);
            double.TryParse(tb_e.Text, out Double e1);
            double.TryParse(tb_m.Text, out Double m);
            List<Point3d> points = new List<Point3d>();
            double L = w + m * 2 + k * 2 - d * 2;
            double H = dh + m * 2 + k * 2 - f * 2;
            //------------------------------------------------------------toạ độ lỗ đế LD
            var holes = GetHoleBorderPoints(xStart: x - L / 2, xLength: L, yStart: y - H / 2, yLength: H, fixedStartOffset: 40, fixedEndOffset: 110, maxSpacing: 350);
            //var holes2 = GetHoleBorderPoints(xStart: x - L / 2, xLength: L, yStart: y - H / 2, yLength: H, fixedStartOffset: 109, fixedEndOffset: 100, maxSpacing: 350);
            points.AddRange(holes);
            //  points.AddRange(holes2);
            //Point3d p1 = new Point3d(x-L/2+40, y-H/2+40, 0);
            //Point3d p2 = new Point3d(x - L/2 + 40, y + H/2 - 40, 0);
            //Point3d p3 = new Point3d(x + L/2 - 40, y + H/2 - 40, 0);
            //Point3d p4 = new Point3d(x + L / 2 - 40, y - H / 2 + 40, 0);
            //points.AddRange(new[] { p1, p2, p3, p4 });
            //double LD1= H - 40 * 2;
            //double LD2 = L - 40 * 2;
            //double Ind1 = Math.Ceiling(LD1 / 350);
            //double Ind2 = Math.Ceiling(LD2 / 350);

            //for(int i = 0; i < Ind1; i++)
            //{
            //    double yt = y - H / 2 + 40 + LD1 / Ind1 * (i+1);
            //   // double yp = y - H / 2 + 40 + LD1 / Ind1 * (i + 1);
            //    Point3d Pld1 = new Point3d(x - L / 2 + 40, yt, 0);
            //    Point3d Pld2 = new Point3d(x + L / 2 - 40, yt, 0);
            //    points.AddRange(new[] { Pld1, Pld2 });
            //}
            //for (int i = 0; i < Ind2; i++)
            //{
            //    //double xt = y - H / 2 + 40 + LD1 / Ind1 * (i + 1);
            //    double xd = x - L / 2 + 40 + LD2 / Ind2 * (i + 1);
            //    Point3d Pld1 = new Point3d(xd, y - H / 2 + 40, 0);
            //    Point3d Pld2 = new Point3d(xd, y + H / 2 - 40, 0);
            //    points.AddRange(new[] { Pld1, Pld2 });
            //}
            //------------------------------------------------------------



            return points;
        }
        public static List<double> GetHolex(double start, double length, double fixedStartOffset, double fixedEndOffset, double maxSpacing = 350)
        {
            var coords = new List<double>();

            double usableLength = length - 2 * fixedStartOffset - 2 * fixedEndOffset;
            int segments = (int)Math.Ceiling(usableLength / maxSpacing);
            double spacing = usableLength / segments;
            double giaso = spacing - Math.Floor(spacing);
            coords.Add(start + fixedStartOffset); // lỗ đầu

            for (int i = 0; i <= segments; i++)
            {
                double py = start + fixedStartOffset + fixedEndOffset + i * spacing;
                // coords.Add(start+fixedStartOffset + fixedEndOffset + i * spacing);

                if ((py < start + length / 2) && (i != 0 || i != segments))
                {
                    py = py - i * giaso;
                }
                else if ((py > start + length / 2) && (i != 0 || i != segments))
                {
                    py = py + (segments - i) * giaso;
                }

                coords.Add(py);
            }

            coords.Add(start + fixedStartOffset + 2 * fixedEndOffset + usableLength); // lỗ cuối

            return coords;
        }
        public static List<double> GetHolex2(double start, double length, double fixedStartOffset, double fixedEndOffset, double maxSpacing = 350)
        {
            var coords = new List<double>();

            double usableLength = length - 2 * fixedStartOffset - 2 * fixedEndOffset;
            int segments = (int)Math.Ceiling(usableLength / maxSpacing);
            double spacing = usableLength / segments;
            double giaso = spacing - Math.Floor(spacing);
            coords.Add(start + fixedStartOffset); // lỗ đầu

            for (int i = 0; i <= segments; i++)
            {
                double py = start + fixedStartOffset + fixedEndOffset + i * spacing;
                // coords.Add(start+fixedStartOffset + fixedEndOffset + i * spacing);

                //if ((py < start + length / 2) && (i != 0 || i != segments))
                //{
                //    py = py - i * giaso;
                //}
                //else if ((py > start + length / 2) && (i != 0 || i != segments))
                //{
                //    py = py + (segments - i) * giaso;
                //}

                coords.Add(py);
            }

            coords.Add(start + fixedStartOffset + 2 * fixedEndOffset + usableLength); // lỗ cuối

            return coords;
        }
        public static List<double> GetHoley(double start, double length, double fixedStartOffset, double maxSpacing = 350)
        {
            var coords = new List<double>();

            double usableLength = length - 2 * fixedStartOffset;
            int segments = (int)Math.Ceiling(usableLength / maxSpacing);
            double spacing = usableLength / segments;
            double giaso = spacing - Math.Floor(spacing);
            coords.Add(start + fixedStartOffset); // lỗ đầu

            for (int i = 1; i < segments; i++)
            {
                double py = start + fixedStartOffset + i * spacing;
                if (py < start + length / 2)
                {
                    py = py - i * giaso;
                }
                else if (py > start + length / 2)
                {
                    py = py + (segments - i) * giaso;
                }

                coords.Add(py);
            }

            coords.Add(start + fixedStartOffset + usableLength); // lỗ cuối

            return coords;
        }
        public static List<double> GetHoley2(double start, double length, double fixedStartOffset, double fixedEndOffset, double maxSpacing = 350)
        {
            var coords = new List<double>();

            double usableLength = length - 2 * fixedStartOffset - 2 * fixedEndOffset;
            int segments = (int)Math.Ceiling(usableLength / maxSpacing);
            double spacing = usableLength / segments;

            coords.Add(start + fixedStartOffset); // lỗ đầu

            for (int i = 0; i <= segments; i++)
            {
                coords.Add(start + fixedStartOffset + fixedEndOffset + i * spacing);
            }

            coords.Add(start + fixedStartOffset + 2 * fixedEndOffset + usableLength); // lỗ cuối

            return coords;
        }
        //-------------------------------------------------------
        public static List<double> GetHoley3(double start, double length, double fixedStartOffset, double fixedEndOffset, double maxSpacing = 350)
        {
            var coords = new List<double>();

            double usableLength = length / 2 - fixedEndOffset - fixedStartOffset - 150;
            int segments = (int)Math.Ceiling(usableLength / maxSpacing);
            double spacing = usableLength / segments;
            //double giaso = spacing - Math.Floor(spacing);
            coords.Add(start + fixedStartOffset); // lỗ đầu

            for (int i = 0; i <= segments; i++)
            {
                double py = start + fixedStartOffset + fixedEndOffset + i * spacing;
                // coords.Add(start+fixedStartOffset + fixedEndOffset + i * spacing);


                coords.Add(py);
            }
            for (int i = 0; i <= segments; i++)
            {
                double py = start + length / 2 + 150 + i * spacing;
                // coords.Add(start+fixedStartOffset + fixedEndOffset + i * spacing);


                coords.Add(py);
            }

            coords.Add(start + length - fixedStartOffset); // lỗ cuối

            return coords;
        }

        //-------------------------------------------------------
        public static List<Point3d> GetHoleBorderPoints(
    double xStart, double xLength,
    double yStart, double yLength,
    double fixedStartOffset,
    double fixedEndOffset,
    double maxSpacing = 350)
        {
            var xList = GetHolex(xStart, xLength, fixedStartOffset, fixedEndOffset, maxSpacing);
            var yList = GetHoley(yStart, yLength, fixedStartOffset, maxSpacing);

            var points = new List<Point3d>();

            foreach (var x in xList)
            {
                foreach (var y in yList)
                {
                    // Giữ điểm ở viền
                    if (x == xList.First() || x == xList.Last() || y == yList.First() || y == yList.Last())
                    {
                        points.Add(new Point3d(x, y, 0));
                    }
                }
            }

            return points;
        }
        public static List<Point3d> GetHoleBorderPoints2(
   double xStart, double xLength,
   double yStart, double yLength,
   double fixedStartOffset,
   double fixedEndOffset,
   double maxSpacing = 350)
        {
            var xList = GetHolex2(xStart, xLength, fixedStartOffset, fixedEndOffset, maxSpacing);
            var yList = GetHoley2(yStart, yLength, fixedStartOffset, fixedEndOffset, maxSpacing);

            var points = new List<Point3d>();

            foreach (var x in xList)
            {
                foreach (var y in yList)
                {
                    bool isOnXEdge = x == xList.First() || x == xList.Last();
                    bool isOnYEdge = y == yList.First() || y == yList.Last();

                    // Giữ điểm nằm trên viền nhưng không phải góc
                    if (isOnXEdge ^ isOnYEdge)  // XOR: chỉ một trong hai là true
                    {
                        points.Add(new Point3d(x, y, 0));
                    }
                }
            }

            return points;

            // return points;
        }
        public static List<Point3d> GetHoleBorderPoints3(
  double xStart, double xLength,
  double yStart, double yLength,
  double fixedStartOffset,
  double fixedEndOffset,
  double maxSpacing = 350)
        {
            var xList = GetHolex2(xStart, xLength, fixedStartOffset, fixedEndOffset, maxSpacing);
            var yList = GetHoley3(yStart, yLength, fixedStartOffset, fixedEndOffset, maxSpacing);

            var points = new List<Point3d>();

            foreach (var x in xList)
            {
                foreach (var y in yList)
                {
                    bool isOnXEdge = x == xList.First() || x == xList.Last();
                    bool isOnYEdge = y == yList.First() || y == yList.Last();

                    // Giữ điểm nằm trên viền nhưng không phải góc
                    if (isOnXEdge ^ isOnYEdge)  // XOR: chỉ một trong hai là true
                    {
                        points.Add(new Point3d(x, y, 0));
                    }
                }
            }
            points.Sort((p1, p2) => p1.Y.CompareTo(p2.Y));
            return points;

            // return points;
        }

        private void DrawLid(double x, double y)
        {
            Double.TryParse(tb_w.Text, out Double w);
            Double.TryParse(tb_dh.Text, out Double dh);
            Double.TryParse(tb_c.Text, out Double c);
            Double.TryParse(tb_j.Text, out Double j);
            Double.TryParse(tb_deg.Text, out Double deg);
            Double.TryParse(tb_b.Text, out Double b);
            Double.TryParse(tb_f.Text, out Double f);
            Double.TryParse(tb_g.Text, out Double g);
            double.TryParse(tb_k.Text, out Double k);
            double.TryParse(tb_a.Text, out Double a);
            double.TryParse(tb_i.Text, out Double i1);
            double.TryParse(tb_d.Text, out Double d);
            double.TryParse(tb_e.Text, out Double e1);
            double.TryParse(tb_i.Text, out Double i);
            double.TryParse(tb_m.Text, out Double m);
            Double ct = j + c + ((dh + m * 2 + i * 2) / 2.0) * Math.Tan(deg * Math.PI / 180.0);//Math.Ceiling(j + c + ((dh + m * 2) / 2.0) * Math.Tan(deg * Math.PI / 180.0) + 3);
            double x1 = x;
            double y1 = y + ct;
            double x2 = x1 + b / 2.0 + i1;
            double y2 = y1 - (b / 2.0 + i1) * Math.Tan(deg * Math.PI / 180.0);
            //  MessageBox.Show($"x1: {x1}, y1: {y1}, x2: {x2}, y2: {y2},tan: {Math.Tan(deg * Math.PI / 180.0)}");
            double x3 = x2 - 30 * Math.Sin(deg * Math.PI / 180.0);
            double y3 = y2 - 30 * Math.Cos(deg * Math.PI / 180.0);
            double x4 = x1;
            double y4 = y1 - 30 * Math.Cos(deg * Math.PI / 180.0);
            double x5 = 2 * x1 - x3;
            double y5 = y3;
            double x6 = 2 * x1 - x2; ;
            double y6 = y2;
            double cp1 = ct - (dh / 2 + m) * Math.Tan(deg * Math.PI / 180.0) - 30 / Math.Cos(deg * Math.PI / 180.0) - j;
            Point2d p1 = new Point2d(x1, y1);
            Point2d p2 = new Point2d(x2, y2);
            Point2d p3 = new Point2d(x3, y3);
            Point2d p4 = new Point2d(x4, y4);
            Point2d p5 = new Point2d(x5, y5);
            Point2d p6 = new Point2d(x6, y6);
            List<Point2d> points = new List<Point2d> { p1, p2, p3, p4, p5, p6, p1 };
            CreatePolyline2D(points, "外形線", LineWeight.LineWeight013);
            // insert rib------------------------------------------------------------
            var lookup = new RibLookup();
            var result = lookup.GetRibCount(int.Parse(tb_a.Text), int.Parse(tb_b.Text));
            if (result.HasValue)
            {
                for (int i_de = 0; i_de < result.Value.RibX / 2; i_de++)
                {
                    double xn1 = x1 + (i_de + 1) * (x2 - x1) / (result.Value.RibX / 2 + 1);
                    double yn1 = y1 + (i_de + 1) * (y2 - y1) / (result.Value.RibX / 2 + 1);
                    double xn3 = x1 + (i_de + 1) * (x6 - x1) / (result.Value.RibX / 2 + 1);
                    double yn3 = y1 + (i_de + 1) * (y6 - y1) / (result.Value.RibX / 2 + 1);

                    Insertdynamicblock.InsertBlock(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\HB-CD.dwg", "ND-01", new Point3d(xn1, yn1, 0), new List<string>(), scale, (360 - deg) * Math.PI / 180.0);
                    Insertdynamicblock.InsertBlock(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\HB-CD.dwg", "ND-01", new Point3d(xn3, yn3, 0), new List<string>(), scale, deg * Math.PI / 180.0);
                }
                //label3.Text = $"RIB X: {result.Value.RibX}, RIB Y: {result.Value.RibY}";
                // Console.WriteLine($"RIB X: {result.Value.RibX}, RIB Y: {result.Value.RibY}");

            }
            //double xn1= x1 + 0.28*(x2 - x1);
            //double yn1 = y1 + 0.28 * (y2 - y1);
            //double xn2 = x1 + 0.65 * (x2 - x1);
            //double yn2 = y1 + 0.65 * (y2 - y1);
            //double xn3 = x1 + 0.28 * (x6 - x1);
            //double yn3 = y1 + 0.28 * (y6 - y1);
            //double xn4 = x1 + 0.65 * (x6 - x1);
            //double yn4 = y1 + 0.65 * (y6 - y1);

            //Insertdynamicblock.InsertBlock(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\HB-CD.dwg", "ND-01", new Point3d(xn1, yn1, 0), new List<string>(), scale,(360- deg) * Math.PI / 180.0);
            //Insertdynamicblock.InsertBlock(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\HB-CD.dwg", "ND-01", new Point3d(xn2, yn2, 0), new List<string>(), scale, (360 - deg) * Math.PI / 180.0);
            //Insertdynamicblock.InsertBlock(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\HB-CD.dwg", "ND-01", new Point3d(xn3, yn3, 0), new List<string>(), scale, deg * Math.PI / 180.0);
            //Insertdynamicblock.InsertBlock(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\HB-CD.dwg", "ND-01", new Point3d(xn4, yn4, 0), new List<string>(), scale, deg * Math.PI / 180.0);
            // insert text----------------------------------------------------------
            Point3d[] point3ds1 = new Point3d[] { new Point3d(x6, y6, 0), new Point3d(x - b / 2, y + j + cp1, 0), new Point3d(x + b / 2, y + j + cp1, 0), new Point3d(x2, y2, 0) };
            Point3d[] point3ds2 = new Point3d[] { new Point3d(x6, y6, 0), new Point3d(x2, y2, 0) };
            Point3d[] point3ds3 = new Point3d[] { new Point3d(x4, y4, 0), new Point3d(x3, y3, 0), new Point3d(x5, y5, 0) };
            createDim(point3ds1, new Point3d(0, y + ct + 130 * 2, 0), dimstyle_text, 0, new Point3d(x6 - 150, y + ct + 130 * 2, 0), new Point3d(x2 + 150, y + ct + 130 * 2, 0), true);
            createDim(point3ds2, new Point3d(0, y + ct + 3 * 130, 0), dimstyle_text, 0, new Point3d(0, 0, 0), new Point3d(0, 0, 0), false);
            createDimagu(point3ds3, new Point3d(x4, y4 - 110, 0), dimstyle_text);

        }
        public static void createDimagu(Point3d[] points, Point3d dimLinePos, ObjectId dimstyle)
        {
            Document doc = acadApp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            //// 4 điểm liên tiếp
            //Point3d p1 = new Point3d(0, 0, 0);
            //Point3d p2 = new Point3d(100, 0, 0);
            //Point3d p3 = new Point3d(180, 0, 0);
            //Point3d p4 = new Point3d(260, 0, 0);

            //// Khoảng cách từ đoạn cần dim lên vị trí đường dim (theo hướng vuông góc)
            //double dimOffset = 20;

            //// Hướng đo (ở đây là trục X, nên hướng dim là vuông góc => trục Y)
            //Vector3d normal = Vector3d.ZAxis;
            //Vector3d dimDir = (p2 - p1).GetNormal();
            //Vector3d offsetDir = dimDir.GetPerpendicularVector().GetNormal();

            //// Tính điểm đặt dimension line đầu tiên
            //Point3d dimLinePos = p1 + offsetDir * dimOffset;
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
            //using (package)
            //    if (Directory.Exists(outputFolder))
            //    {
            //        string filePath2 = Path.Combine(outputFolder, $"HB特注原価計算 {tb_ken.Text} {tb_date.Text.Replace("/", "").Substring(2)}.xlsx");                                                                         // Lưu file Excel
            //        package.SaveAs(filePath2);
            //    }
            //    else { return; }
        }
        private void writedata(string outputfolder)
        {



            if (File.Exists(filePath))
            {
                try
                {
                    string tempFolder = Path.GetTempPath();
                    string fileName = $"HB特注原価計算 {DateTime.Now:yyMMddHHmmss}.xlsx";
                    string tempFilePath = Path.Combine(tempFolder, fileName);

                    File.Copy(filePath, tempFilePath, true);

                    Excel.Application excelApp = new Excel.Application();
                    Excel.Workbook workbook = excelApp.Workbooks.Open(tempFilePath);
                    Excel.Worksheet worksheet1 = workbook.Sheets[1];
                    Excel.Worksheet worksheet2 = workbook.Sheets[2];

                    // Gán dữ liệu vào các ô
                    worksheet1.Cells[18, 7].Value = tb_a.Text;
                    worksheet1.Cells[19, 7].Value = tb_b.Text;
                    worksheet1.Cells[20, 7].Value = tb_c.Text;
                    worksheet1.Cells[21, 7].Value = tb_j.Text;

                    worksheet1.Cells[18, 10].Value = tb_d.Text;
                    worksheet1.Cells[19, 10].Value = tb_e.Text;
                    worksheet1.Cells[20, 10].Value = tb_f.Text;
                    worksheet1.Cells[21, 10].Value = tb_g.Text;

                    worksheet1.Cells[18, 13].Value = tb_h.Text;
                    worksheet1.Cells[19, 13].Value = tb_i.Text;
                    worksheet1.Cells[20, 13].Value = tb_k.Text;
                    worksheet1.Cells[21, 13].Value = tb_deg.Text;

                    // Tính toán lại (nếu có công thức)
                    workbook.RefreshAll();
                    excelApp.Calculate();

                    // Lấy dữ liệu kết quả sau tính toán
                    lb_tong.Text = worksheet1.Cells[17, 15].Text?.ToString();
                    lb_nap.Text = worksheet1.Cells[17, 16].Text?.ToString();
                    lb_than.Text = worksheet1.Cells[17, 17].Text?.ToString();
                    lb_de.Text = worksheet1.Cells[17, 18].Text?.ToString();

                    // Lưu & đóng
                    workbook.Save();
                    workbook.Close(false);
                    excelApp.Quit();

                    // Giải phóng COM
                    Marshal.ReleaseComObject(worksheet1);
                    Marshal.ReleaseComObject(workbook);
                    Marshal.ReleaseComObject(excelApp);

                    // Xoá file tạm
                    if (File.Exists(tempFilePath))
                        File.Delete(tempFilePath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi khi xử lý Excel: " + ex.Message);
                }
            }
            else
            {
                MessageBox.Show("Đường dẫn file không tồn tại.");
            }



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
            //    var lookup = new RibLookup();
            //    var result = lookup.GetRibCount(int.Parse(tb_a.Text), int.Parse(tb_b.Text));
            //    if (result.HasValue)
            //    {
            //        label3.Text = $"RIB X: {result.Value.RibX}, RIB Y: {result.Value.RibY}";
            //       // Console.WriteLine($"RIB X: {result.Value.RibX}, RIB Y: {result.Value.RibY}");
            //    }
            //    else
            //    {
            //    //   MessageBox.Show("Không tìm thấy kết quả phù hợp với kích thước đã nhập.");
            //    }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            RegistryHelper.LoadAllSettings("RWtool-01", this);
            tb_date.Text = DateTime.Now.ToString("yyyy/MM/dd");
            tb_ken.Text = "";
            // rdb_4.Checked = true;
            SelectView1();
             liststylename = new List<string>() { "WKV_20", "WKV_30" , "WKV_35", "WKV_40", "WKV_45", "WKV_50" };
            DimStyleImporter.ImportDimStyle(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\RW.dwg", liststylename);
            btn_ex.ContextMenuStrip = contextMenuStrip1;
            cb_pdf.ContextMenuStrip = contextMenuStrip2;
            PrinterUtility.ListPlotDevices(cbb_printer);
            PrinterUtility.additemcombobox(cbb_plotstyle, PrinterUtility.PlotStyleList());
            string printer = (string)RegistryHelper.GetSetting("RWtool-01", "Printer", "");
            var item = cbb_printer.Items.Cast<object>().FirstOrDefault(i => i.ToString().Equals(printer, StringComparison.OrdinalIgnoreCase));
            if (item != null)
            {
                cbb_printer.SelectedItem = item;
            }
            string plotstyle = (string)RegistryHelper.GetSetting("RWtool-01", "Plotstyle", "");
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
                string papersize = (string)RegistryHelper.GetSetting("RWtool-01", "Papersize", "");
                var item4 = cbb_papersize.Items.Cast<object>().FirstOrDefault(i => i.ToString().Equals(papersize, StringComparison.OrdinalIgnoreCase));
                if (item4 != null)
                {
                    cbb_papersize.SelectedItem = item4;
                }

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            folderBrowserDialog.Description = "Select the file output folder";

            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {

                comboBox1.Text = folderBrowserDialog.SelectedPath;
                // readdata();

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

            // Cho phép: số, phím điều khiển (Backspace...), dấu trừ đầu tiên, dấu chấm thập phân nếu chưa có
            //if (!char.IsControl(e.KeyChar) &&
            //    !char.IsDigit(e.KeyChar) &&
            //    (e.KeyChar != '.') &&
            //    (e.KeyChar != '-'))
            //{
            //    e.Handled = true;
            //}

            //// Chỉ cho phép một dấu chấm (thập phân)
            //if (e.KeyChar == '.' && tb.Text.Contains("."))
            //{
            //    e.Handled = true;
            //}

            //// Dấu trừ chỉ được đứng đầu
            //if (e.KeyChar == '-' && (tb.SelectionStart != 0 || tb.Text.Contains("-")))
            //{
            //    e.Handled = true;
            //}
        }

        private void maskedTextBox1_Leave(object sender, EventArgs e)
        {
            //if (DateTime.TryParseExact(tb_date.Text, "yyyy/MM/dd",
            //                 CultureInfo.InvariantCulture,
            //                 DateTimeStyles.None, out DateTime date))
            //{
            //    // Ngày hợp lệ
            //    Console.WriteLine("Ngày nhập: " + date.ToShortDateString());
            //}
            //else
            //{
            //    MessageBox.Show("Ngày không hợp lệ. Vui lòng nhập lại (yyyy/MM/dd).");
            //    tb_date.Focus();
            //}
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
            if (btn_front.BackColor == Color.White)
            {
                btn_front.BackColor = Color.FromArgb(255, 224, 192);
            }
            else
            {
                btn_front.BackColor = Color.White;
            }
        }

        private void btn_left_Click(object sender, EventArgs e)
        {
            if (btn_left.BackColor == Color.White)
            {
                btn_left.BackColor = Color.FromArgb(255, 224, 192);
            }
            else
            {
                btn_left.BackColor = Color.White;
            }
        }

        private void btn_right_Click(object sender, EventArgs e)
        {
            if (btn_right.BackColor == Color.White)
            {
                btn_right.BackColor = Color.FromArgb(255, 224, 192);
            }
            else
            {
                btn_right.BackColor = Color.White;
            }
        }

        private void btn_top_Click(object sender, EventArgs e)
        {
            if (btn_top.BackColor == Color.White)
            {
                btn_top.BackColor = Color.FromArgb(255, 224, 192);
            }
            else
            {
                btn_top.BackColor = Color.White;
            }
        }

        private void btn_parent_MouseHover(object sender, EventArgs e)
        {
            toolTip1.SetToolTip(btn_front, "Mặt trước");
        }

        private void btn_top_MouseHover(object sender, EventArgs e)
        {
            toolTip1.SetToolTip(btn_top, "Mặt trên");
        }

        private void btn_right_MouseHover(object sender, EventArgs e)
        {
            toolTip1.SetToolTip(btn_right, "Mặt bên phải");
        }

        private void btn_left_MouseHover(object sender, EventArgs e)
        {
            toolTip1.SetToolTip(btn_left, "Mặt bên trái");
        }
        private void btn_back_MouseHover(object sender, EventArgs e)
        {
            toolTip1.SetToolTip(btn_left, "Mặt sau");
        }
        private void button4_Click(object sender, EventArgs e)
        {
            tb_date.Text = DateTime.Now.ToString("yyyy/MM/dd");
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            tb_date.Text = DateTime.Now.ToString("yyyy/MM/dd");
        }

        //private void button5_Click(object sender, EventArgs e)
        //{
        //    NativeMethods.SetForegroundWindow(acadHwnd); values.Clear();



        //    var doc = acadApp.DocumentManager.MdiActiveDocument;
        //    var db = doc.Database;
        //    var ed = doc.Editor;
        //    // int PP = dtb_gr1.RowCount;
        //    //int p = 1;
        //    //  dtb_gr1.Rows.RemoveAt(PP-1);
        //    PromptPointResult pointResult = ed.GetPoint("Chọn vị trí để chèn Block Reference: ");
        //    DrawLid(pointResult.Value.X, pointResult.Value.Y);
        //}

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
            if (btn_back.BackColor == Color.White)
            {
                btn_back.BackColor = Color.FromArgb(255, 224, 192);
            }
            else
            {
                btn_back.BackColor = Color.White;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {

        }

        private void btn_update_Click(object sender, EventArgs e)
        {
            bt_updateinput.PerformClick();
            writedata(comboBox1.Text);
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            //RecalculateExcel(tempFilePath);
            // ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //// ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            // using (var package = new ExcelPackage(new FileInfo(tempFilePath)))
            // {
            //    // System.Diagnostics.Process.Start(tempFilePath);

            //     var worksheet1 = package.Workbook.Worksheets[0];
            //     //var worksheet2 = package.Workbook.Worksheets.Count > 1 ? package.Workbook.Worksheets[1] : null;

            //     lb_tong.Text = worksheet1.Cells[17, 15].Text;//worksheet1.Cells[17, 15].Text;
            //     lb_nap.Text = worksheet1.Cells[17, 16].Text;
            //     lb_than.Text = worksheet1.Cells[17, 17].Text;
            //     lb_de.Text = worksheet1.Cells[17, 18].Text;
            //    // MessageBox.Show($"dữ liệu: {worksheet1.Cells[17, 15].Text}");
            // }
        }

        private void rbt_lapsau_CheckedChanged(object sender, EventArgs e)
        {
            if (rbt_lapsau.Checked == true)
            {
                tb_hbname.Text = "HB-特注 (あと施工)";
            }
            else
            {
                tb_hbname.Text = "HB-特注（埋設施工)";
            }
        }

        private void button5_Click_1(object sender, EventArgs e)
        {  // btn_ex.PerformClick();
            NativeMethods.SetForegroundWindow(acadHwnd);
            LoadingForm loadingForm = new LoadingForm();
            Task showFormTask = Task.Run(() => loadingForm.ShowDialog());
            btn_update.PerformClick();
            loadingForm.Invoke(new Action(() => loadingForm.Close()));
            Double.TryParse(tb_w.Text, out Double w);
            Double.TryParse(tb_dh.Text, out Double dh);
            Double.TryParse(tb_f.Text, out Double f);
            double.TryParse(tb_k.Text, out Double k);
            double.TryParse(tb_d.Text, out Double d);
            double.TryParse(tb_m.Text, out Double m);

            double x = 0;
            double y = 0;

            double L = w + m * 2 + k * 2 - d * 2;
            double H = dh + m * 2 + k * 2 - f * 2;
            double fixedStartOffset;
            if (rbt_lapsau.Checked == true)
            {
                fixedStartOffset = 25;
            }
            else
            {
                fixedStartOffset = 40;
            }
            var holes = GetHoleBorderPoints(xStart: x - L / 2, xLength: L, yStart: y - H / 2, yLength: H, fixedStartOffset, fixedEndOffset: 110, maxSpacing: 350);

            var holes2 = GetHoleBorderPoints2(xStart: x - L / 2, xLength: L, yStart: y - H / 2, yLength: H, fixedStartOffset: 109, fixedEndOffset: 100, maxSpacing: 350);

            var holes3 = GetHoleBorderPoints3(xStart: x - L / 2, xLength: L, yStart: y - H / 2, yLength: H, fixedStartOffset: 36.6, fixedEndOffset: 100, maxSpacing: 350);

            List<Point3d> ListPy = new List<Point3d>();
            List<Point3d> ListPx = new List<Point3d>();
            Dimentionholes1(holes, x - L / 2, y - H / 2, fixedStartOffset, out ListPx, out ListPy);//đế
            // Fix for CS0022: Wrong number of indices inside []; expected 2
            // The issue occurs because `array_hole` is declared as a 2D array, but the code is trying to access it as if it were a jagged array.
            // Correcting the indexing to use two indices for the 2D array.
            List<Point3d> ListPy2 = new List<Point3d>();
            List<Point3d> ListPx2 = new List<Point3d>();
            Dimentionholes2(holes2, x - L / 2, y - H / 2, 110, out ListPx2, out ListPy2);//thân
            List<Point3d> ListPy3 = new List<Point3d>();
            List<Point3d> ListPx3 = new List<Point3d>();
            Dimentionholes2(holes3, x - L / 2, y - H / 2, 40, out ListPx3, out ListPy3);//nắp
            array_hole[0, 0] = ListPx.Count - 4;
            array_hole[1, 0] = ListPy.Count - 2;
            array_hole[0, 1] = ListPx2.Count;
            array_hole[1, 1] = ListPy2.Count;
            array_hole[0, 2] = ListPx3.Count;
            array_hole[1, 2] = ListPy3.Count;

            List<bool> listfile = new List<bool>() { cb_dwg.Checked, cb_dxf.Checked };

            if (File.Exists(filePath))
            {
                //try
                //{
                string dt = DateTime.Now.ToString("HH:mm:ss").Replace(":", "");
                if (cb_xlsx.Checked)
                {
                    string tempFolder = comboBox1.Text;
                    if (Directory.Exists(tempFolder))
                    {

                        string fileName = $"HB特注原価計算(＠230)( {tb_ken.Text} ) {tb_date.Text.Replace("/", "").Substring(2)}.xlsx";
                        tempFilePath = Path.Combine(tempFolder, fileName);

                        if (File.Exists(tempFilePath))
                        {
                            fileName = $"HB特注原価計算(＠230)( {tb_ken.Text} ) {tb_date.Text.Replace("/", "").Substring(2)} {dt}.xlsx";
                            tempFilePath = Path.Combine(tempFolder, fileName);
                            //File.Delete(tempFilePath);
                        }
                        File.Copy(filePath, tempFilePath, true);

                        Excel.Application excelApp = new Excel.Application();
                        Excel.Workbook workbook = excelApp.Workbooks.Open(tempFilePath);
                        Excel.Worksheet worksheet1 = workbook.Sheets[1];
                        Excel.Worksheet worksheet2 = workbook.Sheets.Count >= 2 ? workbook.Sheets[2] : null;

                        // Gán giá trị vào các ô
                        worksheet1.Cells[18, 7].Value = tb_a.Text;
                        worksheet1.Cells[19, 7].Value = tb_b.Text;
                        worksheet1.Cells[20, 7].Value = tb_c.Text;
                        worksheet1.Cells[21, 7].Value = tb_j.Text;

                        worksheet1.Cells[18, 10].Value = tb_d.Text;
                        worksheet1.Cells[19, 10].Value = tb_e.Text;
                        worksheet1.Cells[20, 10].Value = tb_f.Text;
                        worksheet1.Cells[21, 10].Value = tb_g.Text;

                        worksheet1.Cells[18, 13].Value = tb_h.Text;
                        worksheet1.Cells[19, 13].Value = tb_i.Text;
                        worksheet1.Cells[20, 13].Value = tb_k.Text;
                        worksheet1.Cells[21, 13].Value = tb_deg.Text;

                        // Check lỗ
                        var checker1 = (rbt_lapsau.Checked) ? 1 : 2;
                        worksheet1.Cells[13, 20].Value = checker1;

                        worksheet1.Cells[61, 11].Value = array_hole[0, 0]; // Đế mặt trước
                        worksheet1.Cells[60, 11].Value = array_hole[1, 0]; // Đế mặt bên
                        worksheet1.Cells[59, 11].Value = array_hole[0, 1];
                        worksheet1.Cells[58, 11].Value = array_hole[1, 1];
                        worksheet1.Cells[57, 11].Value = array_hole[0, 2];
                        worksheet1.Cells[56, 11].Value = array_hole[1, 2];
                        worksheet2.Cells[19, 14].Value = total1219;
                        worksheet2.Cells[20, 14].Value = total914;
                        Editor ed = acadApp.DocumentManager.MdiActiveDocument.Editor;
                        ed.WriteMessage($"\nLỗ đế: {array_hole[0, 0]}-{array_hole[1, 0]}\nLỗ nắp: {array_hole[1, 1]}-{array_hole[0, 1]}\nLỗ thân: {array_hole[0, 2]}-{array_hole[1, 2]} ");

                        workbook.Save();
                        workbook.Close(false);
                        excelApp.Quit();

                        // Giải phóng tài nguyên COM
                        Marshal.ReleaseComObject(worksheet1);
                        if (worksheet2 != null) Marshal.ReleaseComObject(worksheet2);
                        Marshal.ReleaseComObject(workbook);
                        Marshal.ReleaseComObject(excelApp);
                    }
                }
                //----------- xuất pdf -----------
                // NativeMethods.SetForegroundWindow(acadHwnd);
                //try
                //{
                if (cb_pdf.Checked)
                {
                    liextent2d = select.ScanBlocks(new List<string>() { "HB-A3-CL", "HB-A3-LS" });


                    Sortextents2d.SortExtentsByPosition(liextent2d, SortDirection.TopToBottomLeftToRight);



                    if (liextent2d.Count > 0)
                    {
                        List<string> filepathsmerge = new List<string>();
                        // string ex = extensions.extension(comboBox1.Text);
                        string fileName3 = $"{tb_ken.Text} {tb_date.Text.Replace("/", "").Substring(2)}.pdf";
                        string outputPath3 = Path.Combine(comboBox1.Text, fileName3);
                        if (File.Exists(outputPath3))
                        {
                            fileName3 = $"{tb_ken.Text} {tb_date.Text.Replace("/", "").Substring(2)} {dt}.pdf";
                            outputPath3 = Path.Combine(comboBox1.Text, fileName3);
                        }
                        //   string filePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\inout.txt";

                        // Savefile(filePath, cbb_ots, comboBox1, comboBox2, comboBox3, comboBox5, cbb_sorting, textBox1, cbb_collate, cbb_merge, cbb_plotfile, comboBox4, cbb_fit, textBox2);
                        //  Printset prr = new Printset();
                        DateTime now = DateTime.Now;
                        string dateTimeString = now.ToString("yyyy MM dd HH mm ss fff").Replace(" ", "");



                        //    Stopwatch stopwatch = new Stopwatch();
                        //  stopwatch.Start();


                        int i = 1;
                        foreach (Extents2d strextent in liextent2d)
                        {
                            double width = Math.Abs(strextent.MaxPoint.X - strextent.MinPoint.X);
                            double height = Math.Abs(strextent.MaxPoint.Y - strextent.MinPoint.Y);
                            //  bool ss = width >= height;
                            string filePath1 = Path.Combine(dateTimeString + "-" + i + ".pdf");
                            Printset.CreateOrEditPageSetup(cbb_printer.Text, cbb_papersize.Text, cbb_plotstyle.Text, strextent, false, Math.Round(height / (5660 / 20)), true, filePath1, width >= height);
                            filepathsmerge.Add(filePath1);
                            i++;
                        }




                        JOINPDF.MergePDFFiles(filepathsmerge, Path.Combine(outputPath3));

                    }
                    else { MessageBox.Show("Please select the print object"); }
                }
                //}


                //catch { MessageBox.Show("Lỗi"); }

                // ---------- xuất file DWG ----------
                string fileName2 = $"{tb_ken.Text} {tb_date.Text.Replace("/", "").Substring(2)}.dwg";
                string outputPath2 = Path.Combine(comboBox1.Text, fileName2);
                if (File.Exists(outputPath2))
                {
                    fileName2 = $"{tb_ken.Text} {tb_date.Text.Replace("/", "").Substring(2)} {dt}.dwg";
                    outputPath2 = Path.Combine(comboBox1.Text, fileName2);
                }
                ExportToDwg(outputPath2, listfile);

                DialogResult result = MessageBox.Show($"Đã xuất file thành công vào thư mục {comboBox1.Text}.\n",
"Xuất thành công", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    string folderPath = Path.GetDirectoryName(comboBox1.Text);
                    if (Directory.Exists(folderPath))
                    {
                        System.Diagnostics.Process.Start("explorer.exe", folderPath);
                    }
                }
                //}
                //catch (Exception ex)
                //{
                //    MessageBox.Show("Đã xảy ra lỗi khi xử lý Excel: " + ex.Message);
                //}
            }
            else
            {

                MessageBox.Show("Đường dẫn file không tồn tại.");
            }

        }

        private void bt_updateinput_Click(object sender, EventArgs e)
        {
            Double.TryParse(tb_i.Text, out Double i);
            Double.TryParse(tb_w.Text, out Double w);
            Double.TryParse(tb_dh.Text, out Double dh);
            Double.TryParse(tb_c.Text, out Double c);
            Double.TryParse(tb_j.Text, out Double j);
            Double.TryParse(tb_deg.Text, out Double deg);
            Double.TryParse(tb_caotong.Text, out Double caotong);
            double.TryParse(tb_m.Text, out Double m);
            double.TryParse(tb_d.Text, out Double d);
            double.TryParse(tb_e.Text, out Double e1);
            double.TryParse(tb_k.Text, out Double k);
            double.TryParse(tb_f.Text, out Double f);
            double.TryParse(tb_g.Text, out Double g);
            double a = (w + m * 2);
            double b = (dh + m * 2);
            double c1 = Math.Round(caotong - j - ((dh + m * 2 + i * 2) / 2.0) * Math.Tan(deg * Math.PI / 180.0), 0, MidpointRounding.AwayFromZero);
            Double.TryParse(tb_h.Text, out Double h);
            tb_a.Text = a.ToString();
            tb_b.Text = b.ToString();
            tb_c.Text = c1.ToString();
            var lookup = new RibLookup();
            var result = lookup.GetRibCount(int.Parse(tb_a.Text), int.Parse(tb_b.Text));
            if (result.HasValue)
            {
                label3.Text = $"RIB X: {result.Value.RibX}, RIB Y: {result.Value.RibY}";
                // Console.WriteLine($"RIB X: {result.Value.RibX}, RIB Y: {result.Value.RibY}");
            }
            double Ld = a - d - e1 + 2 * k;
            double Hd = b - f - g + 2 * k;
            double Ln = a + 2 * i;
            double Hn = b + 2 * i;
            ListZam = new List<ZamPara>
            {
                new ZamPara
                {
                    LengthX =Ld-150+1.6 ,
                    HeightY = j+50+78.9,
                    Quanlity = 2
                },
                new ZamPara
                {
                    LengthX =Hd ,
                    HeightY = j+50+78.9,
                    Quanlity = 2
                },
                new ZamPara
                {
                    LengthX =a+0.8 ,
                    HeightY = c+h+20,
                    Quanlity = 2
                },
                new ZamPara
                {
                    LengthX =b+35+3.2 ,
                    HeightY = c+h+20+(b+35+3.2)*Math.Tan(deg * Math.PI / 180.0)-3.2,
                    Quanlity = 2
                }
                ,
                new ZamPara
                {
                    LengthX =Ln+80+1.6,
                    HeightY =Hn+80+3.2,
                    Quanlity = 1
                }
            };


            //--------------------------------------------------------------
            int kk = 0;
            foreach (var zam in ListZam)
            {
                if (kk < 2) // Loại 1219x2438
                {
                    for (int i1 = 0; i1 < zam.Quanlity; i1++)
                    {
                        items1.Add(new Item { Length = (int)zam.LengthX, Width = (int)zam.HeightY });
                    }
                }
                else // Loại 914x2438
                {
                    for (int i1 = 0; i1 < zam.Quanlity; i1++)
                    {
                        items2.Add(new Item { Length = (int)zam.LengthX, Width = (int)zam.HeightY });
                    }
                }
                kk++;
            }
            //---------------------------------------------------------------

            //  var result1 = SheetArrangement.ArrangeSheets(items1);
            // var result2 = SheetArrangement.ArrangeSheets(items1);
            var bins = new List<Bin>
            {
    new Bin { Id = 1, Length = 2438, Width = 1219 }, // Loại A
    new Bin { Id = 2, Length = 2438, Width = 914 },  // Loại B
            };

            int gap = 5;

            var result12 = MultiBinPacking.Optimize(items1, bins, gap);

            foreach (var kv in result12)
            { MessageBox.Show($"Bin loại {kv.Key}: {kv.Value} tấm đã dùng"); }
            // TB3_1_1.Text = result1.
            // TB3_2_1.Text = Zam914_de.ToString();

            //int index_z = 0;
            //int Zam914_de = 0;
            //int Zam1219_de = 0;
            //int Zam914_than = 0;
            //int Zam1219_than = 0;
            //string resultText = string.Empty;
            //string resultText2 = string.Empty;
            //double totalperusers = 0;
            //int sheetplus914 = 0;
            //int sheetplus1219 = 0;
            //foreach (var zam in ListZam)
            //{
            //    var result1 = SheetArrangement.ArrangeSheets(zam.LengthX, zam.HeightY, zam.Quanlity);
            //    {
            //        if (result1.SheetSizeUsed == "914x2438")
            //        {
            //            resultText = resultText + "-" + $"Thứ{i + 1}:{result1.TotalSheets}-914x2438";
            //            if (index_z < 2)

            //            {
            //                Zam914_de += result1.TotalSheets;

            //            }
            //            else
            //            {
            //                Zam914_than += result1.TotalSheets;
            //            }
            //            list914.Insert(index_z, result1.TotalSheets.ToString());
            //            list1219.Insert(index_z, "0");
            //        }
            //        else if (result1.SheetSizeUsed == "1219x2438")
            //        {
            //            resultText2 = resultText2 + "-" + $"Thứ{i + 1}:{result1.TotalSheets}-1219x2438";
            //            if (index_z < 2)

            //            {
            //                Zam1219_de += result1.TotalSheets;

            //            }
            //            else
            //            {
            //                Zam1219_than += result1.TotalSheets;
            //            }
            //            list1219.Insert(index_z, result1.TotalSheets.ToString());
            //            list914.Insert(index_z, "0");
            //        }
            //        index_z++;
            //        totalperusers = totalperusers + result1.AreaNesting;
            //    }
            //}
            //double Peruser = (totalperusers * 100) / ((Zam914_de + Zam914_than) * (2438 * 914) + (Zam1219_de + Zam1219_than) * (2438 * 1219));
            //double ss = 30;
            //double.TryParse(tb_per.Text, out ss);
            //string tt = (100 - Peruser) > ss ? ">" : "<";
            //lb_per.Text = (100-Peruser).ToString("0.00") + "% " + tt;
            //lb_per2.Text = (100 - Peruser) > ss ? "+1 : 3'x8'" : "+1 : 4'x8'";
            //if (100-Peruser <ss)
            //{
            //    sheetplus914 = 0;
            //    sheetplus1219 = 1;
            //} else 
            //{
            //    sheetplus914 = 1;
            //    sheetplus1219 = 0;
            //}

            //TB3_1_1.Text = Zam1219_de.ToString();
            //TB3_2_1.Text = Zam914_de.ToString();
            //TB3_1_2.Text = Zam1219_than.ToString();
            //TB3_2_2.Text = Zam914_than.ToString();
            //TB3_1_3.Text = (sheetplus1219 ).ToString();
            //TB3_2_3.Text = (sheetplus914).ToString();
            //total914 = (Zam914_de + Zam914_than + sheetplus914);
            //total1219 = (Zam1219_de + Zam1219_than + sheetplus1219);
            //TB3_2_4.Text = total914.ToString();
            //TB3_1_4.Text = total1219.ToString();

            //  lb_per.Text = Peruser.ToString("0.00") + "%";
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
            string folderPath = comboBox1.Text;
            if (Directory.Exists(folderPath))
            {
                System.Diagnostics.Process.Start("explorer.exe", folderPath);
            }
        }

        private void cbb_printer_Click(object sender, EventArgs e)
        {

        }

        private void cbb_printer_SelectedIndexChanged(object sender, EventArgs e)
        {
            PrinterUtility.additemcombobox(cbb_papersize, PrinterUtility.GetPaperSizes(cbb_printer.Text));
        }

        private void btn_preview_Click(object sender, EventArgs e)
        {
            NativeMethods.SetForegroundWindow(acadHwnd);
            liextent2d = select.ScanBlocks(new List<string>() { "HB-A3-CL", "HB-A3-LS" });

            Sortextents2d.SortExtentsByPosition(liextent2d, SortDirection.TopToBottomLeftToRight);
            List<Extents2d> exbuf = new List<Extents2d>();
            List<string> filepathsmerge = new List<string>();
            exbuf = liextent2d;
            DateTime now = DateTime.Now;
            string dateTimeString = now.ToString("yyyy MM dd HH mm ss fff").Replace(" ", "");
            int i = 1;

            if (exbuf.Count != 0)
            {
                // string text = "";
                int num = 0;
                double width;
                double height;
                string fileName3 = $"{tb_ken.Text} {tb_date.Text.Replace("/", "").Substring(2)}.pdf";
                string outputPath3 = Path.Combine(comboBox1.Text, fileName3);
                if (File.Exists(outputPath3))
                {
                    fileName3 = $"{tb_ken.Text} {tb_date.Text.Replace("/", "").Substring(2)} {now.ToString("yyMMdd")}.pdf";
                    outputPath3 = Path.Combine(comboBox1.Text, fileName3);
                }
                // string ex = extensions.extension(comboBox1.Text);
                PreviewEndPlotStatus previewEndPlotStatus = PreviewEndPlotStatus.Next;
                if (previewEndPlotStatus == PreviewEndPlotStatus.Next)
                {
                    while (exbuf.Count > 0 && previewEndPlotStatus != PreviewEndPlotStatus.Cancel && previewEndPlotStatus > 0)
                    {
                        width = Math.Abs(exbuf[num].MaxPoint.X - exbuf[num].MinPoint.X);
                        height = Math.Abs(exbuf[num].MaxPoint.Y - exbuf[num].MinPoint.Y);
                        PreviewEngineFlags previewEngineFlags = PreviewEngineFlags.Plot;

                        if (exbuf.Count > 1)
                        {
                            previewEngineFlags = previewEngineFlags | PreviewEngineFlags.NextSheet | PreviewEngineFlags.PreviousSheet;
                        }
                        previewEndPlotStatus = Printset.preview(cbb_printer.Text, cbb_papersize.Text, cbb_plotstyle.Text, liextent2d[num], false, Math.Round(height / (5660 / 20)), false, "", width >= height, 7);

                        if (previewEndPlotStatus == PreviewEndPlotStatus.Next)
                        {
                            num++;
                        }
                        if (previewEndPlotStatus == PreviewEndPlotStatus.Previous)
                        {
                            num--;
                        }
                        if (previewEndPlotStatus == PreviewEndPlotStatus.Plot)
                        {
                            width = Math.Abs(exbuf[num].MaxPoint.X - exbuf[num].MinPoint.X);
                            height = Math.Abs(exbuf[num].MaxPoint.Y - exbuf[num].MinPoint.Y);
                            string filePath1 = Path.Combine(dateTimeString + "-" + i + ".pdf");
                            Printset.CreateOrEditPageSetup(cbb_printer.Text, cbb_papersize.Text, cbb_plotstyle.Text, liextent2d[num], false, Math.Round(height / (5660 / 20)), true, filePath1, width >= height);
                            filepathsmerge.Add(filePath1);
                            exbuf.Remove(exbuf[num]);
                            i++;
                        }
                        if (num < 0)
                        {
                            num = exbuf.Count - 1;
                        }

                        if (num > exbuf.Count - 1)
                        {
                            num = 0;
                        }

                    }
                    JOINPDF.MergePDFFiles(filepathsmerge, Path.Combine(outputPath3));

                }
            }
            else { MessageBox.Show("Please select the print object"); }
        }

        private void btn_folder_Click(object sender, EventArgs e)
        {
            string folderPath = comboBox1.Text;
            if (Directory.Exists(folderPath))
            {
                System.Diagnostics.Process.Start("explorer.exe", folderPath);
            }
        }

        private void tb_per_MouseHover(object sender, EventArgs e)
        {
            toolTip1.SetToolTip(tb_per, "Nhập ngưỡng % cộng thêm tấm Zam: %S dư lớn hơn: +1 tấm 3'x8', %S dư bé hơn +1 tấm 4'x8'  ");
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
            try
            {
                double.TryParse(tb_OW.Text, out OW);
                double.TryParse(tb_OH.Text, out OH);
                string filePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\TRATHONGSO.xlsM";
                // Khởi tạo
                var service = new ExcelLookupService(filePath);
             //   MessageBox.Show("ddã vào đây1-1");
                // Sử dụng và lấy kết quả
                var result = service.Val(OW, OH);
               // MessageBox.Show("ddã vào đây1-2");
                // Kiểm tra kết quả
                if (result.IsValid)
                {
                    tb_RSow.Text = OW.ToString();
                    tb_RSoh.Text = OH.ToString();
                    A= result.AVal;
                    tb_RSa.Text = A.ToString();
                    B = result.BVal;
                    tb_RSb.Text=B.ToString();
                    C = result.CVal;
                    tb_RSc.Text = C.ToString();
                    D = result.DVal.ToString();
                    tb_RSd.Text = D;
                    motor = result.Value1;
                    tb_motor.Text = motor;
                    gearbox = result.Value2;
                    tb_gb.Text = gearbox;
                    qlypp = result.AdditionalInfo;
                    tb_qlypp.Text = qlypp;
                    spacing = result.Spacing;
                    tb_kc.Text = spacing.ToString();

                    // MessageBox.Show("Đã cập nhật thành công các thông số:\n" +
                    //$"RSa: {result.AVal}\n" +
                    //$"RSb: {result.BVal}\n" +
                    //$"RSc: {result.CVal}\n" +
                    //$"RSd: {result.DVal}");
                    // tb_motor.Text = result.
                    // Sử dụng result.AVal, result.BVal, etc.
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message);
            }
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
            
            //if (tb_ken.Text == "")
            //{
            //    MessageBox.Show("Vui lòng nhập tên công trình");
            //    return;
            //}
            //LoadingForm loadingForm = new LoadingForm();
            //Task showFormTask = Task.Run(() => loadingForm.ShowDialog());
            //btn_update.PerformClick();
            //Length = 0.0;
            //Height = 0.0;



            var doc = acadApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;
            // int PP = dtb_gr1.RowCount;
            //int p = 1;
            //  dtb_gr1.Rows.RemoveAt(PP-1);
            //loadingForm.Invoke(new Action(() => loadingForm.Close()));
            PromptPointResult pointResult = ed.GetPoint("Chọn vị trí để chèn Block Reference: ");
            
            Double ct = OH + double.Parse(D.Substring(0,3));
            //----------------------------------

            //SizeView1 Frontview = new SizeView1()
            //{
            //    Size_L = a / 2 + 225,
            //    Size_R = a / 2 + 440,
            //    Size_H = ct + 500 + 500
            //};
            //SizeView1 Backview = new SizeView1()
            //{
            //    Size_L = a / 2 + 225,
            //    Size_R = a / 2 + 440,
            //    Size_H = ct + 500 + 500
            //};
            //SizeView1 Rightview = new SizeView1()
            //{
            //    Size_L = b / 2 + 235,
            //    Size_R = b / 2 + 320,
            //    Size_H = ct + 500 + 500
            //};
            //SizeView1 Lefttview = new SizeView1()
            //{
            //    Size_L = b / 2 + 235,
            //    Size_R = b / 2 + 320,
            //    Size_H = ct + 500 + 500
            //};
            //SizeView1 Topview = new SizeView1()
            //{
            //    Size_L = a / 2 + 1100,
            //    Size_R = a / 2 + 800,
            //    Size_H = b + 370 + 370
            //};
            //sizeview = new List<SizeView1>()
            //{
            //    Frontview,
            //    Rightview,
            //    Backview,
            //    Lefttview,
            //    Topview
            //};
            //// Màu cần kiểm tra
            //Color targetColor = Color.FromArgb(255, 224, 192);

            //// Tạo list bool theo điều kiện
            //List<bool> colorMatches = new List<bool>
            //{
            // btn_front.BackColor == targetColor,
            // btn_right.BackColor == targetColor,
            // btn_back.BackColor == targetColor,
            // btn_left.BackColor == targetColor,
            // btn_top.BackColor == targetColor
            //};
            //gapx = double.Parse(tb_gapx.Text);
            //gapy = double.Parse(tb_gapy.Text);
            //Length = gapx; // Khởi tạo Length về 0 trước khi tính toán
            //Height = 0.0; // Khởi tạo Height về 0 trước khi tính toán
            //string targetName = "";
            //double heightSum = 0.0; // Biến tạm để tính tổng chiều cao
            //double lengthSum = 0.0; // Biến tạm để tính tổng chiều dài
            //                        // Khoảng cách giữa các kích thước
            //                        // Duyệt qua danh sách sizeview và cộng kích thước tương ứng vào Length và Height
            //int index_view = 0; // Biến để theo dõi chỉ số view hiện tại
            //for (int i_view = 0; i_view < sizeview.Count; i_view++)
            //{
            //    if (colorMatches[i_view] && i_view != sizeview.Count - 1)
            //    {
            //        // Nếu màu trùng, thêm kích thước tương ứng vào Length
            //        Length += sizeview[i_view].Size_L + sizeview[i_view].Size_R + gapx; // Cộng thêm 600 cho khoảng cách
            //        heightSum += sizeview[i_view].Size_H + gapy; // Cộng thêm kích thước chiều cao
            //        index_view += 1; // Tăng chỉ số view hiện tại
            //    }
            //    if (colorMatches[i_view] && i_view == sizeview.Count - 1)
            //    {
            //        // Nếu là view cuối cùng, không cộng thêm khoảng cách
            //        Height += sizeview[i_view].Size_H + gapy; // Cộng thêm kích thước chiều cao
            //    }
            //    if (colorMatches[i_view] && i_view != sizeview.Count - 1 && i_view != 1)
            //    {
            //        lengthSum += sizeview[i_view].Size_L + sizeview[i_view].Size_R + gapx; // Cộng thêm kích thước chiều dài
            //    }
            //}
            //if (index_view != 0)
            //{
            //    Height = heightSum / index_view + Height; // Cập nhật tổng chiều cao chia cho số lượng view đã chọn 
            //}
            //if (Length < 7800 && Height < 4000)
            //{
            //    if (index_view != 0 && (heightSum / index_view - gapy) < 1780 && lengthSum < 5700)
            //    {
            //        scale = 1.0; // Nếu chiều cao trung bình nhỏ hơn 1780, không cần scale
            //        targetName = "WKV_20";
            //    }
            //    else
            //    {
            //        scale = 1.5; // Nếu chiều cao trung bình lớn hơn hoặc bằng 1780, áp dụng scale
            //        targetName = "WKV_30";
            //    }
            //    // scale = 1.0; // Nếu tổng chiều dài và chiều cao nhỏ hơn 7800 và 4000, không cần scale
            //}
            //else
            //{
            //    scale = 1.5; // Nếu tổng chiều dài hoặc chiều cao lớn hơn, áp dụng scale
            //    targetName = "WKV_30";
            //}
         
            //using (DocumentLock m_DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument())
            //{
            //    using (Transaction tr = db.TransactionManager.StartTransaction())
            //    {
            //        DimStyleTable dimTable = (DimStyleTable)tr.GetObject(db.DimStyleTableId, OpenMode.ForRead);
            //        foreach (string targetName in liststylename)
            //        {
            //            if (dimTable.Has(targetName))
            //            {

            //                dimstyle = dimTable[targetName];
            //                // ed.WriteMessage($"\ntest '{targetName}'.");
            //                // DimStyleTableRecord dimStyle = (DimStyleTableRecord)tr.GetObject(dimstyle, OpenMode.ForRead);

            //                //  ed.WriteMessage($"\nTìm thấy DimStyle: {dimstyle.}, ID: {dimstyle1}");
            //            }
            //            else
            //            {
            //                ed.WriteMessage($"\nKhông tìm thấy DimStyle có tên '{targetName}'.");
            //            }
            //        }
            //        tr.Commit();
            //    }
            //}
            //// Tính toán các giá trị cần thiết

            //double x = (index_view == 4) ? (pointResult.Value.X + (7800 - 2125) * scale - sizeview[4].Size_R) : (pointResult.Value.X + (7800 / 2) * scale);
            //double y = pointResult.Value.Y + 1640 * scale + 500;
            //double x1 = x + (sizeview[0].Size_R + sizeview[1].Size_L + gapx);
            //double x2 = x - (sizeview[0].Size_L + sizeview[3].Size_R + gapx);
            //double x3 = x2 - (sizeview[2].Size_R + sizeview[3].Size_L + gapx);
            //double x4 = x - (sizeview[0].Size_L + sizeview[2].Size_R + gapx);
            //double y2 = y + sizeview[0].Size_H - 500 + gapy + sizeview[4].Size_H / 2;
            //double L = w + m * 2 + k * 2 - d * 2;
            //double H = dh + m * 2 + k * 2 - f * 2;
            //Point3d p0 = new Point3d(x, y, 0);
            //Point3d p1 = new Point3d(x1, y, 0);
            //Point3d p2 = new Point3d(x2, y, 0);
            //Point3d p3 = new Point3d(x3, y, 0);
            //Point3d p4 = new Point3d(x, y2, 0);
            //Point3d p5 = new Point3d(x4, y, 0);
            //if (btn_front.BackColor == Color.FromArgb(255, 224, 192))
            //{
            //    mattruoc(p0, ed);

            //}
            //if (btn_top.BackColor == Color.FromArgb(255, 224, 192))
            //{
            //    if (rbt_lapsau.Checked == true)
            //    {
            //        mattren2(p4, ed);
            //    }
            //    else
            //    {
            //        mattren(p4, ed);
            //    }

            //}
            //if (btn_right.BackColor == Color.FromArgb(255, 224, 192))
            //{
            //    matphai(p1, ed);
            //    //  Length = Length + b + i * 2+750;
            //}
            //if (btn_left.BackColor == Color.FromArgb(255, 224, 192) && btn_back.BackColor == Color.FromArgb(255, 224, 192))
            //{
            //    mattrai(p2, ed);
            //    // Length = Length + b + i * 2 + 750;
            //    matsau(p3, ed);
            //    // Length = Length + a + i * 2+750;
            //}
            //else if (btn_back.BackColor == Color.FromArgb(255, 224, 192))
            //{
            //    matsau(p5, ed);
            //    //Length = Length + a + i * 2 + 750;
            //}
            //else if (btn_left.BackColor == Color.FromArgb(255, 224, 192))
            //{
            //    mattrai(p2, ed);
            //    // Length = Length + b + i * 2 + 750;
            //}
            //// Length += 800 * 2;
            //// if (Length > 7800) { scale = 1.5; }


            //HB_A3_CL = new List<string>() { };
            //WKV_KT1 = new List<string>() { tb_date.Text, tb_ken.Text, tb_hbname.Text };

            //double ss1 = a + i * 2;
            //double ss2 = a - d - e1 + 2 * k;
            //string selectedValue1 = (ss1 > ss2) ? ss1.ToString() : ss2.ToString();
            //double ss3 = b + i * 2;
            //double ss4 = b - f - g + 2 * k;
            //string selectedValue2 = (ss3 > ss4) ? ss3.ToString() : ss4.ToString();

            //WKV_KT3 = new List<string> { selectedValue1, selectedValue2, tb_caotong.Text };
            //WKV_KT4 = new List<string>() { lb_tong.Text, lb_de.Text, lb_than.Text, lb_nap.Text };
            //WKV_KT5 = new List<string>() { $"屋根スラブ配管取出最大開口寸法 {w},{dh} x {RoundToNearest(c - 90 - 60)}", "配管箱開口部最大開口寸法 " + w + " x " + dh };
            //WKV_KT6 = new List<string>() { $"1/{Math.Round(scale * 20, 0)}" };
            //WKV_Array = new List<string>[] { HB_A3_CL, WKV_KT1, WKV_KT2, WKV_KT3, WKV_KT4, WKV_KT5, WKV_KT6 };
            //int ind = 0;
            //List<string> listblockname_ins = new List<string>();
            //if (rbt_lapsau.Checked == true)
            //{
            //    listblockname_ins = listblockname2;
            //}
            //else
            //{
            //    listblockname_ins = listblockname;
            //}
            //foreach (string blockname in listblockname_ins)
            //{
            //    Insertdynamicblock.InsertBlock(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\HB-CD.dwg", blockname, new Point3d(pointResult.Value.X, pointResult.Value.Y, 0), WKV_Array[ind], scale, 0);
            //    ind++;
            //}
        }

        private void button5_Click_2(object sender, EventArgs e)
        {
            string filePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\TRATHONGSO.xlsl";
            var processor = new ExcelDataProcessor();
             processor.ProcessDataFromFile(filePath);
        }
    }
}
