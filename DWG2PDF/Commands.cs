using System;
using System.IO;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Windows;
using Autodesk.AutoCAD.Runtime;
using AcAp = Autodesk.AutoCAD.ApplicationServices.Application;
using Microsoft.Win32;
using System.Diagnostics;
using System.Reflection;
using Autodesk.AutoCAD.Interop.Common;
[assembly: CommandClass(typeof(DWG2PDF.Commands))]

namespace DWG2PDF
{
    public class Commands
    {
        public Form1 Px { get; set; }
        public  A3SETTING  PA3 { get; set; }
        
        //--------------------------------------------
        [CommandMethod("CRW")]
        public void exportexcel()
        {
           
                this.Px = new Form1();
                this.Px.FormClosing += Px_FormClosing;
                Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(this.Px);
            

        }
        //--------------------------------------------
        [CommandMethod("AH")]
        public void CreateHoles()
        {
            

           
                var doc = AcAp.DocumentManager.MdiActiveDocument;
                var db = doc.Database;
                var ed = doc.Editor;
                // int PP = dtb_gr1.RowCount;
                //int p = 1;
                //  dtb_gr1.Rows.RemoveAt(PP-1);
                PromptPointResult pointResult = ed.GetPoint("Chọn tâm hình chiếu bằng: ");
                // this.Px = new Form1();
                if (this.Px == null)
                {
                    this.Px = new Form1();
                    this.Px.FormClosing += Px_FormClosing;
                    Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(this.Px);
                }
                else
                {
                    this.Px.BringToFront();
                }
                Px.AutoHoles(pointResult.Value.X, pointResult.Value.Y);
            

        }

        private void Px_FormClosing(object sender, System.Windows.Forms.FormClosingEventArgs e)

        {
            this.Px = null;
        }
        //[CommandMethod("test")]

        //// [CommandMethod("GetLayoutNames")]
        //public void GetLayoutNames()
        //{
        //    Document doc = AcAp.DocumentManager.MdiActiveDocument;
        //    Database db = doc.Database;
        //    Editor ed = doc.Editor;

        //    using (Transaction trx = db.TransactionManager.StartTransaction())
        //    {
        //        DBDictionary layoutDic = trx.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;

        //        foreach (DBDictionaryEntry entry in layoutDic)
        //        {
        //            Layout lay = trx.GetObject(entry.Value, OpenMode.ForRead) as Layout;
        //            ed.WriteMessage("\n" + lay.LayoutName);
        //        }


        //    }
        //}

        //[CommandMethod("A3S")]
        //public void A3SETTING()
        //{
        //    if (!CHECKTIME.CHECKLICENSE(20251215))
        //    {

        //       MessageBox.Show("Giấy phép đã hết hạn\nFacebook: Plugins Autocad", "Thông báo giấy phép");
        //       

        //        // Mở trình duyệt web với liên kết đó
        //        Process.Start(link);


        //    }
        //    else
        //    {
        //        this.PA3 = new A3SETTING();
        //        this.PA3.FormClosing += PA3_FormClosing;
        //        Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(this.PA3);
        //    }

        //}

        //private void PA3_FormClosing(object sender, System.Windows.Forms.FormClosingEventArgs e)

        //{
        //    this.PA3 = null;
        //}

        //[CommandMethod("A3A")]
        //public void A3ADD()
        //{
        //    Document doc1 = AcAp.DocumentManager.MdiActiveDocument;
        // //   Database destDb1 = doc1.Database;
        //    Editor ed = doc1.Editor;
        //    String PATH = ResxFileHelper.ReadFromResxFile(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Properties\Resources.resx", "Pathinputa3");
        //    String GTCBB2 = ResxFileHelper.ReadFromResxFile(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Properties\Resources.resx", "Blockorxref");
        //    if (File.Exists(PATH))
        //    {
        //        if (GTCBB2 == "Block")
        //        {
        //            ADDDWG.ImportDwgObjectsCommand(PATH);
        //           // ADDDWG.WBlockBetweenDataBase(PATH);

        //        }
        //        else
        //        {
        //            ADDDWG.AttachingExternalReference(PATH);
        //        }
        //    }
        //    else
        //    { ed.WriteMessage("Không tìm thấy đường dẫn tới file dwg"); }
        //}

        }

}

