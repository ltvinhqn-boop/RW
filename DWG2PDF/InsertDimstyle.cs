using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using System.Windows.Forms;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
namespace DWG2PDF
{


    public class DimStyleImporter
    {
        [CommandMethod("ImportDimStyle")]
        public static void ImportDimStyle(string sourceFile,List< string> dimStyleName)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database destDb = doc.Database;
            //string sourceFile = @"C:\Path\To\DimStyleTemplate.dwg"; // Đường dẫn tới file chứa DimStyle
            // string dimStyleName = "MyDimStyle"; // Tên DimStyle cần nhập

            using (Database sourceDb = new Database(false, true))
            {
                sourceDb.ReadDwgFile(sourceFile, FileShare.ReadWrite, true, null);

                using (Transaction tr = destDb.TransactionManager.StartTransaction())
                {
                    // Chuẩn bị để chèn DimStyle
                    IdMapping mapping = new IdMapping();
                    sourceDb.WblockCloneObjects(
                        GetDimStyleId(sourceDb, dimStyleName),
                        destDb.DimStyleTableId,
                        mapping,
                        DuplicateRecordCloning.Replace, // hoặc Ignore
                        false
                    );

                    tr.Commit();
                    doc.Editor.WriteMessage($"\nĐã chèn DimStyle '{dimStyleName}' thành công.");
                }
            }
        }

        // Trả về ObjectIdCollection chứa DimStyle cần chèn
        private static ObjectIdCollection GetDimStyleId(Database db,List< string> dimStyleName)
        {
            ObjectIdCollection ids = new ObjectIdCollection();
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                DimStyleTable dst = tr.GetObject(db.DimStyleTableId, OpenMode.ForRead) as DimStyleTable;
                if (dimStyleName.All(name => dst.Has(name)))
                {
                    foreach(string dimname in dimStyleName)
                    {
                        ids.Add(dst[dimname]);
                    }    
                   
                    
                }
                else
                {
                    // throw new System.Exception($"Không tìm thấy DimStyle '{dimStyleName}' trong file nguồn.");
                }
                tr.Commit();
            }
            return ids;
        }
    }

}
