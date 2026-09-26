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
        public static void ImportDimStyle(string sourceFile, string dimStyleName, string dimStyleName2, string dimStyleName3, string dimStyleName4, string dimStyleName5, string dimStyleName6, string dimStyleName7, string dimStyleName8)
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
                        GetDimStyleId(sourceDb, dimStyleName, dimStyleName2, dimStyleName3, dimStyleName4, dimStyleName5, dimStyleName6, dimStyleName7, dimStyleName8),
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
        private static ObjectIdCollection GetDimStyleId(Database db, string dimStyleName, string dimStyleName2, string dimStyleName3, string dimStyleName4, string dimStyleName5, string dimStyleName6, string dimStyleName7, string dimStyleName8)
        {
            ObjectIdCollection ids = new ObjectIdCollection();
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                DimStyleTable dst = tr.GetObject(db.DimStyleTableId, OpenMode.ForRead) as DimStyleTable;
                if (dst.Has(dimStyleName) && dst.Has(dimStyleName2) && dst.Has(dimStyleName3) && dst.Has(dimStyleName4) && dst.Has(dimStyleName5) && dst.Has(dimStyleName6) && dst.Has(dimStyleName7) && dst.Has(dimStyleName8))
                {
                    ids.Add(dst[dimStyleName]);
                    ids.Add(dst[dimStyleName2]);
                    ids.Add(dst[dimStyleName3]);
                    ids.Add(dst[dimStyleName4]);
                    ids.Add(dst[dimStyleName5]);
                    ids.Add(dst[dimStyleName6]);
                    ids.Add(dst[dimStyleName7]);
                    ids.Add(dst[dimStyleName8]);
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
