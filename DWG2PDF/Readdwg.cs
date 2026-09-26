using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
namespace DWG2PDF
{
    public class DWGReader
    {
       
            public static List<List<Extents2d>> GetBlockExtents(string filePath, string blockName)
        {
            List<List<Extents2d>> extentsList = new List<List<Extents2d>>();

            Database db = new Database(false, true);
            db.ReadDwgFile(filePath, System.IO.FileShare.ReadWrite, false, "");

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                // Get ModelSpace extents
               // BlockTableRecord modelSpace = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
              //  List<Extents2d> modelSpaceExtents = GetBlockExtentsFromBlockTableRecord(tr, modelSpace, blockName);
               // extentsList.Add(modelSpaceExtents);
                int i = 0;
                // Get Layout extents
                foreach (ObjectId layoutId in bt)
                {
                 //  
                    BlockTableRecord layout = (BlockTableRecord)tr.GetObject(layoutId, OpenMode.ForRead);
                  
                    
                        List<Extents2d> layoutExtents = GetBlockExtentsFromBlockTableRecord(tr, layout, blockName);
                    if (layoutExtents.Count>0)
                    {
                        extentsList.Add(layoutExtents);
                        i++;
                    }
                   // }
                }
                MessageBox.Show(i.ToString());
                tr.Commit();
            }

            db.Dispose();

            return extentsList;
        }

        private static List<Extents2d> GetBlockExtentsFromBlockTableRecord(Transaction tr, BlockTableRecord btr, string blockName)
        {
            List<Extents2d> extentsList = new List<Extents2d>();

            foreach (ObjectId objId in btr)
            {
                Entity entity = (Entity)tr.GetObject(objId, OpenMode.ForRead);
                BlockReference blockRef = entity as BlockReference;
                if (entity != null && blockRef!=null && blockRef.Name == blockName)
            {
                    Extents3d? extentsNullable = blockRef.Bounds;
                    Extents3d extents = extentsNullable ?? new Extents3d();
                    Point2d minPoint2d = new Point2d(extents.MinPoint.X, extents.MinPoint.Y);
                    Point2d maxPoint2d = new Point2d(extents.MaxPoint.X, extents.MaxPoint.Y);
                    Extents2d extents2d = new Extents2d(minPoint2d, maxPoint2d);
                    extentsList.Add(extents2d);
                }
            }

            return extentsList;
        }




    }
}
