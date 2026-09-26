using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Windows;
using Autodesk.AutoCAD.Runtime;
using AcAp = Autodesk.AutoCAD.ApplicationServices.Application;
using System.Windows.Forms;
namespace DWG2PDF
{
    class select
    {
        public static string SelectBlock()
        {
            Document acDoc = AcAp.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            Editor acEd = acDoc.Editor;
            string blockName="";
            PromptEntityOptions promptOptions = new PromptEntityOptions("\nSelect a block: ");
            promptOptions.SetRejectMessage("Please select only block references.");
            promptOptions.AddAllowedClass(typeof(BlockReference), exactMatch: true);

            PromptEntityResult promptResult = acEd.GetEntity(promptOptions);

            if (promptResult.Status == PromptStatus.OK)
            {
                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    BlockReference blockRef = acTrans.GetObject(promptResult.ObjectId, OpenMode.ForRead) as BlockReference;

                    if (blockRef != null)
                    {
                         blockName = blockRef.Name;
                        
                    }

                    acTrans.Commit();
                }
            }
            return blockName;
        }
        //-----------------------------------
        public static string SelectPolyline()
        {
            Document acDoc = AcAp.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            Editor acEd = acDoc.Editor;
            string layername = "";
            PromptEntityOptions promptOptions = new PromptEntityOptions("\nSelect a polyline: ");
            promptOptions.SetRejectMessage("Please select only polylines.");
            promptOptions.AddAllowedClass(typeof(Polyline), exactMatch: true);

            PromptEntityResult promptResult = acEd.GetEntity(promptOptions);

            if (promptResult.Status == PromptStatus.OK)
            {
                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    Polyline polyline = acTrans.GetObject(promptResult.ObjectId, OpenMode.ForRead) as Polyline;

                    if (polyline != null)
                    {
                         layername = polyline.Layer;

                      
                    }

                    acTrans.Commit();
                }
            }
            return layername;
        }
        //---------------------------------
        public static List<Extents2d> ScanBlocks(List<string> blockNames)
        {
        
            List<Extents2d> blockExtents = new List<Extents2d>();

            Document doc = AcAp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord modelSpace = trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;

                foreach (ObjectId objId in modelSpace)
                {
                    Entity ent = trans.GetObject(objId, OpenMode.ForRead) as Entity;

                    if (ent is BlockReference blockRef)
                    {
                        // Lấy tên thực của block định nghĩa
                        BlockTableRecord blockDef = trans.GetObject(blockRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                        if (blockDef.Name.Equals(blockNames[0], StringComparison.OrdinalIgnoreCase)|| blockDef.Name.Equals(blockNames[1], StringComparison.OrdinalIgnoreCase))
                        {
                            if (blockRef.Bounds.HasValue)
                            {
                                Extents3d extents = blockRef.Bounds.Value;

                                Extents2d extent2d = new Extents2d(
                                    new Point2d(extents.MinPoint.X, extents.MinPoint.Y),
                                    new Point2d(extents.MaxPoint.X, extents.MaxPoint.Y)
                                );

                                blockExtents.Add(extent2d);
                            }
                        }
                    }
                }

                trans.Commit();
            }

            return blockExtents;
        

        }

        private static SelectionFilter GetBlockSelectionFilter(string blockName)
        {
            TypedValue[] values = new TypedValue[]
            {
            new TypedValue((int)DxfCode.Start, "INSERT"),
            new TypedValue((int)DxfCode.BlockName, blockName)
            };

            return new SelectionFilter(values);
        }
        //-------------------------------------------------
        public static List<Extents2d> GetPolylineExtents(string layername)
        {
            List<Extents2d> extentsList = new List<Extents2d>();

            Document doc = AcAp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor editor = doc.Editor;

            PromptSelectionResult selectionResult = editor.GetSelection();
            if (selectionResult.Status != PromptStatus.OK)
            {
                editor.WriteMessage("\nNo objects selected.");
                return extentsList;
            }

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                SelectionSet selectionSet = selectionResult.Value;
                foreach (SelectedObject selectedObject in selectionSet)
                {
                    if (selectedObject.ObjectId.ObjectClass.DxfName == "LWPOLYLINE")
                    {
                        Polyline polyline = tr.GetObject(selectedObject.ObjectId, OpenMode.ForRead) as Polyline;
                        if (polyline != null&& polyline.Layer == layername)
                        {
                            Extents3d extents = polyline.GeometricExtents;
                            Point3d minPoint = extents.MinPoint;
                            Point3d maxPoint = extents.MaxPoint;

                            // Chuyển đổi tọa độ 3D thành tọa độ 2D
                            Point2d minPoint2d = new Point2d(minPoint.X, minPoint.Y);
                            Point2d maxPoint2d = new Point2d(maxPoint.X, maxPoint.Y);

                            Extents2d extent2d = new Extents2d(minPoint2d, maxPoint2d);
                           // blockExtents.Add(extent2d);
                            extentsList.Add(extent2d);
                        }
                    }

                }

                tr.Commit();
            }
            extentsList= Extents2dFilter.RemoveDuplicates(extentsList);
            return extentsList;
        }
        public static List<Extents2d> SCANBLOCKPOLYLINE(string blockName,string layername)
        {
            List<Extents2d> extentsList = new List<Extents2d>();

            Document doc = AcAp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor editor = doc.Editor;

            PromptSelectionResult selectionResult = editor.GetSelection();
            if (selectionResult.Status != PromptStatus.OK)
            {
                editor.WriteMessage("\nNo objects selected.");
                return extentsList;
            }

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                SelectionSet selectionSet = selectionResult.Value;
                foreach (SelectedObject selectedObject in selectionSet)
                {
                    if (selectedObject.ObjectId.ObjectClass.DxfName == "LWPOLYLINE")
                    {
                        Polyline polyline = tr.GetObject(selectedObject.ObjectId, OpenMode.ForRead) as Polyline;
                        if (polyline != null && polyline.Layer == layername)
                        {
                            Extents3d extents = polyline.GeometricExtents;
                            Point3d minPoint = extents.MinPoint;
                            Point3d maxPoint = extents.MaxPoint;

                            // Chuyển đổi tọa độ 3D thành tọa độ 2D
                            Point2d minPoint2d = new Point2d(minPoint.X, minPoint.Y);
                            Point2d maxPoint2d = new Point2d(maxPoint.X, maxPoint.Y);

                            Extents2d extent2d = new Extents2d(minPoint2d, maxPoint2d);
                            // blockExtents.Add(extent2d);
                            extentsList.Add(extent2d);
                        }
                    }
                    if (selectedObject.ObjectId.ObjectClass.DxfName == "INSERT")
                    {
                        BlockReference blockReference = tr.GetObject(selectedObject.ObjectId, OpenMode.ForRead) as BlockReference;

                        if (blockReference != null && blockReference.Name == blockName)
                        {
                            Extents3d? extentsNullable = blockReference.Bounds;
                            Extents3d extents = extentsNullable ?? new Extents3d();
                            Point3d minPoint = extents.MinPoint;
                            Point3d maxPoint = extents.MaxPoint;

                            // Chuyển đổi tọa độ 3D thành tọa độ 2D
                            Point2d minPoint2d = new Point2d(minPoint.X, minPoint.Y);
                            Point2d maxPoint2d = new Point2d(maxPoint.X, maxPoint.Y);

                            Extents2d extent2d = new Extents2d(minPoint2d, maxPoint2d);
                            extentsList.Add(extent2d);
                        }
                    }
                }

                tr.Commit();
            }
            extentsList= Extents2dFilter.RemoveDuplicates(extentsList);
            return extentsList;
        }
        public static List<Extents2d> ScanBlocksfile(string blockName, string layoutname, Document doc)
        {
            List<Extents2d> blockExtents = new List<Extents2d>();

           // Document doc = AcAp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor editor = doc.Editor;

           
                using (Transaction trans = db.TransactionManager.StartOpenCloseTransaction())
            {

                BlockTableRecord currentSpace = trans.GetObject(db.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;

                // Tạo một bộ chọn (SelectionSet) để chứa các đối tượng polyline thuộc layername
                //PromptSelectionResult selectionResult = null;

                // Tạo một bộ chọn (SelectionSet) để chứa các đối tượng polyline thuộc layername
                PromptSelectionResult selectionResult = null;
                TypedValue[] filter = new TypedValue[]
                {
                new TypedValue((int)DxfCode.LayoutName, layoutname)              
              
                };
                SelectionFilter selectionFilter = new SelectionFilter(filter);
                selectionResult = doc.Editor.SelectAll(selectionFilter);
                if (selectionResult.Status == PromptStatus.OK)
                {
                    SelectionSet selectionSet = selectionResult.Value;
                    foreach (SelectedObject selectedObject in selectionSet)
                    {
                        if (selectedObject.ObjectId.ObjectClass.DxfName == "INSERT")
                        {
                            BlockReference blockReference = trans.GetObject(selectedObject.ObjectId, OpenMode.ForRead) as BlockReference;

                            if (blockReference != null && blockReference.Name == blockName)
                            {
                                Extents3d? extentsNullable = blockReference.Bounds;
                                Extents3d extents = extentsNullable ?? new Extents3d();
                                Point3d minPoint = extents.MinPoint;
                                Point3d maxPoint = extents.MaxPoint;

                                // Chuyển đổi tọa độ 3D thành tọa độ 2D
                                Point2d minPoint2d = new Point2d(minPoint.X, minPoint.Y);
                                Point2d maxPoint2d = new Point2d(maxPoint.X, maxPoint.Y);

                                Extents2d extent2d = new Extents2d(minPoint2d, maxPoint2d);
                                blockExtents.Add(extent2d);
                            }
                        }
                    }

                    trans.Commit();
                }
            }
            blockExtents= Extents2dFilter.RemoveDuplicates(blockExtents);
            return blockExtents;
        }
        //------------------------------------------------------------------------------------------------------------------------------------------------
        public static List<Extents2d> GetPolylineExtentsfile(string layername, string layoutname, Document doc)
        {
            List<Extents2d> polylineExtents = new List<Extents2d>();

         
            Database db = doc.Database;

            using (Transaction trans = db.TransactionManager.StartTransaction())
            {


                BlockTableRecord currentSpace = trans.GetObject(db.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;

                // Tạo một bộ chọn (SelectionSet) để chứa các đối tượng polyline thuộc layername
                //PromptSelectionResult selectionResult = null;

                // Tạo một bộ chọn (SelectionSet) để chứa các đối tượng polyline thuộc layername
                PromptSelectionResult selectionResult = null;
                    TypedValue[] filter = new TypedValue[]
                    {
                new TypedValue((int)DxfCode.LayoutName, layoutname),
                new TypedValue((int)DxfCode.Start, "LWPOLYLINE"),
                new TypedValue((int)DxfCode.LayerName, layername)
               
                    };
                    SelectionFilter selectionFilter = new SelectionFilter(filter);
                    selectionResult = doc.Editor.SelectAll(selectionFilter);

                    if (selectionResult.Status == PromptStatus.OK)
                    {
                        SelectionSet selectionSet = selectionResult.Value;
                        foreach (SelectedObject selectedObject in selectionSet)
                        {
                            if (selectedObject.ObjectId.ObjectClass.DxfName == "LWPOLYLINE")
                            {
                                Polyline polyline = trans.GetObject(selectedObject.ObjectId, OpenMode.ForRead) as Polyline;
                                if (polyline != null)
                                {
                                    Extents3d? extentsNullable = polyline.Bounds;
                                    Extents3d extents = extentsNullable ?? new Extents3d();
                                    Point3d minPoint = extents.MinPoint;
                                    Point3d maxPoint = extents.MaxPoint;

                                    // Chuyển đổi tọa độ 3D thành tọa độ 2D
                                    Point2d minPoint2d = new Point2d(minPoint.X, minPoint.Y);
                                    Point2d maxPoint2d = new Point2d(maxPoint.X, maxPoint.Y);

                                    Extents2d extent2d = new Extents2d(minPoint2d, maxPoint2d);
                                    polylineExtents.Add(extent2d);
                                }
                            }
                        }
                    }
                

                trans.Commit();
            }
            polylineExtents = Extents2dFilter.RemoveDuplicates(polylineExtents);
          
            return polylineExtents;
        }
  





















    }

}

                    
                      