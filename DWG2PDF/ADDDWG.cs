using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Reflection;

using System.Threading.Tasks;

using System.Collections.Specialized;
using System.Diagnostics;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

using Autodesk.AutoCAD.Runtime;
using AcAp = Autodesk.AutoCAD.ApplicationServices.Application;
using Autodesk.AutoCAD.PlottingServices;
namespace DWG2PDF
{
    class ADDDWG
    {

        public static void AttachingExternalReference(string PathName)
        {
            // Get the current database and start a transaction
            Document doc = AcAp.DocumentManager.MdiActiveDocument;
            Database acCurDb = doc.Database;
            // Yêu cầu người dùng chọn một điểm trên bản vẽ chính
            PromptPointOptions promptOptions = new PromptPointOptions("\nChọn điểm chèn: ");
            PromptPointResult promptResult = doc.Editor.GetPoint(promptOptions);

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                if (promptResult.Status == PromptStatus.OK)
                {
                    Point3d insertPoint = promptResult.Value;
                    // Create a reference to a DWG file
                    //  string PathName = "C:\\AutoCAD\\Sample\\Sheet Sets\\Architectural\\Res\\Exterior Elevations.dwg";
                    ObjectId acXrefId = acCurDb.AttachXref(PathName, "Exterior Elevations");

                    // If a valid reference is created then continue
                    if (!acXrefId.IsNull)
                    {

                        using (BlockReference acBlkRef = new BlockReference(insertPoint, acXrefId))
                        {
                            BlockTableRecord acBlkTblRec;
                            acBlkTblRec = acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                            acBlkTblRec.AppendEntity(acBlkRef);
                            acTrans.AddNewlyCreatedDBObject(acBlkRef, true);
                        }
                    }

                    // Save the new objects to the database

                }
                acTrans.Commit();
                // Dispose of the transaction
            }

        }
        public static void ImportDwgObjectsCommand(string dwgFilePath)
        {
            Document doc = AcAp.DocumentManager.MdiActiveDocument;
            Database destDb = doc.Database;
            Editor ed = doc.Editor;
            //try
            //{


                // Yêu cầu người dùng chọn một điểm trên bản vẽ chính
                PromptPointOptions promptOptions = new PromptPointOptions("\nChọn điểm chèn: ");
                PromptPointResult promptResult = doc.Editor.GetPoint(promptOptions);

                if (promptResult.Status == PromptStatus.OK)
                {
                    Point3d insertPoint = promptResult.Value;
                    // Prompt the user to enter the scale factor
                    var scaleResult = ed.GetDouble("\nEnter scale factor: ");
                    if (scaleResult.Status != PromptStatus.OK)
                    {
                        ed.WriteMessage("\nInvalid scale factor. File insertion aborted.");
                        return;
                    }
                    var scaleFactor = scaleResult.Value;
                    using (DocumentLock m_DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument())
                    {

                    
                    using (Transaction tr1 = destDb.TransactionManager.StartTransaction())
                    {

                        // Open the block table
                        
                        var bt = (BlockTable)tr1.GetObject(destDb.BlockTableId, OpenMode.ForRead, false);
                        ObjectIdCollection abc = ImportAllObjects(destDb, dwgFilePath, insertPoint);

                        // Create the block reference at the specified insertion point
                      //  BlockTableRecord modelSpace = (BlockTableRecord)tr1.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                        var modelSpace = (BlockTableRecord)tr1.GetObject(destDb.CurrentSpaceId, OpenMode.ForWrite);

                        foreach (ObjectId xyz in abc)
                            {
                                BlockReference blockReference = new BlockReference(insertPoint, xyz);
                                blockReference.ScaleFactors = new Scale3d(scaleFactor);

                                // Add the block reference to the drawing
                                modelSpace.AppendEntity(blockReference);
                                tr1.AddNewlyCreatedDBObject(blockReference, true);
                            }


                            // Save changes to the current document
                          //  doc.TransactionManager.QueueForGraphicsFlush();
                            tr1.Commit();
                        }

                    }
                    ed.WriteMessage("\nImported all objects from the DWG file.");
                }
            //}
            //catch (Autodesk.AutoCAD.Runtime.Exception ex)
            //{
            //    // var ed = AcAp.DocumentManager.MdiActiveDocument.Editor;
            //    MessageBox.Show("\nError during copy: " + ex.Message + "\n" + ex.StackTrace);
            //}
        }
        //--------------------------------------------------------------------
        public static ObjectIdCollection ImportAllObjects(Database destDb, string sourceFilename, Point3d p1)
        {
            var blockIds = new ObjectIdCollection();
            //try
            //{

                if (System.IO.File.Exists(sourceFilename))
                {
                    using (var sourceDb = new Database(false, true))
                    {
                        sourceDb.ReadDwgFile(sourceFilename, FileOpenMode.OpenForReadAndAllShare, true, "");
                        var sourceModelSpaceId = SymbolUtilityServices.GetBlockModelSpaceId(sourceDb);

                        using (var tr = new OpenCloseTransaction())
                        {
                            var sourceModelSpace = (BlockTableRecord)tr.GetObject(sourceModelSpaceId, OpenMode.ForRead);
                            foreach (ObjectId id in sourceModelSpace)
                            {
                                blockIds.Add(id);
                            }


                            //tr.Commit();
                        }

                        var mapping = new IdMapping();
                        sourceDb.WblockCloneObjects(blockIds, destDb.BlockTableId, mapping, DuplicateRecordCloning.Ignore, false);
                       // sourceDb.Dispose();
                        return blockIds;

                    }
                }

            //}
            //catch (Autodesk.AutoCAD.Runtime.Exception ex)
            //{
            //    //var ed = AcAp.DocumentManager.MdiActiveDocument.Editor;
            //    MessageBox.Show("\nError during copy: " + ex.Message + "\n" + ex.StackTrace);
            //}
            return blockIds;
            //-----------------------------------------------------------------

        }
        //public static void WBlockBetweenDataBase(string sourceFilename)
        //{
        //    using (DocumentLock m_DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument())
        //    {
        //        Database SourceDB = new Database(false, true);
        //    SourceDB.ReadDwgFile(sourceFilename, FileOpenMode.OpenForReadAndAllShare, true, "");
        //    var sourceModelSpaceId = SymbolUtilityServices.GetBlockModelSpaceId(SourceDB);
        //    // ErrorStatus Stats = this.GetInsertDwgFile(ref SourceDB);
        //    Document Doc = AcAp.DocumentManager.MdiActiveDocument;
        //    Editor ed = Doc.Editor;
        //    Database TargetDB = Doc.Database;



        //    PromptPointOptions promptOptions = new PromptPointOptions("\nChọn điểm chèn: ");
        //    PromptPointResult promptResult = Doc.Editor.GetPoint(promptOptions);

        //    if (promptResult.Status == PromptStatus.OK)
        //    {
        //        Point3d insertPoint = promptResult.Value;
        //        // Prompt the user to enter the scale factor
        //        var scaleResult = ed.GetDouble("\nEnter scale factor: ");
        //        if (scaleResult.Status != PromptStatus.OK)
        //        {
        //            ed.WriteMessage("\nInvalid scale factor. File insertion aborted.");
        //            return;
        //        }
        //        var scaleFactor = scaleResult.Value;
              


        //            ObjectIdCollection SourceIds = new ObjectIdCollection();
             
        //            using (Transaction Tr = SourceDB.TransactionManager.StartTransaction())
        //            {
        //                BlockTable SourceBt = (BlockTable)Tr.GetObject(SourceDB.BlockTableId, OpenMode.ForRead);
        //                using (Transaction Tr2 = TargetDB.TransactionManager.StartTransaction())
        //                {
        //                    var sourceModelSpace = (BlockTableRecord)Tr.GetObject(sourceModelSpaceId, OpenMode.ForRead);
        //                    foreach (ObjectId id in sourceModelSpace)
        //                    {
        //                        SourceIds.Add(id);
        //                    }

        //                    Tr2.Commit();
        //                }

        //                if (SourceIds.Count != 0)
        //                {
        //                    IdMapping IdMap = new IdMapping();
        //                    TargetDB.WblockCloneObjects(SourceIds, TargetDB.BlockTableId, IdMap, DuplicateRecordCloning.Ignore, false);
        //                    SourceDB.Dispose();
        //                    // return (true);
        //                }
        //                else
        //                {
        //                    SourceDB.Dispose();
        //                    // return (false);
        //                }

        //                BlockTableRecord modelSpace = (BlockTableRecord)Tr.GetObject(SourceBt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
        //                foreach (ObjectId xyz in SourceIds)
        //                {
        //                    BlockReference blockReference = new BlockReference(insertPoint, xyz);
        //                    blockReference.ScaleFactors = new Scale3d(scaleFactor);

        //                    // Add the block reference to the drawing
        //                    modelSpace.AppendEntity(blockReference);
        //                    Tr.AddNewlyCreatedDBObject(blockReference, true);
        //                }


        //                // Save changes to the current document
        //                Doc.TransactionManager.QueueForGraphicsFlush();
        //            }

        //        }
        //    }
        //}
    }
}
