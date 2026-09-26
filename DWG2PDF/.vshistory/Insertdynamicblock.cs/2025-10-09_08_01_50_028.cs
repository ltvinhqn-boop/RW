using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Interop.Common;
using acadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using System.Windows.Forms;
using Autodesk.AutoCAD.Colors;
using System.Globalization;
//using static iTextSharp.text.pdf.events.IndexEvents;
namespace DWG2PDF
{
    // internal class Insertdynamicblock
    // {
    //     public static void InsertDynamicBlock(string blockPath, string blockName, Point3d insertPoint,
    //Dictionary<string, double> dynamicProperties, System.Windows.Forms.ComboBox cbbtest,double scale = 1.0)
    //     {
    //         var doc = acadApp.DocumentManager.MdiActiveDocument;
    //         var db = doc.Database;
    //         var ed = doc.Editor;

    //         using (DocumentLock docLock = doc.LockDocument())
    //         using (Transaction tr = db.TransactionManager.StartTransaction())
    //         {
    //             BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
    //             ObjectId btrId = bt.Has(blockName) ? bt[blockName] : ImportBlock(db, blockName, blockPath);

    //             if (btrId == ObjectId.Null)
    //             {
    //                 ed.WriteMessage($"\nKhông tìm thấy block '{blockName}' trong bản vẽ hoặc file nguồn.");
    //                 return;
    //             }

    //             BlockTableRecord modelSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
    //             BlockReference br = new BlockReference(insertPoint, btrId)
    //             {
    //                 ScaleFactors = new Scale3d(scale)
    //             };
    //             modelSpace.AppendEntity(br);
    //             tr.AddNewlyCreatedDBObject(br, true);

    //             // Chỉ xử lý các thuộc tính động (Dynamic Properties)
    //             if (br.IsDynamicBlock)
    //             {//217536879
    //                 foreach (DynamicBlockReferenceProperty prop in br.DynamicBlockReferencePropertyCollection)
    //                 {
    //                     cbbtest.Items.Add(prop.PropertyName+"-"+prop.Value);
    //                     if (dynamicProperties.TryGetValue(prop.PropertyName, out double val))
    //                     {
    //                         if (!prop.ReadOnly)
    //                         {
    //                             try
    //                             {
    //                                 prop.Value = val;
    //                                // br.RecordGraphicsModified(true);
    //                                // br.DynamicBlockTableRecord = br.DynamicBlockTableRecord; // force update
    //                                 doc.Editor.Regen(); // hoặc acadDoc.SendCommand("BATTUPDATE\n");
    //                             }
    //                             catch
    //                             {
    //                                 ed.WriteMessage($"\nKhông thể gán giá trị '{val}' cho thuộc tính động '{prop.PropertyName}'.");
    //                             }
    //                         }
    //                     }
    //                 }
    //             }


    //             BlockTableRecord blockDef = (BlockTableRecord)tr.GetObject(br.DynamicBlockTableRecord, OpenMode.ForWrite);
    //             // Tạo hatch mới trong block


    //             // Tìm các entity có thể hatch được
    //             ObjectIdCollection loopIds = new ObjectIdCollection();
    //             foreach (ObjectId entId in blockDef)
    //             {
    //                 Entity ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
    //                 if ((ent is Polyline pl && pl.Closed) && ent.ColorIndex == 5)
    //                 {
    //                     loopIds.Add(entId);
    //                 }
    //             }
    //             Hatch hatch = new Hatch();
    //             hatch.ColorIndex = 5; // Màu xanh
    //             hatch.SetDatabaseDefaults();

    //             // 🟢 CHỈ GỌI SetHatchPattern SAU KHI hatch được thêm vào DB
    //             blockDef.AppendEntity(hatch);
    //             tr.AddNewlyCreatedDBObject(hatch, true);

    //             // Gọi SetHatchPattern sau khi hatch đã vào database
    //             hatch.PatternScale = 500; // Tỉ lệ hatch
    //             hatch.SetHatchPattern(HatchPatternType.PreDefined, "ANSI31");

    //             hatch.Associative = true;
    //             // Nếu có ít nhất 1 vòng kín thì tạo hatch
    //             if (loopIds.Count > 0)
    //             {

    //                 hatch.AppendLoop(HatchLoopTypes.Default, loopIds);
    //                 hatch.EvaluateHatch(true);
    //                 doc.Editor.Regen();
    //             }
    //             br.RecordGraphicsModified(true);
    //             tr.Commit();
    //             ForceUpdateBlockRecord(br.ObjectId);
    //             //using (Transaction tr2 = db.TransactionManager.StartTransaction())
    //             //{
    //             //    Hatch hatchEval = tr2.GetObject(hatch.ObjectId, OpenMode.ForWrite) as Hatch;
    //             //    hatchEval.EvaluateHatch(true);
    //             //    tr2.Commit();
    //             //}
    //         }

    //    }
    internal class Insertdynamicblock
    {
        struct TextInfo
        {
            public Point3d Position { get; private set; }
            public Point3d Alignment { get; private set; }
            public bool IsAligned { get; private set; }
            public double Rotation { get; private set; }
            public TextInfo(Point3d position, Point3d alignment, bool aligned, double rotation)
            {
                Position = position;
                Alignment = alignment;
                IsAligned = aligned;
                Rotation = rotation;
            }
        }
        public static void InsertBlock(string blockPath, string blockName, Point3d pointResult, List<string> valuetable, double scale, double ang)
        {
            var doc = acadApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;
            //  PromptPointResult pointResult = ed.GetPoint("Chọn vị trí để chèn Block Reference: ");
            using (DocumentLock m_DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument())
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    ObjectId btrId = bt.Has(blockName) ?
                        bt[blockName] :
                        ImportBlock(db, blockName, blockPath);
                    if (btrId.IsNull)
                    {
                        ed.WriteMessage($"\nBlock '{blockName}' not found.");
                        return;
                    }
                    // btrId.
                    var cSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                    var br = new BlockReference(pointResult, btrId);
                    br.ScaleFactors = new Scale3d(scale, scale, scale);
                    br.Rotation = ang;
                    cSpace.AppendEntity(br);
                    tr.AddNewlyCreatedDBObject(br, true);

                    // add attribute references to the block reference
                    var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                    var attInfos = new Dictionary<string, TextInfo>();
                    if (btr.HasAttributeDefinitions)
                    {
                        int i = 0;
                        foreach (ObjectId id in btr)
                        {
                            if (id.ObjectClass.DxfName == "ATTDEF")
                            {
                                var attDef = (AttributeDefinition)tr.GetObject(id, OpenMode.ForRead);
                                attInfos[attDef.Tag] = new TextInfo(
                                    attDef.Position,
                                    attDef.AlignmentPoint,
                                    attDef.Justify != AttachmentPoint.BaseLeft,
                                    attDef.Rotation);
                                var attRef = new AttributeReference();
                                //  MessageBox.Show(valuetable.Count.ToString());
                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);

                                attRef.TextString = valuetable[i];

                                br.AttributeCollection.AppendAttribute(attRef);
                                tr.AddNewlyCreatedDBObject(attRef, true);
                                // ed.WriteMessage("\n" + attDef.Tag + "-" + attDef.Position + attDef.TextString);
                                i++;

                            }
                        }
                    }

                    tr.Commit();
                }
            }
        }
        private static void ModifyDimensionsInBlock(Transaction tr, BlockTableRecord blockDef, double newTextHeight, double newArrowSize)
        {
            foreach (ObjectId entId in blockDef)
            {
                Entity ent = tr.GetObject(entId, OpenMode.ForWrite) as Entity;
                if (ent is Dimension dim)
                {
                    dim.Dimtxt = newTextHeight;
                    dim.Dimasz = newArrowSize;
                }
                // Nếu có block lồng bên trong, xử lý đệ quy
                else if (ent is BlockReference nestedBr)
                {
                    BlockTableRecord nestedBtr = (BlockTableRecord)tr.GetObject(nestedBr.BlockTableRecord, OpenMode.ForRead);
                    ModifyDimensionsInBlock(tr, nestedBtr, newTextHeight, newArrowSize);
                }
            }
        }

        public static ObjectId InsertDynamicBlock(string blockPath, string blockName, Point3d insertPoint,
        Dictionary<string, double> dynamicProperties/*, System.Windows.Forms.ComboBox cbbtest*//*, double th*/, ObjectId dimstyle1, out Hatch hatch, double scale = 1.0)
        {
            var doc = acadApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;
            //  ObjectId dimstyle1 = ObjectId.Null;
            using (DocumentLock m_DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument())
            {
               // ObjectId btrId;
                BlockReference br;
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    hatch = new Hatch();
                    string patternName = "ANSI31";
                         int colorIndex = 5;
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    ObjectId btrId = bt.Has(blockName) ? bt[blockName] : ImportBlock(db, blockName, blockPath);

                    if (btrId == ObjectId.Null)
                    {
                        ed.WriteMessage($"\nKhông tìm thấy block '{blockName}' trong bản vẽ hoặc file nguồn.");
                        return ObjectId.Null;
                    }

                    BlockTableRecord modelSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                     br = new BlockReference(insertPoint, btrId)
                    {
                        ScaleFactors = new Scale3d(scale)
                    };
                    modelSpace.AppendEntity(br);
                    tr.AddNewlyCreatedDBObject(br, true);
                    BlockTableRecord blockDef = (BlockTableRecord)tr.GetObject(br.DynamicBlockTableRecord, OpenMode.ForWrite);
                    // Kiểm tra nếu đã có hatch, không tạo hatch mới
                   // bool hatchExists = false;

                    // Tìm các entity có thể hatch được
                    // ObjectIdCollection loopIds = new ObjectIdCollection();
                    foreach (ObjectId entId in blockDef)
                    {
                        Entity ent = tr.GetObject(entId, OpenMode.ForWrite) as Entity;
                        if (ent is Dimension dimension)
                        {
                            dimension.DimensionStyle = dimstyle1;
                        }
                        ////  else {  dimstyle1 = ObjectId.Null; }
                        //  if (ent is Hatch hatchEntity && hatchEntity.ObjectId != ObjectId.Null)
                        //  {
                        //      // Nếu có hatch tồn tại, không cần tạo hatch mới
                        //      hatchExists = true;
                        //      break;
                        //  }
                        //  else if (ent is Polyline pl && pl.Closed && (ent.ColorIndex == 5 || ent.ColorIndex == 6))
                        //  {
                        //      if (ent.ColorIndex == 5)
                        //      {
                        //          patternName = "ANSI31";
                        //          colorIndex = 5;
                        //      }
                        //      else if (ent.ColorIndex == 6)
                        //      {
                        //          patternName = "ANSI37";
                        //          colorIndex = 6;
                        //      }
                        //      loopIds.Add(entId);
                        //  }
                    }
                    if (br.IsDynamicBlock)
                    {
                        foreach (DynamicBlockReferenceProperty prop in br.DynamicBlockReferencePropertyCollection)
                        {
                            // cbbtest.Items.Add(prop.PropertyName + "-" + prop.Value);
                            if (dynamicProperties.TryGetValue(prop.PropertyName, out double val))
                            {
                                if (!prop.ReadOnly)
                                {
                                    try
                                    {
                                        prop.Value = val;
                                        // br.RecordGraphicsModified(true);
                                        // br.DynamicBlockTableRecord = br.DynamicBlockTableRecord; // force update
                                        // hoặc acadDoc.SendCommand("BATTUPDATE\n");
                                    }
                                    catch
                                    {
                                        ed.WriteMessage($"\nKhông thể gán giá trị '{val}' cho thuộc tính động '{prop.PropertyName}'.");
                                    }
                                }
                            }
                        }


                        // br.DynamicBlockTableRecord = br.DynamicBlockTableRecord;
                    }
                 //   DBObjectCollection explodedObjs = new DBObjectCollection();
                    string uniqueName = "StaticBlock_" + Guid.NewGuid().ToString("N").Substring(0, 8);
                    br.ConvertToStaticBlock(uniqueName);
                    //------------------------------------------------------------
                    //BlockTableRecord blockDef1 = (BlockTableRecord)tr.GetObject(br.BlockTableRecord, OpenMode.ForWrite);
                    //// Kiểm tra nếu đã có hatch, không tạo hatch mới
                    //bool hatchExists = false;

                    //// Tìm các entity có thể hatch được
                    //ObjectIdCollection loopIds = new ObjectIdCollection();
                    //foreach (ObjectId entId in blockDef1)
                    //{
                    //    Entity ent = tr.GetObject(entId, OpenMode.ForWrite) as Entity;

                    //    //  else {  dimstyle1 = ObjectId.Null; }
                    //    if (ent is Hatch hatchEntity && hatchEntity.ObjectId != ObjectId.Null)
                    //    {
                    //        // Nếu có hatch tồn tại, không cần tạo hatch mới
                    //        hatchExists = true;
                    //        break;
                    //    }
                    //    else if (ent is Polyline pl && pl.Closed && (ent.ColorIndex == 5 || ent.ColorIndex == 6))
                    //    {
                    //        if (ent.ColorIndex == 5)
                    //        {
                    //            patternName = "ANSI31";
                    //            colorIndex = 5;
                    //        }
                    //        else if (ent.ColorIndex == 6)
                    //        {
                    //            patternName = "ANSI37";
                    //            colorIndex = 6;
                    //        }
                    //        loopIds.Add(entId);
                    //    }
                    //}
                    //if (!hatchExists)
                    //{
                    //    // Nếu chưa có hatch, tạo hatch mới

                    //    hatch.ColorIndex = colorIndex; // Màu xanh
                    //    hatch.SetDatabaseDefaults();
                    //    hatch.LineWeight = LineWeight.LineWeight009; // Đặt độ dày đường viền
                    //    // 🟢 CHỈ GỌI SetHatchPattern SAU KHI hatch được thêm vào DB
                    //    blockDef1.AppendEntity(hatch);
                    //    tr.AddNewlyCreatedDBObject(hatch, true);

                    //    // Gọi SetHatchPattern sau khi hatch đã vào database
                    //    hatch.PatternScale = 500; // Tỉ lệ hatch
                    //                              //  if(ent.ColorIndex == 5)
                    //    hatch.SetHatchPattern(HatchPatternType.PreDefined, patternName);

                    //    hatch.Associative = true;

                    //    // Nếu có ít nhất 1 vòng kín thì tạo hatch
                    //    if (loopIds.Count > 0)
                    //    {
                    //        hatch.AppendLoop(HatchLoopTypes.Default, loopIds);
                    //        //  hatch.EvaluateHatch(true);
                    //        // doc.Editor.Regen();
                    //    }
                    //    hatch.EvaluateHatch(true); // Cập nhật đường biên
                    //    br.RecordGraphicsModified(true);
                    //}
                    //------------------------------------------------------------
                    doc.Editor.Regen();
                    br.ExplodeToOwnerSpace(); // X đối tượng chèn // test bật lên khi x
                                              // hatch.UpgradeOpen();

                    //// hatch.EvaluateHatch(true); // Cập nhật đường biên
                    //br.Erase();// test bật lên khi x
                    ////  Entity ent = tr.GetObject(entId, OpenMode.ForWrite) as Entity;
                    ////BlockTableRecord blockDef = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForWrite);
                    //// test bật lên khi x
                    //if (!blockDef.IsErased && blockDef.GetBlockReferenceIds(true, false).Count == 0 && !blockDef.IsFromExternalReference && !blockDef.IsLayout)
                    //{
                    //    blockDef.Erase();
                    //}
                    //br.Explode(explodedObjs);
                    // br.Erase();
                    //ObjectIdCollection loopIds = new ObjectIdCollection();
                    //string patternName = "ANSI31";
                    //int colorIndex = 5;
                    //if(explodedObjs.Count == 0)
                    //{
                    //    ed.WriteMessage("\nKhông có đối tượng nào để nổ block.");
                    //    return ObjectId.Null;
                    //}
                    //foreach (DBObject obj in explodedObjs)
                    //{
                    //    if (obj is Entity ent)
                    //    {
                    //        modelSpace.AppendEntity(ent);
                    //        tr.AddNewlyCreatedDBObject(ent, true);

                    //        if (ent is Polyline pl && pl.Closed && (pl.ColorIndex == 5 || pl.ColorIndex == 6))
                    //        {
                    //            if (pl.ColorIndex == 5)
                    //            {
                    //                patternName = "ANSI31";
                    //                colorIndex = 5;
                    //            }
                    //            else if (pl.ColorIndex == 6)
                    //            {
                    //                patternName = "ANSI37";
                    //                colorIndex = 6;
                    //            }
                    //            loopIds.Add(ent.ObjectId);
                    //        }
                    //    }
                    //}

                    //if (loopIds.Count > 0)
                    //{
                    //    hatch = new Hatch();
                    //    hatch.SetDatabaseDefaults();
                    //    hatch.ColorIndex = colorIndex;
                    //    hatch.PatternScale = 500;
                    //    hatch.SetHatchPattern(HatchPatternType.PreDefined, patternName);
                    //    hatch.Associative = true;

                    //    modelSpace.AppendEntity(hatch);
                    //    tr.AddNewlyCreatedDBObject(hatch, true);

                    //    hatch.AppendLoop(HatchLoopTypes.Default, loopIds);
                    //    hatch.EvaluateHatch(true);
                    //}
                    tr.Commit();
                }
                


                return dimstyle1;
                // ForceUpdateBlockRecord(br.ObjectId);
            }
        }
        //using (DocumentLock m_DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument())
        //{
        //    using (Transaction tr = db.TransactionManager.StartTransaction())
        //{
        //    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
        //    ObjectId btrId = bt.Has(blockName) ? bt[blockName] : ImportBlock(db, blockName, blockPath);

        //    if (btrId == ObjectId.Null)
        //    {
        //        ed.WriteMessage($"\nKhông tìm thấy block '{blockName}' trong bản vẽ hoặc file nguồn.");
        //        return;
        //    }

        //    BlockTableRecord modelSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
        //    BlockReference br = new BlockReference(insertPoint, btrId)
        //    {
        //        ScaleFactors = new Scale3d(scale)
        //    };
        //    modelSpace.AppendEntity(br);
        //    tr.AddNewlyCreatedDBObject(br, true);

        //    // Chỉ xử lý các thuộc tính động (Dynamic Properties)
        //    if (br.IsDynamicBlock)
        //    {
        //        foreach (DynamicBlockReferenceProperty prop in br.DynamicBlockReferencePropertyCollection)
        //        {
        //            cbbtest.Items.Add(prop.PropertyName + "-" + prop.Value);
        //            if (dynamicProperties.TryGetValue(prop.PropertyName, out double val))
        //            {
        //                if (!prop.ReadOnly)
        //                {
        //                    try
        //                    {
        //                        prop.Value = val;
        //                        // br.RecordGraphicsModified(true);
        //                        // br.DynamicBlockTableRecord = br.DynamicBlockTableRecord; // force update
        //                        doc.Editor.Regen(); // hoặc acadDoc.SendCommand("BATTUPDATE\n");
        //                    }
        //                    catch
        //                    {
        //                        ed.WriteMessage($"\nKhông thể gán giá trị '{val}' cho thuộc tính động '{prop.PropertyName}'.");
        //                    }
        //                }
        //            }
        //        }
        //    }

        //    BlockTableRecord blockDef = (BlockTableRecord)tr.GetObject(br.DynamicBlockTableRecord, OpenMode.ForWrite);
        //    // Kiểm tra nếu đã có hatch, không tạo hatch mới
        //    bool hatchExists = false;

        //    // Tìm các entity có thể hatch được
        //    ObjectIdCollection loopIds = new ObjectIdCollection();
        //    foreach (ObjectId entId in blockDef)
        //    {
        //        Entity ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
        //        if (ent is Hatch hatchEntity && hatchEntity.ObjectId != ObjectId.Null)
        //        {
        //            // Nếu có hatch tồn tại, không cần tạo hatch mới
        //            hatchExists = true;
        //            break;
        //        }
        //        else if (ent is Polyline pl && pl.Closed && ent.ColorIndex == 5)
        //        {
        //            loopIds.Add(entId);
        //        }
        //    }

        //    if (!hatchExists)
        //    {
        //        // Nếu chưa có hatch, tạo hatch mới
        //        Hatch hatch = new Hatch();
        //        hatch.ColorIndex = 5; // Màu xanh
        //        hatch.SetDatabaseDefaults();

        //        // 🟢 CHỈ GỌI SetHatchPattern SAU KHI hatch được thêm vào DB
        //        blockDef.AppendEntity(hatch);
        //        tr.AddNewlyCreatedDBObject(hatch, true);

        //        // Gọi SetHatchPattern sau khi hatch đã vào database
        //        hatch.PatternScale = 500; // Tỉ lệ hatch
        //        hatch.SetHatchPattern(HatchPatternType.PreDefined, "ANSI31");

        //        hatch.Associative = true;

        //        // Nếu có ít nhất 1 vòng kín thì tạo hatch
        //        if (loopIds.Count > 0)
        //        {
        //            hatch.AppendLoop(HatchLoopTypes.Default, loopIds);
        //            hatch.EvaluateHatch(true);
        //            doc.Editor.Regen();
        //        }

        //        br.RecordGraphicsModified(true);
        //    }

        //    tr.Commit();
        //    ForceUpdateBlockRecord(br.ObjectId);
        //    }
        //}
    



    
        private static bool IsBlue(Color color)
        {
            if (color.IsByColor)
                return color.ColorValue == System.Drawing.Color.Blue;
            else if (color.ColorMethod == ColorMethod.ByAci)
                return color.ColorIndex == 5;
            return false;
        }
        private static ObjectId ImportBlock(Database destDb, string blockName, string sourceFileName)
        {
            if (System.IO.File.Exists(sourceFileName))
            {
                using (var sourceDb = new Database(false, true))
                {
                    try
                    {
                        // Read the DWG into a side database
                        sourceDb.ReadDwgFile(sourceFileName, FileOpenMode.OpenForReadAndAllShare, true, "");

                        // Create a variable to store the block identifier
                        var id = ObjectId.Null;
                        using (var tr = new OpenCloseTransaction())
                        {
                            // Open the block table
                            var bt = (BlockTable)tr.GetObject(sourceDb.BlockTableId, OpenMode.ForRead, false);

                            // if the block table contains 'blockName', store it into the variable
                            if (bt.Has(blockName))
                                id = bt[blockName];
                        }
                        // if the variable is not null (i.e. the block was found)
                        if (!id.IsNull)
                        {
                            // Copy the block deinition from source to destination database
                            var blockIds = new ObjectIdCollection();
                            blockIds.Add(id);
                            var mapping = new IdMapping();
                            sourceDb.WblockCloneObjects(blockIds, destDb.BlockTableId, mapping, DuplicateRecordCloning.Replace, false);
                            // if the copy succeeded, return the ObjectId of the clone
                            if (mapping[id].IsCloned)
                                return mapping[id].Value;
                        }
                    }
                    catch (Autodesk.AutoCAD.Runtime.Exception ex)
                    {
                        var ed = acadApp.DocumentManager.MdiActiveDocument.Editor;
                        ed.WriteMessage("\nError during copy: " + ex.Message + "\n" + ex.StackTrace);
                    }
                }
            }
            return ObjectId.Null;
        }
        public static void ForceUpdateBlockRecord(ObjectId blockReferenceId)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            using (DocumentLock docLock = doc.LockDocument())
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockReference br = tr.GetObject(blockReferenceId, OpenMode.ForRead) as BlockReference;
                if (br == null)
                {
                   // doc.Editor.WriteMessage("\nKhông thể đọc BlockReference.");
                    return;
                }

                // Mở BlockTableRecord gốc (không phải instance trong ModelSpace)
                BlockTableRecord btr = tr.GetObject(br.DynamicBlockTableRecord, OpenMode.ForWrite) as BlockTableRecord;
                if (btr == null)
                {
                   // doc.Editor.WriteMessage("\nKhông thể đọc BlockTableRecord.");
                    return;
                }

                // Thêm một dòng dummy rồi xoá nó - ép AutoCAD đánh dấu thay đổi
                Line dummy = new Line(Point3d.Origin, new Point3d(1, 1, 0));
                btr.AppendEntity(dummy);
                tr.AddNewlyCreatedDBObject(dummy, true);
                dummy.Erase();

                tr.Commit();
            }

            // Ép Regen toàn bộ
            doc.Editor.Regen();
        }

    }


}
