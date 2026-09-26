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
using Autodesk.AutoCAD.EditorInput;
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
        Dictionary<string, double> dynamicProperties/*, System.Windows.Forms.ComboBox cbbtest*//*, double th*/, ObjectId dimstyle1, double scale = 1.0)
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
                   // hatch = new Hatch();
                    //string patternName = "ANSI31";
                    //     int colorIndex = 5;
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
                    //  Dimension abc= null;
                    //  abc.DimensionStyle = dimstyle1;
                    DimStyleTableRecord dimRec = tr.GetObject(dimstyle1, OpenMode.ForRead) as DimStyleTableRecord;
                    
                    foreach (ObjectId entId in blockDef)
                    {
                        
                        Entity ent = tr.GetObject(entId, OpenMode.ForWrite) as Entity;
                        if (ent is Dimension dimension)
                        {
                            dimension.DimensionStyle = dimstyle1;
                           // abc = dimension;
                        }
                        else if (ent is DBText dBText)
                        {
                            dBText.Height = dimRec.Dimtxt*dimRec.Dimscale;
                        }
                        else if (ent is MText mText)
                        {
                            mText.Height = dimRec.Dimtxt * dimRec.Dimscale;
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
                   
                    //------------------------------------------------------------
                    doc.Editor.Regen();
                    br.ExplodeToOwnerSpace(); // X đối tượng chèn // test bật lên khi x
                                              // hatch.UpgradeOpen();
                    if (!blockDef.IsErased && blockDef.GetBlockReferenceIds(true, false).Count == 0 && !blockDef.IsFromExternalReference && !blockDef.IsLayout)
                    {
                        blockDef.Erase();
                    }
                    //// hatch.EvaluateHatch(true); // Cập nhật đường biên
                    br.Erase();// test bật lên khi x
                    
                    tr.Commit();
                }
                


                return dimstyle1;
                // ForceUpdateBlockRecord(br.ObjectId);
            }
        }
       
        //--------------------------------------------------
        public static ObjectId Insertbelt(string blockPath, string blockName,List< Point3d> insertPoint,double cout,
            Dictionary<string, double> dynamicProperties/*, System.Windows.Forms.ComboBox cbbtest*//*, double th*/, ObjectId dimstyle1, double scale = 1.0)
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
                    // hatch = new Hatch();
                    //string patternName = "ANSI31";
                    //     int colorIndex = 5;
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    ObjectId btrId = bt.Has(blockName) ? bt[blockName] : ImportBlock(db, blockName, blockPath);

                    if (btrId == ObjectId.Null)
                    {
                        ed.WriteMessage($"\nKhông tìm thấy block '{blockName}' trong bản vẽ hoặc file nguồn.");
                        return ObjectId.Null;
                    }
                    BlockTableRecord modelSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                  //  MessageBox.Show(cout.ToString());
                    for (int i = 0; i < cout; i++)
                    {
                        //MessageBox.Show(i.ToString());
                       
                        br = new BlockReference(insertPoint[i], btrId)
                        {
                            ScaleFactors = new Scale3d(scale)
                        };
                        modelSpace.AppendEntity(br);
                        tr.AddNewlyCreatedDBObject(br, true);
                        BlockTableRecord blockDef = (BlockTableRecord)tr.GetObject(br.DynamicBlockTableRecord, OpenMode.ForWrite);

                        DimStyleTableRecord dimRec = tr.GetObject(dimstyle1, OpenMode.ForRead) as DimStyleTableRecord;

                        foreach (ObjectId entId in blockDef)
                        {

                            Entity ent = tr.GetObject(entId, OpenMode.ForWrite) as Entity;
                            if (ent is Dimension dimension)
                            {
                                dimension.DimensionStyle = dimstyle1;
                                // abc = dimension;
                            }
                            else if (ent is DBText dBText)
                            {
                                dBText.Height = dimRec.Dimtxt * dimRec.Dimscale;
                            }
                            else if (ent is MText mText)
                            {
                                mText.Height = dimRec.Dimtxt * dimRec.Dimscale;
                            }

                        }
                        //****

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

                        doc.Editor.Regen();
                        br.ExplodeToOwnerSpace(); // X đối tượng chèn // test bật lên khi x
                        br.Erase();                         // hatch.UpgradeOpen();

                        //***
                  /*      if (!blockDef.IsErased && blockDef.GetBlockReferenceIds(true, false).Count == 0 && !blockDef.IsFromExternalReference && !blockDef.IsLayout)
                        {
                            blockDef.Erase();
                        }*/
                        //// hatch.EvaluateHatch(true); // Cập nhật đường biên
                   //     br.Erase();// test bật lên khi x
                    }
                    tr.Commit();
                }



                return dimstyle1;
                // ForceUpdateBlockRecord(br.ObjectId);
            }
        }


        public static void InsertBlockMleader(string blockPath, string blockName, string newText, Point3d pointResult, double scale)

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

                    cSpace.AppendEntity(br);
                    tr.AddNewlyCreatedDBObject(br, true);

                    // add attribute references to the block reference
                    var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                    var attInfos = new Dictionary<string, TextInfo>();
                    foreach (ObjectId id in btr)
                    {
                        if (id.ObjectClass.DxfName == "MULTILEADER")
                        {
                          //  var mLeader = (MLeader)tr.GetObject(id, OpenMode.ForRead);
                            var mLeader = (MLeader)tr.GetObject(id, OpenMode.ForWrite);
                            // int mleaderindex = mLeader.
                            // Lấy ArrowSize hiện tại
                            //  double currentArrowSize = mLeader.ArrowSize;

                            // Nhân theo scale
                          //  //   mLeader.ArrowSize = mLeader.ArrowSize * scale;
                           // MessageBox.Show(mLeader.GetArrowSize(0).ToString());
                            mLeader.SetArrowSize(0,75*scale);
                          //  mLeader.SetOverrideProperty(MLeaderProperty.ArrowSize, true);
                            // Lấy MText từ MLeader
                            if (mLeader.ContentType == ContentType.MTextContent)
                            {
                                
                                MText mtext = mLeader.MText;
                                if (mtext != null)
                                {
                                    // Clone MLeader để modify
                                    mLeader.UpgradeOpen();

                                    // Thay đổi text của MText
                                    //mtext.Contents = newText;
                                    //mLeader.MText = mtext;

                                    ed.WriteMessage($"\nĐã cập nhật text của MLeader thành: {newText}");
                                }
                            }
                        }
                    }
                    doc.Editor.Regen();
                    br.ExplodeToOwnerSpace(); // X đối tượng chèn // test bật lên khi x
                                              // hatch.UpgradeOpen();

                    // hatch.EvaluateHatch(true); // Cập nhật đường biên
                    br.Erase();// test bật lên khi x
                    tr.Commit();
                }
            }
        }
        //-------------------------------------------
        public static void InsertBlockMleader5(string blockPath, string blockName, string newText, Point3d pointResult, double scale)

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

                    cSpace.AppendEntity(br);
                    tr.AddNewlyCreatedDBObject(br, true);

                    // add attribute references to the block reference
                    var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                    var attInfos = new Dictionary<string, TextInfo>();
                    foreach (ObjectId id in btr)
                    {
                        if (id.ObjectClass.DxfName == "MULTILEADER")
                        {
                            var mLeader = (MLeader)tr.GetObject(id, OpenMode.ForRead);

                            // Lấy MText từ MLeader
                            if (mLeader.ContentType == ContentType.MTextContent)
                            {
                                MText mtext = mLeader.MText;
                                if (mtext != null)
                                {
                                    // Clone MLeader để modify
                                    mLeader.UpgradeOpen();

                                    // Thay đổi text của MText
                                    mtext.Contents = newText;
                                    mLeader.MText = mtext;

                                    ed.WriteMessage($"\nĐã cập nhật text của MLeader thành: {newText}");
                                }
                            }
                        }
                    }
                    doc.Editor.Regen();
                    br.ExplodeToOwnerSpace(); // X đối tượng chèn // test bật lên khi x
                                              // hatch.UpgradeOpen();

                    // hatch.EvaluateHatch(true); // Cập nhật đường biên
                    br.Erase();// test bật lên khi x
                    tr.Commit();
                }
            }
        }
        public static void InsertBlockMleader1(
     string blockPath,
     string blockName,
     string newText,
     Point3d insertPoint,
     double scale,
     Point3d newSecondPoint 
 )
        {
            var doc = acadApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            using (var docLock = doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                ObjectId btrId = bt.Has(blockName)
                    ? bt[blockName]
                    : ImportBlock(db, blockName, blockPath);

                if (btrId.IsNull)
                {
                    ed.WriteMessage("\nBlock không tồn tại hoặc không import được.");
                    return;
                }

                var cSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                var br = new BlockReference(insertPoint, btrId);
                br.ScaleFactors = new Scale3d(scale);

                cSpace.AppendEntity(br);
                tr.AddNewlyCreatedDBObject(br, true);

                var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);

                foreach (ObjectId id in btr)
                {
                    if (id.ObjectClass.DxfName != "MULTILEADER")
                        continue;

                    var mLeader = (MLeader)tr.GetObject(id, OpenMode.ForWrite);

                    // 1. Thay text
                    if (mLeader.ContentType == ContentType.MTextContent)
                    {
                        MText mt = mLeader.MText;
                        if (mt != null)
                        {
                            mt.Contents = newText;
                            mLeader.MText = mt;
                        }
                    }

                    // 2. Thay vertex thứ 2
                    if (mLeader.LeaderCount > 0)
                    {
                        int leaderIndex = 0;  // đa số MLeader chỉ có
                    }
                }
            }
        }




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
