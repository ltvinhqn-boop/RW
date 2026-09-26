using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using acap= Autodesk.AutoCAD.ApplicationServices.Application;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.PlottingServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Interop.Common;
using Autodesk.AutoCAD.EditorInput;

namespace DWG2PDF
{
    public class Printset
    {
        public static PreviewEndPlotStatus  CreateOrEditPageSetup( string Plottername, string papersize, string plotstyle, Extents2d plotarea, bool cb1, double scale, bool tofile, string path, bool ori)
        {
            PreviewEndPlotStatus result = 0;
        
            Document acDoc = acap.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            using (DocumentLock m_DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument())
           {
                short num = Convert.ToInt16(acap.GetSystemVariable("BACKGROUNDPLOT"));
                acap.SetSystemVariable("BACKGROUNDPLOT", 0);
                short num1 = Convert.ToInt16(acap.GetSystemVariable("INSUNITS"));
                acap.SetSystemVariable("INSUNITS", 4);
                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {

                    DBDictionary plSets = acTrans.GetObject(acCurDb.PlotSettingsDictionaryId,
                                                            OpenMode.ForRead) as DBDictionary;
                    DBDictionary vStyles = acTrans.GetObject(acCurDb.VisualStyleDictionaryId,
                                                             OpenMode.ForRead) as DBDictionary;

                    PlotSettings acPlSet = default(PlotSettings);
                    bool createNew = false;

                    // Reference the Layout Manager
                    LayoutManager acLayoutMgr = LayoutManager.Current;

                    // Get the current layout and output its name in the Command Line window
                    Layout acLayout = acTrans.GetObject(acLayoutMgr.GetLayoutId(acLayoutMgr.CurrentLayout),
                                                        OpenMode.ForRead) as Layout;
                    using ( PlotInfo acPlInfo = new PlotInfo())
                    {
                        acPlInfo.Layout = acLayout.ObjectId;
                        PlotSettingsValidator acPlSetVdr;
               if (acLayout.ModelType == true)
                  {
                            if (plSets.Contains("Vteams") == false)
                            {
                                createNew = true;

                                // Create a new PlotSettings object: 
                                //    True - model space, False - named layout
                                acPlSet = new PlotSettings(acLayout.ModelType);
                                acPlSet.CopyFrom(acLayout);

                                acPlSet.PlotSettingsName = "Vteams";
                                acPlSet.AddToPlotSettingsDictionary(acCurDb);
                                acTrans.AddNewlyCreatedDBObject(acPlSet, true);
                            }
                            else
                            {
                                acPlSet = plSets.GetAt("Vteams").GetObject(OpenMode.ForWrite) as PlotSettings;
                            }

                            // Update the PlotSettings object
                            // try
                            //{
                             acPlSetVdr = PlotSettingsValidator.Current;

                            // Set the Plotter and page size
                            acPlSetVdr.SetPlotConfigurationName(acPlSet, Plottername, papersize);

                            // Set to plot to the current display
                           
                                acPlSetVdr.SetPlotWindowArea(acPlSet, plotarea);
                                acPlSetVdr.SetPlotType(acPlSet, Autodesk.AutoCAD.DatabaseServices.PlotType.Window);
                          
                        



                            acPlSetVdr.SetPlotOrigin(acPlSet, new Point2d(0, 0));
                            acPlSetVdr.SetPlotCentered(acPlSet, true);
                            if (cb1)
                            {
                                acPlSetVdr.SetUseStandardScale(acPlSet, true);
                                acPlSetVdr.SetStdScaleType(acPlSet, StdScaleType.ScaleToFit);
                                acPlSetVdr.SetPlotPaperUnits(acPlSet, PlotPaperUnit.Millimeters);
                            }
                            else
                            {
                                acPlSetVdr.SetUseStandardScale(acPlSet, false);
                               CustomScale  customScale = new CustomScale(1.0, scale);
                                //  cbb1.Text = scalefit(cbb1)[0] + " - " + scalefit(cbb1)[1];
                                acPlSetVdr.SetCustomPrintScale(acPlSet, customScale);
                                acPlSetVdr.SetPlotPaperUnits(acPlSet, PlotPaperUnit.Millimeters);

                            }

                           // acPlSet.ScaleLineweights = true;

                            // Specify if plot styles should be displayed on the layout
                            acPlSet.ShowPlotStyles = true;

                            acPlSet.PrintLineweights = true;
                            acPlSet.PlotTransparency = false;
                            acPlSet.PlotPlotStyles = true;
                            acPlSet.DrawViewportsFirst = true;
                            //acPlSet.
                 }
                        // Check to see if the page setup exists
                       else 
                           {
                            if (plSets.Contains("VteamsLayout") == false)
                            {
                                createNew = true;

                                // Create a new PlotSettings object: 
                                //    True - model space, False - named layout
                                acPlSet = new PlotSettings(acLayout.ModelType);
                                acPlSet.CopyFrom(acLayout);
                                
                                acPlSet.PlotSettingsName = "VteamsLayout";
                                acPlSet.AddToPlotSettingsDictionary(acCurDb);
                                acTrans.AddNewlyCreatedDBObject(acPlSet, true);
                            }
                            else
                            {
                                acPlSet = plSets.GetAt("VteamsLayout").GetObject(OpenMode.ForWrite) as PlotSettings;
                            }

                            // Update the PlotSettings object
                            // try
                            //{
                             acPlSetVdr = PlotSettingsValidator.Current;

                            // Set the Plotter and page size
                            acPlSetVdr.SetPlotConfigurationName(acPlSet, Plottername, papersize);

                            // Set to plot to the current display
                          
                                acPlSetVdr.SetPlotWindowArea(acPlSet, plotarea);
                                acPlSetVdr.SetPlotType(acPlSet, Autodesk.AutoCAD.DatabaseServices.PlotType.Window);
                       

                            



                            acPlSetVdr.SetPlotOrigin(acPlSet, new Point2d(0, 0));
                            acPlSetVdr.SetPlotCentered(acPlSet, true);
                            if (cb1)
                            {
                                // MessageBox.Show("đã lọt vào");
                                // Set the plot scale
                                acPlSetVdr.SetUseStandardScale(acPlSet, true);
                                acPlSetVdr.SetStdScaleType(acPlSet, StdScaleType.ScaleToFit);
                                acPlSetVdr.SetPlotPaperUnits(acPlSet, PlotPaperUnit.Millimeters);
                            }
                            else
                            {
                                acPlSetVdr.SetUseStandardScale(acPlSet, false);
                                CustomScale customScale = new CustomScale(1, scale);
                                //  cbb1.Text = scalefit(cbb1)[0] + " - " + scalefit(cbb1)[1];
                                acPlSetVdr.SetCustomPrintScale(acPlSet, customScale);
                                acPlSetVdr.SetPlotPaperUnits(acPlSet, PlotPaperUnit.Millimeters);

                            }

                            acPlSet.ScaleLineweights = true;

                            // Specify if plot styles should be displayed on the layout
                            acPlSet.ShowPlotStyles = true;

                            acPlSet.PrintLineweights = true;
                            acPlSet.PlotTransparency = false;
                            acPlSet.PlotPlotStyles = true;
                            acPlSet.DrawViewportsFirst = true;
                            
                        }
             
                        double mediaPageWidth = acPlSet.PlotPaperSize.X;
                        double mediaPageHeight = acPlSet.PlotPaperSize.Y;
                      //  cbb1.Text = mediaPageWidth + " - " + mediaPageHeight;
                        bool ori2 = mediaPageWidth >= mediaPageHeight;
                        if (ori)
                        {
                            if(ori2)
                            {
                                acPlSetVdr.SetPlotRotation(acPlSet, PlotRotation.Degrees000);
                            }
                            else
                            {
                                acPlSetVdr.SetPlotRotation(acPlSet, PlotRotation.Degrees270);
                            }

                        }
                        else
                        {
                            if (ori2)
                            {
                                acPlSetVdr.SetPlotRotation(acPlSet, PlotRotation.Degrees090);
                            }
                            else
                            {
                                acPlSetVdr.SetPlotRotation(acPlSet, PlotRotation.Degrees000);
                            }
                        }
                        acPlSetVdr.SetCurrentStyleSheet(acPlSet, plotstyle);
                    
                        acPlSetVdr.SetZoomToPaperOnUpdate(acPlSet, true);


                        

                        acTrans.Commit();
                       
                        acPlInfo.OverrideSettings = acPlSet;
                      
                          
                        
                        using (PlotInfoValidator acPlInfoVdr = new PlotInfoValidator())
                        {
                            acPlInfoVdr.MediaMatchingPolicy = MatchingPolicy.MatchEnabled;
                            acPlInfoVdr.Validate(acPlInfo);
                           
                            // Check to see if a plot is already in progress
                            if (PlotFactory.ProcessPlotState == ProcessPlotState.NotPlotting)
                            {
                                using (PlotEngine acPlEng = PlotFactory.CreatePublishEngine())
                                {
                                    // Track the plot progress with a Progress dialog
                                    using (PlotProgressDialog acPlProgDlg = new PlotProgressDialog(false, 1, true))
                                    {
                                        using ((acPlProgDlg))
                                        {
                                            // Define the status messages to display 
                                            // when plotting starts
                                            acPlProgDlg.set_PlotMsgString(PlotMessageIndex.DialogTitle, "Plot Progress");
                                            acPlProgDlg.set_PlotMsgString(PlotMessageIndex.CancelJobButtonMessage, "Cancel Job");
                                            acPlProgDlg.set_PlotMsgString(PlotMessageIndex.CancelSheetButtonMessage, "Cancel Sheet");
                                            acPlProgDlg.set_PlotMsgString(PlotMessageIndex.SheetSetProgressCaption, "Sheet Set Progress");
                                            acPlProgDlg.set_PlotMsgString(PlotMessageIndex.SheetProgressCaption, "Sheet Progress");
                                        
                                            // Set the plot progress range
                                            acPlProgDlg.LowerPlotProgressRange = 0;
                                            acPlProgDlg.UpperPlotProgressRange = 100;
                                            acPlProgDlg.PlotProgressPos = 0;

                                            // Display the Progress dialog
                                            acPlProgDlg.OnBeginPlot();
                                            acPlProgDlg.IsVisible = true;

                                            // Start to plot the layout
                                            acPlEng.BeginPlot(acPlProgDlg, null);

                                            // Define the plot output
                                            acPlEng.BeginDocument(acPlInfo, acDoc.Name, null, 1, tofile, path);

                                            // Display information about the current plot
                                            acPlProgDlg.set_PlotMsgString(PlotMessageIndex.Status, "Plotting: " + acDoc.Name + " - " + acLayout.LayoutName);

                                            // Set the sheet progress range
                                            acPlProgDlg.OnBeginSheet();
                                            acPlProgDlg.LowerSheetProgressRange = 0;
                                            acPlProgDlg.UpperSheetProgressRange = 100;
                                            acPlProgDlg.SheetProgressPos = 0;

                                            // Plot the first sheet/layout
                                            using (PlotPageInfo acPlPageInfo = new PlotPageInfo())
                                            {
                                                acPlEng.BeginPage(acPlPageInfo, acPlInfo, true, null);
                                            }

                                            acPlEng.BeginGenerateGraphics(null);
                                            acPlEng.EndGenerateGraphics(null);

                                            // Finish plotting the sheet/layout
                                            PreviewEndPlotInfo previewEndPlotInfo = new PreviewEndPlotInfo();
                                            acPlEng.EndPage(previewEndPlotInfo);
                                            result = previewEndPlotInfo.Status;
                                           // acPlEng.EndPage(null);
                                            acPlProgDlg.SheetProgressPos = 100;
                                            acPlProgDlg.OnEndSheet();

                                            // Finish plotting the document
                                            acPlEng.EndDocument(null);

                                            // Finish the plot
                                            acPlProgDlg.PlotProgressPos = 100;
                                            acPlProgDlg.OnEndPlot();
                                            acPlEng.EndPlot(null);
                                        }
                                    }
                                }


                            }

                        }
                        //   result.acLayout = acLayout;
                        //   result.acPlInfo = acPlInfo;
                        //  return result;
                        
                        if (createNew == true)
                        {
                            acPlSet.Dispose();
                        }
                       // acPlSetVdr.RefreshLists(acPlSet);// co the xoa
                    }

                    acap.SetSystemVariable("BACKGROUNDPLOT", num);
                    acap.SetSystemVariable("INSUNITS", num1);
                  
                }
            }
            return result;
        }

        //---------------------------------------------------------------------------------------------------------------------------------------------------------------------

        public static PreviewEndPlotStatus preview(string Plottername, string papersize, string plotstyle, Extents2d plotarea, bool cb1, double cbb1, bool tofile, string path, bool ori,int flagview)
        {
            PreviewEndPlotStatus result = 0;
            Document acDoc = acap.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            using (DocumentLock m_DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument())
            {
                short num = Convert.ToInt16(acap.GetSystemVariable("BACKGROUNDPLOT"));
                acap.SetSystemVariable("BACKGROUNDPLOT", 0);
                short num1 = Convert.ToInt16(acap.GetSystemVariable("INSUNITS"));
                acap.SetSystemVariable("INSUNITS", 4);
                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {

                    DBDictionary plSets = acTrans.GetObject(acCurDb.PlotSettingsDictionaryId,
                                                            OpenMode.ForRead) as DBDictionary;
                    DBDictionary vStyles = acTrans.GetObject(acCurDb.VisualStyleDictionaryId,
                                                             OpenMode.ForRead) as DBDictionary;

                    PlotSettings acPlSet = default(PlotSettings);
                   
                    bool createNew = false;

                    // Reference the Layout Manager
                    LayoutManager acLayoutMgr = LayoutManager.Current;

                    // Get the current layout and output its name in the Command Line window
                    Layout acLayout = acTrans.GetObject(acLayoutMgr.GetLayoutId(acLayoutMgr.CurrentLayout),
                                                        OpenMode.ForRead) as Layout;
                    using (PlotInfo acPlInfo = new PlotInfo())
                    {
                        acPlInfo.Layout = acLayout.ObjectId;

                        PlotSettingsValidator acPlSetVdr;
                        if (acLayout.ModelType == true)
                        {
                            if (plSets.Contains("Vteams") == false)
                            {
                                createNew = true;

                                // Create a new PlotSettings object: 
                                //    True - model space, False - named layout
                                acPlSet = new PlotSettings(acLayout.ModelType);
                                acPlSet.CopyFrom(acLayout);

                                acPlSet.PlotSettingsName = "Vteams";
                                acPlSet.AddToPlotSettingsDictionary(acCurDb);
                                acTrans.AddNewlyCreatedDBObject(acPlSet, true);
                            }
                            else
                            {
                                acPlSet = plSets.GetAt("Vteams").GetObject(OpenMode.ForWrite) as PlotSettings;
                            }

                            // Update the PlotSettings object
                            // try
                            //{
                            acPlSetVdr = PlotSettingsValidator.Current;

                            // Set the Plotter and page size
                            acPlSetVdr.SetPlotConfigurationName(acPlSet, Plottername, papersize);

                            // Set to plot to the current display

                            acPlSetVdr.SetPlotWindowArea(acPlSet, plotarea);
                            acPlSetVdr.SetPlotType(acPlSet, Autodesk.AutoCAD.DatabaseServices.PlotType.Window);





                            acPlSetVdr.SetPlotOrigin(acPlSet, new Point2d(0, 0));
                            acPlSetVdr.SetPlotCentered(acPlSet, true);
                            if (cb1)
                            {
                                
                                acPlSetVdr.SetUseStandardScale(acPlSet, true);
                                acPlSetVdr.SetStdScaleType(acPlSet, StdScaleType.ScaleToFit);
                                acPlSetVdr.SetPlotPaperUnits(acPlSet, PlotPaperUnit.Millimeters);
                            }
                            else
                            {
                                acPlSetVdr.SetUseStandardScale(acPlSet, false);
                                CustomScale customScale = new CustomScale(1.0, cbb1);
                                //  cbb1.Text = scalefit(cbb1)[0] + " - " + scalefit(cbb1)[1];
                                acPlSetVdr.SetCustomPrintScale(acPlSet, customScale);
                                acPlSetVdr.SetPlotPaperUnits(acPlSet, PlotPaperUnit.Millimeters);

                            }

                            acPlSet.ScaleLineweights = true;

                            // Specify if plot styles should be displayed on the layout
                            acPlSet.ShowPlotStyles = true;

                            // Rebuild plotter, plot style, and canonical media lists 
                            // (must be called before setting the plot style)
                            // acPlSetVdr.RefreshLists(acPlSet);
                            // cbb1.Text = acPlSet.PlotPaperUnits.ToString();
                            // Specify the shaded viewport options
                            // acPlSet.ShadePlot = PlotSettingsShadePlotType.AsDisplayed;

                            //  acPlSet.ShadePlotResLevel = ShadePlotResLevel.Normal;

                            // Specify the plot options
                            acPlSet.PrintLineweights = true;
                            acPlSet.PlotTransparency = false;
                            acPlSet.PlotPlotStyles = true;
                            acPlSet.DrawViewportsFirst = true;
                        }
                        // Check to see if the page setup exists
                        else
                        {
                            if (plSets.Contains("VteamsLayout") == false)
                            {
                                createNew = true;

                                // Create a new PlotSettings object: 
                                //    True - model space, False - named layout
                                acPlSet = new PlotSettings(acLayout.ModelType);
                                acPlSet.CopyFrom(acLayout);

                                acPlSet.PlotSettingsName = "VteamsLayout";
                                acPlSet.AddToPlotSettingsDictionary(acCurDb);
                                acTrans.AddNewlyCreatedDBObject(acPlSet, true);
                            }
                            else
                            {
                                acPlSet = plSets.GetAt("VteamsLayout").GetObject(OpenMode.ForWrite) as PlotSettings;
                            }

                            // Update the PlotSettings object
                            // try
                            //{
                            acPlSetVdr = PlotSettingsValidator.Current;

                            // Set the Plotter and page size
                            acPlSetVdr.SetPlotConfigurationName(acPlSet, Plottername, papersize);

                            // Set to plot to the current display

                            acPlSetVdr.SetPlotWindowArea(acPlSet, plotarea);
                            acPlSetVdr.SetPlotType(acPlSet, Autodesk.AutoCAD.DatabaseServices.PlotType.Window);






                            acPlSetVdr.SetPlotOrigin(acPlSet, new Point2d(0, 0));
                            acPlSetVdr.SetPlotCentered(acPlSet, true);
                            if (cb1)
                            {
                                // MessageBox.Show("đã lọt vào");
                                // Set the plot scale
                                acPlSetVdr.SetUseStandardScale(acPlSet, true);
                                acPlSetVdr.SetStdScaleType(acPlSet, StdScaleType.ScaleToFit);
                                acPlSetVdr.SetPlotPaperUnits(acPlSet, PlotPaperUnit.Millimeters);
                            }
                            else
                            {
                                acPlSetVdr.SetUseStandardScale(acPlSet, false);
                                CustomScale customScale = new CustomScale(1.0,cbb1);
                                //  cbb1.Text = scalefit(cbb1)[0] + " - " + scalefit(cbb1)[1];
                                acPlSetVdr.SetCustomPrintScale(acPlSet, customScale);
                                acPlSetVdr.SetPlotPaperUnits(acPlSet, PlotPaperUnit.Millimeters);

                            }

                            acPlSet.ScaleLineweights = true;

                            // Specify if plot styles should be displayed on the layout
                            acPlSet.ShowPlotStyles = true;

                            acPlSet.PrintLineweights = true;
                            acPlSet.PlotTransparency = false;
                            acPlSet.PlotPlotStyles = true;
                            acPlSet.DrawViewportsFirst = true;
                           // acPlSet.
                        }

                        //  PlotRotation paperOrientation = acPlInfo.GetPlotPaperOrientation();
                        //  double mediaPageWidth = acLayout.PlotPaperSize.X;
                        double mediaPageWidth = acPlSet.PlotPaperSize.X;
                        double mediaPageHeight = acPlSet.PlotPaperSize.Y;
                        //  cbb1.Text = mediaPageWidth + " - " + mediaPageHeight;
                        bool ori2 = mediaPageWidth >= mediaPageHeight;
                        if (ori)
                        {
                            if (ori2)
                            {
                                acPlSetVdr.SetPlotRotation(acPlSet, PlotRotation.Degrees000);
                            }
                            else
                            {
                                acPlSetVdr.SetPlotRotation(acPlSet, PlotRotation.Degrees270);
                            }

                        }
                        else
                        {
                            if (ori2)
                            {
                                acPlSetVdr.SetPlotRotation(acPlSet, PlotRotation.Degrees090);
                            }
                            else
                            {
                                acPlSetVdr.SetPlotRotation(acPlSet, PlotRotation.Degrees000);
                            }
                        }
                        acPlSetVdr.SetCurrentStyleSheet(acPlSet, plotstyle);

                        acPlSetVdr.SetZoomToPaperOnUpdate(acPlSet, true);

                        


                        acTrans.Commit();

                        acPlInfo.OverrideSettings = acPlSet;



                        using (PlotInfoValidator acPlInfoVdr = new PlotInfoValidator())
                        {
                            acPlInfoVdr.MediaMatchingPolicy = MatchingPolicy.MatchEnabled;
                            acPlInfoVdr.Validate(acPlInfo);

                            // Check to see if a plot is already in progress
                            if (PlotFactory.ProcessPlotState == ProcessPlotState.NotPlotting)
                            {

                                PlotEngine plotEngine = null;
                                plotEngine = PlotFactory.CreatePreviewEngine(flagview);
                                using (plotEngine)
                                {
                                    using (PlotProgressDialog plotProgressDialog = new PlotProgressDialog(true, 1, true))
                                    {
                                        using (plotProgressDialog)
                                        {
                                          
                                            plotProgressDialog.set_PlotMsgString(PlotMessageIndex.CancelJobButtonMessage, "Plot Progress");
                                            plotProgressDialog.set_PlotMsgString(PlotMessageIndex.SheetSetProgressCaption, "Cancel Job");
                                            plotProgressDialog.set_PlotMsgString(PlotMessageIndex.SheetProgressCaption, "Cancel Sheet");
                                            plotProgressDialog.set_PlotMsgString(PlotMessageIndex.MessageCount, "Sheet Set Progress");
                                            plotProgressDialog.set_PlotMsgString(PlotMessageIndex.MessageCancelingCurrent, "Sheet Progress");
                                            plotProgressDialog.LowerPlotProgressRange = 0;
                                            plotProgressDialog.UpperPlotProgressRange = 100;
                                            plotProgressDialog.PlotProgressPos = 0;
                                            plotProgressDialog.OnBeginPlot();
                                            plotProgressDialog.IsVisible = true;
                                            //plotProgressDialog.LO
                                            plotEngine.BeginPlot(plotProgressDialog, null);
                                            plotEngine.BeginDocument(acPlInfo, acDoc.Name, null, 1, false, path);
                                            plotProgressDialog.set_PlotMsgString(PlotMessageIndex.MessageCanceling, "Plotting: " + acDoc.Name + " - " + acLayout.LayoutName);
                                         
                                            plotProgressDialog.OnBeginSheet();
                                            plotProgressDialog.LowerSheetProgressRange = 0;
                                            plotProgressDialog.UpperSheetProgressRange = 100;
                                            plotProgressDialog.SheetProgressPos = 0;
                                            using (PlotPageInfo plotPageInfo = new PlotPageInfo())
                                            {
                                                plotEngine.BeginPage(plotPageInfo, acPlInfo, true, null);

                                            }
                                            plotEngine.BeginGenerateGraphics(null);
                                            plotEngine.EndGenerateGraphics(null);
                                            PreviewEndPlotInfo previewEndPlotInfo = new PreviewEndPlotInfo();
                                            plotEngine.EndPage(previewEndPlotInfo);
                                            result = previewEndPlotInfo.Status;
                                            plotProgressDialog.SheetProgressPos = 100;
                                            plotProgressDialog.OnEndSheet();
                                            plotEngine.EndDocument(null);
                                            plotProgressDialog.PlotProgressPos = 100;
                                            plotProgressDialog.OnEndPlot();
                                            plotEngine.EndPlot(null);
                                        }
                                    }
                                }


                            }

                        }
                        //   result.acLayout = acLayout;
                        //   result.acPlInfo = acPlInfo;
                        //  return result;

                        if (createNew == true)
                        {
                            acPlSet.Dispose();
                        }
                     //   acPlSetVdr.RefreshLists(acPlSet);// co the xoa
                    }
                    acap.SetSystemVariable("BACKGROUNDPLOT", num);
                    acap.SetSystemVariable("INSUNITS", num1);

                }
            }
            return result;
        }
        //----------------------------------------------------------------------------------------------------------------------------------------------------
        public static PreviewEndPlotStatus CreateOrEditPageSetupmulti(Document newDocument, string Plottername, string papersize, string plotstyle, Extents2d plotarea, bool cb1, ComboBox cbb1, bool tofile, string path, bool ori)
        {
            

                // Document newDocument= acap.DocumentManager.MdiActiveDocument;
                PreviewEndPlotStatus result = 0;
       


                short num = Convert.ToInt16(acap.GetSystemVariable("BACKGROUNDPLOT"));
                acap.SetSystemVariable("BACKGROUNDPLOT", 0);
                short num1 = Convert.ToInt16(acap.GetSystemVariable("INSUNITS"));
                acap.SetSystemVariable("INSUNITS", 4);


                using (DocumentLock m_DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument())
                {
                    using (Transaction acTrans = newDocument.Database.TransactionManager.StartTransaction())
                    {
                    
                        DBDictionary plSets = acTrans.GetObject(newDocument.Database.PlotSettingsDictionaryId,
                                                                OpenMode.ForRead) as DBDictionary;
                        DBDictionary vStyles = acTrans.GetObject(newDocument.Database.VisualStyleDictionaryId,
                                                                 OpenMode.ForRead) as DBDictionary;

                        PlotSettings acPlSet = default(PlotSettings);
                        bool createNew = false;

                        // Reference the Layout Manager
                        LayoutManager acLayoutMgr = LayoutManager.Current;

                        // Get the current layout and output its name in the Command Line window
                        Layout acLayout = acTrans.GetObject(acLayoutMgr.GetLayoutId(acLayoutMgr.CurrentLayout),
                                                            OpenMode.ForRead) as Layout;
                        using (PlotInfo acPlInfo = new PlotInfo())
                        {
                            acPlInfo.Layout = acLayout.ObjectId;
                            PlotSettingsValidator acPlSetVdr;

                            if (acLayout.ModelType == true)
                            {
                                if (plSets.Contains("Vteams") == false)
                                {
                                    createNew = true;

                                    // Create a new PlotSettings object: 
                                    //    True - model space, False - named layout
                                    acPlSet = new PlotSettings(acLayout.ModelType);
                                    acPlSet.CopyFrom(acLayout);

                                    acPlSet.PlotSettingsName = "Vteams";
                                    acPlSet.AddToPlotSettingsDictionary(newDocument.Database);
                                    acTrans.AddNewlyCreatedDBObject(acPlSet, true);
                                }
                                else
                                {
                                    acPlSet = plSets.GetAt("Vteams").GetObject(OpenMode.ForWrite) as PlotSettings;
                                }

                                // Update the PlotSettings object
                                // try
                                //{
                                acPlSetVdr = PlotSettingsValidator.Current;

                                // Set the Plotter and page size
                                acPlSetVdr.SetPlotConfigurationName(acPlSet, Plottername, papersize);

                                // Set to plot to the current display

                                acPlSetVdr.SetPlotWindowArea(acPlSet, plotarea);
                                acPlSetVdr.SetPlotType(acPlSet, Autodesk.AutoCAD.DatabaseServices.PlotType.Window);





                                acPlSetVdr.SetPlotOrigin(acPlSet, new Point2d(0, 0));
                                acPlSetVdr.SetPlotCentered(acPlSet, true);
                                if (cb1)
                                {
                                    // MessageBox.Show("đã lọt vào");
                                    // Set the plot scale
                                    acPlSetVdr.SetUseStandardScale(acPlSet, true);
                                    acPlSetVdr.SetStdScaleType(acPlSet, StdScaleType.ScaleToFit);
                                    acPlSetVdr.SetPlotPaperUnits(acPlSet, PlotPaperUnit.Millimeters);
                                }
                                else
                                {
                                    acPlSetVdr.SetUseStandardScale(acPlSet, false);
                                    CustomScale customScale = new CustomScale(scalefit(cbb1)[0], scalefit(cbb1)[1]);
                                    //  cbb1.Text = scalefit(cbb1)[0] + " - " + scalefit(cbb1)[1];
                                    acPlSetVdr.SetCustomPrintScale(acPlSet, customScale);
                                    acPlSetVdr.SetPlotPaperUnits(acPlSet, PlotPaperUnit.Millimeters);

                                }

                                acPlSet.ScaleLineweights = true;

                                // Specify if plot styles should be displayed on the layout
                                acPlSet.ShowPlotStyles = true;

                                // Rebuild plotter, plot style, and canonical media lists 
                                // (must be called before setting the plot style)
                                // acPlSetVdr.RefreshLists(acPlSet);
                                // cbb1.Text = acPlSet.PlotPaperUnits.ToString();
                                // Specify the shaded viewport options
                                // acPlSet.ShadePlot = PlotSettingsShadePlotType.AsDisplayed;

                                //  acPlSet.ShadePlotResLevel = ShadePlotResLevel.Normal;

                                // Specify the plot options
                                acPlSet.PrintLineweights = true;
                                acPlSet.PlotTransparency = false;
                                acPlSet.PlotPlotStyles = true;
                                acPlSet.DrawViewportsFirst = true;
                            }
                            // Check to see if the page setup exists
                            else
                            {
                                if (plSets.Contains("VteamsLayout") == false)
                                {
                                    createNew = true;

                                    // Create a new PlotSettings object: 
                                    //    True - model space, False - named layout
                                    acPlSet = new PlotSettings(acLayout.ModelType);
                                    acPlSet.CopyFrom(acLayout);

                                    acPlSet.PlotSettingsName = "VteamsLayout";
                                    acPlSet.AddToPlotSettingsDictionary(newDocument.Database);
                                    acTrans.AddNewlyCreatedDBObject(acPlSet, true);
                                }
                                else
                                {
                                    acPlSet = plSets.GetAt("VteamsLayout").GetObject(OpenMode.ForWrite) as PlotSettings;
                                }

                                // Update the PlotSettings object
                                // try
                                //{
                                acPlSetVdr = PlotSettingsValidator.Current;

                                // Set the Plotter and page size
                                acPlSetVdr.SetPlotConfigurationName(acPlSet, Plottername, papersize);

                                // Set to plot to the current display

                                acPlSetVdr.SetPlotWindowArea(acPlSet, plotarea);
                                acPlSetVdr.SetPlotType(acPlSet, Autodesk.AutoCAD.DatabaseServices.PlotType.Window);






                                acPlSetVdr.SetPlotOrigin(acPlSet, new Point2d(0, 0));
                                acPlSetVdr.SetPlotCentered(acPlSet, true);
                                if (cb1)
                                {
                                    // MessageBox.Show("đã lọt vào");
                                    // Set the plot scale
                                    acPlSetVdr.SetUseStandardScale(acPlSet, true);
                                    acPlSetVdr.SetStdScaleType(acPlSet, StdScaleType.ScaleToFit);
                                    acPlSetVdr.SetPlotPaperUnits(acPlSet, PlotPaperUnit.Millimeters);
                                }
                                else
                                {
                                    acPlSetVdr.SetUseStandardScale(acPlSet, false);
                                    CustomScale customScale = new CustomScale(scalefit(cbb1)[0], scalefit(cbb1)[1]);
                                    //  cbb1.Text = scalefit(cbb1)[0] + " - " + scalefit(cbb1)[1];
                                    acPlSetVdr.SetCustomPrintScale(acPlSet, customScale);
                                    acPlSetVdr.SetPlotPaperUnits(acPlSet, PlotPaperUnit.Millimeters);

                                }

                                acPlSet.ScaleLineweights = true;

                                // Specify if plot styles should be displayed on the layout
                                acPlSet.ShowPlotStyles = true;

                                // Rebuild plotter, plot style, and canonical media lists 
                                // (must be called before setting the plot style)
                                // acPlSetVdr.RefreshLists(acPlSet);
                                // cbb1.Text = acPlSet.PlotPaperUnits.ToString();
                                // Specify the shaded viewport options
                                // acPlSet.ShadePlot = PlotSettingsShadePlotType.AsDisplayed;

                                //  acPlSet.ShadePlotResLevel = ShadePlotResLevel.Normal;

                                // Specify the plot options
                                acPlSet.PrintLineweights = true;
                                acPlSet.PlotTransparency = false;
                                acPlSet.PlotPlotStyles = true;
                                acPlSet.DrawViewportsFirst = true;

                            }
                            //tới đây
                            //  PlotRotation paperOrientation = acPlInfo.GetPlotPaperOrientation();
                            //  double mediaPageWidth = acLayout.PlotPaperSize.X;
                            double mediaPageWidth = acPlSet.PlotPaperSize.X;
                            double mediaPageHeight = acPlSet.PlotPaperSize.Y;
                            //  cbb1.Text = mediaPageWidth + " - " + mediaPageHeight;
                            bool ori2 = mediaPageWidth >= mediaPageHeight;
                            if (ori)
                            {
                                if (ori2)
                                {
                                    acPlSetVdr.SetPlotRotation(acPlSet, PlotRotation.Degrees000);
                                }
                                else
                                {
                                    acPlSetVdr.SetPlotRotation(acPlSet, PlotRotation.Degrees270);
                                }

                            }
                            else
                            {
                                if (ori2)
                                {
                                    acPlSetVdr.SetPlotRotation(acPlSet, PlotRotation.Degrees090);
                                }
                                else
                                {
                                    acPlSetVdr.SetPlotRotation(acPlSet, PlotRotation.Degrees000);
                                }
                            }
                            acPlSetVdr.SetCurrentStyleSheet(acPlSet, plotstyle);

                            acPlSetVdr.SetZoomToPaperOnUpdate(acPlSet, true);




                            acTrans.Commit();

                            acPlInfo.OverrideSettings = acPlSet;



                            using (PlotInfoValidator acPlInfoVdr = new PlotInfoValidator())
                            {

                                acPlInfoVdr.MediaMatchingPolicy = MatchingPolicy.MatchEnabled;
                                acPlInfoVdr.Validate(acPlInfo);

                                // Check to see if a plot is already in progress
                                if (PlotFactory.ProcessPlotState == ProcessPlotState.NotPlotting)
                                {
                                    using (PlotEngine acPlEng = PlotFactory.CreatePublishEngine())
                                    {
                                        // Track the plot progress with a Progress dialog
                                        using (PlotProgressDialog acPlProgDlg = new PlotProgressDialog(false, 1, true))
                                        {
                                            using ((acPlProgDlg))
                                            {
                                                // Define the status messages to display 
                                                // when plotting starts
                                                acPlProgDlg.set_PlotMsgString(PlotMessageIndex.DialogTitle, "Plot Progress");
                                                acPlProgDlg.set_PlotMsgString(PlotMessageIndex.CancelJobButtonMessage, "Cancel Job");
                                                acPlProgDlg.set_PlotMsgString(PlotMessageIndex.CancelSheetButtonMessage, "Cancel Sheet");
                                                acPlProgDlg.set_PlotMsgString(PlotMessageIndex.SheetSetProgressCaption, "Sheet Set Progress");
                                                acPlProgDlg.set_PlotMsgString(PlotMessageIndex.SheetProgressCaption, "Sheet Progress");

                                                // Set the plot progress range
                                                acPlProgDlg.LowerPlotProgressRange = 0;
                                                acPlProgDlg.UpperPlotProgressRange = 100;
                                                acPlProgDlg.PlotProgressPos = 0;

                                                // Display the Progress dialog
                                                acPlProgDlg.OnBeginPlot();
                                                acPlProgDlg.IsVisible = true;

                                                // Start to plot the layout
                                                acPlEng.BeginPlot(acPlProgDlg, null);

                                                // Define the plot output
                                                acPlEng.BeginDocument(acPlInfo, newDocument.Name, null, 1, tofile, path);

                                                // Display information about the current plot
                                                acPlProgDlg.set_PlotMsgString(PlotMessageIndex.Status, "Plotting: " + newDocument.Name + " - " + acLayout.LayoutName);

                                                // Set the sheet progress range
                                                acPlProgDlg.OnBeginSheet();
                                                acPlProgDlg.LowerSheetProgressRange = 0;
                                                acPlProgDlg.UpperSheetProgressRange = 100;
                                                acPlProgDlg.SheetProgressPos = 0;

                                                // Plot the first sheet/layout
                                                using (PlotPageInfo acPlPageInfo = new PlotPageInfo())
                                                {
                                                    acPlEng.BeginPage(acPlPageInfo, acPlInfo, true, null);
                                                }

                                                acPlEng.BeginGenerateGraphics(null);
                                                acPlEng.EndGenerateGraphics(null);

                                                // Finish plotting the sheet/layout
                                                PreviewEndPlotInfo previewEndPlotInfo = new PreviewEndPlotInfo();
                                                acPlEng.EndPage(previewEndPlotInfo);
                                                result = previewEndPlotInfo.Status;
                                                // acPlEng.EndPage(null);
                                                acPlProgDlg.SheetProgressPos = 100;
                                                acPlProgDlg.OnEndSheet();

                                                // Finish plotting the document
                                                acPlEng.EndDocument(null);

                                                // Finish the plot
                                                acPlProgDlg.PlotProgressPos = 100;
                                                acPlProgDlg.OnEndPlot();
                                                acPlEng.EndPlot(null);

                                            }
                                        }
                                    }


                                }

                            }
                       

                            if (createNew == true)
                            {
                                acPlSet.Dispose();
                            }
                            // acPlSetVdr.RefreshLists(acPlSet);// co the xoa
                        }

                        acap.SetSystemVariable("BACKGROUNDPLOT", num);
                        acap.SetSystemVariable("INSUNITS", num1);



                    }




                }
                 return result;
               // newDocument.CloseAndDiscard();
            
        }

        //---------------------------------------------------------------------------------------------------------------------------------------------------------------------

      
        //------------------------------------------------------------------------------------------------------------------------------------------------------------------
        public static List<double> scalefit(ComboBox cbb2)
        {
            List<double> scales = new List<double>();

            if (cbb2.Text.Contains(":"))
            {
                string[] parts = cbb2.Text.Split(':');

                // Chuyển đổi phần tử đầu tiên thành double
                double dividend = double.Parse(parts[0]);

                // Chuyển đổi phần tử thứ hai thành double
                double divisor = double.Parse(parts[1]);

                // Thực hiện phép chia
                scales.Add(dividend);
                scales.Add(divisor);
            }
            else
            {
                int dividend;
                int divisor;

                GetFraction(double.Parse(cbb2.Text), out dividend, out divisor);
                scales.Add(dividend);
                scales.Add(divisor);
            }
            return scales;
        }
        ///---------------------------------------------
        public static void GetFraction(double number, out int numerator, out int denominator)
        {
            // Chuyển đổi số double thành phân số tối giản
            double epsilon = 1e-6; // Độ chính xác epsilon để so sánh với số thực

            // Xác định tử số ban đầu và mẫu số ban đầu
            numerator = 1;
            denominator = 1;

            // Tìm ước số chung lớn nhất (GCD) của tử số và mẫu số
            double factor = 1.0;
            double error = Math.Abs(factor - number);

            for (int i = 1; i <= 1000; i++)
            {
                for (int j = 1; j <= 1000; j++)
                {
                    double currentFactor = (double)i / j;
                    double currentError = Math.Abs(currentFactor - number);

                    if (currentError < error)
                    {
                        factor = currentFactor;
                        error = currentError;
                        numerator = i;
                        denominator = j;
                    }

                    if (currentError < epsilon)
                    {
                        // Đạt được độ chính xác đủ, thoát vòng lặp
                        break;
                    }
                }
            }

            // Đơn giản hóa phân số bằng cách chia tử số và mẫu số cho GCD
            int gcd = GetGCD(numerator, denominator);
            numerator /= gcd;
            denominator /= gcd;
        }

        public static int GetGCD(int a, int b)
        {
            // Thuật toán Euclid để tìm ước số chung lớn nhất (GCD)
            while (b != 0)
            {
                int temp = b;
                b = a % b;
                a = temp;
            }
            return a;
        }

        ///--------------------------------------

        public static List<string> GetLayoutNames()
        {
            List<string> laname = new List<string>();
            Document doc = acap.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction trx = db.TransactionManager.StartTransaction())
            {
                DBDictionary layoutDic = trx.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;

                foreach (DBDictionaryEntry entry in layoutDic)
                {
                    Layout lay = trx.GetObject(entry.Value, OpenMode.ForRead) as Layout;
                    laname.Add(lay.LayoutName);
                }


            }
            return laname;
        }
        public static void SortLayout(List<string> myList)
        {     

            // Tìm vị trí của đối tượng có tên "Model"
            int index = myList.FindIndex(x => x == "Model");

            if (index >= 0)
            {                
                string model = myList[index];
                myList.RemoveAt(index);
                myList.Insert(0, model);
            }

         
        }
        public static void SetCurrentLayoutTab(Document doc,string tab)
        {
            //var doc = Application.DocumentManager.MdiActiveDocument;
            using (doc.LockDocument())
                LayoutManager.Current.CurrentLayout = tab;
        }
    }

}
