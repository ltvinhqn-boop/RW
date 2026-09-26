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
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.PlottingServices;
using System.Collections.Specialized;
//using Autodesk.AutoCAD.Interop;
namespace DWG2PDF
{
    public class PrinterUtility
    {

        static public void ListPlotDevices(ToolStripComboBox comboBox)

        {

            Document doc = AcAp.DocumentManager.MdiActiveDocument;

            Editor ed = doc.Editor;



            PlotConfigInfoCollection devices = PlotConfigManager.Devices;

            foreach (PlotConfigInfo info in devices)

            {
                comboBox.Items.Add(info.DeviceName);

                // ed.WriteMessage(info.DeviceName + "\n");

            }

        }
        static public StringCollection GetPaperSizes(string plotConfigName)
        {
            PlotSettingsValidator plotSettingsValidator = PlotSettingsValidator.Current;
            PlotSettings plotSettings = new PlotSettings(true);
            //StringCollection result = new StringCollection();
            StringCollection collection = new StringCollection();

            using (plotSettings)
            {
                plotSettingsValidator.SetPlotConfigurationName(plotSettings, plotConfigName, null);
                plotSettingsValidator.RefreshLists(plotSettings);
                collection = plotSettingsValidator.GetCanonicalMediaNameList(plotSettings);
                // Chuyển StringCollection thành mảng string
                string[] array = new string[collection.Count];
                collection.CopyTo(array, 0);

                // Sắp xếp mảng theo thứ tự bảng chữ cái
                Array.Sort(array);
                List<string> filteredList = new List<string>();

                foreach (string item in array)
                {
                    if (item.IndexOf("user", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        filteredList.Add(item);
                    }
                }
                array = array.Except(filteredList).ToArray();
                // Chuyển mảng trở lại StringCollection (nếu cần)
                collection.Clear();
                collection.AddRange(array);
            }

            return collection;
        }
        //---------------------------------------------
        public static StringCollection PlotStyleList()
        {
            StringCollection collection = new StringCollection();
            foreach (string plotStyle in PlotSettingsValidator.Current.GetPlotStyleSheetList())
            {
                //string plotStyleDirectory = PlotSettingsValidator.GetCanonicalMediaName("Plot Styles");
                //  if (plotStyle.IndexOf(".stb", StringComparison.OrdinalIgnoreCase) < 0)
                //  {
                collection.Add(plotStyle);
              //  }
            }
            return collection;
        }
        static public void additemcombobox(ToolStripComboBox cbb, StringCollection strs)
        {
            cbb.Items.Clear();       
            foreach (string str in strs)
            {
                if (str.IndexOf("in", StringComparison.OrdinalIgnoreCase) >= 0|| str.IndexOf("Envelope", StringComparison.OrdinalIgnoreCase) >= 0|| str.IndexOf("10x14", StringComparison.OrdinalIgnoreCase) >= 0|| str.IndexOf("11x17", StringComparison.OrdinalIgnoreCase) >= 0)
                {

                    continue;
                }
                cbb.Items.Add(str);
            }
        }
        //--------------------------------------------------

        //
    }
}
