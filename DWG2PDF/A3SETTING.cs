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
    public partial class A3SETTING : Form
    {
        IntPtr acadHwnd = new IntPtr(AcAp.MainWindow.Handle.ToInt64());
        public A3SETTING()
        {
            InitializeComponent();
        }

        private void vbButton1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "DWG Files (*.dwg)|*.dwg|All Files (*.*)|*.*";
            openFileDialog.Title = "Chọn file DWG";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string selectedFilePath = openFileDialog.FileName;
                // flatCombobox1.Items.Add(selectedFilePath);
                if (flatCombobox1.Items.Count > 0)
                {
                    //textBox2.Text = folderBrowserDialog.SelectedPath;
                    flatCombobox1.Items[0] = selectedFilePath;

                }
                else
                {
                    flatCombobox1.Items.Add(selectedFilePath);
                    flatCombobox1.SelectedIndex = 0;
                }
            }
        }

        private void vbButton2_Click(object sender, EventArgs e)
        {
            Dictionary<string, string> keyValuePairs = new Dictionary<string, string>();
            keyValuePairs.Add("Pathinputa3", flatCombobox1.SelectedItem.ToString());
            keyValuePairs.Add("Blockorxref", flatCombobox2.SelectedItem.ToString());

            ResxFileHelper.WriteToResxFile(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Properties\Resources.resx", keyValuePairs);
            //  ResxFileHelper.WriteToResxFile(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Properties\Resources.resx", "Pathinputa3", "ddd");
            //  ResxFileHelper.WriteToResxFile(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Properties\Resources.resx", "Blockorxref", flatCombobox2.SelectedItem.ToString());
        }

        private void A3SETTING_Load(object sender, EventArgs e)
        {
            String GTCBB1 = ResxFileHelper.ReadFromResxFile(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Properties\Resources.resx", "Pathinputa3");
            String GTCBB2 = ResxFileHelper.ReadFromResxFile(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Properties\Resources.resx", "Blockorxref");
            flatCombobox2.SelectedIndex = flatCombobox2.FindString(GTCBB2);
            flatCombobox1.Items.Add(GTCBB1);
            flatCombobox1.SelectedIndex = 0;


        }

        private void vbButton1_MouseHover(object sender, EventArgs e)
        {
            toolTip1.SetToolTip(vbButton1, "Chọn đường dẫn tới block khung tên mẫu");
        }

        private void vbButton2_MouseHover(object sender, EventArgs e)
        {
            toolTip1.SetToolTip(vbButton2, "Lưu thiết lập");
        }

        /*  private void vbButton3_Click(object sender, EventArgs e)
          {
              NativeMethods.SetForegroundWindow(acadHwnd);
              if (File.Exists(flatCombobox1.Text))
              {
                  using (Document newDocument = AcAp.DocumentManager.Open(flatCombobox1.Text, false))
                  {  if (select.SelectBlock() != null)
                      {
                          flatCombobox3.Items.Insert(0,select.SelectBlock());
                          flatCombobox3.SelectedIndex = 0;
                      }
                      newDocument.CloseAndDiscard();
                  }

              }
          }*/

    }
}
