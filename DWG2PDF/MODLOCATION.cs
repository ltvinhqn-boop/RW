using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;
using System.Collections.Specialized;
using System.Diagnostics;
namespace DWG2PDF
{
    class MODLOCATION
    {
        public static void Regenlocationplus(Form fr, TabControl tabctr,Panel pn, List<Panel> pnels, List<TableLayoutPanel> tblps, int hei)
        {
            fr.Size = new Size(fr.Size.Width, fr.Size.Height + hei);
           // tabctr.Size = new Size(tabctr.Size.Width, tabctr.Size.Height + hei);
            pn.Size = new Size(pn.Size.Width, pn.Size.Height + hei);
            foreach (Panel pnel in pnels)
            {
                pnel.Location =new Point(pnel.Location.X, pnel.Location.Y + hei);
            }
            foreach (TableLayoutPanel tblp in tblps)
                {
                tblp.Location = new Point(tblp.Location.X, tblp.Location.Y + hei);
            }

             

                
                
          
        }
            public static void Regenlocationsub(Form fr,TabControl tabctr, Panel pn, List<Panel> pnels, List<TableLayoutPanel> tblps, int hei)
        {

            foreach (TableLayoutPanel tblp in tblps)
            {
                tblp.Location = new Point(tblp.Location.X, tblp.Location.Y - hei);
            }

            foreach (Panel pnel in pnels)
            {
                pnel.Location = new Point(pnel.Location.X, pnel.Location.Y - hei);
            }
            pn.Size = new Size(pn.Size.Width, pn.Size.Height - hei);
           // tabctr.Size = new Size(tabctr.Size.Width, tabctr.Size.Height - hei);
                fr.Size = new Size(fr.Size.Width, fr.Size.Height - hei);

            
        
        }
    }
}
