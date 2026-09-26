using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DWG2PDF
{
    public class CustomTableLayoutPanel : TableLayoutPanel
    {
        private Color cellboderColor = Color.LightBlue;
        private int cellBorderWidth = 2;

        public Color CellBoderColor
        {
            get { return cellboderColor; }
            set { cellboderColor = value; }
        }

        public int CellBorderWidth
        {
            get { return cellBorderWidth; }
            set { cellBorderWidth = value; }
        }

        protected override void OnCellPaint(TableLayoutCellPaintEventArgs e)
        {
            base.OnCellPaint(e);

        

            // Vẽ border cho ô
            ControlPaint.DrawBorder(e.Graphics, e.CellBounds, cellboderColor, cellBorderWidth, ButtonBorderStyle.Solid,
                                    cellboderColor, cellBorderWidth, ButtonBorderStyle.Solid,
                                   cellboderColor, cellBorderWidth, ButtonBorderStyle.Solid,
                                    cellboderColor, cellBorderWidth, ButtonBorderStyle.Solid);
        }
    }
}
