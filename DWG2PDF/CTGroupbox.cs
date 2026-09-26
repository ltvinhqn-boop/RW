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
    public partial class CTGroupbox : UserControl
    {
        private Color innerBorderColor = Color.FromArgb(50,50,50); // Màu sắc của đường viền phía trong
        private Color outerBorderColor = Color.Blue; // Màu sắc của đường viền bên ngoài
        private int cornerRadius = 10;
        private int innerBorderWidth = 10;
        private int wtext = 10;// Độ cong của góc bo
        private string groupText = "GroupBox"; // Văn bản hiển thị trên border
        public CTGroupbox()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.DoubleBuffer, true);
            UpdateStyles();
        }

        public Color InnerBorderColor
        {
            get { return innerBorderColor; }
            set
            {
                innerBorderColor = value;
                Invalidate();
            }
        }

        public Color OuterBorderColor
        {
            get { return outerBorderColor; }
            set
            {
                outerBorderColor = value;
                Invalidate();
            }
        }
        public string GroupText
        {
            get { return groupText; }
            set
            {
                groupText = value;
                Invalidate(); // Gọi lại phương thức OnPaint để vẽ lại UserControl
            }
        }
        public int CornerRadius
        {
            get { return cornerRadius; }
            set
            {
                cornerRadius = value;
                Invalidate(); // Gọi lại phương thức OnPaint để vẽ lại UserControl
            }
        }
        public int widthText
        {
            get { return wtext; }
            set
            {
                wtext = value;
                Invalidate(); // Gọi lại phương thức OnPaint để vẽ lại UserControl
            }
        }
        public int BorderWidth
        {
            get { return innerBorderWidth; }
            set
            {
                innerBorderWidth = value;
                Invalidate(); // Gọi lại phương thức OnPaint để vẽ lại UserControl
            }
        }
        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);

            using (Pen innerBorderPen = new Pen(innerBorderColor,innerBorderWidth))
            using (Pen outerBorderPen = new Pen(BackColor))
            using (SolidBrush textBrush = new SolidBrush(outerBorderColor))
            {
                // Vẽ đường viền phía trong
                int innerBorderWidth = 6; // Độ dày của innerBorder
                Rectangle innerBorderRect = new Rectangle(innerBorderWidth, innerBorderWidth, Width - innerBorderWidth * 2, Height - innerBorderWidth * 2);
                DrawRoundedRectangle2(e.Graphics, innerBorderPen, innerBorderRect, cornerRadius,wtext,innerBorderWidth);

                // Vẽ đường viền bên ngoài và tạo vùng vẽ cho văn bản
                Rectangle outerBorderRect = new Rectangle(0, 0, Width - 1, Height - 1);
                DrawRoundedRectangle(e.Graphics, outerBorderPen, outerBorderRect, cornerRadius);

                // Vẽ văn bản trong vùng vẽ
               // string text = "Grouptext";
                using (StringFormat format = new StringFormat())
                {
                    format.Alignment = StringAlignment.Near;
                    format.LineAlignment = StringAlignment.Near;

                    // Tính toán vị trí mới của văn bản
                    int textX =  cornerRadius*2; // Vị trí ngang giữa border cộng với offsetX
                    int textY = -1; // Vị trí dọc giữa border cộng với offsetY

                    // Vẽ văn bản tại vị trí mới
                    PointF textPosition = new PointF(textX, textY);

                    e.Graphics.DrawString(groupText, Font, new SolidBrush(ForeColor), textPosition, format);
                }
            }
        }

        private void DrawRoundedRectangle(Graphics graphics, Pen pen, Rectangle rect, int cornerRadius)
        {
            int diameter = cornerRadius * 2;
            Size size = new Size(diameter, diameter);
            Rectangle arcRect = new Rectangle(rect.Location, size);

            // Vẽ các góc bo
            graphics.DrawArc(pen, arcRect, 180, 90);
            arcRect.X = rect.Right - diameter;
            graphics.DrawArc(pen, arcRect, 270, 90);
            arcRect.Y = rect.Bottom - diameter;
            graphics.DrawArc(pen, arcRect, 0, 90);
            arcRect.X = rect.Left;
            graphics.DrawArc(pen, arcRect, 90, 90);

            // Vẽ các đoạn thẳng
            graphics.DrawLine(pen, rect.Left + cornerRadius, rect.Top, rect.Right - cornerRadius, rect.Top);
            graphics.DrawLine(pen, rect.Right, rect.Top + cornerRadius, rect.Right, rect.Bottom - cornerRadius);
            graphics.DrawLine(pen, rect.Right - cornerRadius, rect.Bottom, rect.Left + cornerRadius, rect.Bottom);
            graphics.DrawLine(pen, rect.Left, rect.Bottom - cornerRadius, rect.Left, rect.Top + cornerRadius);
        }
        private void DrawRoundedRectangle2(Graphics g, Pen pen, Rectangle rect, int cornerRadius, int wtext, int innerBorderWidth)
        {
            int diameter = cornerRadius * 2;
            Rectangle arcRect = new Rectangle(rect.Left, rect.Top, diameter, diameter);

            // Vẽ các góc bo
            g.DrawArc(pen, arcRect, 180, 90);
            arcRect.X = rect.Right - diameter;
            g.DrawArc(pen, arcRect, 270, 90);
            arcRect.Y = rect.Bottom - diameter;
            g.DrawArc(pen, arcRect, 0, 90);
            arcRect.X = rect.Left;
            g.DrawArc(pen, arcRect, 90, 90);

            // Vẽ các đoạn thẳng
            g.DrawLine(pen, rect.Left + cornerRadius+wtext, rect.Top, rect.Right - cornerRadius+ innerBorderWidth/4, rect.Top);
            g.DrawLine(pen, rect.Right, rect.Top + cornerRadius, rect.Right, rect.Bottom - cornerRadius+ innerBorderWidth / 4);
            g.DrawLine(pen, rect.Right - cornerRadius+ innerBorderWidth /4, rect.Bottom, rect.Left + cornerRadius, rect.Bottom);
            g.DrawLine(pen, rect.Left, rect.Bottom - cornerRadius + innerBorderWidth / 4, rect.Left, rect.Top + cornerRadius + innerBorderWidth / 4);
        }
    }
}
