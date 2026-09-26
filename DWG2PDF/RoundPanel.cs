using System.Drawing;
using System.Windows.Forms;
using System.Drawing.Drawing2D;
public class CustomGroupBox : GroupBox
{
    private Color borderColor = Color.Blue; // Màu viền tùy chỉnh

    public Color BorderColor
    {
        get { return borderColor; }
        set
        {
            borderColor = value;
            Invalidate(); // Gọi lại phương thức OnPaint để vẽ lại GroupBox
        }
    }
    protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e);

        int cornerRadius = 10; // Độ cong của góc bo (có thể điều chỉnh)

        using (Pen borderPen = new Pen(Color.Black, 2)) // Đặt màu và độ dày của đường viền
        {
            GraphicsPath path = new GraphicsPath();
            Rectangle borderRect = new Rectangle(0, 0, Width - 1, Height - 1);
            int diameter = cornerRadius * 2;

            // Vẽ góc bo trên bên trái
            path.AddArc(borderRect.Left, borderRect.Top, diameter, diameter, 180, 90);

            // Vẽ đường ngang bên trên
            path.AddLine(borderRect.Left + cornerRadius, borderRect.Top, borderRect.Right - cornerRadius, borderRect.Top);

            // Vẽ góc bo trên bên phải
            path.AddArc(borderRect.Right - diameter, borderRect.Top, diameter, diameter, 270, 90);

            // Vẽ đường dọc bên phải
            path.AddLine(borderRect.Right, borderRect.Top + cornerRadius, borderRect.Right, borderRect.Bottom - cornerRadius);

            // Vẽ góc bo dưới bên phải
            path.AddArc(borderRect.Right - diameter, borderRect.Bottom - diameter, diameter, diameter, 0, 90);

            // Vẽ đường ngang bên dưới
            path.AddLine(borderRect.Right - cornerRadius, borderRect.Bottom, borderRect.Left + cornerRadius, borderRect.Bottom);

            // Vẽ góc bo dưới bên trái
            path.AddArc(borderRect.Left, borderRect.Bottom - diameter, diameter, diameter, 90, 90);

            // Vẽ đường dọc bên trái
            path.AddLine(borderRect.Left, borderRect.Bottom - cornerRadius, borderRect.Left, borderRect.Top + cornerRadius);

            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias; // Kích hoạt chế độ làm mờ
            e.Graphics.DrawPath(borderPen, path);
        }
    }
}