using System.Windows.Forms;

public class CustomTextBox : TextBox
{
    private int groupText = 28;
    public int GroupText
    {
        get { return groupText; }
        set
        {
            groupText = value;
            Invalidate(); // Gọi lại phương thức OnPaint để vẽ lại UserControl
        }
    }
    protected override void OnCreateControl()
    {
        base.OnCreateControl();

        // Đặt chiều cao mặc định cho TextBox
        
        this.Height = groupText;
    }
}