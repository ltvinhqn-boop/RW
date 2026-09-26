using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DWG2PDF
{
    public partial class CBBT2 : UserControl
    {
        
        public CBBT2()
        {
            InitializeComponent();
          
        }
       
      
        public string NameRemove
        {
            get { return bt_rm.Name; }
            set { bt_rm.Name = value; }
        }
        public string NamePrint
        {
            get { return bt_print.Name; }
            set { bt_print.Name = value; }
        }
        public string TextNumber
        {
            get { return tb_cbbt3.Text; }
            set { tb_cbbt3.Text = value; }
        }
        public string TextDrawing
        {
            get { return tb_cbbt2.Text; }
            set { tb_cbbt2.Text = value; }
        }
     

        public Color ColorButtonPrint
        {
            get { return bt_print.ForeColor; }
            set { bt_print.ForeColor = value; }
        }
        private void bt_la1_Click(object sender, EventArgs e)
        {
            //vb10bool = !vb10bool;
            //if (vb10bool == true)
            //{
            //    bt_la1.ForeColor = Color.Yellow;
            //}
            //else
            //{
            //    bt_la1.ForeColor = Color.White;
            //}
            OnBtla1Clicked();
        }
    

        public void bt_rm_Click(object sender, EventArgs e)
        {
            OnBtrmClicked();
        }

   

     
        private void bt_rm_MouseEnter(object sender, EventArgs e)
        {
            bt_rm.BorderColor = Color.FromArgb(55, 55, 55);
            bt_rm.BackColor = Color.FromArgb(55, 55, 55);
            bt_rm.BackgroundColor = Color.FromArgb(55, 55, 55);

        }

        private void bt_rm_MouseLeave(object sender, EventArgs e)
        {
            bt_rm.BorderColor = Color.FromArgb(50, 50, 50);
            bt_rm.BackColor = Color.FromArgb(50, 50, 50);
            bt_rm.BackgroundColor = Color.FromArgb(50, 50, 50);
        }


        

        public event EventHandler BtMd1Clicked;
        public event EventHandler Btla1Clicked;
        public event EventHandler BtrmClicked;
        public event EventHandler BtprintClicked;
        public event EventHandler Textbox1Clicked;
        public event EventHandler Textbox2Clicked;
        // Khai báo các sự kiện cho các button khác (nếu có)

        private void OnBtMd1Clicked()
        {
            BtMd1Clicked?.Invoke(this, EventArgs.Empty);
        }

        private void OnBtla1Clicked()
        {
            Btla1Clicked?.Invoke(this, EventArgs.Empty);
        }
        private void OnBtrmClicked()
        {
            BtrmClicked?.Invoke(this, EventArgs.Empty);
        }
        private void OnBtprintClicked()
        {
            BtprintClicked?.Invoke(this, EventArgs.Empty);
        }
        private void OnTextbox1Clicked()
        {
           Textbox1Clicked?.Invoke(this, EventArgs.Empty);
        }
        private void OnTextbox2Clicked()
        {
            Textbox2Clicked?.Invoke(this, EventArgs.Empty);
        }
        private void bt_mdel_Click(object sender, EventArgs e)
        {
          //  vb9bool1 = !vb9bool1;
          //  if (vb9bool1 == true)
          //  {
         //       bt_mdel.ForeColor = Color.Yellow;
         //   }
           // else
          //  {
          //      bt_mdel.ForeColor = Color.White;
         //   }
            OnBtMd1Clicked();
        }

     

        private void bt_print_Click(object sender, EventArgs e)
        {
            //vb11bool = !vb11bool;
            //if (vb11bool == true)
            //{
            //    bt_print.ForeColor = Color.Yellow;
            //}
            //else
            //{
            //    bt_print.ForeColor = Color.White;
            //}
            OnBtprintClicked();
        }

        private void bt_print_MouseEnter(object sender, EventArgs e)
        {
            bt_print.BorderColor = Color.FromArgb(55, 55, 55);
            bt_print.BackColor = Color.FromArgb(55, 55, 55);
            bt_print.BackgroundColor = Color.FromArgb(55, 55, 55);
        }

        private void bt_print_MouseLeave(object sender, EventArgs e)
        {
            bt_print.BorderColor = Color.FromArgb(50, 50, 50);
            bt_print.BackColor = Color.FromArgb(50, 50, 50);
            bt_print.BackgroundColor = Color.FromArgb(50, 50, 50);
        }

        private void tb_cbbt2_Click(object sender, EventArgs e)
        {
            OnTextbox1Clicked();
        }

        private void tb_cbbt3_Click(object sender, EventArgs e)
        {
            OnTextbox2Clicked();
        }

        private void tb_cbbt3_MouseEnter(object sender, EventArgs e)
        {
            tb_cbbt3.ForeColor = Color.Yellow;
        }

        private void tb_cbbt3_MouseLeave(object sender, EventArgs e)
        {
            tb_cbbt3.ForeColor = Color.White;
        }

        private void tb_cbbt2_MouseEnter(object sender, EventArgs e)
        {
            tb_cbbt2.ForeColor = Color.Yellow;
        }

        private void tb_cbbt2_MouseLeave(object sender, EventArgs e)
        {
            tb_cbbt2.ForeColor = Color.White;
        }

        private void bt_rm_MouseHover(object sender, EventArgs e)
        {
            toolTip2.SetToolTip(bt_rm, "Xóa Layout");
        }

        private void bt_print_MouseHover(object sender, EventArgs e)
        {
            if (bt_print.ForeColor == Color.Yellow)
            {
                toolTip2.SetToolTip(bt_print, "Tắt Layout");
            }
            else
            {
                toolTip2.SetToolTip(bt_print, "Bật Layout");
            }
        }
    }
}
