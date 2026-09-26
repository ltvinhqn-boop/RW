using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DWG2PDF
{
    public partial class Form2 : Form
    {
        private Form1 _form1;

        public Form2(Form1 form1)
        {
            InitializeComponent();
            _form1 = form1;
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            List<ZamPara> ListZam2 = _form1.ListZam;
            //----------------------------------------
            lb_12.Text = ListZam2[0].LengthX.ToString("0.00");
            lb_13.Text = ListZam2[0].HeightY.ToString("0.00");
            lb_14.Text = ListZam2[0].Quanlity.ToString();
            //----------------------------------------
            //----------------------------------------
            lb_22.Text = ListZam2[1].LengthX.ToString("0.00");
            lb_23.Text = ListZam2[1].HeightY.ToString("0.00");
            lb_24.Text = ListZam2[1].Quanlity.ToString();
            //----------------------------------------
            //----------------------------------------
            lb_32.Text = ListZam2[2].LengthX.ToString("0.00");
            lb_33.Text = ListZam2[2].HeightY.ToString("0.00");
            lb_34.Text = ListZam2[2].Quanlity.ToString();
            //----------------------------------------
            //----------------------------------------
            lb_42.Text = ListZam2[3].LengthX.ToString("0.00");
            lb_43.Text = ListZam2[3].HeightY.ToString("0.00");
            lb_44.Text = ListZam2[3].Quanlity.ToString();
            //----------------------------------------
            //----------------------------------------
            lb_52.Text = ListZam2[4].LengthX.ToString("0.00");
            lb_53.Text = ListZam2[4].HeightY.ToString("0.00");
            lb_54.Text = ListZam2[4].Quanlity.ToString();
            //----------------------------------------
            lb_15.Text = _form1.list914[0].ToString();
            lb_25.Text = _form1.list914[1].ToString();
            lb_35.Text = _form1.list914[2].ToString();
            lb_45.Text = _form1.list914[3].ToString();
            lb_55.Text = _form1.list914[4].ToString();
            //----------------------------------------
            //----------------------------------------
            lb_16.Text = _form1.list1219[0].ToString();
            lb_26.Text = _form1.list1219[1].ToString();
            lb_36.Text = _form1.list1219[2].ToString();
            lb_46.Text = _form1.list1219[3].ToString();
            lb_56.Text = _form1.list1219[4].ToString();
            //----------------------------------------
        }
    }
}
