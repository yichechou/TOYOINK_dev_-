using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TOYOINK_dev.SingleQuery
//namespace TOYOINK_dev
{
    public partial class fm_Query_INVPL_Order : Form
    {
        public fm_Query_INVPL_Order()
        {
            InitializeComponent();
        }

        private void btn_Search_Click(object sender, EventArgs e)
        {

        }

        private void btn_Clear_Click(object sender, EventArgs e)
        {
            //尋找panel1內的控制鍵
            foreach (Control p1 in panel1.Controls)
            {
                //判別為Textbox
                if (p1 is TextBox)
                {
                    TextBox tb = p1 as TextBox;
                    tb.Text = "";
                }
                //判別為Combobox
                else if (p1 is ComboBox)
                {
                    ComboBox cob = p1 as ComboBox;
                    //判別名稱前八碼為為[cob_Cond]回到預設
                    if (cob.Name.Substring(0,8) == "cbo_Cond") 
                    {
                        cob.SelectedIndex = 0;

                    }
                }

            }

        }
    }
}
