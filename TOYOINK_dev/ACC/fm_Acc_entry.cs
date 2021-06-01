using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Myclass;
using ClosedXML.Excel;
using System.Globalization;
using System.IO;

namespace TOYOINK_dev.ACC
{
    public partial class fm_Acc_entry : Form
    {
        public MyClass MyCode;
        月曆 fm_月曆;
        string save_as_5aMonth = "", save_as_5aTotal = "", temp_excel_5a, temp_excel_8a, save_as_8aMonth = "", save_as_8aTotal = "";
        string createday = DateTime.Now.ToString("yyyy/MM/dd");

        string str_date_s, str_date_m_s, str_date_s_CloseOut, str_date_m_s_CloseOut;
        string str_date_e, str_date_m_e, str_date_e_ym, str_date_y_e, str_date_e_CloseOut, str_date_m_e_CloseOut, str_date_e_ym_CloseOut, str_date_y_e_CloseOut;

        string defaultfilePath = "";

        DateTime date_s, date_e, date_s_CloseOut, date_e_CloseOut;

        DataTable dt_8aCOPTH = new DataTable();  //8a品種彙總表
        DataTable dt_5aCOPTH = new DataTable();  //5a明細表

   

        DataTable dt_COPTH = new DataTable();  //銷貨單


        public fm_Acc_entry()
        {
            InitializeComponent();
            MyCode = new Myclass.MyClass();
            MyCode.strDbCon = "packet size=4096;user id=pwuser;password=sqlmis003;data source=192.168.128.219;persist security info=False;initial catalog=A01A;";
            //MyCode.strDbCon = "packet size=4096;user id=yj.chou;password=yjchou3369;data source=192.168.128.219;persist security info=False;initial catalog=Leader;";
            temp_excel_5a = @"\\192.168.128.219\Company\MIS自開發主檔\會計報表公版\銷貨成本分析月報5a_temp.xlsx";
        }

       

        private void Btn_date_s_Click(object sender, EventArgs e)
        {
            str_date_s = txt_date_s.Text.Trim();
            this.fm_月曆 = new 月曆(this.txt_date_s, this.Btn_date_s, "單據起始日期");

        }


        private void Btn_date_e_Click(object sender, EventArgs e)
        {
            str_date_e = txt_date_e.Text.Trim();
            this.fm_月曆 = new 月曆(this.txt_date_e, this.Btn_date_e, "單據結束日期");
            str_date_m_e = txt_date_e.Text.Trim().Substring(0, 6);
        }

        private void Btn_date_s_CloseOut_Click(object sender, EventArgs e)
        {
            str_date_s_CloseOut = txt_date_s_CloseOut.Text.Trim();
            this.fm_月曆 = new 月曆(this.txt_date_s_CloseOut, this.Btn_date_s_CloseOut, "單據起始日期");
        }

        private void Btn_date_e_CloseOut_Click(object sender, EventArgs e)
        {
            str_date_e_CloseOut = txt_date_e_CloseOut.Text.Trim();
            this.fm_月曆 = new 月曆(this.txt_date_e_CloseOut, this.Btn_date_e_CloseOut, "單據結束日期");
            str_date_m_e_CloseOut = txt_date_e_CloseOut.Text.Trim().Substring(0, 6);
        }

        private void btn_down_CloseOut_Click(object sender, EventArgs e)
        {
            date_s_CloseOut = DateTime.ParseExact(txt_date_s_CloseOut.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);
            date_e_CloseOut = DateTime.ParseExact(txt_date_e_CloseOut.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);

            txt_date_s_CloseOut.Text = DateTime.Parse(date_s_CloseOut.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyyMMdd");
            txt_date_e_CloseOut.Text = DateTime.Parse(date_e_CloseOut.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd");
        }

        private void btn_up_CloseOut_Click(object sender, EventArgs e)
        {
            date_s_CloseOut = DateTime.ParseExact(txt_date_s_CloseOut.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);
            date_e_CloseOut = DateTime.ParseExact(txt_date_e_CloseOut.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);

            txt_date_s_CloseOut.Text = DateTime.Parse(date_s_CloseOut.ToString("yyyy-MM-01")).AddMonths(1).ToString("yyyyMMdd");
            txt_date_e_CloseOut.Text = DateTime.Parse(date_e_CloseOut.ToString("yyyy-MM-01")).AddMonths(2).AddDays(-1).ToString("yyyyMMdd");
        }

        private void btn_down_Click(object sender, EventArgs e)
        {
            date_s = DateTime.ParseExact(txt_date_s.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);
            date_e = DateTime.ParseExact(txt_date_e.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);

            txt_date_s.Text = DateTime.Parse(date_s.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyyMMdd");
            txt_date_e.Text = DateTime.Parse(date_e.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd");
        }

        private void btn_up_Click(object sender, EventArgs e)
        {
            date_s = DateTime.ParseExact(txt_date_s.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);
            date_e = DateTime.ParseExact(txt_date_e.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);

            txt_date_s.Text = DateTime.Parse(date_s.ToString("yyyy-MM-01")).AddMonths(1).ToString("yyyyMMdd");
            txt_date_e.Text = DateTime.Parse(date_e.ToString("yyyy-MM-01")).AddMonths(2).AddDays(-1).ToString("yyyyMMdd");
        }
        private void btn_file_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();

            //首次defaultfilePath为空，按FolderBrowserDialog默认设置（即桌面）选择
            if (defaultfilePath != "")
            {
                //设置此次默认目录为上一次选中目录
                dialog.SelectedPath = defaultfilePath;
            }

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                //记录选中的目录
                defaultfilePath = dialog.SelectedPath;
                txt_path.Text = defaultfilePath;
            }
        }

        private void btn_fileopen_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process prc = new System.Diagnostics.Process();
            prc.StartInfo.FileName = txt_path.Text.ToString();
            prc.Start();
        }

        private void DtAndDgvClear()
        {
            //單月
            //dt_8aCOPTH_m.Clear();   //8a品種彙總表
            //dt_5aCOPTH_m.Clear();   //5a明細表
            //dt_COPTH_m.Clear();   //銷貨單

            ////累計
            //dt_8aCOPTH.Clear();   //8a品種彙總表
            //dt_5aCOPTH.Clear();   //5a明細表
            //dt_COPTH.Clear();   //銷貨單

            //dgv_8aCOPTH.DataSource = null;
            //dgv_5aCOPTH.DataSource = null;
            //dgv_COPTH.DataSource = null;
            //dgv_COPTJ.DataSource = null;

            BtnFalse();
        }

        private void BtnFalse()
        {
            btn_5aMonth.Enabled = false;
            btn_5aTotal.Enabled = false;
            btn_5aMT.Enabled = false;

        }
        private void BtnTrue()
        {
            btn_5aMonth.Enabled = true;
            btn_5aTotal.Enabled = true;
            btn_5aMT.Enabled = true;

        }

        private void btn_search_CloseOut_Click(object sender, EventArgs e)
        {

        }
        private void btn_search_Click(object sender, EventArgs e)
        {

        }


        bool IsToForm1 = false; //紀錄是否要回到Form1
        protected override void OnClosing(CancelEventArgs e) //在視窗關閉時觸發
        {
            //DialogResult dr = MessageBox.Show("\"是\"回到主畫面 \r\n \"否\"關閉程式", "是否要關閉程式"
            //    , MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            //if (dr == DialogResult.Yes)
            //{
            IsToForm1 = true;
            //}
            //else if (dr == DialogResult.Cancel) 
            //{

            //}

            base.OnClosing(e);
            if (IsToForm1) //判斷是否要回到Form1
            {
                this.DialogResult = DialogResult.Yes; //利用DialogResult傳遞訊息
                fm_menu fm_menu = (fm_menu)this.Owner; //取得父視窗的參考
            }
            else
            {
                this.DialogResult = DialogResult.No;
            }
        }

    }
}
