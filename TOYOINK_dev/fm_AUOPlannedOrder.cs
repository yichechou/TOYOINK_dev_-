using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Transactions;
using Myclass;

namespace TOYOINK_dev
{
    public partial class fm_AUOPlannedOrder : Form
    {
        public MyClass MyCode;

        月曆 fm_月曆;
        string str_Line;
        public fm_AUOPlannedOrder()
        {
            InitializeComponent();
            MyCode = new Myclass.MyClass();
            MyCode.strDbCon = "packet size=4096;user id=pwuser;password=sqlmis003;data source=192.168.128.219;persist security info=False;initial catalog=A01A;";
            //MyCode.strDbCon = "packet size=4096;user id=yj.chou;password=yjchou3369;data source=192.168.128.219;persist security info=False;initial catalog=Leader;";
        }

        private void fm_AUOPlannedOrder_Load(object sender, EventArgs e)
        {
            //fm_AUOPlannedOrder.WriteLog("恢復預設值");
        }

        bool IsToForm1 = false; //紀錄是否要回到Form1
        protected override void OnClosing(CancelEventArgs e) //在視窗關閉時觸發
        {
            DialogResult dr = MessageBox.Show("\"是\"回到主畫面 \r\n \"否\"關閉程式", "是否要關閉程式"
                , MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dr == DialogResult.Yes)
            {
                IsToForm1 = true;
            }
             
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

        public static void WriteLog(string message)
        {
            string DIRNAME = Application.StartupPath + @"\Log\";
            string FILENAME = DIRNAME + DateTime.Now.ToString("yyyyMMdd") + ".txt";

            if (!Directory.Exists(DIRNAME))
                Directory.CreateDirectory(DIRNAME);

            if (!File.Exists(FILENAME))
            {
                // The File.Create method creates the file and opens a FileStream on the file. You neeed to close it.
                File.Create(FILENAME).Close();
            }
            using (StreamWriter sw = File.AppendText(FILENAME))
            {
                Log(message, sw);
            }
        }

        private static void Log(string logMessage, TextWriter w)
        {
            w.Write("Log Entry : ");
            w.WriteLine("{0} {1}", DateTime.Now.ToLongDateString(), DateTime.Now.ToLongTimeString());
            w.WriteLine("手動備份:{0}", logMessage);
            w.WriteLine("-------------------------------");
        }

        private void btn_file_Click(object sender, EventArgs e)
        {
            if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                txt_path.Text = this.openFileDialog1.FileName;
            }
            else
            {
                return;
            }

            str_Line = cob_InputExcel.Text.ToString(); //線別名稱

            if (str_Line == "") 
            {
                MessageBox.Show("請先選擇【線別】", "未選擇【線別】", MessageBoxButtons.OK,MessageBoxIcon.Error);
                txt_path.Text = "";
                return;
            }

            DataTable dt_InputExcel = new DataTable();
            DataTable dt_AUO_ERP_NO = new DataTable(); //AUO與ERP品號對照表及公版
            
            
            MyCode.Sql_dt("select * from CT_AUO_ERPNO", dt_AUO_ERP_NO);

            dt_InputExcel = MyClass.ReadExcelSheetToTable("fm_AUOPlannedOrder", txt_path.Text.ToString(), str_Line + "上傳", "1=1");
            dgv_InputExcel.DataSource = dt_InputExcel;

            if (dt_InputExcel == null) 
            {
                 MessageBox.Show("找不到 " + str_Line, "請重新選擇 Excel 或更改【線別】", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
              string Col_name = dgv_InputExcel.Columns[0].HeaderCell.Value.ToString();
        }

        private void button_單據日期_Click(object sender, EventArgs e)
        {
            //TODO:單頭及單身若不為空值，表示已轉換ERP格式，需重新轉換 或 資料已上傳ERP，需重新選擇日期
            //資料上傳ERP後，dgv_excel會清空
            //if (dgv_tc.DataSource != null || dgv_td.DataSource != null || dgv_excel.DataSource != null)
            //if (btn_toerp.Enabled == true || btn_erpup.Enabled == true)

            //{
            //    DialogResult Result = MessageBox.Show("修改 單據日期 後，需重新【選擇檔案】", "Excel檔案已匯入", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);

            //    if (Result == DialogResult.OK)
            //    {
            //        lab_status.Text = "請 選擇檔案";
            //        txt_path.Text = "";
            //        dgv_excel.DataSource = null;
            //        dgv_cfipo.DataSource = null;
            //        tabCtl_data.SelectedIndex = 0;
            //        btn_toerp.Enabled = false;
            //        btn_toerp.BackColor = System.Drawing.SystemColors.Control;
            //        btn_toerp.ForeColor = System.Drawing.SystemColors.ControlText;
            //        btn_erpup.Enabled = false;
            //        btn_erpup.BackColor = System.Drawing.SystemColors.Control;
            //        btn_erpup.ForeColor = System.Drawing.SystemColors.ControlText;
            //        dgv_tc.DataSource = null;
            //        dgv_td.DataSource = null;

            //        txterr.Text += Environment.NewLine +
            //                   DateTime.Now.ToString() + Environment.NewLine +
            //                   " 修改單據日期，請重新【選擇檔案】" + Environment.NewLine +
            //                   "===========";

            //        this.fm_月曆 = new 月曆(this.textBox_單據日期, this.button_單據日期, "單據日期");
            //    }
            //    else if (Result == DialogResult.Cancel)
            //    {
            //        return;
            //    }
            //}
            //else
            //{
                this.fm_月曆 = new 月曆(this.textBox_單據日期, this.button_單據日期, "單據日期");
                btn_file.Enabled = true;
                 cob_InputExcel.Focus();
                 lab_status.Text = "請 選擇【線別】";
            //}
        }

        private void cob_InputExcel_SelectedValueChanged(object sender, EventArgs e)
        {
            btn_file.Focus();
            lab_status.Text = "請 選擇【Excel檔案】";
        }
    }
}
