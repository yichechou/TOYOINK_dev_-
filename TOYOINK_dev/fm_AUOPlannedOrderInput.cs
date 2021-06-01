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
using Myclass;
using ClosedXML.Excel;
using NPOI;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

namespace TOYOINK_dev
{
    public partial class fm_AUOPlannedOrderInput : Form
    {
        public MyClass MyCode;
        string defaultfilePath_from = "", defaultfilePath_to = "";
        string str_廠別 = "A01A", str_建立者ID = "", str_建立者GP = "", str_建立日期 = "";
        月曆 fm_月曆;
        
        DataTable dt_建立者;
        public fm_AUOPlannedOrderInput()
        {
            InitializeComponent();
        }

        private void fm_AUOPlannedOrderInput_Load(object sender, EventArgs e)
        {

            //fm_AUOPlannedOrder.WriteLog("恢復預設值");
        }

        bool IsToForm1 = false; //紀錄是否要回到Form1
        protected override void OnClosing(CancelEventArgs e) //在視窗關閉時觸發
        {
            ////DialogResult dr = MessageBox.Show("\"是\"回到主畫面 \r\n \"否\"關閉程式", "是否要關閉程式"
            ////    , MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            ////if (dr == DialogResult.Yes)
            ////{
            ////    IsToForm1 = true;
            ////}

            //base.OnClosing(e);
            //if (IsToForm1) //判斷是否要回到Form1
            //{
            //    this.DialogResult = DialogResult.Yes; //利用DialogResult傳遞訊息
            //    fm_menu fm_menu = (fm_menu)this.Owner; //取得父視窗的參考
            //}
            //else
            //{
            //    this.DialogResult = DialogResult.No;
            //}
            Environment.Exit(Environment.ExitCode);
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

        private void btn_fileopen_from_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process prc = new System.Diagnostics.Process();
            prc.StartInfo.FileName = txt_path_from.Text.ToString();
            prc.Start();
        }

        private void btn_file_to_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process prc = new System.Diagnostics.Process();
            prc.StartInfo.FileName = txt_path_to.Text.ToString();
            prc.Start();
        }

        private void btn_fileopen_to_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();

            //首次defaultfilePath为空，按FolderBrowserDialog默认设置（即桌面）选择
            if (defaultfilePath_to != "")
            {
                //设置此次默认目录为上一次选中目录
                dialog.SelectedPath = defaultfilePath_to;
            }

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                //记录选中的目录
                defaultfilePath_to = dialog.SelectedPath;
                txt_path_from.Text = defaultfilePath_to;
            }
        }

        string save_as_Trademark = "", temp_excel;

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void btn_DateInput_Click(object sender, EventArgs e)
        {
            //TODO:單頭及單身若不為空值，表示已轉換ERP格式，需重新轉換 或 資料已上傳ERP，需重新選擇日期
            //資料上傳ERP後，dgv_excel會清空
            //if (dgv_tc.DataSource != null || dgv_td.DataSource != null || dgv_excel.DataSource != null)
            if (btn_ToERP.Enabled == true || btn_UpERP.Enabled == true)

            {
                DialogResult Result = MessageBox.Show("修改 單據日期 後，需重新【選擇檔案】", "Excel檔案已匯入", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);

                if (Result == DialogResult.OK)
                {
                    //lab_status.Text = "請 選擇檔案";
                    //txt_path.Text = "";
                    //dgv_excel.DataSource = null;
                    //dgv_cfipo.DataSource = null;
                    //tabCtl_data.SelectedIndex = 0;
                    btn_ToERP.Enabled = false;
                    btn_ToERP.BackColor = System.Drawing.SystemColors.Control;
                    btn_ToERP.ForeColor = System.Drawing.SystemColors.ControlText;
                    btn_UpERP.Enabled = false;
                    btn_UpERP.BackColor = System.Drawing.SystemColors.Control;
                    btn_UpERP.ForeColor = System.Drawing.SystemColors.ControlText;
                    //dgv_tc.DataSource = null;
                    //dgv_td.DataSource = null;

                    txterr.Text += Environment.NewLine +
                               DateTime.Now.ToString() + Environment.NewLine +
                               " 修改單據日期，請重新【選擇檔案】" + Environment.NewLine +
                               "===========";

                    this.fm_月曆 = new 月曆(this.txt_DateInput, this.btn_DateInput, "單據日期");
                }
                else if (Result == DialogResult.Cancel)
                {
                    return;
                }
            }
            else
            {
                this.fm_月曆 = new 月曆(this.txt_DateInput, this.btn_DateInput, "單據日期");
                //btn_file.Enabled = true;
                //lab_status.Text = "請 選擇檔案";

            }
        }

        private void btn_ToERP_Click(object sender, EventArgs e)
        {
            
        }

        private void btn_UpERP_Click(object sender, EventArgs e)
        {

        }

        private void btn_ToExcel_Click(object sender, EventArgs e)
        {

        }

        private void cob_Creater_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void btn_search_Click(object sender, EventArgs e)
        {
            //txt_path_from.Text = @"D:\test_excel\VISFCSTSupply_AUO MS Yvonne 20200333.xls";
            //dgv_def.DataSource = MyClass.ReadExcelToTable(txt_path_from.Text.ToString(),"FAB");
            //取得資料夾內的檔名

            // Create a reference to the current directory.
            DirectoryInfo File = new DirectoryInfo(txt_path_from.Text.ToString());
            // Create an array representing the files in the current directory.
            FileInfo[] filenameList = File.GetFiles();

            int sheetnum = 1;
            //temp_excel = @"\\192.168.128.219\Company\會計\商標權報表\商標權報表_temp.xlsx";

            //資料夾內Excel合併
            
                //for (int i = 0; i < filenameList.Length; i++)
                //{
                ////temp_excel = filenameList[0].FullName;

                ////打开一个文档
                ////HSSFWorkbook workbook;
                //HSSFWorkbook workbook = new HSSFWorkbook(filenameList[0].FullName);
                ////using (FileStream stream = File.OpenRead(filenameList[0].FullName))
                ////{
                ////    workbook = new HSSFWorkbook(stream);
                ////}

                //int SheetCount = workbook.NumberOfSheets;//获取表的数量
                //string[] SheetName = new string[SheetCount];//保存表的名称
                //for (int j = 0; j < SheetCount; j++)
                //    SheetName[j] = workbook.GetSheetName(j);
                //foreach (string j in SheetName)//测试读取状态
                //    txterr.Text += j;
                //workbook.Clear();

                //}


        }

        private void btn_file_from_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();

            //首次defaultfilePath为空，按FolderBrowserDialog默认设置（即桌面）选择
            if (defaultfilePath_from != "")
            {
                //设置此次默认目录为上一次选中目录
                dialog.SelectedPath = defaultfilePath_from;
            }

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                //记录选中的目录
                defaultfilePath_from = dialog.SelectedPath;
                txt_path_from.Text = defaultfilePath_from;
            }
        }
    }
}
