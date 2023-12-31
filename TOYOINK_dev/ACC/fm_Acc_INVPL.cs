﻿using System;
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
    public partial class fm_Acc_INVPL : Form
    {
        public MyClass MyCode;
        月曆 fm_月曆;
        string save_as_5aMonth = "", save_as_5aTotal = "", temp_excel_5a, temp_excel_8a, save_as_8aMonth = "", save_as_8aTotal = "";
        string createday = DateTime.Now.ToString("yyyy/MM/dd");
        int opencode = 0;

        string str_date_s, str_date_m_s, str_date_ym_s;
        string str_date_e, str_date_m_e, str_date_ym_e, str_date_y_e;

        string defaultfilePath = "";

        DateTime date_s, date_e;

        DataTable dt_8aCOPTH = new DataTable();  //8a品種彙總表
        DataTable dt_5aCOPTH = new DataTable();  //5a明細表
        DataTable dt_COPTH = new DataTable();  //銷貨單
        public fm_Acc_INVPL()
        {
            InitializeComponent();
            MyCode = new Myclass.MyClass();

            //MyCode.strDbCon = MyCode.strDbConLeader;
            //this.sqlConnection1.ConnectionString = MyCode.strDbConLeader;

            MyCode.strDbCon = MyCode.strDbConTemp;
            //this.sqlConnection1.ConnectionString = MyCode.strDbConLeader;

            //MyCode.strDbCon = MyCode.strDbConA01A;
            //this.sqlConnection1.ConnectionString = MyCode.strDbConA01A;

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


        private void btn_down_Click(object sender, EventArgs e)
        {
            date_s = DateTime.ParseExact(txt_date_s.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);
            date_e = DateTime.ParseExact(txt_date_e.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);

            txt_date_s.Text = DateTime.Parse(date_s.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyyMMdd");
            txt_date_e.Text = DateTime.Parse(date_e.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd");

            DtAndDgvClear(); 
        }

        private void btn_up_Click(object sender, EventArgs e)
        {
            date_s = DateTime.ParseExact(txt_date_s.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);
            date_e = DateTime.ParseExact(txt_date_e.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);

            txt_date_s.Text = DateTime.Parse(date_s.ToString("yyyy-MM-01")).AddMonths(1).ToString("yyyyMMdd");
            txt_date_e.Text = DateTime.Parse(date_e.ToString("yyyy-MM-01")).AddMonths(2).AddDays(-1).ToString("yyyyMMdd");

            DtAndDgvClear();
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
        private void btn_search_Click(object sender, EventArgs e)
        {
            if (MyClass.DateIntervalCheck(txt_date_s, txt_date_e) is false)
            {
                return;
            }

            DtAndDgvClear();

            str_date_s = txt_date_s.Text.Trim();
            str_date_ym_s = txt_date_s.Text.Trim().Substring(0, 6);
            str_date_e = txt_date_e.Text.Trim();
            str_date_ym_e = txt_date_e.Text.Trim().Substring(0, 6);

            ////銀行存款明細帳細項_評價後及匯入暫存表CT_F22_1_SGLDT_After_Temp
            //string sql_str_Insert_CT_F22_1_SGLDT_After_Temp = String.Format(@"
            //        ", str_date_ym_e);
            //MyCode.sqlExecuteNonQuery(sql_str_Insert_CT_F22_1_SGLDT_After_Temp, "S2008X64");

            //string sql_str_CT_F22_1_SGLDT_After_Temp = String.Format(@"
            //        select * from CT_F22_1_SGLDT_After_Temp order by 銀行代號,日期");
            //MyCode.Sql_dgv(sql_str_CT_F22_1_SGLDT_After_Temp, dt_SGL_Detail, dgv_SGL_Detail);

            BtnTrue();
        }

        //private void btn_ToE_F22_1_M_Click(object sender, EventArgs e)
        //{
        //    BtnFalse();

        //    using (XLWorkbook wb_F22_1_Month = new XLWorkbook())
        //    {
        //        using (var templateWB = new XLWorkbook(temp_excel_F22_1))
        //        {
        //            var ws = templateWB.Worksheet("F22-1_銀行口座一覧表TAST");
        //            var ws2 = templateWB.Worksheet("明細帳(評價前)");
        //            var ws3 = templateWB.Worksheet("明細帳(評價後)");
        //            var ws4 = templateWB.Worksheet("評價表");
        //            var ws5 = templateWB.Worksheet("明細帳細項");

        //            ws.CopyTo(wb_F22_1_Month, "F22-1_銀行口座一覧表TAST");
        //            ws2.CopyTo(wb_F22_1_Month, "明細帳(評價前)");
        //            ws3.CopyTo(wb_F22_1_Month, "明細帳(評價後)");
        //            ws4.CopyTo(wb_F22_1_Month, "評價表");
        //            ws5.CopyTo(wb_F22_1_Month, "明細帳細項");
        //        }

        //        var wsheet_F22_1_m = wb_F22_1_Month.Worksheet("F22-1_銀行口座一覧表TAST");
        //        var wsheet_SGL_Before = wb_F22_1_Month.Worksheet("明細帳(評價前)");
        //        var wsheet_SGL_After = wb_F22_1_Month.Worksheet("明細帳(評價後)");
        //        var wsheet_ADFOR = wb_F22_1_Month.Worksheet("評價表");
        //        var wsheet_SGL_Detail = wb_F22_1_Month.Worksheet("明細帳細項");

        //        //=== F22-1_銀行口座一覧表TAST ==========================================
        //        //wsheet_F22_1_m.Cell(2, 1).Value = "月份區間:" + str_date_ym_s + "~" + str_date_ym_e; //查詢月份區間
        //        //wsheet_F22_1_m.Cell(3, 1).Value = "製表日期:" + DateTime.Now.ToString("yyyy/MM/dd"); //會計年度

        //        ////== 明細帳(評價前).明細帳(評價後).評價表 =======
        //        ///ERP_DTInputExcel(wsheet_8aCOPTH, dt_8aCOPTH, str_date_y_e + "01");
        //        ERP_DTInputExcel(wsheet_SGL_Before, dt_SGL_Before, 5, 1, str_date_ym_s, "", "本幣期末餘額");
        //        ERP_DTInputExcel(wsheet_SGL_After, dt_SGL_After, 5, 1, str_date_ym_s, "", "本幣期末餘額");
        //        ERP_DTInputExcel(wsheet_ADFOR, dt_ADFOR, 5, 1, str_date_ym_s, "幣別", "原幣存款金額");
        //        ERP_DTInputExcel(wsheet_SGL_Detail, dt_SGL_Detail, 5, 1, str_date_ym_s, "", "");
        //        //ERP_DTInputExcel(wsheet_ADFOR, dt_ADFOR, str_date_ym_s, "幣別", "原幣存款金額;本幣存款金額;重估本幣金額;匯兌損失;淨(損)益");

        //        save_as_F22_1_Month = txt_path.Text.ToString().Trim() + "\\" + str_date_ym_e + @"_F22-1_銀行口座一覧表TAST_" + DateTime.Now.ToString("yyyyMMdd") + @".xlsx";
        //        wb_F22_1_Month.SaveAs(save_as_F22_1_Month);

        //        //打开文件
        //        if (opencode != 1)
        //        {
        //            System.Diagnostics.Process.Start(save_as_F22_1_Month);
        //        }
        //    }
        //    BtnTrue();
        //}

        //private void ERP_DTInputExcel(ClosedXML.Excel.IXLWorksheet wsheet, DataTable dt, int i_col, int j_row, string str_date, string str_SubTotal, string str_Total)
        //{
        //    int i = 0;
        //    int rows_count_dt = dt.Rows.Count;
        //    int col_count_dt = dt.Columns.Count;
        //    string str_SubTotal_Name = "";

        //    wsheet.Cell(2, 2).Value = str_date + "~" + str_date_ym_e; //查詢月份區間
        //    wsheet.Cell(3, 2).Style.NumberFormat.Format = "@";
        //    wsheet.Cell(3, 2).Value = DateTime.Now.ToString("yyyy/MM/dd"); //製表日期

        //    foreach (DataRow row in dt.Rows)
        //    {
        //        int j = 0;

        //        if (str_SubTotal.Length > 0 && str_Total.Length > 0)
        //        {
        //            if (str_SubTotal_Name.ToString() != "" && row[str_SubTotal].ToString() != str_SubTotal_Name.ToString())
        //            {
        //                wsheet.Range("A" + (i + i_col) + ":K" + (i + i_col)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
        //                wsheet.Range("A" + (i + i_col) + ":K" + (i + i_col)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
        //                wsheet.Range("A" + (i + i_col) + ":K" + (i + i_col)).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;
        //                wsheet.Cell(i + i_col, 2).Value = str_SubTotal_Name;
        //                wsheet.Cell(i + i_col, 4).Value = "小計";
        //                wsheet.Range("E" + (i + i_col) + ":K" + (i + i_col)).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
        //                wsheet.Cell(i + i_col, 5).FormulaA1 = "=SUMIFS(E:E,$A:$A,\"" + str_SubTotal_Name + "\")";
        //                wsheet.Cell(i + i_col, 7).FormulaA1 = "=SUMIFS(G:G,$A:$A,\"" + str_SubTotal_Name + "\")";
        //                wsheet.Cell(i + i_col, 8).FormulaA1 = "=SUMIFS(H:H,$A:$A,\"" + str_SubTotal_Name + "\")";
        //                wsheet.Cell(i + i_col, 9).FormulaA1 = "=SUMIFS(I:I,$A:$A,\"" + str_SubTotal_Name + "\")";
        //                wsheet.Cell(i + i_col, 10).FormulaA1 = "=SUMIFS(J:J,$A:$A,\"" + str_SubTotal_Name + "\")";

        //                i++;
        //            }
        //        }

        //        foreach (DataColumn Column in dt.Columns)
        //        {
        //            switch (Column.ColumnName.ToString())
        //            {
        //                case "銀行帳號":
        //                case "銀行代號":
        //                    wsheet.Cell(i + i_col, j + j_row).Style.NumberFormat.Format = "@";
        //                    break;
        //                case "本幣期初餘額":
        //                case "本幣入帳金額":
        //                case "本幣出帳金額":
        //                case "本幣期末餘額":
        //                case "本幣存款金額":
        //                case "重估本幣金額":
        //                case "匯兌收益":
        //                case "匯兌損失":
        //                    wsheet.Cell(i + i_col, j + j_row).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
        //                    break;
        //                case "原幣進貨金額":
        //                case "原幣期初餘額":
        //                case "原幣入帳金額":
        //                case "原幣出帳金額":
        //                case "原幣期末餘額":
        //                case "原幣存款金額":
        //                case "單位製費成本":
        //                    wsheet.Cell(i + i_col, j + j_row).Style.NumberFormat.Format = "#,##0.00";
        //                    break;
        //                case "匯率":
        //                case "重估匯率":
        //                case "平均匯率":
        //                    wsheet.Cell(i + i_col, j + j_row).Style.NumberFormat.Format = "#,##0.0000";
        //                    break;
        //                default:
        //                    break;
        //            }
        //            wsheet.Cell(i + i_col, j + j_row).Value = row[j];
        //            j++;
        //        }

        //        if (str_SubTotal.Length > 0 && str_Total.Length > 0)
        //        {
        //            str_SubTotal_Name = row[str_SubTotal].ToString().Trim();
        //        }

        //        if ((rows_count_dt - 1) == dt.Rows.IndexOf(row)) //資料列結尾運算
        //        {
        //            if (str_SubTotal.Length == 0 && str_Total.Length > 0)
        //            //if (wsheet.ToString() == "明細帳(評價前)" || wsheet.ToString() == "明細帳(評價後)")
        //            {
        //                i++;
        //                wsheet.Range("A" + (i + i_col) + ":L" + (i + i_col)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
        //                wsheet.Range("A" + (i + i_col) + ":L" + (i + i_col)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
        //                wsheet.Range("A" + (i + i_col) + ":L" + (i + i_col)).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;
        //                wsheet.Cell(i + i_col, col_count_dt - 2).Value = "小計";
        //                wsheet.Cell(i + i_col, j + j_row - 1).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
        //                wsheet.Cell(i + i_col, j + j_row - 1).FormulaA1 = "=sum(L" + i_col + ":L" + (i + i_col - 1) + ")";
        //                wsheet.SheetView.ZoomScale = 80;

        //            }
        //            //if (wsheet.ToString() == "評價表")
        //            if (str_SubTotal.Length > 0 && str_Total.Length > 0)
        //            {
        //                i++;
        //                wsheet.Range("A" + (i + i_col) + ":K" + (i + i_col)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
        //                wsheet.Range("A" + (i + i_col) + ":K" + (i + i_col)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
        //                wsheet.Range("A" + (i + i_col) + ":K" + (i + i_col)).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;
        //                wsheet.Cell(i + i_col, 2).Value = str_SubTotal_Name;
        //                wsheet.Cell(i + i_col, 4).Value = "小計";
        //                wsheet.Range("E" + (i + i_col) + ":K" + (i + i_col)).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
        //                wsheet.Cell(i + i_col, 5).FormulaA1 = "=SUMIFS(E:E,$A:$A,\"" + str_SubTotal_Name + "\")";
        //                wsheet.Cell(i + i_col, 7).FormulaA1 = "=SUMIFS(G:G,$A:$A,\"" + str_SubTotal_Name + "\")";
        //                wsheet.Cell(i + i_col, 8).FormulaA1 = "=SUMIFS(H:H,$A:$A,\"" + str_SubTotal_Name + "\")";
        //                wsheet.Cell(i + i_col, 9).FormulaA1 = "=SUMIFS(I:I,$A:$A,\"" + str_SubTotal_Name + "\")";
        //                wsheet.Cell(i + i_col, 10).FormulaA1 = "=SUMIFS(J:J,$A:$A,\"" + str_SubTotal_Name + "\")";

        //                i++;
        //                wsheet.Range("A" + (i + i_col) + ":K" + (i + i_col)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
        //                wsheet.Range("A" + (i + i_col) + ":K" + (i + i_col)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
        //                wsheet.Range("A" + (i + i_col) + ":K" + (i + i_col)).Style.Fill.BackgroundColor = XLColor.Honeydew;
        //                wsheet.Cell(i + i_col, 4).Value = "總計";
        //                wsheet.Range("E" + (i + i_col) + ":K" + (i + i_col)).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
        //                wsheet.Cell(i + i_col, 5).FormulaA1 = "=SUMIFS(E:E,$D:$D,\"小計\")";
        //                wsheet.Cell(i + i_col, 7).FormulaA1 = "=SUMIFS(G:G,$D:$D,\"小計\")";
        //                wsheet.Cell(i + i_col, 8).FormulaA1 = "=SUMIFS(H:H,$D:$D,\"小計\")";
        //                wsheet.Cell(i + i_col, 9).FormulaA1 = "=SUMIFS(I:I,$D:$D,\"小計\")";
        //                wsheet.Cell(i + i_col, 10).FormulaA1 = "=SUMIFS(J:J,$D:$D,\"小計\")";
        //                wsheet.Cell(i + i_col, 11).FormulaA1 = "=I" + (i + i_col) + "-J" + (i + i_col);

        //            }
        //        }

        //        i++;
        //    }

        //}
        public string loginid = "";
        public string loginName = "", LoginFormName = "", loginDep = "";

     

        public void show_fmlogin_FormName(string data_LoginFormName)
        {
            LoginFormName = data_LoginFormName;
        }

        public string QP_Item = "", QP_Value = "", QP_SQL = "";

        //開航日期
        private void btn_IP_SDate_Click(object sender, EventArgs e)
        {
            this.fm_月曆 = new 月曆(this.txt_IP_SDate, this.btn_IP_ODate, "開航日期");
        }

        private void tspbtn_IP_Build_Click(object sender, EventArgs e)
        {
            panel_IP_Search.Enabled = true;
            tspbtn_IP_Build.Enabled = false;
        }

        private void tspbtn_IP_Save_Click(object sender, EventArgs e)
        {
            panel_IP_Search.Enabled = false;
            tspbtn_IP_Build.Enabled = true;
        }

        private void txt_IP_CustID_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                QP_SQL_Return("客戶代號", "MA001");
            }
           
        }

        private void tspbtn_Order_Add_Click(object sender, EventArgs e)
        {
            SingleQuery.fm_Query_INVPL_Order fm_QueryPublic = new SingleQuery.fm_Query_INVPL_Order();
            fm_QueryPublic.ShowDialog(this);
        }

        //單據日期
        private void btn_IP_ODate_Click(object sender, EventArgs e)
        {
            this.fm_月曆 = new 月曆(this.txt_IP_ODate, this.btn_IP_ODate, "單據日期");
        }

        private void btn_IP_Ship_Click(object sender, EventArgs e)
        {
            QP_SQL_Return("運輸方式代號", "NJ001");
        }

        public Dictionary<string, string> QP_dict_Item = new Dictionary<string, string>();
        public Dictionary<string, string> QP_dict_Result = new Dictionary<string, string>();
        public void show_fm_QueryPublic_QP_Item(Dictionary<string, string> data_QP_dict_Item)
        {
            QP_dict_Item = data_QP_dict_Item;
        }
        public void show_fm_QueryPublic_QP_dict_Result(Dictionary<string, string> data_QP_dict_Result)
        {
            QP_dict_Result = data_QP_dict_Result;
        }
        //public void show_fm_QueryPublic_QP_Item(string data_QP_Item)
        //{
        //    QP_Item = data_QP_Item;
        //}
        public void show_fm_QueryPublic_QP_Value(string data_QP_Value)
        {
            QP_Value = data_QP_Value;
        }
        public void show_fm_QueryPublic_QP_SQL(string data_QP_SQL)
        {
            QP_SQL = data_QP_SQL;
        }

        private void btn_IP_CustID_Click(object sender, EventArgs e)
        {
            QP_SQL_Return("客戶代號", "MA001");
            //            QP_dict_Item.Clear();
            //            QP_Value = txt_IP_CustID.Text.ToString();
            //            QP_dict_Item.Add("客戶代號","MA001");

            //            QP_SQL = @"select MA001 as '客戶代號',MA002 as '客戶簡稱',MA083 as '付款條件代號',NA003 as '付款條件'
            //,MA048 as '運輸方式代號',NJ002 as '運輸方式中文',MA051 as '目的地',MA052 as '海運港口',MA109 as '交易條件代號' ,NK002 as '交易條件名稱' 
            //from COPMA 
            //left join CMSNJ on NJ001 = MA048 
            //left join CMSNA on NA002 = MA083 and NA001 = '2'
            //left join CMSNK on NK001 = MA109 where 1=1";

            //            SingleQuery.fm_QueryPublic fm_QueryPublic = new SingleQuery.fm_QueryPublic();
            //            fm_QueryPublic.show_fm_QueryPublic_QP_Item(QP_dict_Item);
            //            fm_QueryPublic.show_fm_QueryPublic_QP_Value(QP_Value);
            //            fm_QueryPublic.show_fm_QueryPublic_QP_SQL(QP_SQL);
            //            fm_QueryPublic.ShowDialog(this);

            //            //QP_dict_Result
            //            foreach (var OneItem in QP_dict_Result)
            //            {
            //                switch (OneItem.Key)
            //                {
            //                    case "客戶代號":
            //                        txt_IP_CustID.Text = OneItem.Value;
            //                        break;
            //                    case "客戶簡稱":
            //                        lbl_IP_CustID.Text = OneItem.Value;
            //                        break;
            //                    case "付款條件代號":
            //                        txt_IP_Pay.Text = OneItem.Value;
            //                        break;
            //                    case "付款條件":
            //                        lbl_IP_Pay.Text = OneItem.Value;
            //                        break;
            //                    case "運輸方式代號":
            //                        txt_IP_Ship.Text = OneItem.Value;
            //                        break;
            //                    case "運輸方式中文":
            //                        lbl_IP_Ship.Text = OneItem.Value;
            //                        break;
            //                    case "交易條件代號":
            //                        txt_IP_Trade.Text = OneItem.Value;
            //                        break;
            //                    case "交易條件名稱":
            //                        lbl_IP_Trade.Text = OneItem.Value;
            //                        break;
            //                    case "目的地":
            //                        txt_IP_Destn.Text = OneItem.Value;
            //                        break;
            //                    case "海運港口":
            //                        lbl_IP_Trade.Text = OneItem.Value;
            //                        break;
            //                    case "目的港口":
            //                        txt_IP_SD.Text = OneItem.Value;
            //                        break;
            //                    case "出貨港口":
            //                        txt_IP_SO.Text = OneItem.Value;
            //                        break;
            //                }
            //            }

            //txt_IP_CustID.Text = "";
            //lbl_IP_CustID.Text = "";
            //txt_IP_Trade.Text = "";
            //lbl_IP_Trade.Text = "";
            //txt_IP_Pay.Text = "";
            //lbl_IP_Pay.Text = "";
            //txt_IP_Ship.Text = "";
            //lbl_IP_Ship.Text = "";
            //txt_IP_Destn.Text = "";
            //txt_IP_SD.Text = "";
            //txt_IP_SO.Text = "";


        }

        private void btn_IP_FROM_Click(object sender, EventArgs e)
        {
            QP_SQL_Return("出口地全名", "NS003");
        }

        private void btn_IP_TO_Click(object sender, EventArgs e)
        {
            QP_SQL_Return("目的地全名", "NS003");
        }

        private void btn_IP_CITY_Click(object sender, EventArgs e)
        {
            QP_SQL_Return("到貨城市全名", "NS003");
        }




        public void QP_SQL_Return(string str_Name, string str_ID)
        {
            QP_dict_Item.Clear();
            //QP_Value = txt_IP_CustID.Text.ToString();
            QP_dict_Item.Add(str_Name, str_ID);

            switch (str_Name)
            {
                case "客戶代號":
                    QP_Value = txt_IP_CustID.Text.ToString();
                    QP_SQL = @"select TOP(50) MA001 as '客戶代號',MA002 as '客戶簡稱',MA083 as '付款條件代號',NA003 as '付款條件'
,MA048 as '運輸方式代號',NJ003 as '運輸方式英文'
,(select NS001 from CMSNS where NS001 like 'AT%' and NS003 = trim(MA051)) as '目的地代號',MA051 as '目的地全名'
,(select NS001 from CMSNS where NS001 like 'AF%' and NS003 = trim(MA052)) as '出口地代號',MA052 as '出口地全名'
,(select NS001 from CMSNS where NS001 like 'AC%' and NS003 = trim(MA053)) as '到貨城市代號',MA053 as '到貨城市全名'
,MA109 as '交易條件代號' ,NK002 as '交易條件名稱' 
from COPMA 
left join CMSNJ on NJ001 = MA048 
left join CMSNA on NA002 = MA083 and NA001 = '2'
left join CMSNK on NK001 = MA109 where 1=1";
                    break;
                case "運輸方式代號":
                    QP_Value = txt_IP_Ship.Text.ToString();
                    QP_SQL = @"select TOP(50) NJ001 as '運輸方式代號',NJ002 as '運輸方式中文',NJ003 as '運輸方式英文'
from CMSNJ 
where 1=1";
                    break;
                    break;
                case "目的地全名":
                    QP_Value = lbl_IP_TO.Text.ToString();
                    QP_SQL = @"select TOP(50) NS001 as '目的地代號', NS002 as '目的地簡稱', NS003 as '目的地全名'
, NS004 as '國家別代號', (select MR003 from CMSMR where MR001 = '4' and MR002 = NS004) as 國家別名稱
, NS005 as '海空港代號', (case NS005　when '1' then '空運'　when '2' then '海運' END) as '海空港名稱' from CMSNS 
where 1=1 and NS001 like 'AT%'";
                    break;
                case "出口地全名":
                    QP_Value = lbl_IP_FROM.Text.ToString();
                    QP_SQL = @"select TOP(50) NS001 as '出口地代號', NS002 as '出口地簡稱', NS003 as '出口地全名'
, NS004 as '國家別代號', (select MR003 from CMSMR where MR001 = '4' and MR002 = NS004) as 國家別名稱
, NS005 as '海空港代號', (case NS005　when '1' then '空運'　when '2' then '海運' END) as '海空港名稱' from CMSNS 
where 1=1 and NS001 like 'AF%'";
                    break;
                case "到貨城市全名":
                    QP_Value = lbl_IP_CITY.Text.ToString();
                    QP_SQL = @"select TOP(50) NS001 as '到貨城市代號', NS002 as '到貨城市簡稱', NS003 as '到貨城市全名'
, NS004 as '國家別代號', (select MR003 from CMSMR where MR001 = '4' and MR002 = NS004) as 國家別名稱
, NS005 as '海空港代號', (case NS005　when '1' then '空運'　when '2' then '海運' END) as '海空港名稱' from CMSNS 
where 1=1 and NS001 like 'AC%'";
                    break;
            }


            SingleQuery.fm_QueryPublic fm_QueryPublic = new SingleQuery.fm_QueryPublic();
            fm_QueryPublic.show_fm_QueryPublic_QP_Item(QP_dict_Item);
            fm_QueryPublic.show_fm_QueryPublic_QP_Value(QP_Value);
            fm_QueryPublic.show_fm_QueryPublic_QP_SQL(QP_SQL);
            fm_QueryPublic.ShowDialog(this);

            //QP_dict_Result
            foreach (var OneItem in QP_dict_Result)
            {
                switch (OneItem.Key)
                {
                    case "客戶代號":
                        txt_IP_CustID.Text = OneItem.Value;
                        break;
                    case "客戶簡稱":
                        lbl_IP_CustID.Text = OneItem.Value;
                        break;
                    case "付款條件代號":
                        txt_IP_Pay.Text = OneItem.Value;
                        break;
                    case "付款條件":
                        lbl_IP_Pay.Text = OneItem.Value;
                        break;
                    case "運輸方式代號":
                        txt_IP_Ship.Text = OneItem.Value;
                        break;
                    case "運輸方式英文":
                        lbl_IP_Ship.Text = OneItem.Value;
                        break;
                    case "交易條件代號":
                        txt_IP_Trade.Text = OneItem.Value;
                        break;
                    case "交易條件名稱":
                        lbl_IP_Trade.Text = OneItem.Value;
                        break;
                    case "目的地代號":
                        txt_IP_TO.Text = OneItem.Value;
                        break;
                    case "目的地全名":
                        lbl_IP_TO.Text = OneItem.Value;
                        break;
                    case "出口地代號":
                        txt_IP_FROM.Text = OneItem.Value;
                        break;
                    case "出口地全名":
                        lbl_IP_FROM.Text = OneItem.Value;
                        break;
                    //case "海運港口":
                    //    lbl_IP_Trade.Text = OneItem.Value;
                    //    break;
                    case "到貨城市代號":
                        txt_IP_CITY.Text = OneItem.Value;
                        break;
                    case "到貨城市全名":
                        lbl_IP_CITY.Text = OneItem.Value;
                        break;
                    
                }
            }
        }

        bool IsToForm1 = false; //紀錄是否要回到Form1
        protected override void OnClosing(CancelEventArgs e) //在視窗關閉時觸發
        {
            //DialogResult dr = MessageBox.Show("\"是\"回到主畫面 \r\n \"否\"關閉程式", "是否要關閉程式"
            //    , MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            //if (dr == DialogResult.Yes)
            //{
            //IsToForm1 = true;
            ////}
            ////else if (dr == DialogResult.Cancel) 
            ////{

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


    }
}
