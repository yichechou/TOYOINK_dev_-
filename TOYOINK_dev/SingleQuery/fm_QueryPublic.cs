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

namespace TOYOINK_dev.SingleQuery
{
    public partial class fm_QueryPublic : Form
    {
        public MyClass MyCode;
        public fm_QueryPublic()
        {
            InitializeComponent();

            MyCode = new Myclass.MyClass();
            MyCode.strDbCon = MyCode.strDbConLeader;
            //this.sqlConnection1.ConnectionString = MyCode.strDbConLeader;

            //MyCode.strDbCon = MyCode.strDbConA01A;
            //this.sqlConnection1.ConnectionString = MyCode.strDbConA01A;
        }

        //接收fm_Acc_INVPK資料，並顯示
        public string str_QP_Cond = "", str_QP_Value = "";
        public string QP_ItemKey = "", QP_ItemValue = "", QP_Value = "", QP_SQL = "";
        public Dictionary<string, string> QP_dict_Item = new Dictionary<string, string>();
        public Dictionary<string, string> QP_dict_Result { get; set; } = new Dictionary<string, string>();
        public Dictionary<string, string> QP_dict_Result_Temp { get; set; } = new Dictionary<string, string>();
        public void show_fm_QueryPublic_QP_Item(Dictionary<string, string> data_QP_dict_Item)
        {
            QP_dict_Item = data_QP_dict_Item;
        }

        private void btn_Save_Click(object sender, EventArgs e)
        {
            QP_dict_Result = QP_dict_Result_Temp;
            ACC.fm_Acc_INVPL fm_Acc_INVPL = (ACC.fm_Acc_INVPL)this.Owner;
            fm_Acc_INVPL.show_fm_QueryPublic_QP_dict_Result(QP_dict_Result);
            //Application.Exit();
            this.Close();
        }

        private void dgv_Result_RowHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            QP_dict_Result_Temp.Clear();
            int a = dgv_Result.CurrentRow.Index;
            int i = 0;

            foreach (DataGridViewColumn dgv_Result_col in dgv_Result.Columns)
            {
                QP_dict_Result_Temp.Add(dgv_Result.Columns[i].HeaderText.ToString(),dgv_Result.Rows[a].Cells[i].Value.ToString().Trim());
                i += 1;
            }
            QP_dict_Result = QP_dict_Result_Temp;
        }

        public void show_fm_QueryPublic_QP_dict_Result(Dictionary<string, string> data_QP_dict_Result)
        {
            QP_dict_Result = data_QP_dict_Result;
        }
        public void show_fm_QueryPublic_QP_Value(string data_QP_Value)
        {
            QP_Value = data_QP_Value;
        }

        public void show_fm_QueryPublic_QP_SQL(string data_QP_SQL)
        {
            QP_SQL = data_QP_SQL;
        }

        private void fm_QueryPublic_Load(object sender, EventArgs e)
        {
            foreach (var OneItem in QP_dict_Item)
            {
                cbo_Item.Items.Add(OneItem.Key);
                //Console.WriteLine("Key = " + OneItem.Key + ", Value = " + OneItem.Value);
            }
            cbo_Item.SelectedIndex = 0;
            txt_Value.Text = QP_Value;

            //查詢結果
            DataTable dt_QP_Result = new DataTable();

            if (QP_dict_Item.ContainsKey(cbo_Item.SelectedItem.ToString()))
            {
                QP_ItemValue = String.Format(@" and {0} {1} '{2}'"
                                , QP_dict_Item[cbo_Item.SelectedItem.ToString()],cbo_Cond.Text.ToString(), txt_Value.Text.ToString());
            }

            string sql_QP_Result = QP_SQL + QP_ItemValue;

            MyCode.Sql_dgv(sql_QP_Result, dt_QP_Result, dgv_Result);
            dgv_Result.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
        }

        private void btn_Search_Click(object sender, EventArgs e)
        {
            //查詢結果
            DataTable dt_QP_Result = new DataTable();

            switch (cbo_Cond.Text.ToString())
            {
                case ">=":
                    str_QP_Cond = cbo_Cond.Text.ToString();
                    str_QP_Value = txt_Value.Text.ToString();
                    break;
                case "<=":
                    str_QP_Cond = cbo_Cond.Text.ToString();
                    str_QP_Value = txt_Value.Text.ToString();
                    break;
                case "=":
                    str_QP_Cond = cbo_Cond.Text.ToString();
                    str_QP_Value = txt_Value.Text.ToString();
                    break;
                case "like%":
                    str_QP_Cond = "like";
                    str_QP_Value = txt_Value.Text.ToString() + "%";
                    break;
                case "%like":
                    str_QP_Cond = "like";
                    str_QP_Value = "%" + txt_Value.Text.ToString();
                    break;
                case "%like%":
                    str_QP_Cond = "like";
                    str_QP_Value = "%" + txt_Value.Text.ToString() + "%";
                    break;
            }

            if (QP_dict_Item.ContainsKey(cbo_Item.SelectedItem.ToString()))
            {
                QP_ItemValue = String.Format(@" and {0} {1} '{2}'"
                                , QP_dict_Item[cbo_Item.SelectedItem.ToString()], str_QP_Cond, str_QP_Value);
            }

            string sql_QP_Result = QP_SQL + QP_ItemValue;

            MyCode.Sql_dgv(sql_QP_Result, dt_QP_Result, dgv_Result);
            dgv_Result.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
        }

        //protected override void OnClosing(CancelEventArgs e) //在視窗關閉時觸發
        //{
        //    Environment.Exit(Environment.ExitCode);
        //}
    }
}
