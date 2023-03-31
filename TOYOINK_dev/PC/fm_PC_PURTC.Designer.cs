namespace TOYOINK_dev
{
    partial class fm_PC_PURTC
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(fm_PC_PURTC));
            this.lab_status = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.btn_FormDate = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.txterr = new System.Windows.Forms.TextBox();
            this.btn_erpup = new System.Windows.Forms.Button();
            this.btn_toerp = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.lab_loginIDName = new System.Windows.Forms.Label();
            this.txt_TC001 = new System.Windows.Forms.TextBox();
            this.txt_TC002 = new System.Windows.Forms.TextBox();
            this.cbo_ExcelType = new System.Windows.Forms.ComboBox();
            this.label6 = new System.Windows.Forms.Label();
            this.cbo_Currency = new System.Windows.Forms.ComboBox();
            this.label_交易幣別 = new System.Windows.Forms.Label();
            this.cbo_Supplier = new System.Windows.Forms.ComboBox();
            this.label_供應廠商 = new System.Windows.Forms.Label();
            this.btn_NeedDate = new System.Windows.Forms.Button();
            this.txt_NeedDate = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.txt_path = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btn_file = new System.Windows.Forms.Button();
            this.txt_FormDate = new System.Windows.Forms.TextBox();
            this.sqlDataAdapter1 = new System.Data.SqlClient.SqlDataAdapter();
            this.sqlInsertCommand1 = new System.Data.SqlClient.SqlCommand();
            this.sqlConnection1 = new System.Data.SqlClient.SqlConnection();
            this.sqlCommand1 = new System.Data.SqlClient.SqlCommand();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.dgv_excel = new System.Windows.Forms.DataGridView();
            this.tabCtl_data = new System.Windows.Forms.TabControl();
            this.tabPage_excel = new System.Windows.Forms.TabPage();
            this.tabPage_trans = new System.Windows.Forms.TabPage();
            this.dgv_trans = new System.Windows.Forms.DataGridView();
            this.tabPage_tc = new System.Windows.Forms.TabPage();
            this.dgv_tc = new System.Windows.Forms.DataGridView();
            this.tabPage_td = new System.Windows.Forms.TabPage();
            this.dgv_td = new System.Windows.Forms.DataGridView();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_excel)).BeginInit();
            this.tabCtl_data.SuspendLayout();
            this.tabPage_excel.SuspendLayout();
            this.tabPage_trans.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_trans)).BeginInit();
            this.tabPage_tc.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_tc)).BeginInit();
            this.tabPage_td.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_td)).BeginInit();
            this.SuspendLayout();
            // 
            // lab_status
            // 
            this.lab_status.BackColor = System.Drawing.SystemColors.Info;
            this.lab_status.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lab_status.ForeColor = System.Drawing.SystemColors.ControlText;
            this.lab_status.Location = new System.Drawing.Point(565, 54);
            this.lab_status.Name = "lab_status";
            this.lab_status.Size = new System.Drawing.Size(201, 87);
            this.lab_status.TabIndex = 20;
            this.lab_status.Text = " 請先選擇 單據日期";
            this.lab_status.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label2.Location = new System.Drawing.Point(841, 14);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(92, 25);
            this.label2.TabIndex = 23;
            this.label2.Text = "建  立  者";
            // 
            // btn_FormDate
            // 
            this.btn_FormDate.Image = ((System.Drawing.Image)(resources.GetObject("btn_FormDate.Image")));
            this.btn_FormDate.Location = new System.Drawing.Point(223, 101);
            this.btn_FormDate.Name = "btn_FormDate";
            this.btn_FormDate.Size = new System.Drawing.Size(35, 33);
            this.btn_FormDate.TabIndex = 18;
            this.btn_FormDate.Click += new System.EventHandler(this.btn_FormDate_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label5.Location = new System.Drawing.Point(12, 105);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(92, 25);
            this.label5.TabIndex = 15;
            this.label5.Text = "單據日期";
            // 
            // txterr
            // 
            this.txterr.Font = new System.Drawing.Font("微軟正黑體", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txterr.Location = new System.Drawing.Point(933, 53);
            this.txterr.Multiline = true;
            this.txterr.Name = "txterr";
            this.txterr.ReadOnly = true;
            this.txterr.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txterr.Size = new System.Drawing.Size(232, 143);
            this.txterr.TabIndex = 0;
            this.txterr.TextChanged += new System.EventHandler(this.txterr_TextChanged);
            // 
            // btn_erpup
            // 
            this.btn_erpup.BackColor = System.Drawing.SystemColors.Control;
            this.btn_erpup.Enabled = false;
            this.btn_erpup.Font = new System.Drawing.Font("微軟正黑體", 13.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_erpup.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btn_erpup.Location = new System.Drawing.Point(790, 124);
            this.btn_erpup.Name = "btn_erpup";
            this.btn_erpup.Size = new System.Drawing.Size(137, 72);
            this.btn_erpup.TabIndex = 6;
            this.btn_erpup.Text = "上傳至ERP";
            this.btn_erpup.UseVisualStyleBackColor = false;
            this.btn_erpup.Click += new System.EventHandler(this.btn_erpup_Click);
            // 
            // btn_toerp
            // 
            this.btn_toerp.BackColor = System.Drawing.SystemColors.Control;
            this.btn_toerp.Enabled = false;
            this.btn_toerp.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_toerp.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btn_toerp.Location = new System.Drawing.Point(790, 49);
            this.btn_toerp.Name = "btn_toerp";
            this.btn_toerp.Size = new System.Drawing.Size(137, 72);
            this.btn_toerp.TabIndex = 7;
            this.btn_toerp.Text = "轉為ERP格式";
            this.btn_toerp.UseVisualStyleBackColor = false;
            this.btn_toerp.Click += new System.EventHandler(this.btn_toerp_Click);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.lab_loginIDName);
            this.panel1.Controls.Add(this.txt_TC001);
            this.panel1.Controls.Add(this.txterr);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.btn_erpup);
            this.panel1.Controls.Add(this.txt_TC002);
            this.panel1.Controls.Add(this.btn_toerp);
            this.panel1.Controls.Add(this.cbo_ExcelType);
            this.panel1.Controls.Add(this.label6);
            this.panel1.Controls.Add(this.cbo_Currency);
            this.panel1.Controls.Add(this.label_交易幣別);
            this.panel1.Controls.Add(this.cbo_Supplier);
            this.panel1.Controls.Add(this.label_供應廠商);
            this.panel1.Controls.Add(this.btn_NeedDate);
            this.panel1.Controls.Add(this.txt_NeedDate);
            this.panel1.Controls.Add(this.label7);
            this.panel1.Controls.Add(this.lab_status);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.txt_path);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.btn_file);
            this.panel1.Controls.Add(this.btn_FormDate);
            this.panel1.Controls.Add(this.txt_FormDate);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Location = new System.Drawing.Point(16, 12);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1176, 201);
            this.panel1.TabIndex = 17;
            // 
            // lab_loginIDName
            // 
            this.lab_loginIDName.BackColor = System.Drawing.SystemColors.ScrollBar;
            this.lab_loginIDName.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lab_loginIDName.Location = new System.Drawing.Point(933, 12);
            this.lab_loginIDName.Name = "lab_loginIDName";
            this.lab_loginIDName.Size = new System.Drawing.Size(231, 31);
            this.lab_loginIDName.TabIndex = 118;
            this.lab_loginIDName.Text = "0123456 台灣東洋先端";
            this.lab_loginIDName.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txt_TC001
            // 
            this.txt_TC001.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.txt_TC001.Cursor = System.Windows.Forms.Cursors.No;
            this.txt_TC001.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txt_TC001.Location = new System.Drawing.Point(59, 53);
            this.txt_TC001.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txt_TC001.Name = "txt_TC001";
            this.txt_TC001.ReadOnly = true;
            this.txt_TC001.Size = new System.Drawing.Size(58, 34);
            this.txt_TC001.TabIndex = 117;
            this.txt_TC001.Text = "330";
            // 
            // txt_TC002
            // 
            this.txt_TC002.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.txt_TC002.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.txt_TC002.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txt_TC002.Location = new System.Drawing.Point(173, 53);
            this.txt_TC002.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txt_TC002.Name = "txt_TC002";
            this.txt_TC002.Size = new System.Drawing.Size(125, 34);
            this.txt_TC002.TabIndex = 116;
            this.txt_TC002.Text = "200000000";
            // 
            // cbo_ExcelType
            // 
            this.cbo_ExcelType.AutoCompleteCustomSource.AddRange(new string[] {
            "TVS顏料"});
            this.cbo_ExcelType.Enabled = false;
            this.cbo_ExcelType.FormattingEnabled = true;
            this.cbo_ExcelType.Items.AddRange(new object[] {
            "",
            "H10-C4A",
            "H11-C5D.C6C",
            "H14-C5E"});
            this.cbo_ExcelType.Location = new System.Drawing.Point(100, 8);
            this.cbo_ExcelType.Name = "cbo_ExcelType";
            this.cbo_ExcelType.Size = new System.Drawing.Size(364, 37);
            this.cbo_ExcelType.TabIndex = 115;
            // 
            // label6
            // 
            this.label6.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label6.Location = new System.Drawing.Point(9, 12);
            this.label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(99, 31);
            this.label6.TabIndex = 114;
            this.label6.Text = "匯入格式";
            // 
            // cbo_Currency
            // 
            this.cbo_Currency.Enabled = false;
            this.cbo_Currency.FormattingEnabled = true;
            this.cbo_Currency.Items.AddRange(new object[] {
            "",
            "H10-C4A",
            "H11-C5D.C6C",
            "H14-C5E"});
            this.cbo_Currency.Location = new System.Drawing.Point(399, 54);
            this.cbo_Currency.Name = "cbo_Currency";
            this.cbo_Currency.Size = new System.Drawing.Size(131, 37);
            this.cbo_Currency.TabIndex = 113;
            // 
            // label_交易幣別
            // 
            this.label_交易幣別.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label_交易幣別.Location = new System.Drawing.Point(308, 58);
            this.label_交易幣別.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label_交易幣別.Name = "label_交易幣別";
            this.label_交易幣別.Size = new System.Drawing.Size(96, 25);
            this.label_交易幣別.TabIndex = 112;
            this.label_交易幣別.Text = "交易幣別";
            // 
            // cbo_Supplier
            // 
            this.cbo_Supplier.Enabled = false;
            this.cbo_Supplier.FormattingEnabled = true;
            this.cbo_Supplier.Items.AddRange(new object[] {
            "",
            "H10-C4A",
            "H11-C5D.C6C",
            "H14-C5E"});
            this.cbo_Supplier.Location = new System.Drawing.Point(568, 8);
            this.cbo_Supplier.Name = "cbo_Supplier";
            this.cbo_Supplier.Size = new System.Drawing.Size(198, 37);
            this.cbo_Supplier.TabIndex = 111;
            this.cbo_Supplier.SelectedIndexChanged += new System.EventHandler(this.cbo_Supplier_SelectedIndexChanged);
            // 
            // label_供應廠商
            // 
            this.label_供應廠商.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label_供應廠商.Location = new System.Drawing.Point(477, 12);
            this.label_供應廠商.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label_供應廠商.Name = "label_供應廠商";
            this.label_供應廠商.Size = new System.Drawing.Size(99, 31);
            this.label_供應廠商.TabIndex = 110;
            this.label_供應廠商.Text = "供應廠商";
            // 
            // btn_NeedDate
            // 
            this.btn_NeedDate.Image = ((System.Drawing.Image)(resources.GetObject("btn_NeedDate.Image")));
            this.btn_NeedDate.Location = new System.Drawing.Point(519, 101);
            this.btn_NeedDate.Name = "btn_NeedDate";
            this.btn_NeedDate.Size = new System.Drawing.Size(35, 33);
            this.btn_NeedDate.TabIndex = 29;
            this.btn_NeedDate.Click += new System.EventHandler(this.btn_NeedDate_Click);
            // 
            // txt_NeedDate
            // 
            this.txt_NeedDate.BackColor = System.Drawing.Color.LightGoldenrodYellow;
            this.txt_NeedDate.Cursor = System.Windows.Forms.Cursors.No;
            this.txt_NeedDate.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txt_NeedDate.Location = new System.Drawing.Point(399, 100);
            this.txt_NeedDate.Name = "txt_NeedDate";
            this.txt_NeedDate.ReadOnly = true;
            this.txt_NeedDate.Size = new System.Drawing.Size(114, 34);
            this.txt_NeedDate.TabIndex = 30;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label7.Location = new System.Drawing.Point(308, 105);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(92, 25);
            this.label7.TabIndex = 28;
            this.label7.Text = "需求日期";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label3.Location = new System.Drawing.Point(9, 58);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(52, 25);
            this.label3.TabIndex = 11;
            this.label3.Text = "單別";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label4.Location = new System.Drawing.Point(119, 58);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(52, 25);
            this.label4.TabIndex = 12;
            this.label4.Text = "單號";
            // 
            // txt_path
            // 
            this.txt_path.Location = new System.Drawing.Point(99, 155);
            this.txt_path.Name = "txt_path";
            this.txt_path.ReadOnly = true;
            this.txt_path.Size = new System.Drawing.Size(533, 38);
            this.txt_path.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label1.Location = new System.Drawing.Point(7, 160);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(92, 25);
            this.label1.TabIndex = 0;
            this.label1.Text = "檔案路徑";
            // 
            // btn_file
            // 
            this.btn_file.Location = new System.Drawing.Point(638, 154);
            this.btn_file.Name = "btn_file";
            this.btn_file.Size = new System.Drawing.Size(128, 37);
            this.btn_file.TabIndex = 2;
            this.btn_file.Text = "選擇檔案";
            this.btn_file.UseVisualStyleBackColor = true;
            this.btn_file.Click += new System.EventHandler(this.btn_file_Click);
            // 
            // txt_FormDate
            // 
            this.txt_FormDate.BackColor = System.Drawing.Color.LightGoldenrodYellow;
            this.txt_FormDate.Cursor = System.Windows.Forms.Cursors.No;
            this.txt_FormDate.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txt_FormDate.Location = new System.Drawing.Point(103, 100);
            this.txt_FormDate.Name = "txt_FormDate";
            this.txt_FormDate.ReadOnly = true;
            this.txt_FormDate.Size = new System.Drawing.Size(114, 34);
            this.txt_FormDate.TabIndex = 19;
            // 
            // sqlDataAdapter1
            // 
            this.sqlDataAdapter1.InsertCommand = this.sqlInsertCommand1;
            this.sqlDataAdapter1.SelectCommand = this.sqlCommand1;
            this.sqlDataAdapter1.TableMappings.AddRange(new System.Data.Common.DataTableMapping[] {
            new System.Data.Common.DataTableMapping("Table", "CFIPO", new System.Data.Common.DataColumnMapping[] {
                        new System.Data.Common.DataColumnMapping("線別", "線別"),
                        new System.Data.Common.DataColumnMapping("Sample", "Sample"),
                        new System.Data.Common.DataColumnMapping("Number", "Number"),
                        new System.Data.Common.DataColumnMapping("Item", "Item"),
                        new System.Data.Common.DataColumnMapping("Item Description", "Item Description"),
                        new System.Data.Common.DataColumnMapping("UOM", "UOM"),
                        new System.Data.Common.DataColumnMapping("Shipment Amount", "Shipment Amount"),
                        new System.Data.Common.DataColumnMapping("Quantity", "Quantity"),
                        new System.Data.Common.DataColumnMapping("Supplier", "Supplier"),
                        new System.Data.Common.DataColumnMapping("Supplier Site", "Supplier Site"),
                        new System.Data.Common.DataColumnMapping("Currency", "Currency"),
                        new System.Data.Common.DataColumnMapping("Buyer", "Buyer"),
                        new System.Data.Common.DataColumnMapping("Needdate", "Needdate"),
                        new System.Data.Common.DataColumnMapping("備註", "備註"),
                        new System.Data.Common.DataColumnMapping("序號", "序號")})});
            // 
            // sqlInsertCommand1
            // 
            this.sqlInsertCommand1.CommandText = resources.GetString("sqlInsertCommand1.CommandText");
            this.sqlInsertCommand1.Connection = this.sqlConnection1;
            this.sqlInsertCommand1.Parameters.AddRange(new System.Data.SqlClient.SqlParameter[] {
            new System.Data.SqlClient.SqlParameter("@線別", System.Data.SqlDbType.NVarChar, 0, "線別"),
            new System.Data.SqlClient.SqlParameter("@Sample", System.Data.SqlDbType.NVarChar, 0, "Sample"),
            new System.Data.SqlClient.SqlParameter("@Number", System.Data.SqlDbType.NVarChar, 0, "Number"),
            new System.Data.SqlClient.SqlParameter("@Item", System.Data.SqlDbType.NVarChar, 0, "Item"),
            new System.Data.SqlClient.SqlParameter("@Item_Description", System.Data.SqlDbType.NVarChar, 0, "Item Description"),
            new System.Data.SqlClient.SqlParameter("@UOM", System.Data.SqlDbType.NVarChar, 0, "UOM"),
            new System.Data.SqlClient.SqlParameter("@Shipment_Amount", System.Data.SqlDbType.Decimal, 0, System.Data.ParameterDirection.Input, false, ((byte)(18)), ((byte)(2)), "Shipment Amount", System.Data.DataRowVersion.Current, null),
            new System.Data.SqlClient.SqlParameter("@Quantity", System.Data.SqlDbType.Decimal, 0, System.Data.ParameterDirection.Input, false, ((byte)(18)), ((byte)(2)), "Quantity", System.Data.DataRowVersion.Current, null),
            new System.Data.SqlClient.SqlParameter("@Supplier", System.Data.SqlDbType.NVarChar, 0, "Supplier"),
            new System.Data.SqlClient.SqlParameter("@Supplier_Site", System.Data.SqlDbType.NVarChar, 0, "Supplier Site"),
            new System.Data.SqlClient.SqlParameter("@Currency", System.Data.SqlDbType.NVarChar, 0, "Currency"),
            new System.Data.SqlClient.SqlParameter("@Buyer", System.Data.SqlDbType.NVarChar, 0, "Buyer"),
            new System.Data.SqlClient.SqlParameter("@Needdate", System.Data.SqlDbType.NVarChar, 0, "Needdate"),
            new System.Data.SqlClient.SqlParameter("@備註", System.Data.SqlDbType.NVarChar, 0, "備註"),
            new System.Data.SqlClient.SqlParameter("@序號", System.Data.SqlDbType.NVarChar, 0, "序號")});
            // 
            // sqlConnection1
            // 
            this.sqlConnection1.ConnectionString = "Data Source=192.168.128.219;Initial Catalog=A01A;User ID=pwuser;Password=sqlmis00" +
    "3";
            this.sqlConnection1.FireInfoMessageEventOnUserErrors = false;
            // 
            // sqlCommand1
            // 
            this.sqlCommand1.CommandText = "SELECT  *\r\n  FROM CFIPO";
            this.sqlCommand1.Connection = this.sqlConnection1;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // dgv_excel
            // 
            this.dgv_excel.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dgv_excel.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_excel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgv_excel.Location = new System.Drawing.Point(0, 0);
            this.dgv_excel.Name = "dgv_excel";
            this.dgv_excel.ReadOnly = true;
            this.dgv_excel.RowHeadersWidth = 51;
            this.dgv_excel.RowTemplate.Height = 27;
            this.dgv_excel.Size = new System.Drawing.Size(1177, 495);
            this.dgv_excel.TabIndex = 4;
            // 
            // tabCtl_data
            // 
            this.tabCtl_data.Controls.Add(this.tabPage_excel);
            this.tabCtl_data.Controls.Add(this.tabPage_trans);
            this.tabCtl_data.Controls.Add(this.tabPage_tc);
            this.tabCtl_data.Controls.Add(this.tabPage_td);
            this.tabCtl_data.Location = new System.Drawing.Point(12, 221);
            this.tabCtl_data.Name = "tabCtl_data";
            this.tabCtl_data.SelectedIndex = 0;
            this.tabCtl_data.Size = new System.Drawing.Size(1185, 537);
            this.tabCtl_data.TabIndex = 16;
            // 
            // tabPage_excel
            // 
            this.tabPage_excel.Controls.Add(this.dgv_excel);
            this.tabPage_excel.Location = new System.Drawing.Point(4, 38);
            this.tabPage_excel.Name = "tabPage_excel";
            this.tabPage_excel.Size = new System.Drawing.Size(1177, 495);
            this.tabPage_excel.TabIndex = 0;
            this.tabPage_excel.Text = "來源Excel";
            this.tabPage_excel.UseVisualStyleBackColor = true;
            // 
            // tabPage_trans
            // 
            this.tabPage_trans.Controls.Add(this.dgv_trans);
            this.tabPage_trans.Location = new System.Drawing.Point(4, 38);
            this.tabPage_trans.Name = "tabPage_trans";
            this.tabPage_trans.Size = new System.Drawing.Size(1177, 495);
            this.tabPage_trans.TabIndex = 3;
            this.tabPage_trans.Text = "整理結果";
            this.tabPage_trans.UseVisualStyleBackColor = true;
            // 
            // dgv_trans
            // 
            this.dgv_trans.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dgv_trans.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.dgv_trans.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_trans.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgv_trans.Location = new System.Drawing.Point(0, 0);
            this.dgv_trans.Name = "dgv_trans";
            this.dgv_trans.ReadOnly = true;
            this.dgv_trans.RowHeadersWidth = 51;
            this.dgv_trans.RowTemplate.Height = 27;
            this.dgv_trans.Size = new System.Drawing.Size(1177, 495);
            this.dgv_trans.TabIndex = 0;
            // 
            // tabPage_tc
            // 
            this.tabPage_tc.Controls.Add(this.dgv_tc);
            this.tabPage_tc.Location = new System.Drawing.Point(4, 38);
            this.tabPage_tc.Name = "tabPage_tc";
            this.tabPage_tc.Size = new System.Drawing.Size(1177, 495);
            this.tabPage_tc.TabIndex = 1;
            this.tabPage_tc.Text = "單頭 PURTC";
            this.tabPage_tc.UseVisualStyleBackColor = true;
            // 
            // dgv_tc
            // 
            this.dgv_tc.AllowUserToDeleteRows = false;
            this.dgv_tc.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dgv_tc.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_tc.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgv_tc.Location = new System.Drawing.Point(0, 0);
            this.dgv_tc.Name = "dgv_tc";
            this.dgv_tc.ReadOnly = true;
            this.dgv_tc.RowHeadersWidth = 51;
            this.dgv_tc.RowTemplate.Height = 27;
            this.dgv_tc.Size = new System.Drawing.Size(1177, 495);
            this.dgv_tc.TabIndex = 4;
            // 
            // tabPage_td
            // 
            this.tabPage_td.Controls.Add(this.dgv_td);
            this.tabPage_td.Location = new System.Drawing.Point(4, 38);
            this.tabPage_td.Name = "tabPage_td";
            this.tabPage_td.Size = new System.Drawing.Size(1177, 495);
            this.tabPage_td.TabIndex = 2;
            this.tabPage_td.Text = "單身 PURTD";
            this.tabPage_td.UseVisualStyleBackColor = true;
            // 
            // dgv_td
            // 
            this.dgv_td.AllowUserToDeleteRows = false;
            this.dgv_td.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dgv_td.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_td.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgv_td.Location = new System.Drawing.Point(0, 0);
            this.dgv_td.Name = "dgv_td";
            this.dgv_td.ReadOnly = true;
            this.dgv_td.RowHeadersWidth = 51;
            this.dgv_td.RowTemplate.Height = 27;
            this.dgv_td.Size = new System.Drawing.Size(1177, 495);
            this.dgv_td.TabIndex = 5;
            // 
            // fm_PC_PURTC
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(14F, 29F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1204, 770);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.tabCtl_data);
            this.Font = new System.Drawing.Font("微軟正黑體", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(5);
            this.Name = "fm_PC_PURTC";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "採購單匯入 顏料分散體 For TVS (20220914 1350)";
            this.Load += new System.EventHandler(this.fm_PC_PURTC_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_excel)).EndInit();
            this.tabCtl_data.ResumeLayout(false);
            this.tabPage_excel.ResumeLayout(false);
            this.tabPage_trans.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgv_trans)).EndInit();
            this.tabPage_tc.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgv_tc)).EndInit();
            this.tabPage_td.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgv_td)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Label lab_status;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btn_FormDate;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txterr;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txt_path;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btn_file;
        private System.Windows.Forms.Button btn_toerp;
        private System.Windows.Forms.TextBox txt_FormDate;
        private System.Windows.Forms.Button btn_erpup;
        public System.Data.SqlClient.SqlDataAdapter sqlDataAdapter1;
        private System.Data.SqlClient.SqlCommand sqlInsertCommand1;
        public System.Data.SqlClient.SqlConnection sqlConnection1;
        private System.Data.SqlClient.SqlCommand sqlCommand1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.DataGridView dgv_excel;
        private System.Windows.Forms.TabControl tabCtl_data;
        private System.Windows.Forms.TabPage tabPage_excel;
        private System.Windows.Forms.TabPage tabPage_trans;
        private System.Windows.Forms.DataGridView dgv_trans;
        private System.Windows.Forms.TabPage tabPage_tc;
        private System.Windows.Forms.DataGridView dgv_tc;
        private System.Windows.Forms.TabPage tabPage_td;
        private System.Windows.Forms.DataGridView dgv_td;
        private System.Windows.Forms.Button btn_NeedDate;
        private System.Windows.Forms.TextBox txt_NeedDate;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ComboBox cbo_Supplier;
        private System.Windows.Forms.Label label_供應廠商;
        private System.Windows.Forms.ComboBox cbo_Currency;
        private System.Windows.Forms.Label label_交易幣別;
        private System.Windows.Forms.ComboBox cbo_ExcelType;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txt_TC001;
        private System.Windows.Forms.TextBox txt_TC002;
        private System.Windows.Forms.Label lab_loginIDName;
    }
}