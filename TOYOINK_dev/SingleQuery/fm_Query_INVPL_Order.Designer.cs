namespace TOYOINK_dev.SingleQuery
{
    partial class fm_Query_INVPL_Order
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(fm_Query_INVPL_Order));
            this.panel1 = new System.Windows.Forms.Panel();
            this.txt_Value3 = new System.Windows.Forms.TextBox();
            this.cbo_Cond3 = new System.Windows.Forms.ComboBox();
            this.cbo_Item3 = new System.Windows.Forms.ComboBox();
            this.txt_Value2 = new System.Windows.Forms.TextBox();
            this.cbo_Cond2 = new System.Windows.Forms.ComboBox();
            this.cbo_Item2 = new System.Windows.Forms.ComboBox();
            this.btn_Search = new System.Windows.Forms.Button();
            this.txt_Value = new System.Windows.Forms.TextBox();
            this.cbo_Cond = new System.Windows.Forms.ComboBox();
            this.cbo_Item = new System.Windows.Forms.ComboBox();
            this.panel2 = new System.Windows.Forms.Panel();
            this.dgv_Result = new System.Windows.Forms.DataGridView();
            this.btn_Save = new System.Windows.Forms.Button();
            this.btn_Exit = new System.Windows.Forms.Button();
            this.btn_Clear = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_Result)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.btn_Clear);
            this.panel1.Controls.Add(this.txt_Value3);
            this.panel1.Controls.Add(this.cbo_Cond3);
            this.panel1.Controls.Add(this.cbo_Item3);
            this.panel1.Controls.Add(this.txt_Value2);
            this.panel1.Controls.Add(this.cbo_Cond2);
            this.panel1.Controls.Add(this.cbo_Item2);
            this.panel1.Controls.Add(this.btn_Search);
            this.panel1.Controls.Add(this.txt_Value);
            this.panel1.Controls.Add(this.cbo_Cond);
            this.panel1.Controls.Add(this.cbo_Item);
            this.panel1.Location = new System.Drawing.Point(3, 3);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(693, 121);
            this.panel1.TabIndex = 4;
            // 
            // txt_Value3
            // 
            this.txt_Value3.Location = new System.Drawing.Point(272, 81);
            this.txt_Value3.Name = "txt_Value3";
            this.txt_Value3.Size = new System.Drawing.Size(288, 34);
            this.txt_Value3.TabIndex = 9;
            // 
            // cbo_Cond3
            // 
            this.cbo_Cond3.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.cbo_Cond3.FormattingEnabled = true;
            this.cbo_Cond3.Items.AddRange(new object[] {
            "=",
            ">=",
            "<=",
            "like%",
            "%like",
            "%like%"});
            this.cbo_Cond3.Location = new System.Drawing.Point(168, 81);
            this.cbo_Cond3.Name = "cbo_Cond3";
            this.cbo_Cond3.Size = new System.Drawing.Size(98, 33);
            this.cbo_Cond3.TabIndex = 8;
            this.cbo_Cond3.Text = "=";
            // 
            // cbo_Item3
            // 
            this.cbo_Item3.FormattingEnabled = true;
            this.cbo_Item3.Location = new System.Drawing.Point(5, 81);
            this.cbo_Item3.Name = "cbo_Item3";
            this.cbo_Item3.Size = new System.Drawing.Size(157, 33);
            this.cbo_Item3.TabIndex = 7;
            this.cbo_Item3.Text = "單據日期";
            // 
            // txt_Value2
            // 
            this.txt_Value2.Location = new System.Drawing.Point(272, 42);
            this.txt_Value2.Name = "txt_Value2";
            this.txt_Value2.Size = new System.Drawing.Size(288, 34);
            this.txt_Value2.TabIndex = 6;
            // 
            // cbo_Cond2
            // 
            this.cbo_Cond2.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.cbo_Cond2.FormattingEnabled = true;
            this.cbo_Cond2.Items.AddRange(new object[] {
            "=",
            ">=",
            "<=",
            "like%",
            "%like",
            "%like%"});
            this.cbo_Cond2.Location = new System.Drawing.Point(168, 42);
            this.cbo_Cond2.Name = "cbo_Cond2";
            this.cbo_Cond2.Size = new System.Drawing.Size(98, 33);
            this.cbo_Cond2.TabIndex = 5;
            this.cbo_Cond2.Text = "=";
            // 
            // cbo_Item2
            // 
            this.cbo_Item2.FormattingEnabled = true;
            this.cbo_Item2.Location = new System.Drawing.Point(5, 42);
            this.cbo_Item2.Name = "cbo_Item2";
            this.cbo_Item2.Size = new System.Drawing.Size(157, 33);
            this.cbo_Item2.TabIndex = 4;
            this.cbo_Item2.Text = "客戶代號";
            // 
            // btn_Search
            // 
            this.btn_Search.Location = new System.Drawing.Point(568, 3);
            this.btn_Search.Name = "btn_Search";
            this.btn_Search.Size = new System.Drawing.Size(115, 72);
            this.btn_Search.TabIndex = 3;
            this.btn_Search.Text = "查詢";
            this.btn_Search.UseVisualStyleBackColor = true;
            this.btn_Search.Click += new System.EventHandler(this.btn_Search_Click);
            // 
            // txt_Value
            // 
            this.txt_Value.Location = new System.Drawing.Point(272, 3);
            this.txt_Value.Name = "txt_Value";
            this.txt_Value.Size = new System.Drawing.Size(288, 34);
            this.txt_Value.TabIndex = 2;
            // 
            // cbo_Cond
            // 
            this.cbo_Cond.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.cbo_Cond.FormattingEnabled = true;
            this.cbo_Cond.Items.AddRange(new object[] {
            "=",
            ">=",
            "<=",
            "like%",
            "%like",
            "%like%"});
            this.cbo_Cond.Location = new System.Drawing.Point(168, 3);
            this.cbo_Cond.Name = "cbo_Cond";
            this.cbo_Cond.Size = new System.Drawing.Size(98, 33);
            this.cbo_Cond.TabIndex = 1;
            this.cbo_Cond.Text = "=";
            // 
            // cbo_Item
            // 
            this.cbo_Item.FormattingEnabled = true;
            this.cbo_Item.Location = new System.Drawing.Point(5, 3);
            this.cbo_Item.Name = "cbo_Item";
            this.cbo_Item.Size = new System.Drawing.Size(157, 33);
            this.cbo_Item.TabIndex = 0;
            this.cbo_Item.Text = "客戶單號";
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.dgv_Result);
            this.panel2.Location = new System.Drawing.Point(3, 124);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(694, 341);
            this.panel2.TabIndex = 5;
            // 
            // dgv_Result
            // 
            this.dgv_Result.AllowUserToAddRows = false;
            this.dgv_Result.AllowUserToDeleteRows = false;
            this.dgv_Result.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dgv_Result.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_Result.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgv_Result.Location = new System.Drawing.Point(0, 0);
            this.dgv_Result.Name = "dgv_Result";
            this.dgv_Result.ReadOnly = true;
            this.dgv_Result.RowHeadersWidth = 51;
            this.dgv_Result.RowTemplate.Height = 27;
            this.dgv_Result.Size = new System.Drawing.Size(694, 341);
            this.dgv_Result.TabIndex = 2;
            // 
            // btn_Save
            // 
            this.btn_Save.Image = ((System.Drawing.Image)(resources.GetObject("btn_Save.Image")));
            this.btn_Save.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btn_Save.Location = new System.Drawing.Point(198, 471);
            this.btn_Save.Name = "btn_Save";
            this.btn_Save.Size = new System.Drawing.Size(85, 37);
            this.btn_Save.TabIndex = 7;
            this.btn_Save.Text = "確認";
            this.btn_Save.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btn_Save.UseVisualStyleBackColor = true;
            // 
            // btn_Exit
            // 
            this.btn_Exit.Image = ((System.Drawing.Image)(resources.GetObject("btn_Exit.Image")));
            this.btn_Exit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btn_Exit.Location = new System.Drawing.Point(393, 471);
            this.btn_Exit.Name = "btn_Exit";
            this.btn_Exit.Size = new System.Drawing.Size(85, 37);
            this.btn_Exit.TabIndex = 6;
            this.btn_Exit.Text = "取消";
            this.btn_Exit.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btn_Exit.UseVisualStyleBackColor = true;
            // 
            // btn_Clear
            // 
            this.btn_Clear.Location = new System.Drawing.Point(568, 79);
            this.btn_Clear.Name = "btn_Clear";
            this.btn_Clear.Size = new System.Drawing.Size(115, 34);
            this.btn_Clear.TabIndex = 10;
            this.btn_Clear.Text = "清除欄位";
            this.btn_Clear.UseVisualStyleBackColor = true;
            this.btn_Clear.Click += new System.EventHandler(this.btn_Clear_Click);
            // 
            // fm_Query_INVPL_Order
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(705, 513);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.btn_Save);
            this.Controls.Add(this.btn_Exit);
            this.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "fm_Query_INVPL_Order";
            this.Text = "客戶訂單查詢";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgv_Result)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.TextBox txt_Value3;
        private System.Windows.Forms.ComboBox cbo_Cond3;
        private System.Windows.Forms.ComboBox cbo_Item3;
        private System.Windows.Forms.TextBox txt_Value2;
        private System.Windows.Forms.ComboBox cbo_Cond2;
        private System.Windows.Forms.ComboBox cbo_Item2;
        private System.Windows.Forms.Button btn_Search;
        private System.Windows.Forms.TextBox txt_Value;
        private System.Windows.Forms.ComboBox cbo_Cond;
        private System.Windows.Forms.ComboBox cbo_Item;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.DataGridView dgv_Result;
        private System.Windows.Forms.Button btn_Save;
        private System.Windows.Forms.Button btn_Exit;
        private System.Windows.Forms.Button btn_Clear;
    }
}