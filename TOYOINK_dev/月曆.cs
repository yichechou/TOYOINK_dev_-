using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Windows.Threading;
using System.Globalization;

namespace TOYOINK_dev
{
	/// <summary>
	/// ��� ���K�n�y�z�C
	/// </summary>
	public class ��� : System.Windows.Forms.Form
	{
		private System.Windows.Forms.MonthCalendar monthCalendar1;
		private System.ComponentModel.IContainer components;
		private System.Windows.Forms.Timer timer1;
		private TextBox f_textbox;
		private Button f_button;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Button button1;
        string str_title;

		public ���()
		{
			//
			// Windows Form �]�p�u��䴩�����n��
			//
			InitializeComponent();

            //
            // TODO: �b InitializeComponent �I�s����[�J����غc�禡�{���X
            //
        }

		public ���(TextBox f_textbox,Button f_button,string str_title)
		{
			//
			// Windows Form �]�p�u��䴩�����n��
			//
			InitializeComponent();

			//
			// TODO: �b InitializeComponent �I�s����[�J����غc�禡�{���X
			//
			this.f_textbox = f_textbox;
			this.f_button = f_button;
			this.str_title = str_title;
			this.Show();
		}

		/// <summary>
		/// �M������ϥΤ����귽�C
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form �]�p�u�㲣�ͪ��{���X
		/// <summary>
		/// �����]�p�u��䴩�ҥ�������k - �ФŨϥε{���X�s�边�ק�
		/// �o�Ӥ�k�����e�C
		/// </summary>
		private void InitializeComponent()
		{
            this.components = new System.ComponentModel.Container();
            this.monthCalendar1 = new System.Windows.Forms.MonthCalendar();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.label1 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // monthCalendar1
            // 
            this.monthCalendar1.Font = new System.Drawing.Font("�L�n������", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.monthCalendar1.Location = new System.Drawing.Point(13, 6);
            this.monthCalendar1.Name = "monthCalendar1";
            this.monthCalendar1.TabIndex = 0;
            this.monthCalendar1.DateChanged += new System.Windows.Forms.DateRangeEventHandler(this.monthCalendar1_DateChanged);
            this.monthCalendar1.MouseDown += new System.Windows.Forms.MouseEventHandler(this.monthCalendar1_MouseDown);
            // 
            // timer1
            // 
            this.timer1.Tick += new System.EventHandler(this.Timer1_Tick);
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("�L�n������", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label1.Location = new System.Drawing.Point(1, 210);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(211, 39);
            this.label1.TabIndex = 1;
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("�L�n������", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.button1.Location = new System.Drawing.Point(219, 215);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(59, 29);
            this.button1.TabIndex = 2;
            this.button1.Text = "OK";
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // ���
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(8, 20);
            this.ClientSize = new System.Drawing.Size(285, 255);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.monthCalendar1);
            this.Font = new System.Drawing.Font("�L�n������", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Name = "���";
            this.Text = "���";
            this.Load += new System.EventHandler(this.���_Load);
            this.ResumeLayout(false);

		}
		#endregion

		private void monthCalendar1_DateChanged(object sender, System.Windows.Forms.DateRangeEventArgs e)
		{
			this.label1.Text = e.Start.ToShortDateString();
		}

        private void Timer1_Tick(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void ���_Load(object sender, System.EventArgs e)
		{
            this.timer1.Interval = 100;
			this.Text = this.str_title;
            if (f_textbox.Text.Trim() == "") 
            {
                f_textbox.Text = DateTime.Now.ToString("yyyyMMdd");
            }
            DateTime time = DateTime.ParseExact(f_textbox.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);
            monthCalendar1.SelectionStart = DateTime.ParseExact(f_textbox.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);

            this.label1.Text = time.ToShortDateString();
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			if(this.label1.Text.Length == 0)
			{
				return;
			}
			string str_year = "",str_month = "",str_day="",str_short;
			str_short = this.label1.Text;
			for(int i = 0; i<str_short.Length;i++)
			{
				if(str_short[i] != '/' )
				{
					str_year += str_short[i];
				}
				else
				{
					break;
				}
			}

			for(int i = 1; i<=2 ; i++)
			{
				if(str_short[str_year.Length+i] != '/')
				{
					str_month += str_short[str_year.Length+i];
				}
				else
				{
					break;
				}

			}

			while(str_month.Length <2)
			{
				str_month = "0" + str_month;
			}

			for(int i = 1; i<=2;i++)
			{
				if(str_short[str_short.Length -i] != '/')
				{
					str_day = str_short[str_short.Length -i] + str_day;
				}
				else
				{
					break;
				}

			}

			while(str_day.Length <2)
			{
				str_day = "0" + str_day;
			}
			this.f_textbox.Text = str_year + str_month + str_day;
			this.f_button.Enabled = true;
		  
			this.timer1.Enabled = true;
		}

        int i = 0;

        private void monthCalendar1_MouseDown(object sender, MouseEventArgs e)
        {
            i += 1;

            DispatcherTimer timer = new DispatcherTimer();
            timer.Interval = new TimeSpan(0, 0, 0, 0, 300);
            timer.Tick += (s, e1) => { timer.IsEnabled = false; i = 0; };
            timer.IsEnabled = true;

            if (i % 2 == 0)
            {
                timer.IsEnabled = false;
                i = 0;

                if (this.label1.Text.Length == 0)
                {
                    return;
                }
                string str_year = "", str_month = "", str_day = "", str_short;
                str_short = this.label1.Text;
                for (int i = 0; i < str_short.Length; i++)
                {
                    if (str_short[i] != '/')
                    {
                        str_year += str_short[i];
                    }
                    else
                    {
                        break;
                    }
                }

                for (int i = 1; i <= 2; i++)
                {
                    if (str_short[str_year.Length + i] != '/')
                    {
                        str_month += str_short[str_year.Length + i];
                    }
                    else
                    {
                        break;
                    }
                }

                while (str_month.Length < 2)
                {
                    str_month = "0" + str_month;
                }

                for (int i = 1; i <= 2; i++)
                {
                    if (str_short[str_short.Length - i] != '/')
                    {
                        str_day = str_short[str_short.Length - i] + str_day;
                    }
                    else
                    {
                        break;
                    }
                }

                while (str_day.Length < 2)
                {
                    str_day = "0" + str_day;
                }
                this.f_textbox.Text = str_year + str_month + str_day;
                this.f_button.Enabled = true;

                this.timer1.Enabled = true;
                this.Close();
            }
        }
    }
}
