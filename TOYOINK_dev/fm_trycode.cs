using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TOYOINK_dev
{
    public partial class fm_trycode : Form
    {
        public fm_trycode()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DataTable dtA = new DataTable();
            dtA.Columns.Add("id", typeof(int));
            dtA.Columns.Add("price", typeof(string));
            dtA.Rows.Add(1, "111");
            dtA.Rows.Add(2, "222");
            dtA.Rows.Add(3, "333");
            dtA.Rows.Add(4, "444");
            dtA.Rows.Add(5, "555");

            DataTable dtB = dtA.Clone();
            dtB.Rows.Add(1, "121");
            dtB.Rows.Add(2, "221");
            dtB.Rows.Add(3, "331");

            DataTable dtC = dtA.Clone();
            dtC.Columns.Add("price_excel");

            var query = from a in dtA.AsEnumerable()
                        join b in dtB.AsEnumerable()
                        on a.Field<int>("id") equals b.Field<int>("id") into g
                        from b in g.DefaultIfEmpty()
                        select new
                        {
                            id = a.Field<int>("id"),
                            price = a.Field<string>("price"),
                            price_excel = b == null ? "None" : b.Field<string>("price")
                        };

            query.ToList().ForEach(q => dtC.Rows.Add(q.id, q.price, q.price_excel));
            dataGridView1.DataSource = dtC;
        }
    }
}
