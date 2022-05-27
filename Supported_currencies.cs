using System;
using System.Windows.Forms;

namespace Exchangerate
{
    public partial class Supported_currencies : Form
    {
        Form1 frm1 = new Form1();
        public Supported_currencies()
        {
            InitializeComponent();
            Form1 frm1 = new Form1();

            dynamic d_curr_list = frm1.get_all_supported_currencies();

            int i = 1;
            foreach (var s in d_curr_list["symbols"])
            {
                dataGridView1.Rows.Add(i.ToString(), s.Value["code"].ToString(), s.Value["description"].ToString());
                i++;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            frm1.export_to_xls(dataGridView1);
        }


    }
}