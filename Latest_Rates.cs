using System;
using System.Windows.Forms;

namespace Exchangerate
{
    public partial class Latest_Rates : Form
    {
        Form1 frm1 = new Form1();

        public Latest_Rates()
        {
            InitializeComponent();

            listBox1.Items.Add("All currencies");
            dynamic d_curr_list = frm1.get_all_supported_currencies();
            foreach (var s in d_curr_list["symbols"])
            {
                string items = s.Value["code"].ToString();
                listBox1.Items.Add(items);
                comboBox1.Items.Add(items);
            }
            comboBox1.SelectedItem = "EUR";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int selected_items = listBox1.SelectedIndices.Count;
            if (selected_items == 0) { MessageBox.Show("Please select at least one currency"); }
            else
            {
                dataGridView1.Rows.Clear();
                dataGridView1.Refresh();

                double amount = double.Parse(textBox1.Text);
                int places = Convert.ToInt32(numericUpDown1.Value);

                string[] output_currencies = new string[selected_items];
                for (int i = 0; i < selected_items; i++)
                {
                    output_currencies[i] = listBox1.Items[listBox1.SelectedIndices[i]].ToString();
                }

                dynamic d_latest = frm1.get_latest(comboBox1.Text, amount, output_currencies, places);
                label13.Text = d_latest["success"].ToString();
                label15.Text = d_latest["base"].ToString();
                label16.Text = d_latest["date"].ToString();

                int k = 1;
                foreach (var s in d_latest["rates"])
                {
                    dataGridView1.Rows.Add(k.ToString(), s.Name, s.Value.ToString());
                    k++;
                }
                dataGridView1.ClearSelection();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            frm1.export_to_xls(dataGridView1);
        }
    }
}