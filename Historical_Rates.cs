using System;
using System.Drawing;
using System.Windows.Forms;

namespace Exchangerate
{
    public partial class Historical_Rates : Form
    {
        Form1 frm1 = new Form1();

        public Historical_Rates()
        {
            InitializeComponent();

            DateTime now = DateTime.Now;
            dateTimePicker1.MaxDate = now.AddDays(-1);


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

                double amount = double.Parse(textBox1.Text);
                int places = Convert.ToInt32(numericUpDown1.Value);

                string[] output_currencies = new string[selected_items];
                for (int x = 0; x < selected_items; x++)
                {
                    output_currencies[x] = listBox1.Items[listBox1.SelectedIndices[x]].ToString();
                }

                dataGridView1.Rows.Clear();
                dataGridView1.Refresh();

                dynamic d_latest = frm1.get_latest(comboBox1.Text, amount, output_currencies, places);
                dynamic d_historical = frm1.get_historical_rates(dateTimePicker1.Value,comboBox1.Text, amount, output_currencies, places);

                label13.Text = d_latest["success"].ToString();
                label15.Text = d_latest["base"].ToString();
                label16.Text = d_latest["date"].ToString();
                label11.Text = d_historical["success"].ToString();
                label4.Text = d_historical["base"].ToString();
                label7.Text = d_historical["date"].ToString();

                int k = 0;
                foreach (var s in d_latest["rates"])
                {k++;}

                string[] code = new string[k];
                double[] latest_rate = new double[k];
                double[] hist_rate = new double[k];


                int m = 0;
                foreach (var s in d_latest["rates"])
                {
                    code[m] = s.Name.ToString();
                    latest_rate[m]=s.Value;
                    m++;
                }


                int p = 0;
                foreach (var s in d_historical["rates"])
                {
                    hist_rate[p] = s.Value;
                    p++;
                }


                for (int i = 0; i < k; i++) {
                    double change = Math.Round(((latest_rate[i] / hist_rate[i]) - 1) * 100, places) ;
                    dataGridView1.Rows.Add((i+1).ToString(), code[i], latest_rate[i], change+"%");
                }

                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    string s_change_percentage = row.Cells[3].Value.ToString().Replace("%","");
                    double change_percentage = double.Parse(s_change_percentage);
                    if (change_percentage > 0) { row.DefaultCellStyle.ForeColor = Color.DarkGreen; }
                    else if (change_percentage < 0) { row.DefaultCellStyle.ForeColor = Color.Red; }
                    else { }
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
