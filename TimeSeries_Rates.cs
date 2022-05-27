using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace Exchangerate
{
    public partial class TimeSeries_Rates : Form
    {
        Form1 frm1 = new Form1();

        public TimeSeries_Rates()
        {
            InitializeComponent();

            DateTime now = DateTime.Now;
            dateTimePicker1.MaxDate = now.AddDays(-1);
            dateTimePicker1.Value = now.AddDays(-7);
            dateTimePicker2.MaxDate = now;

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
                DateTime start = dateTimePicker1.Value;
                DateTime end = dateTimePicker2.Value;

                string[] output_currencies = new string[selected_items];
                for (int x = 0; x < selected_items; x++)
                {output_currencies[x] = listBox1.Items[listBox1.SelectedIndices[x]].ToString();}

                dataGridView1.Rows.Clear();
                dataGridView1.Columns.Clear();
                dataGridView1.Refresh();
                dataGridView1.SelectionMode = DataGridViewSelectionMode.CellSelect;

                dynamic d_timeseries_rates = frm1.get_timeseries_rates(start, end, comboBox1.Text, amount, output_currencies, places);

                label13.Text = d_timeseries_rates["success"].ToString();
                label14.Text = d_timeseries_rates["timeseries"].ToString();
                label15.Text = d_timeseries_rates["base"].ToString();
                label16.Text = d_timeseries_rates["start_date"].ToString();
                label17.Text = d_timeseries_rates["end_date"].ToString();

                int days = (end - start).Days+1;
                int currencies_count = 0;

                foreach (var s in d_timeseries_rates["rates"][start.ToString("yyyy-MM-dd")])
                { currencies_count++; }

                string[] currencies_array = new string[currencies_count];
                string[,] series_data = new string[days, currencies_count];
                string date_step = "";


                for (int i = 0; i < days; i++)
                {
                    date_step = start.AddDays(i).ToString("yyyy-MM-dd");
                    int t = 0;
                    foreach (var s in d_timeseries_rates["rates"][date_step])
                    {
                        if(i==0)currencies_array[t] = s.Name;
                        series_data[i,t] = s.Value.ToString();
                        t++;
                    }
                }

                for (int d=0; d< currencies_count; d++) {
                    if (d == 0){dataGridView1.Columns.Add("Date", "Date");}
                    dataGridView1.Columns.Add(currencies_array[d], currencies_array[d]);
                }

                for (int i = 0; i < days; i++)
                {
                    DataGridViewRow row = new DataGridViewRow();
                    row.CreateCells(this.dataGridView1);
                    for (int j = 0; j < currencies_count; j++)
                    {
                        if (j == 0)
                        {row.Cells[j].Value = start.AddDays(i).ToString("yyyy-MM-dd");}
                        row.Cells[j+1].Value= series_data[i, j];
                    }
                    dataGridView1.Rows.Add(row);
                }

                foreach (DataGridViewColumn col in dataGridView1.Columns)
                {
                    col.SortMode = DataGridViewColumnSortMode.NotSortable;
                    col.Selected = false;
                }
                dataGridView1.SelectionMode = DataGridViewSelectionMode.FullColumnSelect;
            }
        }


        private void button2_Click(object sender, EventArgs e)
        {
            frm1.export_to_xls(dataGridView1);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            int days = dataGridView1.Rows.Count;
            if (days == 0) { MessageBox.Show("There is no data to view chart."); }
            else {
                chart1.Series.Clear();
                int selected_column = dataGridView1.CurrentCell.ColumnIndex;
                string currency_name = dataGridView1.CurrentCell.OwningColumn.Name;
                Series series = chart1.Series.Add(currency_name);
                series.Color = Color.Black;
                series.ChartType = SeriesChartType.Line;

                if (selected_column == 0) { MessageBox.Show("Please select a valid column."); }
                else
                {
                    DateTime[] date_array = new DateTime[days];
                    double[] currency_array = new double[days];

                    for (int i = 0; i < days; i++)
                    {
                        date_array[i] = Convert.ToDateTime(dataGridView1.Rows[i].Cells[0].Value);
                        currency_array[i] = Convert.ToDouble(dataGridView1.Rows[i].Cells[selected_column].Value);
                    }

                    chart1.Series[0].XValueType = ChartValueType.DateTime;
                    chart1.ChartAreas[0].AxisX.LabelStyle.Format = "yyyy-MM-dd";
                    chart1.ChartAreas[0].AxisX.IntervalType = DateTimeIntervalType.Days;
                    DateTime start = date_array[0];
                    DateTime end = date_array[days - 1];
                    chart1.ChartAreas[0].AxisX.Minimum = start.ToOADate();
                    chart1.ChartAreas[0].AxisX.Maximum = end.ToOADate();
                    chart1.ChartAreas[0].AxisY.Minimum = currency_array.Min();
                    chart1.ChartAreas[0].AxisY.Maximum = currency_array.Max();

                    series.Points.AddXY(date_array[0], currency_array[0]);
                    for (int i = 1; i < days; i++)
                    {
                        int p = series.Points.AddXY(date_array[i], currency_array[i]);

                        if (series.Points[p - 1].YValues[0] < series.Points[p].YValues[0]) { series.Points[p].Color = Color.DarkGreen; }
                        else if (series.Points[p - 1].YValues[0] > series.Points[p].YValues[0]) { series.Points[p].Color = Color.Red; }
                        else { series.Points[p].Color = Color.Black; }
                    }
                }
            }
        }
    }
}
