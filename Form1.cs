using System;
using System.Windows.Forms;
using System.IO;
using System.Net;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;


namespace Exchangerate
{
    public partial class Form1 : Form
    {
        public string main_url = "https://api.exchangerate.host", base_0 = "EUR", format = "", source = "";


        public Form1()
        {
            InitializeComponent();


        }



        private void supportedCurrenciesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Supported_currencies S_C = new Supported_currencies();
            S_C.Show();
        }

        private void latestRatesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Latest_Rates L_R = new Latest_Rates();
            L_R.Show();
        }

        private void historicalRatesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Historical_Rates H_R = new Historical_Rates();
            H_R.Show();
        }

        private void timeSeriesRatesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TimeSeries_Rates TS_R = new TimeSeries_Rates();
            TS_R.Show();
        }








        public dynamic get_all_supported_currencies()
        {
            WebRequest request = HttpWebRequest.Create(main_url + "/symbols");
            WebResponse response = request.GetResponse();
            StreamReader reader = new StreamReader(response.GetResponseStream());
            string currencies_list = reader.ReadToEnd();
            dynamic d_curr_list = JsonConvert.DeserializeObject(currencies_list);
            return (d_curr_list);
        }









        public dynamic get_latest(String base_0, double amount, String[] output_currencies, int places)
        {
            String requested_currencies = "";
            if (output_currencies[0] == "All currencies") { }
            else
            {
                for (int i = 0; i < output_currencies.Length; i++)
                { requested_currencies = requested_currencies + output_currencies[i] + ","; }
                requested_currencies = "&symbols=" + requested_currencies;
            }
            WebRequest request = HttpWebRequest.Create(main_url + "/latest?base=" + base_0 + "&amount=" + amount + "&places=" + places + requested_currencies);
            WebResponse response = request.GetResponse();
            StreamReader reader = new StreamReader(response.GetResponseStream());
            string latest_rates = reader.ReadToEnd();
            dynamic d_latest_rates = JsonConvert.DeserializeObject(latest_rates);
            return (d_latest_rates);
        }








        public dynamic get_historical_rates(DateTime date, String base_0, double amount, String[] output_currencies, int places)
        {
            String requested_currencies = "";
            if (output_currencies[0] == "All currencies") { }
            else
            {
                for (int i = 0; i < output_currencies.Length; i++)
                { requested_currencies = requested_currencies + output_currencies[i] + ","; }
                requested_currencies = "&symbols=" + requested_currencies;
            }
            WebRequest request = HttpWebRequest.Create(main_url + "/" + date.ToString("yyyy-MM-dd") + "?base=" + base_0 + "&amount=" + amount + "&places=" + places + requested_currencies);
            WebResponse response = request.GetResponse();
            StreamReader reader = new StreamReader(response.GetResponseStream());
            string spe_rates = reader.ReadToEnd();
            dynamic d_spe_rates = JObject.Parse(spe_rates);
            return (d_spe_rates);
        }







        public dynamic get_timeseries_rates(DateTime start, DateTime end, String base_0, double amount, String[] output_currencies, int places)
        {
            String requested_currencies = "";
            if (output_currencies[0] == "All currencies") { }
            else
            {
                for (int i = 0; i < output_currencies.Length; i++)
                { requested_currencies = requested_currencies + output_currencies[i] + ","; }
                requested_currencies = "&symbols=" + requested_currencies;
            }
            WebRequest request = HttpWebRequest.Create(main_url + "/timeseries?start_date=" + start.ToString("yyyy-MM-dd") + "&end_date=" + end.ToString("yyyy-MM-dd") + "&base=" + base_0 + "&amount=" + amount + "&places=" + places + requested_currencies);
            WebResponse response = request.GetResponse();
            StreamReader reader = new StreamReader(response.GetResponseStream());
            string spe_rates = reader.ReadToEnd();
            dynamic d_spe_rates = JObject.Parse(spe_rates);
            return (d_spe_rates);
        }






        public void export_to_xls(DataGridView data_t){

            if (data_t.Rows.Count == 0) { MessageBox.Show("There is no data to export."); }
            else
            {
                SaveFileDialog save = new SaveFileDialog();
                save.Title = "Select save location";
                save.Filter = "XLS files (*.xls)|*.*";
                if (save.ShowDialog() == DialogResult.OK)
                {
                    data_t.SelectAll();
                    DataObject dataObj = data_t.GetClipboardContent();
                    Clipboard.SetDataObject(dataObj);
                    data_t.ClearSelection();
                    string file_data = Clipboard.GetText(TextDataFormat.Text);
                    File.WriteAllText(save.FileName + ".xls", file_data);
                    DialogResult dialogResult = MessageBox.Show("Do you want to open it now?", "Successfully saved", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        System.Diagnostics.Process.Start(save.FileName + ".xls");
                    }
                    else if (dialogResult == DialogResult.No) { }
                }
            }

        }











    }
}
