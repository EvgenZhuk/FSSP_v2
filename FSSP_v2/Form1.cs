using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Npgsql;
using System.Net;
using System.IO;
using System.Text.RegularExpressions;
using System.Globalization;


namespace FSSP_v2
{
    /*class ListDataRow
    {
        public string firstname;   // Имя
        public string lastname;    // Фамилия
        public string secondname;  // Отчество
        public string birthdate;   // День рожденье

        public void GetListViolators(string A, string B, string C, string D )
        {
            firstname = A;
            lastname = B;
            secondname = C;
            birthdate = D;

            LICA += "{\n" +
                  "\t  \"type\": 1,\n" +
                  "\t  \"params\": {\n" +
                  "\t     \"firstname\": \"" + firstname + "\",\n" +
                  "\t     \"lastname\": \"" + lastname + "\",\n" +
                  "\t     \"secondname\": \"" + secondname + "\",\n" +
                  "\t     \"region\": \"77\",\n" +
                  "\t     \"birthdate\": \"" + birthdate + "\"\n" +
                  "\t   }\n" +
                  "\t },\n";
        }
        

    }
    */
    public partial class Form1 : Form
    {
        private DataSet ds = new DataSet();
        private DataTable dt = new DataTable();
        int CountViolators; // количество нарушителей
        String TaskFromFsspApi;
        string ZaprosToFssp;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                NpgsqlConnection con = new NpgsqlConnection("Host=172.17.75.4;Username=postgres;Password=postgres;Database=ums");
                con.Open();


                string sql = ("select " +
                                "upper(v.last_name) as Фамилия, upper(v.first_name) as Имя, upper(v.patronymic) as Отчество, " +
                                "(to_char(v.birthdate, 'DD.MM.YYYY')) as \"Дата рождения\" , " +
                                "(md5(concat(upper(v.last_name), upper(v.first_name), upper(v.patronymic), v.birthdate::date))) as \"Контрольная Сумма\", " +
                                "(to_char(c.creation_date, 'YYYY-MM-DD hh24:mi:ss')) as \"Дата визита\", " +
                                "o.address as \"Адрес СУ\", " +
                                "ct.number as \"№СУ\", " +
                                "concat(CASE WHEN mia_check_result = 1 THEN 'МВД'  end, " +
                                "case when fssp_check_result = 1 then ' ФССП' end, " +
                                "case when covid_check_result = 1 then ' КОВИД КАРАНТИН'   end, " +
                                "case when covid_check_result = 2  then  ' КОВИД БОЛЕН' end, " +
                                "(to_char(c.covid_quarantine_end_date, ' до DD.MM.YYYY'))) as \"Реестр\" " +
                                "FROM visitor_violation_checks AS c " +
                                "RIGHT JOIN visitors AS v ON c.visitor_id = v.id " +
                                "RIGHT JOIN court_objects AS o ON v.court_object_id = o.id " +
                                "left join visitor_to_court_station on visitor_to_court_station.visitor_id = v.id " +
                                "left join court_stations ct on ct.id = visitor_to_court_station.court_station_id " +
                                "WHERE v.court_object_id not IN(173, 174) " +
                                "AND(mia_check_result = 1 OR fssp_check_result = 1 or covid_check_result <> 0) " +
                                $"AND v.creation_date >= '{dateTimePicker1.Text}' AND v.creation_date <= '{dateTimePicker2.Text} 23:59:59' ORDER BY v.creation_date desc");

                //DataSet ds = new DataSet();      // если будет необходимость доступа из вне 
                //DataTable dt = new DataTable();  // просто надо будет раскомментировать эти строки выше
                NpgsqlDataAdapter da = new NpgsqlDataAdapter(sql, con);
                ds.Reset();
                da.Fill(ds);
                dt = ds.Tables[0];
                dt.Columns.Add("Задолженность", typeof(String));
                dataGridView1.DataSource = dt;
                con.Close();

                for (int i = 0; i < dataGridView1.RowCount - 1; i++) // нумерация 
                {
                    dataGridView1.Rows[i].HeaderCell.Value = (i + 1).ToString();
                }

                CountViolators = dt.Rows.Count;
                label3.Visible = true;
                label3.Text = "Отобрано записей: " + CountViolators.ToString();
                //MessageBox.Show(dt.Rows[0][1].ToString());
                //dataGridView1.Columns.Add(new DataGridViewTextBoxColumn() { Name = "dgvAge", HeaderText = "Задолженность", Width = 100 });
            }
            catch
            {
                MessageBox.Show("Неудалось осуществить подключение к базе UMS! " +
                    "Проверьте подключена ли сеть или VPN!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int Flag = 0;
            string LICA = "";
            ZaprosToFssp = "";
            for (int i = 0; i < CountViolators; i++)
            {
                LICA += "{\n" +
                "\t  \"type\": 1,\n" +
                "\t  \"params\": {\n" +
                "\t     \"firstname\": \"" + dt.Rows[i][1].ToString() + "\",\n" +
                "\t     \"lastname\": \"" + dt.Rows[i][0].ToString() + "\",\n" +
                "\t     \"secondname\": \"" + dt.Rows[i][2].ToString() + "\",\n" +
                "\t     \"region\": \"77\",\n" +
                "\t     \"birthdate\": \"" + dt.Rows[i][3].ToString() + "\"\n" +
                "\t   }\n" +
                "\t },\n";

                Flag++;


                if (Flag == 50)
                {
                    LICA = LICA.Remove(LICA.LastIndexOf(','));

                    ZaprosToFssp = "{\n" +
                    "\t\"token\": \"aMYRXmjGwPFm\",\n" +
                    "\t\"request\": [\n" +
                    LICA + "\n" +
                    "\t]\n" +
                    "}";
                    richTextBox1.Text = ZaprosToFssp;
                    Flag = 0; LICA = "";

                    //MessageBox.Show(ZaprosToFssp);
                    SendPostToFSSP();
                    DownloadDocsMainPageAsync();
                    ParsingResponse();
                }

            }

            if (LICA != "")
            {
                LICA = LICA.Remove(LICA.LastIndexOf(','));

                ZaprosToFssp = "{\n" +
                "\t\"token\": \"aMYRXmjGwPFm\",\n" +
                "\t\"request\": [\n" +
                LICA + "\n" +
                "\t]\n" +
                "}";
                richTextBox2.Text = ZaprosToFssp;
            }

            SendPostToFSSP();
            DownloadDocsMainPageAsync();
            //new System.Threading.Thread(DownloadDocsMainPageAsync).Start();
            //Task.Run(()=>DownloadDocsMainPageAsync()).RunSynchronously();
            ParsingResponse();

        }

        private void button3_Click(object sender, EventArgs e)
        {            
            int j = 0;
            while (dataGridView1.Rows[j].Cells[9].Value.ToString() != "")
            {
                MessageBox.Show("Строка " + j.ToString() + " не пустая");
                j++;
            }
            MessageBox.Show("Найдена пустая строка: " + j.ToString() + " пустая");
        }


        private void DownloadDocsMainPageAsync()
        {
            //System.Threading.Thread.Sleep(5000);
            string Status = "";
            do
            {
                var request = WebRequest.Create("https://api-ip.fssprus.ru/api/v1.0/result?token=aMYRXmjGwPFm&task=" + TaskFromFsspApi);
                var response = request.GetResponseAsync().Result;
                using (Stream dataStream = response.GetResponseStream())
                {
                    StreamReader reader = new StreamReader(dataStream);
                    string responseFromServer = reader.ReadToEnd();
                    var result = Regex.Replace(responseFromServer, @"\\[Uu]([0-9A-Fa-f]{4})", m => char.ToString((char)ushort.Parse(m.Groups[1].Value, NumberStyles.AllowHexSpecifier)));
                    richTextBox1.Text = result;
                    int A = result.IndexOf("status", 15);
                    Status = result.Substring(A + 8, 1);
                    System.Threading.Thread.Sleep(5000);
                    //await Task.Delay(5000);
                }
                response.Close();

            } while (Status != "0");
                       
        }


        public void SendPostToFSSP()
        {
            WebRequest request = WebRequest.Create("https://api-ip.fssprus.ru/api/v1.0/search/group");
            request.Method = "POST";
            byte[] DATA = Encoding.UTF8.GetBytes(ZaprosToFssp);
            request.ContentType = "application/json; charset=utf-8";
            request.ContentLength = DATA.Length;
            Stream dataStream = request.GetRequestStream();
            dataStream.Write(DATA, 0, DATA.Length);
            dataStream.Close();

            WebResponse response = request.GetResponse();
            using (dataStream = response.GetResponseStream())
            {
                StreamReader reader = new StreamReader(dataStream);
                string responseFromServer = reader.ReadToEnd();
                richTextBox1.Text = responseFromServer;

                int A = responseFromServer.IndexOf("task");
                textBox1.Text = responseFromServer.Substring(A + 7, 36);
                TaskFromFsspApi = responseFromServer.Substring(A + 7, 36);
                //System.Threading.Thread.Sleep(5000);
            }
            response.Close();
        }



        public void ParsingResponse() 
        {
            
        

        }

    }
}
