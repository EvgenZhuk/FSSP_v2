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
using Excel = Microsoft.Office.Interop.Excel;


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
        string result;
        int Flag;

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
            Flag = 0;
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
                    LICA = "";

                    //MessageBox.Show(ZaprosToFssp);
                    SendPostToFSSP();
                    DownloadDocsMainPageAsync();
                    ParsingResponse();
                    Flag = 0;
                    //MessageBox.Show("Ура! Посчитал!");
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
            //MessageBox.Show("Ура! Посчитал!");
            Flag = 0;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //Объявляем приложение
            Excel.Application ex = new Microsoft.Office.Interop.Excel.Application();

            //Отобразить Excel
            ex.Visible = true;

            // Отображается в полноэкранном режиме
            ex.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlMaximized;

            //Количество листов в рабочей книге
            ex.SheetsInNewWorkbook = 1;

            //Добавить рабочую книгу
            Excel.Workbook workBook = ex.Workbooks.Add(Type.Missing);

            //Отключить отображение окон с сообщениями
            ex.DisplayAlerts = false;

            //Получаем первый лист документа (счет начинается с 1)
            Excel.Worksheet sheet = (Excel.Worksheet)ex.Worksheets.get_Item(1);

            //Название листа (вкладки снизу)
            sheet.Name = "Отчет";

            sheet.Cells[1, 1] = "Фамилия";
            sheet.Cells[1, 2] = "Имя";
            sheet.Cells[1, 3] = "Отчество";
            sheet.Cells[1, 4] = "Дата рождения";
            sheet.Cells[1, 5] = "Контрольная сумма";
            sheet.Cells[1, 6] = "Дата визита";
            sheet.Cells[1, 7] = "Адрес СУ";
            sheet.Cells[1, 8] = "№ СУ";
            sheet.Cells[1, 9] = "Реестр";
            sheet.Cells[1, 10] = "Задолженность";

            sheet.Columns[1].ColumnWidth = 23;
            sheet.Columns[2].ColumnWidth = 20;
            sheet.Columns[3].ColumnWidth = 20;
            sheet.Columns[4].ColumnWidth = 15;
            sheet.Columns[5].ColumnWidth = 40;
            sheet.Columns[6].ColumnWidth = 20;
            sheet.Columns[7].ColumnWidth = 40;
            sheet.Columns[8].ColumnWidth = 10;
            sheet.Columns[9].ColumnWidth = 40;
            sheet.Columns[10].ColumnWidth = 20;

            sheet.Rows[1].RowHeight = 25;

            sheet.get_Range("A1", "J2").Font.Bold = true;       // жирный курсив
            sheet.get_Range("A1", "J1").VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            sheet.get_Range("A1", "J1").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            //sheet.get_Range("A1", "J1").Interior.Color = ColorTranslator.ToOle(Color.FromArgb(0xFF, 0xFF, 0xCC));
            //sheet.get_Range("A2", "J2").Interior.Color = ColorTranslator.ToOle(Color.FromArgb(0x0, 0xFF, 0xFF));
            sheet.get_Range("A1", "J1").Interior.Color = Color.NavajoWhite;
            sheet.get_Range("A2", "J2").Interior.Color = Color.DarkTurquoise;
            sheet.get_Range("A1", "J1").Cells.Font.Name = "Times New Roman";
            sheet.get_Range("A1", "J1").Cells.Font.Size = 10;



            int BadBoysCount = 0;
            //Пример заполнения ячеек                        
            for (int i = 0; i < dataGridView1.RowCount - 1; i++) // строки
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++) // столбцы
                {
                    //sheet.Cells[i + 3, j + 1] = String.Format(dataGridView1.Rows[i].Cells[j].Value.ToString());
                    sheet.Cells[i + 3, j + 1] = dataGridView1.Rows[i].Cells[j].Value;
                }
                BadBoysCount++;
            }
            sheet.Cells[2, 1] = "Всего нарушителей - " + BadBoysCount.ToString();


            int FsspGuys = 0;
            int MvdGuys = 0;
            int CovidGuys = 0;


            for (int i = 3; i < sheet.UsedRange.Rows.Count + 1; i++)
            {
                if (sheet.Cells[i, 9].Value == "МВД")
                {
                    sheet.Cells[i, 9].Interior.Color = Color.Tomato;
                    MvdGuys++;
                    //sheet.get_Range(5, 7).Interior.Color = ColorTranslator.ToOle(Color.FromArgb(0x8B, 0x0, 0x0));
                }

                if (sheet.Cells[i, 9].Value == " ФССП")
                {
                    sheet.Cells[i, 9].Interior.Color = Color.LightGreen;
                    FsspGuys++;
                }

                if (sheet.Cells[i, 9].Value.StartsWith(" КОВИД"))
                {
                    sheet.Cells[i, 9].Interior.Color = Color.Yellow;
                    CovidGuys++;
                }
            }
            sheet.Cells[2, 9] = "ФССП - " + FsspGuys.ToString() + "\n" + "МВД - " + MvdGuys.ToString() + "\n" + "Covid - " + CovidGuys.ToString();

            sheet.UsedRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            sheet.UsedRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

          
            ///////////////////////// переводим в денежный формат
            for (int i = 3; i <= sheet.UsedRange.Rows.Count; i++)
            {
                sheet.Cells[i, 10] = Convert.ToDecimal(sheet.Cells[i, 10].Value);
            }

            sheet.get_Range("J3", "J" + (sheet.UsedRange.Rows.Count)).NumberFormat = "#,##0.00 $";
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
                    result = Regex.Replace(responseFromServer, @"\\[Uu]([0-9A-Fa-f]{4})", m => char.ToString((char)ushort.Parse(m.Groups[1].Value, NumberStyles.AllowHexSpecifier)));
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

            // ВОТ ТУТ НАЧИНАЕТСЯ ПАРСИНГ
            string AllText = result;
            //MessageBox.Show("Начался парсинг !");
            int A = AllText.IndexOf('[');
            int B = AllText.LastIndexOf(']');
            AllText = AllText.Remove(B);
            AllText = AllText.Remove(0, A + 1);
            AllText = AllText.Trim();

            int A1 = AllText.IndexOf('[');
            int B1 = AllText.LastIndexOf(']');
            AllText = AllText.Remove(B1 + 1);
            AllText = AllText.Remove(0, A1 - 1);


            for (int i = 1; i <= Flag; i++)

            {
                int A2 = AllText.IndexOf('[');
                int B2 = AllText.IndexOf(']');
                string NEGODYAY = AllText.Substring(A2, B2 + 1);


                int tochkaOtscheta = 0;
                decimal SUMMA;
                decimal ZADOLZHENOST = 0;

                while (NEGODYAY.IndexOf(" руб", tochkaOtscheta) != -1)
                {
                    int RUB = NEGODYAY.IndexOf(" руб", tochkaOtscheta);
                    int DVOETOCH = NEGODYAY.LastIndexOf(':', RUB);
                    string DOLG = NEGODYAY.Substring(DVOETOCH + 2, RUB - DVOETOCH - 2);
                    SUMMA = Decimal.Parse(DOLG.Replace('.', ','));
                    ZADOLZHENOST += SUMMA;
                    tochkaOtscheta = RUB + 1;

                }
                
                int j = 0;
                while (dataGridView1.Rows[j].Cells[9].Value.ToString() != "")
                {                    
                    j++;
                }
                dataGridView1.Rows[j].Cells[9].Value = ZADOLZHENOST;    //.ToString();



            AllText = AllText.Remove(0, B2 + 1);
                if (AllText != "")
                {
                    AllText = AllText.Remove(0, AllText.IndexOf('['));
                }


            }
            // ВОТ ТУТ ЗАКАНЧИВАЕТСЯ ПАРСИНГ



        }

        private void button4_Click(object sender, EventArgs e)
        {

        }
    }
}
