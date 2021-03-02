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

        private DataSet ds2 = new DataSet();
        private DataTable dt2 = new DataTable();

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

                string sql2 = ("select " +
                                "a.\"Фамилия\", a.\"Имя\", a.\"Отчество\", a.\"Дата рождения\", a.\"Дата визита\", a.\"Адрес СУ\", a.\"№СУ\", a.\"РЕЕСТР\",a.\"Контрольная Сумма\" " +
                                "from(SELECT ROW_NUMBER() OVER( " +
                                "PARTITION BY(md5(concat(upper(v.last_name), upper(v.first_name), upper(v.patronymic), v.birthdate::date))) " +
                                "order by(to_char(c.creation_date, 'YYYY-MM-DD hh24:mi:ss')) " +
                                ") as \"П/П\", upper(v.last_name) \"Фамилия\", upper(v.first_name) \"Имя\", upper(v.patronymic) \"Отчество\", " +
                                "v.birthdate::date \"Дата рождения\", (to_char(c.creation_date, 'YYYY-MM-DD hh24:mi:ss')) as \"Дата визита\", " +
                                "o.address as \"Адрес СУ\", string_agg(distinct ct.\"number\"::varchar, ', ') as \"№СУ\", " +
                                "concat(CASE WHEN mia_check_result = 1 THEN ' МВД'  end, " +
                                "case when fssp_check_result = 1 then ' ФССП' end, " +
                                "case when covid_check_result = 1 then ' КОВИД КАРАНТИН' end, " +
                                "case when covid_check_result = 2  then  ' КОВИД КОНТАКТ' end, " +
                                "(to_char(c.covid_quarantine_end_date, ' до DD.MM.YYYY'))) as \"РЕЕСТР\", " +
                                "(md5(concat(upper(v.last_name), upper(v.first_name), upper(v.patronymic), v.birthdate::date))) as \"Контрольная Сумма\" " +
                                "FROM visitor_violation_checks AS c " +
                                "left JOIN visitors AS v  ON c.visitor_id = v.id " +
                                "left JOIN court_objects AS o ON v.court_object_id = o.id " +
                                "left join visitor_to_court_station on visitor_to_court_station.visitor_id = v.id " +
                                "left join court_stations ct on ct.id = visitor_to_court_station.court_station_id " +
                                "WHERE v.court_object_id not IN(173, 174) " +
                                "AND(mia_check_result = 1 OR fssp_check_result = 1 or covid_check_result <> 0) and v.id  not in (1307421)" +
                                //$"and v.id  not in (1307421) AND v.creation_date >= '{dateTimePicker1.Text}' and v.creation_date <= '{dateTimePicker2.Text} 23:59:59' " +
                                "group by \"Фамилия\", \"Имя\", \"Отчество\", \"Дата рождения\", \"Дата визита\", \"Адрес СУ\", \"РЕЕСТР\", \"Контрольная Сумма\" " +
                                "ORDER BY \"Дата визита\" desc) as a where a.\"П/П\" = 1 limit 10");


                //richTextBox3.Text = sql_2;

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




                NpgsqlDataAdapter da2 = new NpgsqlDataAdapter(sql2, con);
                ds2.Reset();
                da2.Fill(ds2);
                dt2 = ds2.Tables[0];
                dt2.Columns.Add("Задолженность", typeof(String));
                dataGridView2.DataSource = dt2;
                con.Close();

                for (int i = 0; i < dataGridView2.RowCount - 1; i++) // нумерация 
                {
                    dataGridView2.Rows[i].HeaderCell.Value = (i + 1).ToString();
                }                
                label4.Text = "Отобрано записей: " + dt2.Rows.Count.ToString();
                label4.Visible = true;
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


            /////////// НАЧНЕМ ВТОРУЮ ТАБЛИЦУ ///////////////////
            ///
            Flag = 0;
            string LICA2 = "";
            ZaprosToFssp = "";
            for (int i = 0; i < dt2.Rows.Count; i++)
            {
                LICA2 += "{\n" +
                "\t  \"type\": 1,\n" +
                "\t  \"params\": {\n" +
                "\t     \"firstname\": \"" + dt2.Rows[i][1].ToString() + "\",\n" +
                "\t     \"lastname\": \"" + dt2.Rows[i][0].ToString() + "\",\n" +
                "\t     \"secondname\": \"" + dt2.Rows[i][2].ToString() + "\",\n" +
                "\t     \"region\": \"77\",\n" +
                "\t     \"birthdate\": \"" + dt2.Rows[i][3].ToString() + "\"\n" +
                "\t   }\n" +
                "\t },\n";

                Flag++;


                if (Flag == 50)
                {
                    LICA2 = LICA2.Remove(LICA2.LastIndexOf(','));

                    ZaprosToFssp = "{\n" +
                    "\t\"token\": \"aMYRXmjGwPFm\",\n" +
                    "\t\"request\": [\n" +
                    LICA2 + "\n" +
                    "\t]\n" +
                    "}";
                    richTextBox1.Text = ZaprosToFssp;
                    LICA2 = "";

                    SendPostToFSSP();
                    DownloadDocsMainPageAsync();
                    ParsingResponse2();
                    Flag = 0;
                }

            }

            if (LICA2 != "")
            {
                LICA2 = LICA2.Remove(LICA2.LastIndexOf(','));

                ZaprosToFssp = "{\n" +
                "\t\"token\": \"aMYRXmjGwPFm\",\n" +
                "\t\"request\": [\n" +
                LICA2 + "\n" +
                "\t]\n" +
                "}";
                richTextBox2.Text = ZaprosToFssp;
            }

            SendPostToFSSP();
            DownloadDocsMainPageAsync();
            //new System.Threading.Thread(DownloadDocsMainPageAsync).Start();
            //Task.Run(()=>DownloadDocsMainPageAsync()).RunSynchronously();
            ParsingResponse2();
            //MessageBox.Show("Ура! Посчитал!");
            Flag = 0;

            for (int i = 0; i < dataGridView1.RowCount - 1; i++) 
            {
                

                if (dataGridView1.Rows[i].Cells[8].Value.ToString().IndexOf("КОВИД") == 1 || dataGridView1.Rows[i].Cells[8].Value.ToString().IndexOf("МВД") == 1)
                {
                    dataGridView1.Rows[i].Cells[9].Value = "0";
                }

                if (dataGridView2.Rows[i].Cells[7].Value.ToString().IndexOf("КОВИД") == 1 || dataGridView2.Rows[i].Cells[7].Value.ToString().IndexOf("МВД") == 1)
                {
                    dataGridView2.Rows[i].Cells[9].Value = "0";
                }

            }
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
            ex.SheetsInNewWorkbook = 3;

            //Добавить рабочую книгу
            Excel.Workbook workBook = ex.Workbooks.Add(Type.Missing);

            //Отключить отображение окон с сообщениями
            ex.DisplayAlerts = false;

            //Получаем первый лист документа (счет начинается с 1)
            Excel.Worksheet sheet = (Excel.Worksheet)ex.Worksheets.get_Item(1);

            //Получаем второй лист документа (счет начинается с 1)
            Excel.Worksheet sheet2 = (Excel.Worksheet)ex.Worksheets.get_Item(2);

            //Получаем третий лист документа (счет начинается с 1)
            Excel.Worksheet sheet3 = (Excel.Worksheet)ex.Worksheets.get_Item(3);

            //Название листа (вкладки снизу)
            sheet.Name = "Отчет";
            sheet2.Name = "Для руководства";
            sheet3.Name = "  ФИО  ";

            sheet3.Cells[1, 1] = "Фамилия";
            sheet3.Cells[1, 2] = "Имя";
            sheet3.Cells[1, 3] = "Отчество";
            sheet3.Cells[1, 4] = "Дата рождения";
            sheet3.Cells[1, 5] = "Контрольная сумма";
            sheet3.Cells[1, 6] = "Дата визита";
            sheet3.Cells[1, 7] = "Адрес СУ";
            sheet3.Cells[1, 8] = "№ СУ";
            sheet3.Cells[1, 9] = "Реестр";
            sheet3.Cells[1, 10] = "Задолженность";

            sheet2.Cells[1, 1] = "ДАТА";
            sheet2.Cells[1, 2] = "АДРЕС";
            sheet2.Cells[1, 3] = "УЧАСТОК";
            sheet2.Cells[1, 4] = "РЕЕСТР";
            sheet2.Cells[1, 5] = "Комментарий";
            sheet2.Cells[1, 6] = "Задержаний";
            sheet2.Cells[1, 7] = "Сумма взысканий";
            sheet2.Cells[2, 1] = "ИТОГО";
            sheet2.Cells[2, 2] = "Нарушителей: " + dt2.Rows.Count.ToString();

            sheet.Cells[1, 1] = "ДАТА";
            sheet.Cells[1, 2] = "АДРЕС";
            sheet.Cells[1, 3] = "УЧАСТОК";
            sheet.Cells[1, 4] = "РЕЕСТР";
            sheet.Cells[1, 5] = "Контрольная Сумма";
            sheet.Cells[1, 6] = "Комментарий";
            sheet.Cells[1, 7] = "Задержаний";
            sheet.Cells[1, 8] = "Сумма взысканий";
            sheet.Cells[2, 1] = "ИТОГО";
            sheet.Cells[2, 2] = "Нарушителей: "+ dt.Rows.Count.ToString();



            sheet3.Columns[1].ColumnWidth = 23;
            sheet3.Columns[2].ColumnWidth = 20;
            sheet3.Columns[3].ColumnWidth = 20;
            sheet3.Columns[4].ColumnWidth = 15;
            sheet3.Columns[5].ColumnWidth = 40;
            sheet3.Columns[6].ColumnWidth = 20;
            sheet3.Columns[7].ColumnWidth = 40;
            sheet3.Columns[8].ColumnWidth = 10;
            sheet3.Columns[9].ColumnWidth = 40;
            sheet3.Columns[10].ColumnWidth = 20;

            sheet2.Columns[1].ColumnWidth = 20;
            sheet2.Columns[2].ColumnWidth = 40;
            sheet2.Columns[3].ColumnWidth = 10;
            sheet2.Columns[4].ColumnWidth = 30;
            sheet2.Columns[5].ColumnWidth = 30;
            sheet2.Columns[6].ColumnWidth = 15;
            sheet2.Columns[7].ColumnWidth = 20;            

            sheet.Columns[1].ColumnWidth = 20;
            sheet.Columns[2].ColumnWidth = 40;
            sheet.Columns[3].ColumnWidth = 10;
            sheet.Columns[4].ColumnWidth = 30;
            sheet.Columns[5].ColumnWidth = 40;
            sheet.Columns[6].ColumnWidth = 30;
            sheet.Columns[7].ColumnWidth = 15;
            sheet.Columns[8].ColumnWidth = 20;



            sheet3.Rows[1].RowHeight = 25;
            sheet2.Rows[1].RowHeight = 25;
            sheet.Rows[1].RowHeight = 25;

            sheet3.get_Range("A1", "J2").Font.Bold = true;       // жирный курсив
            sheet2.get_Range("A1", "G2").Font.Bold = true;       // жирный курсив
            sheet.get_Range("A1", "H2").Font.Bold = true;       // жирный курсив

            sheet3.get_Range("A1", "J2").VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            sheet2.get_Range("A1", "G2").VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            sheet.get_Range("A1", "H2").VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

            sheet3.get_Range("A1", "J2").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            sheet2.get_Range("A1", "G2").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            sheet.get_Range("A1", "H2").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            sheet3.get_Range("A1", "J1").Interior.Color = Color.NavajoWhite;
            sheet2.get_Range("A1", "G1").Interior.Color = Color.NavajoWhite;
            sheet.get_Range("A1", "H1").Interior.Color = Color.NavajoWhite;

            sheet3.get_Range("A2", "J2").Interior.Color = Color.DarkTurquoise;
            sheet2.get_Range("A2", "G2").Interior.Color = Color.DarkTurquoise;
            sheet.get_Range("A2", "H2").Interior.Color = Color.DarkTurquoise;

            sheet3.get_Range("A1", "J2").Cells.Font.Name = "Times New Roman";
            sheet2.get_Range("A1", "G2").Cells.Font.Name = "Times New Roman";
            sheet.get_Range("A1", "H2").Cells.Font.Name = "Times New Roman";

            sheet3.get_Range("A1", "J2").Cells.Font.Size = 10;
            sheet2.get_Range("A1", "G2").Cells.Font.Size = 10;
            sheet.get_Range("A1", "H2").Cells.Font.Size = 10;



            int BadBoysCount = 0;
            //Пример заполнения ячеек                        
            for (int i = 0; i < dataGridView1.RowCount - 1; i++) // строки
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++) // столбцы
                {
                    sheet3.Cells[i + 3, j + 1] = dataGridView1.Rows[i].Cells[j].Value;
                }
                BadBoysCount++;
            }
            sheet3.Cells[2, 1] = "Всего нарушителей - " + BadBoysCount.ToString();



            ///////ФОРМИРУЕМ ОТЧЕТ////////////////////////////////
            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
            {
                sheet.Cells[i + 3, 1] = dataGridView1.Rows[i].Cells[5].Value;
            }

            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
            {
                sheet.Cells[i + 3, 2] = dataGridView1.Rows[i].Cells[6].Value;
            }

            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
            {   
                sheet.Cells[i + 3, 3] = dataGridView1.Rows[i].Cells[7].Value;
            }

            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
            {
                sheet.Cells[i + 3, 4] = dataGridView1.Rows[i].Cells[8].Value;
            }

            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
            {
                sheet.Cells[i + 3, 5] = dataGridView1.Rows[i].Cells[4].Value;
            }

            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
            {
                sheet.Cells[i + 3, 6] = "Сообщили приставу.";
            }

            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
            {
                sheet.Cells[i + 3, 8] = dataGridView1.Rows[i].Cells[9].Value;
            }

            ///////ФОРМИРУЕМ ДЛЯ РУКОВОДСТВА///////////////
            for (int i = 0; i < dataGridView2.RowCount - 1; i++)
            {
                sheet2.Cells[i + 3, 1] = dataGridView2.Rows[i].Cells[4].Value;
            }
            for (int i = 0; i < dataGridView2.RowCount - 1; i++)
            {
                sheet2.Cells[i + 3, 2] = dataGridView2.Rows[i].Cells[5].Value;
            }

            for (int i = 0; i < dataGridView2.RowCount - 1; i++)
            {
                sheet2.Cells[i + 3, 3] = dataGridView2.Rows[i].Cells[6].Value;
            }

            for (int i = 0; i < dataGridView2.RowCount - 1; i++)
            {
                sheet2.Cells[i + 3, 4] = dataGridView2.Rows[i].Cells[7].Value;
            }
            for (int i = 0; i < dataGridView2.RowCount - 1; i++)
            {
                sheet2.Cells[i + 3, 7] = dataGridView2.Rows[i].Cells[9].Value;
            }


            int FsspGuys = 0;
            int MvdGuys = 0;
            int CovidGuys = 0;

            ////////РАСКРАШИВАЕМ ТАБЛИЦУ 3///////////////
            for (int i = 3; i < sheet3.UsedRange.Rows.Count + 1; i++)
            {
                if (sheet3.Cells[i, 9].Value == "МВД")
                {
                    sheet3.Cells[i, 9].Interior.Color = Color.Tomato;
                    MvdGuys++;
                    //sheet.get_Range(5, 7).Interior.Color = ColorTranslator.ToOle(Color.FromArgb(0x8B, 0x0, 0x0));
                }

                if (sheet3.Cells[i, 9].Value == " ФССП")
                {
                    sheet3.Cells[i, 9].Interior.Color = Color.LightGreen;
                    FsspGuys++;
                }

                if (sheet3.Cells[i, 9].Value.StartsWith(" КОВИД"))
                {
                    sheet3.Cells[i, 9].Interior.Color = Color.Yellow;
                    CovidGuys++;
                }
            }
            sheet3.Cells[2, 9] = "ФССП - " + FsspGuys.ToString() + "\n" + "МВД - " + MvdGuys.ToString() + "\n" + "Covid - " + CovidGuys.ToString();

            ////////РАСКРАШИВАЕМ ТАБЛИЦУ 2///////////////
            FsspGuys = 0;
            MvdGuys = 0;
            CovidGuys = 0;

            for (int i = 3; i < sheet2.UsedRange.Rows.Count + 1; i++)
            {
                if (sheet2.Cells[i, 4].Value == "МВД")
                {
                    sheet2.Cells[i, 4].Interior.Color = Color.Tomato;
                    MvdGuys++;
                }

                if (sheet2.Cells[i, 4].Value == " ФССП")
                {
                    sheet2.Cells[i, 4].Interior.Color = Color.LightGreen;
                    FsspGuys++;
                }

                if (sheet2.Cells[i, 4].Value.StartsWith(" КОВИД"))
                {
                    sheet2.Cells[i, 4].Interior.Color = Color.Yellow;
                    CovidGuys++;
                }
            }
            sheet2.Cells[2, 4] = "ФССП - " + FsspGuys.ToString() + "\n" + "МВД - " + MvdGuys.ToString() + "\n" + "Covid - " + CovidGuys.ToString();

            ////////РАСКРАШИВАЕМ ТАБЛИЦУ 1///////////////
            FsspGuys = 0;
            MvdGuys = 0;
            CovidGuys = 0;

            for (int i = 3; i < sheet.UsedRange.Rows.Count + 1; i++)
            {
                if (sheet.Cells[i, 4].Value == "МВД")
                {
                    sheet.Cells[i, 4].Interior.Color = Color.Tomato;
                    MvdGuys++;
                }

                if (sheet.Cells[i, 4].Value == " ФССП")
                {
                    sheet.Cells[i, 4].Interior.Color = Color.LightGreen;
                    FsspGuys++;
                }

                if (sheet.Cells[i, 4].Value.StartsWith(" КОВИД"))
                {
                    sheet.Cells[i, 4].Interior.Color = Color.Yellow;
                    CovidGuys++;
                }
            }
            sheet.Cells[2, 4] = "ФССП - " + FsspGuys.ToString() + "\n" + "МВД - " + MvdGuys.ToString() + "\n" + "Covid - " + CovidGuys.ToString();



            sheet3.UsedRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            sheet3.UsedRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

            sheet2.UsedRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            sheet2.UsedRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

            sheet.UsedRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            sheet.UsedRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

            ///////////////////////// переводим в денежный формат 3 таблицу
            for (int i = 3; i <= sheet3.UsedRange.Rows.Count; i++)
            {
                sheet3.Cells[i, 10] = Convert.ToDecimal(sheet3.Cells[i, 10].Value);
            }

            sheet3.get_Range("J3", "J" + (sheet3.UsedRange.Rows.Count)).NumberFormat = "#,##0.00 $";

            ///////////////////////// переводим в денежный формат 2 таблицу
            for (int i = 3; i <= sheet2.UsedRange.Rows.Count; i++)
            {
                sheet2.Cells[i, 7] = Convert.ToDecimal(sheet2.Cells[i, 7].Value);
            }

            sheet2.get_Range("J3", "G" + (sheet.UsedRange.Rows.Count)).NumberFormat = "#,##0.00 $";

            ///////////////////////// переводим в денежный формат 1 таблицу
            for (int i = 3; i <= sheet.UsedRange.Rows.Count; i++)
            {
                sheet.Cells[i, 8] = Convert.ToDecimal(sheet.Cells[i, 8].Value);
            }

            sheet.get_Range("J3", "H" + (sheet.UsedRange.Rows.Count)).NumberFormat = "#,##0.00 $";
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
                //
                //File.AppendAllText(@"C:\Users\zhukea\Desktop\result2.txt", result);
                //File.AppendAllText(@"C:\Users\zhukea\Desktop\AllText2.txt", AllText);
                //MessageBox.Show(AllText.Substring(A2, B2));
                //
                //string NEGODYAY = AllText.Substring(A2, B2 + 1);
                string NEGODYAY = AllText.Substring(A2, B2);

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

        public void ParsingResponse2()
        {

            // ВОТ ТУТ НАЧИНАЕТСЯ ПАРСИНГ 2
            string AllText = result;
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
                while (dataGridView2.Rows[j].Cells[9].Value.ToString() != "")
                {
                    j++;
                }
                dataGridView2.Rows[j].Cells[9].Value = ZADOLZHENOST;    //.ToString();

                AllText = AllText.Remove(0, B2 + 1);
                if (AllText != "")
                {
                    AllText = AllText.Remove(0, AllText.IndexOf('['));
                }


            }
        }

        private void button4_Click(object sender, EventArgs e)
        {

        }
    }
}
