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


namespace FSSP_v2
{
    class ListDataRow
    {
        public string firstname;   // Имя
        public string lastname;    // Фамилия
        public string secondname;  // Отчество
        public string birthdate;   // День рожденье

        public void GetListViolators()
        {

        }

    }

    public partial class Form1 : Form
    {
        private DataSet ds = new DataSet();
        private DataTable dt = new DataTable();
        int CountViolators; // количество нарушителей

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
            if (CountViolators <= 50)
            {
                MessageBox.Show("Меньше 50, либо равно 50");

                string LICA = "";
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
                }
                    if (LICA.Length - 2 == LICA.LastIndexOf(','))
                    {
                        LICA = LICA.Remove(LICA.LastIndexOf(','));
                    }
                
              string PostGroupFssp = "";
              PostGroupFssp = "{\n" +
              "\t\"token\": \"aMYRXmjGwPFm\",\n" +
              "\t\"request\": [\n" +
              LICA +"\n"+
              "\t]\n" +
              "}";
              richTextBox1.Text = PostGroupFssp;

            }
            else
            {
                MessageBox.Show("Больше 50 ");                
                double CountPosts = Math.Ceiling(Convert.ToDouble(CountViolators) / 50);                
                MessageBox.Show(CountPosts.ToString()+"   ");

                for (int i = 1; i <CountPosts ; i++)
                {

                }


                /*
                List <DataRow> batch = new DataRowCollection();
                foreach (var item in dt.Rows)
                {
                    batch.Add(item);

                    if (batch.Count == 50)
                    {
                        // Perform operation on batch
                        batch.Clear();
                    }
                }

                // Process last batch
                if (batch.Any())
                {


                    /*

                     int batchsize = 5;
                    List<string> colection = new List<string> { "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"};
                    for (int x = 0; x < Math.Ceiling((decimal)colection.Count / batchsize); x++)
                    {
                    var t = colection.Skip(x * batchsize).Take(batchsize);
                        }

                     */


                    // Perform operation on batch

                }
            }


    
        
    }
}
