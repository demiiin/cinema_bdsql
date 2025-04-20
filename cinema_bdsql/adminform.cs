using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Net.Sockets;
using System.Net;

namespace cinema_bdsql
{
    public partial class adminform : Form
    {
        string connectionString = ConfigurationManager.ConnectionStrings["MyDbConnection"].ConnectionString;
        private SqlConnection connection;
        // получаем все IPv4-адреса для текущего хоста и выбираем первый
        IPAddress ipAddress = Dns.GetHostAddresses(Dns.GetHostName())
            .FirstOrDefault(a => a.AddressFamily == AddressFamily.InterNetwork);
        public adminform()
        {
            InitializeComponent();
        }

        public void Get_Info_About_Films() // получаем фильмы текстбокс3
        {
            // создаем SQL-запрос для получения списка фильмов
            string query1 = "SELECT Title FROM Films";

            // создаем команду на основе запроса и соединения
            SqlCommand cmd1 = new SqlCommand(query1, connection);

            // выполняем запрос и получаем объект для чтения данных
            SqlDataReader reader1 = cmd1.ExecuteReader();

            // очищаем список в combobox4
            comboBox4.Items.Clear();
            comboBox2.Items.Clear();

            // добавляем фильмы в combobox4
            while (reader1.Read())
            {
                string title = reader1.GetString(0);
                comboBox4.Items.Add(title);
                comboBox2.Items.Add(title);
            }
            reader1.Close();
        }

        private void adminform_Load(object sender, EventArgs e) 
        {
            // Заменяем placeholder на значение IPv4-адреса
            string formattedConnectionString = string.Format(connectionString, ipAddress);

            // Создаем новое подключение к базе данных
            connection = new SqlConnection(formattedConnectionString);
            try
            {
                connection.Open();
            }
            catch
            {
                MessageBox.Show("Ошибка подключения к БД!");
            }

            Get_Info_About_Films();
        }

        private void adminform_FormClosing(object sender, FormClosingEventArgs e)
        {
            connection.Close();
        }

        private void button1_Click(object sender, EventArgs e) // ДОБАВИТЬ ФИЛЬМ
        {
            string query = "INSERT INTO Films (Title, [Show Date], [Base Price]) VALUES (@title, @showDate, @basePrice)";

            using (SqlCommand command = new SqlCommand(query, connection))
            {
                command.Parameters.AddWithValue("@title", textBox5.Text);
                command.Parameters.AddWithValue("@showDate", Convert.ToDateTime(textBox4.Text));
                command.Parameters.AddWithValue("@basePrice", Convert.ToDecimal(textBox3.Text));

                command.ExecuteNonQuery();

                MessageBox.Show("Фильм успешно добавлен!");
                comboBox4.Items.Clear();
                comboBox2.Items.Clear();
                Get_Info_About_Films();
            }
        }

        private void button2_Click(object sender, EventArgs e) // добавить сеанс
        {
            int filmId = 0;
            string filmTitle = comboBox2.Text;
            string date = textBox1.Text;
            int hallId = comboBox1.SelectedIndex + 1;

            string query = "SELECT Film_ID FROM Films WHERE Title = @title";

            using (SqlCommand command = new SqlCommand(query, connection))
            {
                command.Parameters.AddWithValue("@title", filmTitle);

                object result = command.ExecuteScalar();

                if (result != null)
                {
                    filmId = Convert.ToInt32(result);
                }
            }

            string query2 = "INSERT INTO TimeTable (Date, Film_ID, Session_ID, Hall_ID) VALUES (@date, @filmId, @sessionid, @hallId)";

            using (SqlCommand command = new SqlCommand(query2, connection))
            {
                command.Parameters.AddWithValue("@date", date);
                command.Parameters.AddWithValue("@filmId", filmId);
                command.Parameters.AddWithValue("@sessionid", 3);
                command.Parameters.AddWithValue("@hallId", hallId);

                command.ExecuteNonQuery();
            }
        }

        private void button5_Click(object sender, EventArgs e) // удалить фильм по его названию
        {
            // Получаем ID фильма, который нужно удалить
            int filmId = 0;
            string filmTitle = comboBox4.Text;

            string query = "SELECT Film_ID FROM Films WHERE Title = @title";

            using (SqlCommand command = new SqlCommand(query, connection))
            {
                command.Parameters.AddWithValue("@title", filmTitle);

                object result = command.ExecuteScalar();

                if (result != null)
                {
                    filmId = Convert.ToInt32(result);
                }
            }

            // Удаляем связанные записи из таблицы TimeTable
            query = "DELETE FROM TimeTable WHERE Film_ID = @filmId";

            using (SqlCommand command = new SqlCommand(query, connection))
            {
                command.Parameters.AddWithValue("@filmId", filmId);
                command.ExecuteNonQuery();
            }

            // Удаляем запись из таблицы Films
            query = "DELETE FROM Films WHERE Film_ID = @filmId";

            using (SqlCommand command = new SqlCommand(query, connection))
            {
                command.Parameters.AddWithValue("@filmId", filmId);
                command.ExecuteNonQuery();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            mainmenu mainmenu = new mainmenu();
            this.Close();
            mainmenu.Show();
        }
    }
}
