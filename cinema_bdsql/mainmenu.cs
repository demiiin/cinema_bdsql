using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using Microsoft.IdentityModel.Protocols;
using System.Configuration;
using System.Net;
using System.Net.Sockets;

namespace cinema_bdsql
{
    public partial class mainmenu : Form
    {
        string connectionString = ConfigurationManager.ConnectionStrings["MyDbConnection"].ConnectionString;
        private SqlConnection connection;
        // получаем все IPv4-адреса для текущего хоста и выбираем первый
        IPAddress ipAddress = Dns.GetHostAddresses(Dns.GetHostName())
            .FirstOrDefault(a => a.AddressFamily == AddressFamily.InterNetwork);

        public mainmenu()
        {
            InitializeComponent();
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            
        }
        private void mainmenu_Load(object sender, EventArgs e)
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
            LoadDates();
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
           
        }   
        private void LoadDates()
        {
            string query = "DECLARE @dt datetimeoffset = switchoffset(CONVERT(datetimeoffset, GETDATE()), '-04:00');\r\nSELECT DISTINCT [Date] FROM Timetable\r\nWHERE [Date] > @dt\r\nORDER BY [Date] ASC\r\nOPTION (RECOMPILE);";
            SqlCommand command = new SqlCommand(query, connection); 
            try
            {
                SqlDataReader reader = command.ExecuteReader();

                int labelX = 10; // start X pos
                int labelY = 90; // start Y pos

                while (reader.Read())
                {
                    DateTime date = reader.GetDateTime(0);
                    string labelText = date.ToShortDateString();

                    Label label = new Label();
                    label.Font = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);
                    label.Text = labelText;
                    label.Location = new Point(labelX, labelY);
                    label.AutoSize = true;
                    label.Tag = date; // сохраняем дату в свойстве Tag лейбла

                    label.Click += new EventHandler(Label_Click); // добавляем обработчик события Click
                   
                    this.Controls.Add(label);

                    labelX += label.Width + 10; // добавляем лейбл, немного отдаляя его каждый раз
                }
                reader.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }
        private void Label_Click(object sender, EventArgs e)
        {
            panel1.Controls.Clear();

            Label label = (Label)sender; // получаем лейбл, на который нажали                       
            DateTime selectedDate = (DateTime)label.Tag; // получаем дату из свойства Tag

            if (Convert.ToString(selectedDate) == "03.05.2023 0:00:00")
            {
                SqlDataAdapter dataAdapter = new SqlDataAdapter(new SqlCommand("SELECT Poster FROM Films WHERE Film_ID = 7", connection));
                DataSet dataSet = new DataSet();
                dataAdapter.Fill(dataSet);

                if (dataSet.Tables[0].Rows.Count == 1)
                {
                    Byte[] data = new Byte[0];
                    data = (Byte[])(dataSet.Tables[0].Rows[0]["Poster"]);
                    MemoryStream mem = new MemoryStream(data);
                    PictureBox pictureBox = new PictureBox();
                    pictureBox.Image = Image.FromStream(mem);
                    pictureBox.Location = new Point(50, 0);
                    pictureBox.SizeMode = PictureBoxSizeMode.AutoSize;
                    panel1.Controls.Add(pictureBox);
                }
            }
            else if (Convert.ToString(selectedDate) == "04.05.2023 0:00:00")
            {
                SqlDataAdapter dataAdapter = new SqlDataAdapter(new SqlCommand("SELECT Poster FROM Films WHERE Film_ID = 1", connection));
                DataSet dataSet = new DataSet();
                dataAdapter.Fill(dataSet);

                if (dataSet.Tables[0].Rows.Count == 1)
                {
                    Byte[] data = new Byte[0];
                    data = (Byte[])(dataSet.Tables[0].Rows[0]["Poster"]);
                    MemoryStream mem = new MemoryStream(data);
                    PictureBox pictureBox = new PictureBox();
                    pictureBox.Image = Image.FromStream(mem);
                    pictureBox.Location = new Point(50, 0);
                    pictureBox.SizeMode = PictureBoxSizeMode.AutoSize;
                    panel1.Controls.Add(pictureBox);
                }
            }

            string query = @"SELECT
    CASE
        WHEN ROW_NUMBER() OVER (PARTITION BY Films.Title ORDER BY Sessions.Session_start) = 1
        THEN Films.Title
        ELSE ''
    END AS FilmTitle,
    CONVERT(NVARCHAR(MAX), ISNULL(Sessions.Session_start, ''), 8) AS StartTime,
    ISNULL(Timetable.[Date], '') AS ShowDate,
    ISNULL(Halls.[Title hall], '') AS HallTitle,
    Timetable.Timetable_ID,
    CONVERT(NVARCHAR(MAX), Films.[Base price] * Sessions.Koef, 2) AS TicketPrice
FROM Films
INNER JOIN Timetable ON Films.Film_ID = Timetable.Film_ID
INNER JOIN Sessions ON Timetable.Session_ID = Sessions.Session_ID
INNER JOIN Halls ON Timetable.Hall_ID = Halls.Hall_ID
WHERE Timetable.[Date] = @SelectedDate
ORDER BY Films.Title, Sessions.Session_start;";

            SqlCommand command = new SqlCommand(query, connection);
            command.Parameters.AddWithValue("@SelectedDate", selectedDate);

            try
            {
                SqlDataReader reader = command.ExecuteReader();
               
                int labelX = 220; // start X position
                int labelY = 20; // start Y position

                while (reader.Read())
                {
                    Label titleLabel = new Label();
                    titleLabel.Font = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);

                    titleLabel.Text = reader.GetString(0); //НАЗВАНИЕ ФИЛЬМА

                    titleLabel.Location = new Point(labelX, labelY);
                    titleLabel.AutoSize = true;
                    
                    Label startTimeLabel = new Label();
                    startTimeLabel.Font = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);

                    startTimeLabel.Text = reader.GetString(1); // ВРЕМЯ СТАРТА ФИЛЬМА

                    startTimeLabel.Location = new Point(labelX + titleLabel.Width + 150, labelY);
                    startTimeLabel.AutoSize = true;

                    Label hallLabel = new Label();
                    hallLabel.Font = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);

                    hallLabel.Text = reader.IsDBNull(3) ? "" : reader.GetString(3); //ЗАЛЛ ФИЛЬМА

                    hallLabel.Location = new Point(labelX + titleLabel.Width + startTimeLabel.Width + 300, labelY);
                    hallLabel.AutoSize = true;

                    int timetable_id = reader.GetInt32(4); // получение значения из столбца Timetable_ID

                    string ticketPrice = reader.GetString(5); // получение значения из столбца ticketPrice

                    startTimeLabel.Click += (s, ev) =>
                    {
                        string hallName = hallLabel.Text;
                        if (!string.IsNullOrEmpty(hallName))
                        {
                            string formTitle = Convert.ToString(titleLabel.Text + " - " + hallLabel.Text + " - " + startTimeLabel.Text);
                            kinozal kinozal3 = null;
                            kinozal2 kinozal4 = null;

                            switch (hallName)
                            {
                                    case "Маленький зал №3":
                                        kinozal3 = new kinozal();
                                        kinozal3.Text = formTitle;
                                        kinozal3.timetableID = timetable_id;
                                        kinozal3.ticketprice = ticketPrice;
                                        kinozal3.titlelabel = titleLabel.Text;
                                        kinozal3.halltitle = hallLabel.Text;
                                        kinozal3.startTimeLabel = startTimeLabel.Text;
                                        kinozal3.selecteddate = selectedDate.ToString();
                                        kinozal3.Show();
                                    break;
                                    case "Большой зал №4":
                                        kinozal4 = new kinozal2();
                                        kinozal4.Text = formTitle;
                                        kinozal4.timetableID = timetable_id;
                                        kinozal4.ticketprice = ticketPrice;
                                        kinozal4.titlelabel = titleLabel.Text;
                                        kinozal4.halltitle = hallLabel.Text;
                                        kinozal4.startTimeLabel = startTimeLabel.Text;
                                        kinozal4.selecteddate = selectedDate.ToString();
                                    kinozal4.Show();
                                    break;
                                default:
                                        MessageBox.Show("Что-то пошло не так.....");
                                    break;
                            } 
                        }
                    };

                    panel1.Controls.Add(titleLabel); // добавляем на панель
                    panel1.Controls.Add(startTimeLabel); // добавляем на панель
                    
                    panel1.AutoScroll = true;

                    labelY += titleLabel.Height + 150;
                }
                reader.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            Form1 form = new Form1();
            this.Close();
            form.Show();
        } 
    }
}
