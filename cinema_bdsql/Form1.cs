using Microsoft.IdentityModel.Protocols;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.Sockets;

namespace cinema_bdsql
{
    public partial class Form1 : Form
    {
        string connectionString = ConfigurationManager.ConnectionStrings["MyDbConnection"].ConnectionString;
        private SqlConnection connection;
        // получаем все IPv4-адреса для текущего хоста и выбираем первый
        IPAddress ipAddress = Dns.GetHostAddresses(Dns.GetHostName())
            .FirstOrDefault(a => a.AddressFamily == AddressFamily.InterNetwork);

        public Form1()
        {
            InitializeComponent();

            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            textBox2.UseSystemPasswordChar = true;
            textBox2.PasswordChar = '*';
        }
        private bool ValidateUser(string username, string password)
        {
            SqlCommand command = new SqlCommand("SELECT COUNT(*) FROM Users WHERE Username = @UserName AND UserPassword = @UserPassword", connection);
            command.Parameters.AddWithValue("@UserName", username);
            command.Parameters.AddWithValue("@UserPassword", password);

            int result = (int)command.ExecuteScalar();
            return result > 0;
        }
        private void button1_Click(object sender, EventArgs e)
        {

            string username = textBox1.Text;
            string password = textBox2.Text;

            if (ValidateUser(username, password))
            {
                SqlCommand command = new SqlCommand("SELECT ROLE_ID FROM Users WHERE Username = @Username", connection);
                command.Parameters.AddWithValue("@Username", username);

                int roleId = (int)command.ExecuteScalar();

                if (roleId == 2) // for admins
                {
                    adminform adminform = new adminform();
                    this.Hide();
                    adminform.Show();
                }
                else if (roleId == 3) // for users
                {
                    mainmenu mainmenu = new mainmenu();
                    this.Hide();
                    mainmenu.Show();
                }
            }
            else
            {
                MessageBox.Show("Неверное имя пользователя или пароль");
            }
        }
        private void Form1_Load(object sender, EventArgs e)
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
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (connection != null) 
            {
                connection.Close();
            }
            else
            {
                MessageBox.Show("Произошла ошибка, проверьте подключение!");
            }     
        }
    }
}
