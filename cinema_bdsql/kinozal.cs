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
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Rebar;
using System.Configuration;
using System.Net.Mail;
using System.Net;
using System.Xml.Linq;
using System.Diagnostics;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;
using System.Net.Sockets;

namespace cinema_bdsql
{
    public partial class kinozal : Form
    {
        string connectionString = ConfigurationManager.ConnectionStrings["MyDbConnection"].ConnectionString;
        private SqlConnection connection;
        // получаем все IPv4-адреса для текущего хоста и выбираем первый
        IPAddress ipAddress = Dns.GetHostAddresses(Dns.GetHostName())
            .FirstOrDefault(a => a.AddressFamily == AddressFamily.InterNetwork);

        Dictionary<Button, ButtonInfo> buttonInfos = new Dictionary<Button, ButtonInfo>();

        // глобальные переменные для работы во всей этой части кода
        int placeID;
        int TimetableID;
        string typeOfPayment;
        int summaplateja;
        int rowNumber = 0;
        int numberOfSeats = 0;

        public int timetableID { get; set; }
        public string ticketprice { get; set; }
        public string titlelabel { get; set; }
        public string halltitle { get; set; }
        public string startTimeLabel { get; set; }
        public string selecteddate { get; set; }
        public class ButtonInfo
        {
            public int Place_ID { get; set; }
        }
        private void InitializeButtonInfos()
        {
            // ДЛЯ HALL_ID = 3 

            // первый ряд
            buttonInfos.Add(button1, new ButtonInfo { Place_ID = 171 });
            buttonInfos.Add(button3, new ButtonInfo { Place_ID = 170 });
            buttonInfos.Add(button5, new ButtonInfo { Place_ID = 169 });
            buttonInfos.Add(button25, new ButtonInfo { Place_ID = 168 });
            buttonInfos.Add(button33, new ButtonInfo { Place_ID = 167 });
            buttonInfos.Add(button40, new ButtonInfo { Place_ID = 166 });
            buttonInfos.Add(button48, new ButtonInfo { Place_ID = 165 });
            buttonInfos.Add(button56, new ButtonInfo { Place_ID = 164 });
            buttonInfos.Add(button64, new ButtonInfo { Place_ID = 163 });
            buttonInfos.Add(button72, new ButtonInfo { Place_ID = 162});
            buttonInfos.Add(button80, new ButtonInfo { Place_ID = 161 });

            // второй ряд
            buttonInfos.Add(button2, new ButtonInfo { Place_ID = 182 });
            buttonInfos.Add(button4, new ButtonInfo { Place_ID = 181 });
            buttonInfos.Add(button6, new ButtonInfo { Place_ID = 180 });
            buttonInfos.Add(button24, new ButtonInfo { Place_ID = 179 });
            buttonInfos.Add(button32, new ButtonInfo { Place_ID = 178 });
            buttonInfos.Add(button39, new ButtonInfo { Place_ID = 177 });
            buttonInfos.Add(button47, new ButtonInfo { Place_ID = 176 });
            buttonInfos.Add(button55, new ButtonInfo { Place_ID = 175 });
            buttonInfos.Add(button63, new ButtonInfo { Place_ID = 174 });
            buttonInfos.Add(button71, new ButtonInfo { Place_ID = 173 });
            buttonInfos.Add(button79, new ButtonInfo { Place_ID = 172 });

            //третий ряд
            buttonInfos.Add(button7, new ButtonInfo { Place_ID = 191 });
            buttonInfos.Add(button23, new ButtonInfo { Place_ID = 190 });
            buttonInfos.Add(button31, new ButtonInfo { Place_ID = 189 });
            buttonInfos.Add(button38, new ButtonInfo { Place_ID = 188 });
            buttonInfos.Add(button46, new ButtonInfo { Place_ID = 187 });
            buttonInfos.Add(button54, new ButtonInfo { Place_ID = 186 });
            buttonInfos.Add(button62, new ButtonInfo { Place_ID = 185 });
            buttonInfos.Add(button70, new ButtonInfo { Place_ID = 184 });
            buttonInfos.Add(button78, new ButtonInfo { Place_ID = 183 });

            //четвертый ряд
            buttonInfos.Add(button8, new ButtonInfo { Place_ID = 200 });
            buttonInfos.Add(button22, new ButtonInfo { Place_ID = 199 });
            buttonInfos.Add(button30, new ButtonInfo { Place_ID = 198 });
            buttonInfos.Add(button37, new ButtonInfo { Place_ID = 197 });
            buttonInfos.Add(button45, new ButtonInfo { Place_ID = 196 });
            buttonInfos.Add(button53, new ButtonInfo { Place_ID = 195 });
            buttonInfos.Add(button61, new ButtonInfo { Place_ID = 194 });
            buttonInfos.Add(button69, new ButtonInfo { Place_ID = 193 });
            buttonInfos.Add(button77, new ButtonInfo { Place_ID = 192 });

            //пятый ряд
            buttonInfos.Add(button9, new ButtonInfo { Place_ID = 209 });
            buttonInfos.Add(button21, new ButtonInfo { Place_ID = 208 });
            buttonInfos.Add(button29, new ButtonInfo { Place_ID = 207 });
            buttonInfos.Add(button36, new ButtonInfo { Place_ID = 206 });
            buttonInfos.Add(button44, new ButtonInfo { Place_ID = 205 });
            buttonInfos.Add(button52, new ButtonInfo { Place_ID = 204 });
            buttonInfos.Add(button60, new ButtonInfo { Place_ID = 203 });
            buttonInfos.Add(button68, new ButtonInfo { Place_ID = 202 });
            buttonInfos.Add(button76, new ButtonInfo { Place_ID = 201 });

            //шестой ряд
            buttonInfos.Add(button10, new ButtonInfo { Place_ID = 218 });
            buttonInfos.Add(button20, new ButtonInfo { Place_ID = 217 });
            buttonInfos.Add(button28, new ButtonInfo { Place_ID = 216 });
            buttonInfos.Add(button35, new ButtonInfo { Place_ID = 215 });
            buttonInfos.Add(button43, new ButtonInfo { Place_ID = 214 });
            buttonInfos.Add(button51, new ButtonInfo { Place_ID = 213 });
            buttonInfos.Add(button59, new ButtonInfo { Place_ID = 212 });
            buttonInfos.Add(button67, new ButtonInfo { Place_ID = 211 });
            buttonInfos.Add(button75, new ButtonInfo { Place_ID = 210 });

            //седьмой ряд
            buttonInfos.Add(button15, new ButtonInfo { Place_ID = 229 });
            buttonInfos.Add(button16, new ButtonInfo { Place_ID = 228 });
            buttonInfos.Add(button11, new ButtonInfo { Place_ID = 227 });
            buttonInfos.Add(button19, new ButtonInfo { Place_ID = 226 });
            buttonInfos.Add(button27, new ButtonInfo { Place_ID = 225 });
            buttonInfos.Add(button34, new ButtonInfo { Place_ID = 224 });
            buttonInfos.Add(button42, new ButtonInfo { Place_ID = 223 });
            buttonInfos.Add(button50, new ButtonInfo { Place_ID = 222 });
            buttonInfos.Add(button58, new ButtonInfo { Place_ID = 221 });
            buttonInfos.Add(button66, new ButtonInfo { Place_ID = 220 });
            buttonInfos.Add(button74, new ButtonInfo { Place_ID = 219 });

            //восьмой ряд
            buttonInfos.Add(button14, new ButtonInfo { Place_ID = 240 });
            buttonInfos.Add(button13, new ButtonInfo { Place_ID = 239 });
            buttonInfos.Add(button12, new ButtonInfo { Place_ID = 238 });
            buttonInfos.Add(button18, new ButtonInfo { Place_ID = 237 });
            buttonInfos.Add(button26, new ButtonInfo { Place_ID = 236 });
            buttonInfos.Add(button41, new ButtonInfo { Place_ID = 235 });
            buttonInfos.Add(button49, new ButtonInfo { Place_ID = 234 });
            buttonInfos.Add(button57, new ButtonInfo { Place_ID = 233 });
            buttonInfos.Add(button65, new ButtonInfo { Place_ID = 232 });
            buttonInfos.Add(button73, new ButtonInfo { Place_ID = 231 });
            buttonInfos.Add(button81, new ButtonInfo { Place_ID = 230 });
        }

        public kinozal()
        {
            InitializeComponent();
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.BackColor = Color.FromArgb(0, 28, 45);

            foreach (Control control in this.Controls) 
            {
                if (control is Button && control != button82 )
                {
                    control.BackColor = Color.FromArgb(48, 166, 241); 
                }
            }
            button82.BackColor = Color.FromArgb(228, 26, 105); // кнопка "купить билеты"
            
        }
        private void kinozal_Load(object sender, EventArgs e) // покраска
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

            foreach (Control control in this.Controls)
            {
                if (control is Button && control != button82 )
                {
                    control.Tag = false;
                    control.Click += new EventHandler(Button_Click);
                }
            }
            label19.Text = ticketprice + " " + " рублей ";
            InitializeButtonInfos();
            CheckButtonColor();
        }


       
        private void Button_Click(object sender, EventArgs e) // меняем цвет кнопочки при выборе (на красный и обратно - если билет не нужен)
        {
            Button button = sender as Button;
            if (button != null)
            {
                bool isRed = (bool)button.Tag;
                button.BackColor = isRed ? Color.FromArgb(48, 166, 241) : Color.FromArgb(228, 26, 105); // замените на нужные цвета
                button.Tag = !isRed;   
            }
        }
        private int k = 0;
        private void button82_Click(object sender, EventArgs e)
        {
            List<Tuple<int, int>> selectedPlaces = new List<Tuple<int, int>>();

            foreach (Control control in this.Controls)
            {
                if (control is Button button && button != button82 && button.BackColor == Color.FromArgb(228, 26, 105))
                {
                    k++;
                    ButtonInfo buttonInfo = buttonInfos[button];

                    placeID = buttonInfo.Place_ID;
                    TimetableID = timetableID;
                    typeOfPayment = comboBox1.Text;

                    // Получаем информацию о ряде и месте на основе Place_ID
                    string query = "SELECT [Row number], [Number of seats] FROM Places WHERE Place_ID = @PlaceID";
                    SqlCommand command = new SqlCommand(query, connection);
                    command.Parameters.AddWithValue("@PlaceID", placeID);
                    SqlDataReader reader = command.ExecuteReader();

                    if (reader.Read())
                    {
                        rowNumber = reader.GetInt32(0);
                        numberOfSeats = reader.GetInt32(1);

                        // Добавляем ряд и место в список
                        selectedPlaces.Add(new Tuple<int, int>(rowNumber, numberOfSeats));
                    }
                    reader.Close();

                    string query2 = "INSERT INTO Buy (Place_ID, Timetable_ID, Type_of_payment) VALUES (@PlaceID, @TimetableID, @TypeOfPayment)";

                    SqlCommand command2 = new SqlCommand(query2, connection);
                    command2.Parameters.AddWithValue("@PlaceID", placeID);
                    command2.Parameters.AddWithValue("@TimetableID", TimetableID);
                    command2.Parameters.AddWithValue("@TypeOfPayment", typeOfPayment);
                    command2.ExecuteNonQuery();
                    // изменяем цвет кнопки после выполнения команды
                    button.BackColor = Color.Red;
                    button.Enabled = false;
                }
            }

            // Выводим общую стоимость билетов
            int totalPrice = Convert.ToInt32(ticketprice) * selectedPlaces.Count;
            summaplateja = totalPrice;
            PrintWord();
            SendBill();  
        }

        private void CheckButtonColor()
        {
            //получить все записи из таблицы Buy
            string query = "SELECT * FROM Buy";

            // Создаем SqlCommand и выполняем запрос
            SqlCommand command = new SqlCommand(query, connection);
            SqlDataReader reader = command.ExecuteReader();

            // проходимся по каждой записи в таблице Buy
            while (reader.Read())
            {
                int placeID = (int)reader["Place_ID"];

                // ищем кнопку, у которой Place_ID соответствует записи в таблице Buy
                foreach (Control control in this.Controls)
                {
                    if (control is Button button && button != button82 && button != button17)
                    {
                        ButtonInfo buttonInfo = buttonInfos[button];
                        if (buttonInfo.Place_ID == placeID)
                        {
                            // закрашиваем кнопку соответствующим цветом
                            button.BackColor = Color.FromArgb(228, 26, 105);
                            break;
                        }
                    }
                }
            }
            reader.Close();
        }
        public void PrintWord()// sozdaem .pdf 
        {
            string[] data = new string[6];
            var path = System.IO.Path.GetFullPath(@"Doc1.docx"); //расположение путь файла шаблона
            Microsoft.Office.Interop.Word.Application word = new Application();
            Document docW = word.Documents.Open(path);
            Bookmarks wBookmarks = docW.Bookmarks;
            Range wRange;

            data[5] = titlelabel;//title
            data[1] = halltitle; //hall
            data[2] = Convert.ToString(summaplateja) +  " " + "rub"; //price
            data[0] = selecteddate;//date
            data[4] = startTimeLabel; //time
            data[3] = "Ряд:" + " " + rowNumber + " " + "Место:" + " " + numberOfSeats; //row,numder of seats
                      
            int i = 0;
            foreach (Bookmark mark in wBookmarks)//находит закладку и заменяет ее текстом введенным в массиве
            {
                wRange = mark.Range;
                wRange.Text = data[i];
                i++;
            }
            var pathPDF = System.IO.Path.GetFullPath(@"билеты\Чек");//создаёт файл pdf по прописанному пути
            docW.ExportAsFixedFormat(pathPDF, WdExportFormat.wdExportFormatPDF);//Сохраняет документ в формате PDF 
            MessageBox.Show("Билеты куплены!", "Успешно");
            word.Quit(false);
        }

        private void SendBill() 
        {
            int port = 587;
            bool enableSSL = true;

            // Переменные emailFrom, password и emailTo - это строки, которые хранят информацию об отправителе, получателе и учетных данных для отправки email.
            string emailFrom = "testpochtapdf@mail.ru";
            string password = "WvpDBM7r18H6cewwjcmu";
            string emailTo = "rinat.mingalimov21@mail.ru";
            string subject = "Оплата билетов в кинотеатр - family MOVIE NighT";
            string body = $"Спасибо за покупку!";

            string smtpAdress = "smtp.mail.ru";

            System.Net.Mail.MailMessage mail = new System.Net.Mail.MailMessage();

            mail.From = new MailAddress(emailFrom);
            mail.To.Add(emailTo);
            mail.Subject = subject;
            mail.Body = body;
            mail.IsBodyHtml = true;

            // Создаем объект Attachment и добавляем в коллекцию Attachments MailMessage
            var pathPDF1 = System.IO.Path.GetFullPath(@"билеты\Чек.pdf");//создаёт файл pdf по прописанному пути
            Attachment attachment = new Attachment(pathPDF1);//выбираем файл отправляем чек
            mail.Attachments.Add(attachment);

            using (SmtpClient smtp = new SmtpClient(smtpAdress, port))
            {
                smtp.Credentials = new NetworkCredential(emailFrom, password);
                smtp.EnableSsl = enableSSL;
                try
                {
                    smtp.Send(mail);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
            }
        }

        private void button83_Click(object sender, EventArgs e)
        {
            
        }

        private void button49_Click(object sender, EventArgs e)
        {

        }
    }
}
