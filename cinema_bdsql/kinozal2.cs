using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Net.Mail;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Word.Application;
using System.Reflection.Emit;

namespace cinema_bdsql
{
    public partial class kinozal2 : Form
    {
        private SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["MyDbConnection"].ConnectionString);
        Dictionary<Button, ButtonInfo> buttonInfos = new Dictionary<Button, ButtonInfo>();

        int placeID;
        int TimetableID;
        string typeOfPayment;


        //для формирования чека
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
            // ДЛЯ HALL_ID = 4
            // первый ряд
            buttonInfos.Add(button70, new ButtonInfo { Place_ID = 241 });
            buttonInfos.Add(button66, new ButtonInfo { Place_ID = 242 });
            buttonInfos.Add(button68, new ButtonInfo { Place_ID = 243 });
            buttonInfos.Add(button56, new ButtonInfo { Place_ID = 244 });
            buttonInfos.Add(button64, new ButtonInfo { Place_ID = 245 });
            buttonInfos.Add(button40, new ButtonInfo { Place_ID = 246 });
            buttonInfos.Add(button48, new ButtonInfo { Place_ID = 247 });
            buttonInfos.Add(button24, new ButtonInfo { Place_ID = 248 });
            buttonInfos.Add(button32, new ButtonInfo { Place_ID = 249 });
            buttonInfos.Add(button16, new ButtonInfo { Place_ID = 250 });
            buttonInfos.Add(button5, new ButtonInfo { Place_ID = 251 });
            // второй ряд
            buttonInfos.Add(button69, new ButtonInfo { Place_ID = 252 });
            buttonInfos.Add(button65, new ButtonInfo { Place_ID = 253 });
            buttonInfos.Add(button67, new ButtonInfo { Place_ID = 254 });
            buttonInfos.Add(button55, new ButtonInfo { Place_ID = 255 });
            buttonInfos.Add(button63, new ButtonInfo { Place_ID = 256 });
            buttonInfos.Add(button39, new ButtonInfo { Place_ID = 257 });
            buttonInfos.Add(button47, new ButtonInfo { Place_ID = 258 });
            buttonInfos.Add(button23, new ButtonInfo { Place_ID = 259 });
            buttonInfos.Add(button31, new ButtonInfo { Place_ID = 260 });
            buttonInfos.Add(button15, new ButtonInfo { Place_ID = 261 });
            buttonInfos.Add(button6, new ButtonInfo { Place_ID = 262 });
            //третий ряд
            buttonInfos.Add(button72, new ButtonInfo { Place_ID = 263 });
            buttonInfos.Add(button54, new ButtonInfo { Place_ID = 264 });
            buttonInfos.Add(button62, new ButtonInfo { Place_ID = 265 });
            buttonInfos.Add(button38, new ButtonInfo { Place_ID = 266 });
            buttonInfos.Add(button46, new ButtonInfo { Place_ID = 267 });
            buttonInfos.Add(button22, new ButtonInfo { Place_ID = 268 });
            buttonInfos.Add(button30, new ButtonInfo { Place_ID = 269 });
            buttonInfos.Add(button14, new ButtonInfo { Place_ID = 270 });
            buttonInfos.Add(button7, new ButtonInfo { Place_ID = 271 });

            //четвертый ряд
            buttonInfos.Add(button71, new ButtonInfo { Place_ID = 272 });
            buttonInfos.Add(button53, new ButtonInfo { Place_ID = 273 });
            buttonInfos.Add(button61, new ButtonInfo { Place_ID = 274 });
            buttonInfos.Add(button37, new ButtonInfo { Place_ID = 275 });
            buttonInfos.Add(button45, new ButtonInfo { Place_ID = 276 });
            buttonInfos.Add(button21, new ButtonInfo { Place_ID = 277 });
            buttonInfos.Add(button29, new ButtonInfo { Place_ID = 278 });
            buttonInfos.Add(button13, new ButtonInfo { Place_ID = 279 });
            buttonInfos.Add(button8, new ButtonInfo { Place_ID = 280 });

            //пятый ряд
            buttonInfos.Add(button74, new ButtonInfo { Place_ID = 281 });
            buttonInfos.Add(button52, new ButtonInfo { Place_ID = 282 });
            buttonInfos.Add(button60, new ButtonInfo { Place_ID = 283 });
            buttonInfos.Add(button36, new ButtonInfo { Place_ID = 284 });
            buttonInfos.Add(button44, new ButtonInfo { Place_ID = 285 });
            buttonInfos.Add(button20, new ButtonInfo { Place_ID = 286 });
            buttonInfos.Add(button28, new ButtonInfo { Place_ID = 287 });
            buttonInfos.Add(button12, new ButtonInfo { Place_ID = 288 });
            buttonInfos.Add(button9, new ButtonInfo { Place_ID = 289 });
            //шестой ряд
            buttonInfos.Add(button73, new ButtonInfo { Place_ID = 290 });
            buttonInfos.Add(button51, new ButtonInfo { Place_ID = 291 });
            buttonInfos.Add(button59, new ButtonInfo { Place_ID = 292 });
            buttonInfos.Add(button35, new ButtonInfo { Place_ID = 293 });
            buttonInfos.Add(button43, new ButtonInfo { Place_ID = 294 });
            buttonInfos.Add(button19, new ButtonInfo { Place_ID = 295 });
            buttonInfos.Add(button27, new ButtonInfo { Place_ID = 296 });
            buttonInfos.Add(button4, new ButtonInfo { Place_ID = 297 });
            buttonInfos.Add(button10, new ButtonInfo { Place_ID = 298 });
            //седьмой ряд
            buttonInfos.Add(button90, new ButtonInfo { Place_ID = 299 });
            buttonInfos.Add(button92, new ButtonInfo { Place_ID = 300 });
            buttonInfos.Add(button76, new ButtonInfo { Place_ID = 301 });
            buttonInfos.Add(button50, new ButtonInfo { Place_ID = 302 });
            buttonInfos.Add(button58, new ButtonInfo { Place_ID = 303 });
            buttonInfos.Add(button34, new ButtonInfo { Place_ID = 304 });
            buttonInfos.Add(button42, new ButtonInfo { Place_ID = 305 });
            buttonInfos.Add(button18, new ButtonInfo { Place_ID = 306 });
            buttonInfos.Add(button26, new ButtonInfo { Place_ID = 307 });
            buttonInfos.Add(button3, new ButtonInfo { Place_ID = 308 });
            buttonInfos.Add(button11, new ButtonInfo { Place_ID = 309 });
            //восьмой ряд

            buttonInfos.Add(button89, new ButtonInfo { Place_ID = 310 });
            buttonInfos.Add(button91, new ButtonInfo { Place_ID = 311 });
            buttonInfos.Add(button75, new ButtonInfo { Place_ID = 312 });
            buttonInfos.Add(button49, new ButtonInfo { Place_ID = 313 });
            buttonInfos.Add(button57, new ButtonInfo { Place_ID = 314 });
            buttonInfos.Add(button33, new ButtonInfo { Place_ID = 315 });
            buttonInfos.Add(button41, new ButtonInfo { Place_ID = 316 });
            buttonInfos.Add(button17, new ButtonInfo { Place_ID = 317 });
            buttonInfos.Add(button25, new ButtonInfo { Place_ID = 318 });
            buttonInfos.Add(button2, new ButtonInfo { Place_ID = 319 });
            buttonInfos.Add(button1, new ButtonInfo { Place_ID = 320 });
            //девятый ряд

            buttonInfos.Add(button87, new ButtonInfo { Place_ID = 321 });
            buttonInfos.Add(button88, new ButtonInfo { Place_ID = 322 });
            buttonInfos.Add(button85, new ButtonInfo { Place_ID = 323 });
            buttonInfos.Add(button86, new ButtonInfo { Place_ID = 324 });
            buttonInfos.Add(button83, new ButtonInfo { Place_ID = 325 });
            buttonInfos.Add(button84, new ButtonInfo { Place_ID = 326 });
            buttonInfos.Add(button81, new ButtonInfo { Place_ID = 327 });
            buttonInfos.Add(button82, new ButtonInfo { Place_ID = 328 });
            buttonInfos.Add(button79, new ButtonInfo { Place_ID = 329 });
            buttonInfos.Add(button80, new ButtonInfo { Place_ID = 330 });
            buttonInfos.Add(button77, new ButtonInfo { Place_ID = 331 });
            buttonInfos.Add(button78, new ButtonInfo { Place_ID = 332 });
        }

        public kinozal2()
        {
            InitializeComponent();
            connection.Open();
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.BackColor = Color.FromArgb(0, 28, 45);
            foreach (Control control in this.Controls)
            {
                if (control is Button && control != button93)
                {
                    control.BackColor = Color.FromArgb(48, 166, 241);
                }
            }
            button93.BackColor = Color.FromArgb(228, 26, 105);
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
        private void kinozal2_Load(object sender, EventArgs e)
        {
            foreach (Control control in this.Controls)
            {
                if (control is Button && control != button93)
                {
                    control.Tag = false;
                    control.Click += new EventHandler(Button_Click);
                }
            }
            label19.Text = ticketprice + " " + " рублей ";
            InitializeButtonInfos();
            CheckButtonColor();
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
                    if (control is Button button && button != button93)
                    {
                        ButtonInfo buttonInfo = buttonInfos[button];
                        if (buttonInfo.Place_ID == placeID)
                        {
                            // закрашиваем кнопку соответствующим цветом
                            button.BackColor = Color.FromArgb(228, 26, 105);
                            button.Enabled = false;
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
            data[2] = Convert.ToString(summaplateja) + " " + "rub"; //price
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
        private void SendBill() // отправляем пдф на почту
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
            Attachment attachment = new Attachment(@"C:\Users\Ринат\Downloads\cinema_bdsql\cinema_bdsql\cinema_bdsql\bin\Debug\билеты\Чек.pdf");//выбираем файл отправляем чек
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

        private void button93_Click(object sender, EventArgs e)
        {
            List<Tuple<int, int>> selectedPlaces = new List<Tuple<int, int>>();

            foreach (Control control in this.Controls)
            {
                if (control is Button button && button != button93 && button.BackColor == Color.FromArgb(228, 26, 105))
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
    }
}
