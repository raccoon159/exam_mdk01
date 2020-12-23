using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Экзамен01._01
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            textBox2.Enabled = false;
            textBox4.Enabled = false;
            combo();
            comboBox1.Text = "Выберите услугу";
            
        }
        SqlConnection myConnection = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\Database1.mdf;Integrated Security=True");

        double coun;
        double sum;int i =0;
        public void combo()//Вывод из бд в combobox
        {
            myConnection.Open();

            SqlCommand command = new SqlCommand("SELECT * FROM Услуги", myConnection);
            SqlDataReader reader;
            try
            {
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    string name = reader.GetString(1);
                    comboBox1.Items.Add(name);
                }
                myConnection.Close();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }       

        

        private void button1_Click(object sender, EventArgs e)
        {
            
            if (textBox2.Text != "" && textBox3.Text != "")
            {
                double cum = Convert.ToDouble(textBox3.Text) - Convert.ToDouble(textBox1.Text) * coun;
                if (cum < 0)
                    MessageBox.Show("Возмите еще денег \n минимум еще "+ Math.Abs(cum)+" Рублей", "Ошибка");
                else
                    textBox4.Text = cum.ToString();

            }

            else MessageBox.Show("Заполните или выберите","Ошибка");
        }

        private void comboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {            
            myConnection.Open();

            SqlCommand command = new SqlCommand("SELECT Цена FROM Услуги Where Код=@Код", myConnection);
            command.Parameters.AddWithValue("@Код", comboBox1.SelectedIndex);

            try
            {
                coun = Convert.ToDouble(command.ExecuteScalar().ToString());
                textBox2.Text = coun.ToString();//Вывод цены
                myConnection.Close();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox4.Text != "")
            {
                string name = comboBox1.Text;
                DateTime dt = DateTime.Now;

                Properties.Settings.Default.fil++;//Чек
                // Получаем массив байтов из нашего файла
                byte[] textByteArray = File.ReadAllBytes("Test.docx");
                // Массив данных
                string[] data = new string[] {Properties.Settings.Default.fil.ToString(), dt.ToString(), Convert.ToString(name), Convert.ToString(textBox2.Text), Convert.ToString(textBox1.Text), Convert.ToString("="+ Convert.ToDouble(textBox1.Text) * coun), Convert.ToString(textBox3.Text + " Руб."), Convert.ToString(textBox4.Text)+"Руб." };
                // Начинаем работу с потоком
                using (MemoryStream stream = new MemoryStream())
                {
                    // Записываем в поток наш word-файл
                    stream.Write(textByteArray, 0, textByteArray.Length);
                    // Открываем документ из потока с возможностью редактирования
                    using (WordprocessingDocument doc = WordprocessingDocument.Open(stream, true))
                    {
                        // Ищем все закладки в документе
                        var bookMarks = FindBookmarks(doc.MainDocumentPart.Document);

                        int i = 0;
                        foreach (var end in bookMarks)
                        {
                            // В документе встречаются какие-то служебные закладки
                            // Таким способом отфильтровываем всё ненужное
                            // end.Key содержит имена наших закладок
                            if (end.Key != "i" && end.Key != "d" && end.Key != "date" && end.Key != "name" && end.Key != "summa" && end.Key != "kol" && end.Key != "sum" && end.Key != "ves" && end.Key != "ch" ) continue;
                            // Создаём текстовый элемент
                            var textElement = new Text(data[i].ToString());
                            // Далее данный текст добавляем в закладку
                            var runElement = new Run(textElement);
                            end.Value.InsertAfterSelf(runElement);
                            i++;
                        }
                    }
                    // Записываем всё в наш файл 
                    
                    String path = (Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
                    String file = "чек "+ Properties.Settings.Default.fil+ ".docx";
                    String filePath = Path.Combine(path, file);
                    File.WriteAllBytes(filePath, stream.ToArray());
                    MessageBox.Show(file + " находится на рабочем столе");
                }
                Properties.Settings.Default.Save();//сохраняем значения
            }
            else MessageBox.Show("Покупка","Ошибка");
        }
        private static Dictionary<string, BookmarkEnd> FindBookmarks(OpenXmlElement documentPart, Dictionary<string, BookmarkEnd> outs = null, Dictionary<string, string> bStartWithNoEnds = null)
        {
            if (outs == null) { outs = new Dictionary<string, BookmarkEnd>(); }
            if (bStartWithNoEnds == null) { bStartWithNoEnds = new Dictionary<string, string>(); }

            // Проходимся по всем элементам на странице Word-документа
            foreach (var docElement in documentPart.Elements())
            {
                // BookmarkStart определяет начало закладки в рамках документа
                // маркер начала связан с маркером конца закладки
                if (docElement is BookmarkStart)
                {
                    var bookmarkStart = docElement as BookmarkStart;
                    // Записываем id и имя закладки
                    bStartWithNoEnds.Add(bookmarkStart.Id, bookmarkStart.Name);
                }

                // BookmarkEnd определяет конец закладки в рамках документа
                if (docElement is BookmarkEnd)
                {
                    var bookmarkEnd = docElement as BookmarkEnd;
                    foreach (var startName in bStartWithNoEnds)
                    {
                        // startName.Key как раз и содержит id закладки
                        // здесь проверяем, что есть связь между началом и концом закладки
                        if (bookmarkEnd.Id == startName.Key)
                            // В конечный массив добавляем то, что нам и нужно получить
                            outs.Add(startName.Value, bookmarkEnd);
                    }
                }
                // Рекурсивно вызываем данный метод, чтобы пройтись по всем элементам
                // word-документа
                FindBookmarks(docElement, outs, bStartWithNoEnds);
            }
            return outs;
        }
    }
}
