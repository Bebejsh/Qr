using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using QRCoder;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using Application = Microsoft.Office.Interop.Word.Application;
using System.Windows.Documents;
using Section = Microsoft.Office.Interop.Word.Section;
using Microsoft.Azure.Amqp.Framing;

namespace qr1
{
    public partial class Form1 : Form
    {
        private bool isTextBoxEmpty = true;

        public Form1()
        {
            InitializeComponent();
           
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (pictureBox1.Image != null)
            {
                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Filter = "Images|*.png;*.bmp;*.jpg";
                saveDialog.Title = "Save an Image";

                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    pictureBox1.Image.Save(saveDialog.FileName);
                }
            }
            else
            {
                MessageBox.Show("Ошибка", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            var wordApplication = new Application();
            var wordDocument = wordApplication.Documents.Add();
            var sections = wordDocument.Sections;
            // Получение пути к изображению
            var imagePath = pictureBox1.ImageLocation;

            // Вставка изображения в документ Word
            var shape = wordDocument.Shapes.AddPicture(imagePath, false, true);


            foreach (Section section in sections)
            {
                // Получение коллекции всех нижних колонтитулов раздела
                var footers = section.Footers;

                // Установка свойств нижнего колонтитула каждой страницы
                foreach (Footer footer in footers)
                {
                    // Определение диапазона для нижнего колонтитула текущей страницы
                    var range = footer.Range;

                    // Размещение изображения внизу колонтитула текущей страницы
                    var picture = range.InlineShapes.AddPicture(imagePath);
                    picture.Height = 100; // Задайте нужную высоту изображения
                    picture.Width = 100; // Задайте нужную ширину изображения
                }
            }

            wordDocument.SaveAs("C:\\Users\\kab511students\\Desktop\\dvij — копия\\1233.docx" );

            // Закрытие документа и выход из приложения Word
            wordDocument.Close();
            wordApplication.Quit();



        }

        private void textBox1_Click(object sender, EventArgs e)
        {
           
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_MouseEnter(object sender, EventArgs e)
        {
            if (isTextBoxEmpty)
            {
                textBox1.Text = "";
                isTextBoxEmpty = false;
            }
        }

        private void textBox1_MouseLeave(object sender, EventArgs e)
        {
            if (isTextBoxEmpty)
            {
                textBox1.Text = "Напиши сюда";
                isTextBoxEmpty = true;
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            QRCodeGenerator Qg = new QRCodeGenerator();
            var MyData = Qg.CreateQrCode(textBox1.Text, QRCodeGenerator.ECCLevel.L);
            var data = new QRCode(MyData);
            pictureBox1.Image = data.GetGraphic(50);
        }
    }
}
