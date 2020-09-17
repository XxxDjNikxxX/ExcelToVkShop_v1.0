using System;
using System.Xml;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using VkNet;
using VkNet.Enums.Filters;
using VkNet.Model;
using System.Globalization;
using HtmlAgilityPack;
using System.Drawing.Text;
using System.Threading;

namespace ExcelToVkShop_v1._0
{
    public partial class Form1 : Form
    {
        string filePath = string.Empty;
        string cutted_text = string.Empty;

        OpenFileDialog openFileDialog = new OpenFileDialog();

        VkApi api = new VkApi();

        public Form1()
        {
            InitializeComponent();
        }


        private void button1_Click(object sender, EventArgs e) //открываем поиск по директориям
        {
            get_dir();
            
            textBox1.Text = filePath;
            Form2 f2 = new Form2();
            f2.filePath = this.filePath;

            f2.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                api.Authorize(new ApiAuthParams
                {
                    AccessToken = "", 
                });
            }

            catch
            {
                MessageBox.Show("Ошибка авторизации:");
            }

            if(api.IsAuthorized == true)
            {
                MessageBox.Show("Авторизация прошла успешно!" + "\n" + api.Token);
            }
           
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Excel excel = new Excel(filePath, 1);
            cut_html();

            api.Markets.Add(new VkNet.Model.RequestParams.Market.MarketProductParams
                {
                    OwnerId = -1,
                    Name = textBox3.Text,
                    Description = cutted_text,
                    MainPhotoId = PhotoSize.,
                    CategoryId = 1,
                    Price = Decimal.Parse(textBox4.Text),
                });

            var photos = api.Photo.Save(new PhotoSaveParams
            {
                SaveFileResponse = responseFile,
                AlbumId = 123
            });
        }

        public string get_dir ()
        {
            openFileDialog.InitialDirectory = "c:\\";
            openFileDialog.Filter = "Excel files (*.xls,*.xlsx)|*.xls*.xlsx|All files (*.*)|*.*";
            openFileDialog.FilterIndex = 2;
            openFileDialog.RestoreDirectory = true;

            if (openFileDialog.ShowDialog() == DialogResult.OK) //Получение пути из директории
                try
                {

                    filePath = openFileDialog.FileName;
                }
                catch
                {
                    MessageBox.Show("Произошла ошибка чтения таблицы Excel!!!");
                }

            return filePath;
        }

        public string cut_html()
        {

            var doc = new HtmlAgilityPack.HtmlDocument();
            var uncutted_txt = richTextBox1.Text;

            doc.LoadHtml(uncutted_txt);

            var htmlNodes = doc.DocumentNode.SelectNodes("//p/span");

            foreach (var node in htmlNodes)
            {
                cutted_text += node.InnerText;
            }
            return cutted_text;
        }

    }
}
