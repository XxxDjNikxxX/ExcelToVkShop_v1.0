using System;
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

namespace ExcelToVkShop_v1._0
{
    public partial class Form1 : Form
    {
        string filePath = string.Empty;

        OpenFileDialog openFileDialog = new OpenFileDialog();
        VkApi api = new VkApi();

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }



        private void button1_Click(object sender, EventArgs e) //открываем поиск по директориям
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

            textBox1.Text = filePath;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                api.Authorize(new ApiAuthParams
                {
                    AccessToken = textBox2.Text
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
            int k = 1;
            Excel excel = new Excel(filePath, 1);
             api.Markets.Add( new VkNet.Model.RequestParams.Market.MarketProductParams
             {
                 OwnerId = -1,
                 Name = string.Format(excel.Cells(2,)),
                 Description = string.Format(excel.Cells(2,5)),
                 CategoryId = 0,    
                 Price = decimal.Parse(excel.Cells(2,8)), 
             });
   
        }
    }
}
