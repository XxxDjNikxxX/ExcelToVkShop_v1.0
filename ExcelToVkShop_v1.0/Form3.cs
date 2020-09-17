using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelToVkShop_v1._0
{
    public partial class Form3 : Form
    {
        public string twofa;
        public Form3()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            two_fa_auth();
            this.Close();
        }

        public string two_fa_auth()
        {
            string twofa = textBox1.Text;
            return twofa;
        }

        
    }
}
