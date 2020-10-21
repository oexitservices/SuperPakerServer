using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TestWCF.ServiceReference1;

namespace TestWCF
{
    public partial class Form1 : Form
    {

        
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ServiceReference1.SuperPakerServiceSoapClient client = new ServiceReference1.SuperPakerServiceSoapClient();

            ZalogujWynik wynik =  client.Zaloguj("sj", "Balbina8!!", "", "");
            textBox1.Text = wynik.status_opis;

            ZamowienieDaneWynik ws = client.ZamowieniaListaPobierzOdDaty(wynik.sesja, 146, DateTime.Parse("2020-08-25"));
            dataGridView1.AutoGenerateColumns = true;
            dataGridView1.DataSource = ws.dane;
            dataGridView1.DataMember = "Table";

        }
    }
}
