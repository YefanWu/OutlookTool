using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OutlookTool
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnShowSearch_Click(object sender, EventArgs e)
        {
            //Pop up the Search window. 
            Search formSearch = new Search();
            formSearch.Show();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
