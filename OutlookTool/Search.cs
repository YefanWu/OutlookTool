using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = NetOffice.OutlookApi;

namespace OutlookTool
{
    public partial class Search : Form
    {
        public Search()
        {
            InitializeComponent();
        }

        private void Search_Load(object sender, EventArgs e)
        {
            //Get folder names
            
            listFolders.Columns.Add("Folder Name");
            listFolders.Columns[0].Width = listFolders.ClientSize.Width;
            //Outlook.Application app = new Outlook.Application();
            Outlook.Folders folders;
            folders = Outlook.Folders.GetActiveInstance(true);
            foreach (var folder in folders)
            {
                ListViewItem item = new ListViewItem();
                item.Text = folder.Name;
                listFolders.Items.Add(item);
            }
            
        }
    }
}
