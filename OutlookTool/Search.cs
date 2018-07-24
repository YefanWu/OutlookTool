using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

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

        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            Outlook.Application application = new Outlook.Application();

            var ns = application.Session;
            var inbox = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            var item = inbox.Items;
            string searchFor = boxSearchfor.Text.Trim();
            List<string> subject = new List<string>();

            listResult.Columns.Add("Results");

            foreach (var inboxitem in inbox.Items)
            {
                if (inboxitem is Outlook.MailItem) //Item could be meeting invite etc...
                {
                    Outlook.MailItem mail = new Outlook.MailItem(); //bug need to fix.
                    subject.Add(mail.Subject);

                    //Update listview for testing.
                    ListViewItem viewItem = new ListViewItem();
                    viewItem.Text = mail.Subject;
                    listResult.Items.Add(viewItem);
                }
                break;
            }


            Dispose();
        }
    }
}
