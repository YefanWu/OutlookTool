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
            var inbox = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);//Now we can only search in Inbox. Not subfolders.
            var item = inbox.Items;
            string searchFor = boxSearchfor.Text.Trim();
            List<string> subject = new List<string>();

            listResult.Columns.Add("Results");

            foreach (var inboxitem in inbox.Items)
            {
                if (inboxitem is Outlook.MailItem) //Item could be meeting invite etc...
                {
                    Outlook.MailItem mail = inboxitem as Outlook.MailItem;
                    try
                    {
                        if (mail.Subject.Contains(searchFor)) //If the email do not have a subject, will cause exception here. We can ignore.
                        {
                            subject.Add(mail.Subject);
                        }
                    }
                    catch (Exception)
                    {

                        continue;
                    }

                }

            }

            //Update listview
            listResult.BeginUpdate();
            foreach (var result in subject)
            {
                ListViewItem viewItem = new ListViewItem();
                viewItem.Text = result;
                listResult.Items.Add(viewItem);
            }

            listResult.EndUpdate();

        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            listResult.Clear();
        }
    }
}
