using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookTool
{
    public partial class Search : Form
    {
        Outlook.Folders folders; //All folders.
        public Search()
        {
            InitializeComponent();
        }

        private void Search_Load(object sender, EventArgs e)
        {

        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            //The full search function will migrate to Functions.cs and pull the folder object from GetFolders().
            Outlook.Application application = new Outlook.Application();

            var ns = application.Session;
            var SearchFolder = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);//Now we can only search in Inbox. Not subfolders.
            var item = SearchFolder.Items;
            string searchFor = boxSearchfor.Text.Trim();
            List<string> subject = new List<string>();

            listResult.Columns.Add("Results");

            foreach (var inboxitem in SearchFolder.Items)
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
            listResult.Clear(); //Clear results
            listBoxFolders.ClearSelected(); //Clear seleted folders. The complete will be cleared when get folders.
            lbFolderCount.Visible = false; //Hide the text under folder list box.
        }

        private void btnGetfolder_Click(object sender, EventArgs e)
        {
            GetFolderNames();
        }

        private void btnTest_Click(object sender, EventArgs e)
        {
            try
            {
                lbFolderCount.Text = GetFolders().ToString();
                
            }
            catch (Exception)
            {
                
            }
            MessageBox.Show("break point", "Select Folder");
        }
    }
}
