using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
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
            listResult.Clear(); //Clear results
            listBoxFolders.ClearSelected(); //Clear seleted folders. The complete will be cleared when get folders.
            lbFolderCount.Visible = false; //Hide the text under folder list box.
        }

        private void btnGetfolder_Click(object sender, EventArgs e)
        {
            Outlook.Application app = new Outlook.Application();
            var appNS = app.Session;
            var folder = appNS.Folders;
            Outlook.MAPIFolder theFolder;
            var folderNS = folder.Session; //A folder namespace for the current session.
            int folderCount;

            listBoxFolders.Items.Clear(); //Clear the listbox before update.
            folders = folderNS.Folders; //Return all folders in the current session.
            folderCount = folders.Count;
            //the folder we get here is actually the Outlook datafile.
            for (int i = 0; i < folderCount; i++)
            {
                if (i == 0)
                {
                    theFolder = folders.GetFirst();
                    listBoxFolders.Items.Add(theFolder.Name);
                } else if (i > 0 && i < folderCount -1){
                    theFolder = folders.GetNext();
                    listBoxFolders.Items.Add(theFolder.Name);
                }
                else
                {
                    theFolder = folder.GetLast();
                    listBoxFolders.Items.Add(theFolder.Name);
                }
            }

            lbFolderCount.Visible = true;
            lbFolderCount.Text = string.Format("We found {0} folders, please select search range.", folderCount.ToString());

        }
    }
}
