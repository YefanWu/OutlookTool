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
        //Update Folder List.
        public void GetFolderNames()
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
                }
                else if (i > 0 && i < folderCount - 1)
                {
                    theFolder = folders.GetNext();
                    listBoxFolders.Items.Add(theFolder.Name);
                }
                else
                {
                    theFolder = folder.GetLast();
                    listBoxFolders.Items.Add(theFolder.Name);
                }
            }
            //Update the label below folder list. 
            lbFolderCount.Visible = true;
            lbFolderCount.Text = string.Format("We found {0} folders, please select search range.", folderCount.ToString());
        }

        //Pull the folders selected.
        //Still unable to retrun a folder, looking into why.
        public Outlook.MAPIFolder GetFolders()
        {
            Outlook.Application app = new Outlook.Application();
            var tempFolders = app.Session.Folders;
            Outlook.MAPIFolder folder = tempFolders as Outlook.MAPIFolder;
            
            if (listBoxFolders.SelectedItems.Count > 0 ) //Check if folder selected.
            {
                folder.Name = listBoxFolders.SelectedItems.ToString();

                return folder;
            }
            else
            {
                //If forget to select folder then popup a hint.
                //MessageBoxButtons messageBox = MessageBoxButtons.OK;
                MessageBox.Show("Please select one folder.", "Select Folder");

                return null;
            } 
        }
    }
}
