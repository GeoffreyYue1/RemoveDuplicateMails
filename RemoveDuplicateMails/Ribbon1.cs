using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace RemoveDuplicateMails
{
    public partial class Ribbon1
    {
        List<string> MessageIDList;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btn_RemoveDupMail_Click(object sender, RibbonControlEventArgs e)
        {
            try
            { 
            MessageIDList = new List<string>();

            MAPIFolder olFolder = Globals.ThisAddIn.Application.ActiveExplorer().CurrentFolder;
                MAPIFolder targetFolder = olFolder.Store.GetRootFolder().Folders.Add("DupItems_" + DateTime.Now.ToString("yyyyMMddHHmmss"));

            if(MessageBox.Show("You are going to Remove the duplicate mails from Folder : " + olFolder.FullFolderPath+ ", Please back up your pst before clicking OK button!")==DialogResult.OK)
            {
                    MoveDuplicateMailToFolder(olFolder,targetFolder);
                    MessageBox.Show("Duplicate removing done!");
            }

            MessageIDList.Clear();
            }
            catch(System.Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void MoveDuplicateMailToFolder (MAPIFolder srcFolder, MAPIFolder tgtFolder)
        {


            string messageId = string.Empty;

            List<MailItem> mailsToDelete = new List<MailItem>();
            foreach (var item in srcFolder.Items)
            {

                if (item is MailItem)
                {
                    try
                    {

                        messageId
                            = (item as MailItem).PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1035001F");
                    }
                    catch (System.Exception ex)
                    {

                    }

                    if (MessageIDList.Contains(messageId))
                    {
                        mailsToDelete.Add(item as MailItem);
                    }
                    else
                    {

                        MessageIDList.Add(messageId);

                    }


                }

            }

            foreach (MailItem mailToDelete in mailsToDelete)
            {
                mailToDelete.Move(tgtFolder);
            }

            mailsToDelete.Clear();

            if (srcFolder.Folders != null)
            {
                foreach (MAPIFolder subfolder in srcFolder.Folders)
                {
                    MoveDuplicateMailToFolder(subfolder,tgtFolder);
                }

            }

        }
        /*
        void RemoveDuplicateMailOfFolder(MAPIFolder folder)
        {

            string messageId = string.Empty;

            List<MailItem> mailsToDelete = new List<MailItem>();
            foreach (var item in folder.Items)
            {

                if (item is MailItem)
                {
                    try { 
                     
                   messageId
                       = (item as MailItem).PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1035001F");
                    }
                    catch (System.Exception ex)
                    {

                    }

                    if (MessageIDList.Contains(messageId))
                    {
                        mailsToDelete.Add(item as MailItem);
                    }
                    else
                    {
                       
                            MessageIDList.Add(messageId);
                       
                    }


                }

            }
            
            foreach(MailItem mailToDelete in mailsToDelete )
            {
                mailToDelete.Delete();
            }

            mailsToDelete.Clear();

            if (folder.Folders != null)
            {
                foreach (MAPIFolder subfolder in folder.Folders)
                {
                    RemoveDuplicateMailOfFolder(subfolder);
                }

            }

        }*/
    }
}
