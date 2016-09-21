using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.IO;

namespace OutlookAddIn01
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Ribbon1.addin = this;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // 备注: Outlook 不会再遇到这种问题。如果具有
            //关闭 Outlook 时必须运行的代码，请参阅 http://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion

        private void ShowAllFolder()
        {
            Outlook.Folders allBox =
               this.Application.ActiveExplorer().Session.Folders;

            //MessageBox.Show(allBox.Count.ToString());

            foreach (Outlook.Folder Boxes in allBox)
            {
                MessageBox.Show(Boxes.Name);
                String BoxesInfo = "";
                foreach (Outlook.Folder box in Boxes.Folders)
                {
                    BoxesInfo += box.Name + "----" + box.EntryID + "\r\n";

                }
                MessageBox.Show(BoxesInfo);
            }
        }

        private void SetCurrentFolder()
        {
            string folderName = "Test";
            Outlook.MAPIFolder inBox = (Outlook.MAPIFolder)
                this.Application.ActiveExplorer().Session.GetDefaultFolder
                (Outlook.OlDefaultFolders.olFolderInbox);
            try
            {
                this.Application.ActiveExplorer().CurrentFolder = inBox.
                    Folders[folderName];
                this.Application.ActiveExplorer().CurrentFolder.Display();
            }
            catch
            {
                MessageBox.Show("There is no folder named " + folderName +
                    ".", "Find Folder Name");
                //MessageBox.Show(, "Find Folder Name");

            }
        }
        private void CreateCustomFolder()
        {
            Outlook.MAPIFolder inBox = (Outlook.MAPIFolder)
                this.Application.ActiveExplorer().Session.GetDefaultFolder
                (Outlook.OlDefaultFolders.olFolderInbox);
            string userName = (string)this.Application.ActiveExplorer()
                .Session.CurrentUser.Name;
            Outlook.MAPIFolder customFolder = null;
            try
            {
                customFolder = (Outlook.MAPIFolder)inBox.Folders.Add(userName,
                    Outlook.OlDefaultFolders.olFolderInbox);
                MessageBox.Show("You have created a new folder named " +
                    userName + ".");
                inBox.Folders[userName].Display();
            }
            catch (Exception ex)
            {
                MessageBox.Show("The following error occurred: " + ex.Message);
            }
        }
        private void SearchInBox()
        {
            Outlook.MAPIFolder inbox = this.Application.ActiveExplorer().Session.
                GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);

            MessageBox.Show(inbox.Folders.Count.ToString());

            foreach (Outlook.Folder item in inbox.Folders)
            {
                MessageBox.Show(item.Name);
            }

            Outlook.Items items = inbox.Items;
            Outlook.MailItem mailItem = null;
            object folderItem;
            string subjectName = string.Empty;
            string filter = "[Subject] > 's' And [Subject] <'u'";
            folderItem = items.Find(filter);
            while (folderItem != null)
            {
                mailItem = folderItem as Outlook.MailItem;
                if (mailItem != null)
                {
                    subjectName += "\n" + mailItem.Subject;
                }
                folderItem = items.FindNext();
            }
            subjectName = " The following e-mail messages were found: " +
                subjectName;
            MessageBox.Show(subjectName);
        }
    }
}
