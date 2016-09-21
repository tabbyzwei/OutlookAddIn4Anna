using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn01
{
    public class SaveAttachment
    {
        public static void SaveAttachments(string path, Outlook.Folder folder)
        {
            int TotalFile = 0;
            int NewFile = 0;
            foreach (Outlook.MailItem item in folder.Items)
            {
                foreach (Outlook.Attachment attachment in item.Attachments)
                {
                    TotalFile++;
                    if (!File.Exists(path + "\\" +attachment.FileName))
                    {
                        NewFile++;
                        attachment.SaveAsFile(path + "\\" + attachment.FileName);
                    }
                }
            }
            MessageBox.Show(String.Format("总附件数: {0}, 新增附件数: {1}", TotalFile,NewFile));
        }
    }
}
