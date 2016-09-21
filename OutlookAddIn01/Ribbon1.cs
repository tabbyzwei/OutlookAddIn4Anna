using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using AppTools;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn01
{
    public partial class Ribbon1
    {
        ExcelAndSQLite tool = null;

        public static ThisAddIn addin = null;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            this.btnOutlookFolder.Label = "Outlook文件夹：" + AppConfig.ReadConfig("OutlookFolderName");
            this.btnDBPath.Label = "数据库路径：" + AppConfig.ReadConfig("DBPath");
            this.btnXlsxPath.Label = "Xlsx暂存路径：" + AppConfig.ReadConfig("XLSXPath");
        }

        private void btnOutlookFolder_Click(object sender, RibbonControlEventArgs e)
        {
            Outlook.Folder selectBox = (Outlook.Folder)addin.Application.Session.PickFolder();
            AppConfig.WriteConfig("OutlookFolderName", selectBox.FolderPath);
            AppConfig.WriteConfig("OutlookFolderID", selectBox.EntryID);
            this.btnOutlookFolder.Label = "Outlook文件夹：" + AppConfig.ReadConfig("OutlookFolderName");
        }

        private void btnSearch_Click(object sender, RibbonControlEventArgs e)
        {
            SearchForm searchForm = new SearchForm();
            //传递ExcelAndSQLite对象，防止重复生成
            //if (this.tool == null)
            //{
            //    searchForm.tool = new ExcelAndSQLite();
            //}
            searchForm.tool = this.tool;
            //显示窗体
            searchForm.Show();
        }

        private void btnDBPath_Click(object sender, RibbonControlEventArgs e)
        {
            if (fbdDBPath.ShowDialog() == DialogResult.OK)
            {
                AppConfig.WriteConfig("DBPath", fbdDBPath.SelectedPath);
                this.btnDBPath.Label = "数据库路径：" + AppConfig.ReadConfig("DBPath");
            }
        }

        private void btnXlsxPath_Click(object sender, RibbonControlEventArgs e)
        {
            if (fbdXlsxFolder.ShowDialog() == DialogResult.OK)
            {
                AppConfig.WriteConfig("XLSXPath", fbdXlsxFolder.SelectedPath);
                this.btnXlsxPath.Label = "Xlsx暂存路径：" + AppConfig.ReadConfig("XLSXPath");
            }
        }

        private void btnUpdate_Click(object sender, RibbonControlEventArgs e)
        {
            Outlook.Folder box = (Outlook.Folder)addin.Application.Session.GetFolderFromID(AppConfig.ReadConfig("OutlookFolderID"));
            SaveAttachment.SaveAttachments(AppConfig.ReadConfig("XLSXPath"), box);
            //更新数据库
            tool = new ExcelAndSQLite();
            tool.Excel2SQLite();
        }

        private void btnTestSQLite_Click(object sender, RibbonControlEventArgs e)
        {

        }
    }
}
