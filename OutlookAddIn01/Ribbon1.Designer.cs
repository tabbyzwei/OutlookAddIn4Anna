namespace OutlookAddIn01
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.btnSearch = this.Factory.CreateRibbonButton();
            this.btnUpdate = this.Factory.CreateRibbonButton();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnOutlookFolder = this.Factory.CreateRibbonButton();
            this.btnDBPath = this.Factory.CreateRibbonButton();
            this.btnXlsxPath = this.Factory.CreateRibbonButton();
            this.fbdXlsxFolder = new System.Windows.Forms.FolderBrowserDialog();
            this.fbdDBPath = new System.Windows.Forms.FolderBrowserDialog();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.btnTestSQLite = this.Factory.CreateRibbonButton();
            this.btnTestNPOI = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group1.SuspendLayout();
            this.group3.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Label = "Anna的邮件处理";
            this.tab1.Name = "tab1";
            // 
            // group2
            // 
            this.group2.Items.Add(this.btnSearch);
            this.group2.Items.Add(this.btnUpdate);
            this.group2.Label = "搜索";
            this.group2.Name = "group2";
            // 
            // btnSearch
            // 
            this.btnSearch.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSearch.Image = global::OutlookAddIn01.Properties.Resources.Search;
            this.btnSearch.Label = "Search";
            this.btnSearch.Name = "btnSearch";
            this.btnSearch.ShowImage = true;
            this.btnSearch.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSearch_Click);
            // 
            // btnUpdate
            // 
            this.btnUpdate.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnUpdate.Image = global::OutlookAddIn01.Properties.Resources.update;
            this.btnUpdate.Label = "Update";
            this.btnUpdate.Name = "btnUpdate";
            this.btnUpdate.ShowImage = true;
            this.btnUpdate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdate_Click);
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnOutlookFolder);
            this.group1.Items.Add(this.btnDBPath);
            this.group1.Items.Add(this.btnXlsxPath);
            this.group1.Label = "参数";
            this.group1.Name = "group1";
            // 
            // btnOutlookFolder
            // 
            this.btnOutlookFolder.Label = "Outlook文件夹：";
            this.btnOutlookFolder.Name = "btnOutlookFolder";
            this.btnOutlookFolder.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnOutlookFolder_Click);
            // 
            // btnDBPath
            // 
            this.btnDBPath.Label = "数据库路径：";
            this.btnDBPath.Name = "btnDBPath";
            this.btnDBPath.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDBPath_Click);
            // 
            // btnXlsxPath
            // 
            this.btnXlsxPath.Label = "Xlsx暂存路径：";
            this.btnXlsxPath.Name = "btnXlsxPath";
            this.btnXlsxPath.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnXlsxPath_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.btnTestSQLite);
            this.group3.Items.Add(this.btnTestNPOI);
            this.group3.Label = "测试";
            this.group3.Name = "group3";
            // 
            // btnTestSQLite
            // 
            this.btnTestSQLite.Label = "TestSQLite";
            this.btnTestSQLite.Name = "btnTestSQLite";
            this.btnTestSQLite.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnTestSQLite_Click);
            // 
            // btnTestNPOI
            // 
            this.btnTestNPOI.Label = "TestNPOI";
            this.btnTestNPOI.Name = "btnTestNPOI";
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        public Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOutlookFolder;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDBPath;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnXlsxPath;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSearch;
        private System.Windows.Forms.FolderBrowserDialog fbdXlsxFolder;
        private System.Windows.Forms.FolderBrowserDialog fbdDBPath;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdate;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnTestSQLite;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnTestNPOI;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
