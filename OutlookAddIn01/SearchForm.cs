using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace OutlookAddIn01
{
    public partial class SearchForm : Form
    {
        public ExcelAndSQLite tool = null;
        public SearchForm()
        {
            InitializeComponent();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            tool = new ExcelAndSQLite();
            String path = tool.Search(this.tbKeyWord.Text);
            if (path != null)
            {
                ShowResultInExcel(path);
            }
        }

        #region 用MSlib Office显示Excel
        //启动Excel
        public Excel.Application initailExcel()
        {
            Excel.Application app = null;
            //檢查PC有無Excel在執行
            bool flag = false;
            foreach (var item in Process.GetProcesses())
            {
                if (item.ProcessName == "EXCEL")
                {
                    flag = true;
                    break;
                }
            }

            if (!flag)
            {
                app = new Excel.Application();
            }
            else
            {
                object obj = Marshal.GetActiveObject("Excel.Application");//引用已在執行的Excel
                app = obj as Excel.Application;
            }

            app.Visible = true;//显示Excel

            return app;
        }
        //用Excel展示结果
        public void ShowResultInExcel(String path)
        {
            Excel.Application app = initailExcel();
            app.Visible = false;
            Workbook wbks = app.Workbooks.Open(path, Missing.Value,
                    Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                    Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            //Auto fit
            foreach (Worksheet sheet in wbks.Sheets)
            {
                //适配行列宽度
                sheet.get_Range("A1", "AQ1000").EntireColumn.AutoFit();
                //Sort FIN Net Bill USD
                sheet.UsedRange.Sort(sheet.Cells[1, 35], XlSortOrder.xlDescending,
                    Missing.Value, Missing.Value, XlSortOrder.xlAscending,
                    Missing.Value, XlSortOrder.xlAscending,
                    XlYesNoGuess.xlNo, Missing.Value, XlSortOrientation.xlSortColumns);
                //改变颜色
                int count = 0;
                if (int.TryParse(sheet.Name, out count))
                {
                    for (int i = 1; i < count + 2; i++)
                    {
                        sheet.Cells[i, 6].Interior.Color = Color.Yellow;
                        sheet.Cells[i, 8].Interior.Color = Color.Yellow;
                        sheet.Cells[i, 13].Interior.Color = Color.Yellow;
                        sheet.Cells[i, 17].Interior.Color = Color.Yellow;
                        sheet.Cells[i, 24].Interior.Color = Color.Yellow;
                        sheet.Cells[i, 35].Interior.Color = Color.Yellow;
                    }
                }
            }
        }
        #endregion
    }
}
