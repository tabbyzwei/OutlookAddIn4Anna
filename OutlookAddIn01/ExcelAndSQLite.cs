using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using AppTools;
using System.Reflection;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace OutlookAddIn01
{
    public class ExcelAndSQLite
    {
        //数据库和搜索结果路径
        String strRootPath = null;
        String strSQLHeader = null;
        String strDBTableName = null;
        //Excel file
        String strExcelFolder = null;
        String[] headerArray = new String[] {"Sales VP Territory","Top End Cust Parent Key",
            "Top End Cust Parent Text","Sold-to Parent Key"  ,"Sold-to Parent Text",
            "Sold-to Party Key","Sold-to Party Text","Ship-To Party Key","Ship-To Party Text",
            "AIB Customer flag","Cust Pur Order Num","Cust PO Line","Sales Doc Type",
            "Billing type","Sales document","SO Item","Billing document","Created on",
            "Bill Date","Bill Fisc Week","Sales: Agreement ID","Your Reference","Material",
            "Cust Material","AIB Material flag","AMD Brand","AMD External Device",
            "AMD Market Segment","AMD Model","AMD Prod Sub Family","AMD Product Line",
            "Basic material","FIN Net Bill Qty","Unit Price USD","FIN Net Bill USD","Plant",
            "Storage location key","Storage Location Text","Dropship Flag","Delivery",
            "House Way Bill","Master Way Bill","Sort Field"};

        //SQLite
        SQLiteConnection connectSQLite = null;
        List<String> listDBTableName = null;
        List<SQLiteParameter> listSQLiteParameter = new List<SQLiteParameter>();

        IWorkbook wbSource = null;

        public ExcelAndSQLite()
        {
            //读取配置
            strExcelFolder = AppConfig.ReadConfig("XLSXPath");
            strRootPath = AppConfig.ReadConfig("DBPath");
            //打开数据库
            string dbPath = "Data Source =" + strRootPath + @"/AnnaDB.db";
            connectSQLite = new SQLiteConnection(dbPath);//创建数据库实例，指定文件位置  
            connectSQLite.Open();//打开数据库，若文件不存在会自动创建
            //创建数据表
            //CreateSQLiteDB(connectSQLite, strDBTableName);
            //获取表名称列表
            listDBTableName = GetDBTableName(connectSQLite);
        }

        #region SQLite操作
        //创建SQLite数据表
        public void CreateSQLiteDB(SQLiteConnection connect, String strDBTableName)
        {
            #region SQL创建表
            string sql = @"CREATE TABLE IF NOT EXISTS " + strDBTableName + @" ( 
                Id INTEGER PRIMARY KEY UNIQUE,
                Sales_VP_Territory VARCHAR(100),
                Top_End_Cust_Parent_Key VARCHAR(100),
                Top_End_Cust_Parent_Text VARCHAR(100),
                Sold_To_Parent_Key VARCHAR(100),
                Sold_To_Parent_Text VARCHAR(100),
                Sold_To_Party_Key VARCHAR(100),
                Sold_To_Party_Text VARCHAR(100),
                Ship_To_Party_Key VARCHAR(100),
                Ship_To_Party_Text VARCHAR(100),
                AIB_Customer_Flag VARCHAR(100),
                Cust_Pur_Order_Num VARCHAR(100),
                Cust_PO_Line VARCHAR(100),
                Sales_Doc_Type VARCHAR(100),
                Billing_Type VARCHAR(100),
                Sales_Document VARCHAR(100),
                SO_Item VARCHAR(100),
                Billing_Document VARCHAR(100),
                Created_On VARCHAR(100),
                Bill_Date VARCHAR(100),
                Bill_Fisc_Week VARCHAR(100),
                Sales_Agreement_ID VARCHAR(100),
                Your_Reference VARCHAR(100),
                Material VARCHAR(100),
                Cust_Material VARCHAR(100),
                AIB_Material_Flag VARCHAR(100),
                AMD_Brand VARCHAR(100),
                AMD_External_Device VARCHAR(100),
                AMD_Market_Segment VARCHAR(100),
                AMD_Model VARCHAR(100),
                AMD_Prod_Sub_Family VARCHAR(100),
                AMD_Product Line VARCHAR(100),
                Basic_Material VARCHAR(100),
                FIN_Net_Bill_Qty VARCHAR(100),
                Unit_Price_USD VARCHAR(100),
                FIN_Net_Bill_USD VARCHAR(100),
                Plant VARCHAR(100),
                Storage_location_Key VARCHAR(100),	
                Storage_Location_Text VARCHAR(100),
                Dropship_Flag VARCHAR(100),
                Delivery VARCHAR(100),
                House_Way_Bill VARCHAR(100),
                Master_Way_Bill VARCHAR(100),
                Sort_Field VARCHAR(100)
            );";
            #endregion
            SQLiteCommand cmdCreateTable = new SQLiteCommand(sql, connect);
            //如果表不存在，创建数据表
            if (-1 != cmdCreateTable.ExecuteNonQuery())
            {
                //新建表
                strSQLHeader = InitParameter(connect, strDBTableName);
                InsertDB(connectSQLite, ProcessExcelFile(strExcelFolder + @"\Acer weekly billing report_" + strDBTableName + ".xlsx"), strDBTableName);
            }
        }
        //获取SQLite表头数据
        public String InitParameter(SQLiteConnection connection, String strDBTableName)
        {
            String strSQLHeader = null;
            SQLiteDataAdapter dataAdapter = new SQLiteDataAdapter("SELECT * FROM " + strDBTableName + " ;", connection);
            DataTable dataTable = new DataTable(strDBTableName);
            dataAdapter.Fill(dataTable);
            listSQLiteParameter.Clear();
            foreach (DataColumn column in dataTable.Columns)
            {
                strSQLHeader += "@" + column.ColumnName + ",";
                listSQLiteParameter.Add(new SQLiteParameter(column.ColumnName));
            }
            strSQLHeader = strSQLHeader.TrimEnd(',');
            return strSQLHeader;
        }
        //用DataTable更新SQLite指定表
        public void InsertDB(SQLiteConnection connect, DataTable dt, String tableName)
        {
            //批量写入
            using (SQLiteTransaction tran = connect.BeginTransaction())//实例化一个事务  
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    SQLiteCommand cmd = new SQLiteCommand(connect);//实例化SQL命令  
                    cmd.Transaction = tran;
                    cmd.CommandText = "replace into " + tableName + " values(" + strSQLHeader + ")";
                    //设置带参SQL语句  
                    DataRow row = dt.Rows[i];
                    //String[] itemArray = (String[])row.ItemArray;
                    int j = -1;
                    foreach (SQLiteParameter item in listSQLiteParameter)
                    {
                        if (j < 0)
                        {
                            item.Value = i;
                            j++;
                        }
                        else
                        {
                            item.Value = row.ItemArray[j++];
                        }
                    }
                    cmd.Parameters.AddRange(listSQLiteParameter.ToArray());

                    cmd.ExecuteNonQuery();//执行查询  
                }
                tran.Commit();//提交  
            }
        }
        //查找SQLite
        public DataTable SelectDB(SQLiteConnection connect, String keyWord, String strDBTableName)
        {
            SQLiteDataAdapter dataAdapter = new SQLiteDataAdapter("SELECT * FROM " + strDBTableName + " where Cust_Material = '" + keyWord + "' and Sales_Doc_Type = 'Standard Order' ;", connect);
            DataTable dataTable = new DataTable(strDBTableName);
            dataAdapter.Fill(dataTable);
            return dataTable;
        }
        //SQLite表信息
        public List<String> GetDBTableName(SQLiteConnection connect)
        {
            List<String> listDBTableName = new List<string>();
            SQLiteDataAdapter dataAdapter = new SQLiteDataAdapter("select name from sqlite_master where type = 'table' order by name;", connect);
            DataTable dataTable = new DataTable(strDBTableName);
            dataAdapter.Fill(dataTable);
            foreach (DataRow row in dataTable.Rows)
            {
                listDBTableName.Add((String)row.ItemArray[0]);
            }
            return listDBTableName;
        }
        #endregion

        #region 用NPOI提取Excel内容
        //读取指定Excel文件为DataTable(除去首行)
        public DataTable ProcessExcelFile(String filePath)
        {
            FileInfo fiSource = new FileInfo(filePath);
            //if file exist
            if (fiSource.Exists)
            {
                //open file and get file stream

                if (fiSource.Extension.Equals(".xls"))
                {
                    using (FileStream fsSource = fiSource.OpenRead())
                    {
                        wbSource = new HSSFWorkbook(fsSource);
                    }
                }
                else if (fiSource.Extension.Equals(".xlsx"))
                {
                    using (FileStream fsSource = fiSource.OpenRead())
                    {
                        wbSource = new XSSFWorkbook(fsSource);
                    }
                }
            }
            //Convert to DataTable
            ISheet srcSheet = wbSource.GetSheet("BillingMajor");
            DataTable dt = new DataTable();
            //Create Columns
            for (int j = 1; j <= srcSheet.GetRow(0).LastCellNum; j++)
            {
                dt.Columns.Add(Convert.ToChar(((int)'A') + j).ToString());
            }
            //Fill Row
            for (int i = 1; i <= srcSheet.LastRowNum; i++)
            {
                IRow row = srcSheet.GetRow(i);
                List<String> listCellValue = new List<string>();
                foreach (ICell cell in row.Cells)
                {
                    listCellValue.Add(GetCellStringValue(cell));
                }
                dt.Rows.Add(listCellValue.ToArray());
            }
            return dt;
        }
        //获取Excel单元格String值
        public String GetCellStringValue(ICell cell)
        {
            String result = null;
            switch (cell.CellType)
            {
                case CellType.Blank:
                    result = "[null]";
                    break;
                case CellType.Boolean:
                    result = cell.BooleanCellValue.ToString();
                    break;
                case CellType.Numeric:
                    result = cell.ToString();    //This is a trick to get the correct value of the cell. NumericCellValue will return a numeric value no matter the cell value is a date or a number.
                    break;
                case CellType.String:
                    result = cell.StringCellValue;
                    break;
                case CellType.Error:
                    result = cell.ErrorCellValue.ToString();
                    break;
                case CellType.Formula:
                default:
                    result = "=" + cell.CellFormula;
                    break;
            }
            return result;
        }
        //DataTable写入Sheet
        public ISheet DataTable2Sheet(DataTable dt, XSSFWorkbook tempWorkbook)
        {
            int indexOfSheet = 0;
            ISheet tempSheet = tempWorkbook.CreateSheet(dt.TableName);
            //表头
            IRow tempRow = tempSheet.CreateRow(indexOfSheet++);
            if (tempRow != null)
            {
                for (int i = 0; i < headerArray.Length; i++)
                {
                    ICell cell = tempSheet.GetRow(indexOfSheet - 1).CreateCell(i);
                    if (cell != null)
                    {
                        cell.SetCellValue(headerArray[i]);
                    }
                }
            }
            //数据
            foreach (DataRow row in dt.Rows)
            {
                IRow dataRow = tempSheet.CreateRow(indexOfSheet++);
                if (tempRow != null)
                {
                    for (int i = 0; i < row.ItemArray.Length - 1; i++)
                    {
                        ICell cell = tempSheet.GetRow(indexOfSheet - 1).CreateCell(i);
                        if (cell != null)
                        {
                            cell.SetCellValue((String)row.ItemArray[i + 1]);
                        }
                    }
                }
            }
            return tempSheet;
        }
        #endregion

        #region SQLite中导入Excel 
        public void Excel2SQLite()
        {
            DirectoryInfo folder = new DirectoryInfo(strExcelFolder);
            //遍历文件
            foreach (FileInfo file in folder.GetFiles())
            {
                //筛选xlsx文件，排除隐藏文件
                if ((file.Attributes & FileAttributes.Hidden) == 0 && file.Extension.Equals(".xlsx"))
                {
                    Console.WriteLine(file.FullName);
                    Console.WriteLine(file.Name);
                    //"Acer weekly billing report_BillingMajor_20160814_09073758.xlsx"
                    String regRule = @"(BillingMajor_\d{8}_\d{8})";
                    Regex reg = new Regex(regRule);
                    Console.WriteLine(reg.Match(file.Name).Value);
                    //新建表
                    String tableName = reg.Match(file.Name).Value;
                    CreateSQLiteDB(connectSQLite, tableName);
                }
            }
        }
        #endregion

        #region 查询数据导出Excel
        public String Search(String txtKeyWord)
        {
            bool saveFile = false;
            if (!txtKeyWord.Equals(""))
            {
                //创建Excel
                XSSFWorkbook tempWorkbook = new XSSFWorkbook();
                ISheet sheet = null;
                DataTable dt = new DataTable("Result");
                for (int i = 0; i <= 43; i++)
                {
                    dt.Columns.Add("A" + i);
                }
                foreach (String tableName in listDBTableName)
                {
                    DataTable dataTable = SelectDB(connectSQLite, txtKeyWord, tableName);
                    foreach (DataRow row in dataTable.Rows)
                    {
                        if (dt.Rows.IndexOf(row) == -1)
                        {
                            saveFile = true;
                            dt.Rows.Add(row.ItemArray);
                        }
                    }
                }
                //插入Excel
                if (dt.Rows.Count > 0)
                {
                    dt.TableName = dt.Rows.Count.ToString();
                    saveFile = true;
                    sheet = DataTable2Sheet(dt, tempWorkbook);
                }
                //改变颜色
                //for (int i = 1; i < sheet.LastRowNum; i++)
                //{
                //    IRow row = sheet.GetRow(i);
                //    if (row != null)
                //    {
                //        ICell cell = sheet.GetRow(i).GetCell(6);
                //        if (cell != null)
                //        {
                //            XSSFCellStyle style = cell.CellStyle as XSSFCellStyle;
                //            style.SetFillBackgroundColor(new XSSFColor(IndexedColors.Yellow.RGB));
                //            cell.CellStyle = style as ICellStyle;
                //            cell.CellStyle.FillBackgroundColor = 0x0FFF;//IndexedColors.Yellow.;
                //        }
                //    }
                //}

                //保存Excel
                if (saveFile)
                {
                    String path = strRootPath + @"\" + txtKeyWord + @".xlsx";
                    int cnt = 0;
                    while (File.Exists(path))
                    {
                        path = strRootPath + @"\" + txtKeyWord + cnt++ + @".xlsx";
                    }
                    using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write))
                    {
                        tempWorkbook.Write(fs);
                    }
                    return path;
                }
            }
            return null;
        }
        #endregion
    }
}
