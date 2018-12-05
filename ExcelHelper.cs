using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.HSSF.UserModel;
using System.Reflection;
using System.Data;
using System.Data.OleDb;
using Sigo.FIMS.BLL.BusinessEntity;
using System.Web;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Threading;
using System.Security.AccessControl;
using System.IO.MemoryMappedFiles;
using System.Data.Odbc;

namespace Sigo.FIMS.BLL.Helper
{
    /// <summary>
    /// HttpRequest
    /// </summary>
    public class ExcelHelper
    {
        /// <summary>
        /// 生成Workbook
        /// </summary>
        /// <param name="dictTitle">Sheet 表头</param>
        /// <param name="list">数据源</param>
        /// <returns>HSSFWorkbook</returns>
        public HSSFWorkbook CreateSheet<T>(Dictionary<string, string> dictTitle, List<T> list) where T : class
        {
            #region 声明
            HSSFWorkbook hssfworkbook = new HSSFWorkbook();
            HSSFSheet sheet = (HSSFSheet)hssfworkbook.CreateSheet("Sheet1");
            HSSFCellStyle styleTitle = (HSSFCellStyle)hssfworkbook.CreateCellStyle();
            HSSFCellStyle styleContent = (HSSFCellStyle)hssfworkbook.CreateCellStyle();
            int cell = 0, row = 0;  // cell 行索引，row 列索引
            #endregion

            #region 标题行
            if (dictTitle == null || dictTitle.Count == 0)
            {
                dictTitle = new Dictionary<string, string>();

                T t = null;
                foreach (System.Reflection.PropertyInfo p in t.GetType().GetProperties())
                {
                    if (!dictTitle.ContainsKey(p.Name))
                    {
                        dictTitle.Add(p.Name, p.Name);
                    }
                }
            }

            HSSFRow rowTitle = (HSSFRow)sheet.CreateRow(row);
            foreach (var kv in dictTitle)
            {
                HSSFCell cellTitle = (HSSFCell)rowTitle.CreateCell(cell);
                cellTitle.SetCellValue(kv.Value);
                cellTitle.CellStyle = styleTitle;
                cell++;
            }
            #endregion

            #region 内容行
            int pageNum = 1;
            foreach (T t in list)
            {
                // Excel 单个Sheet 最大 65536行,超出部分自动写入下一个sheet。
                row++;
                if (row == 65536)
                {
                    pageNum++;
                    row = 1;
                    sheet = (HSSFSheet)hssfworkbook.CreateSheet(string.Format("Sheet{0}", pageNum.ToString()));
                }

                HSSFRow rowContent = (HSSFRow)sheet.CreateRow(row);
                int i = 0;
                foreach (var kv in dictTitle)
                {
                    HSSFCell cellContent = (HSSFCell)rowContent.CreateCell(i);

                    Type type = t.GetType();
                    PropertyInfo pi = type.GetProperty(kv.Key);
                    if (pi == null)
                    {
                        cellContent.SetCellType(NPOI.SS.UserModel.CellType.String);
                        cellContent.SetCellValue("NULL");
                        continue;
                    }

                    object obj = pi.GetValue(t, null);
                    switch (pi.PropertyType.FullName)
                    {
                        case "System.Int32":
                            cellContent.SetCellType(NPOI.SS.UserModel.CellType.Numeric);
                            cellContent.SetCellValue(Convert.ToInt32(obj));
                            break;
                        case "System.Decimal":
                            cellContent.SetCellType(NPOI.SS.UserModel.CellType.Numeric);
                            cellContent.SetCellValue(Convert.ToDouble(obj));
                            styleContent.DataFormat = HSSFDataFormat.GetBuiltinFormat("#,##0.00");
                            cellContent.CellStyle = styleContent;
                            break;
                        case "System.String":
                            cellContent.SetCellType(NPOI.SS.UserModel.CellType.String);
                            cellContent.SetCellValue(Convert.ToString(obj));
                            break;
                        case "System.DateTime":
                            cellContent.SetCellType(NPOI.SS.UserModel.CellType.String);
                            cellContent.SetCellValue(Convert.ToDateTime(obj).ToString("yyyy-MM-dd HH:mm:ss"));
                            break;

                        case "System.Boolean":
                            cellContent.SetCellType(NPOI.SS.UserModel.CellType.String);
                            cellContent.SetCellValue(Convert.ToBoolean(obj) ? "是" : "否");
                            break;
                        default:
                            cellContent.SetCellType(NPOI.SS.UserModel.CellType.String);
                            cellContent.SetCellValue(Convert.ToString(obj));
                            break;
                    }

                    i++;
                }
            }
            #endregion

            #region 自适应列宽度
            for (int i = 0; i < dictTitle.Count; i++)
            {
                sheet.AutoSizeColumn(i);
            }
            #endregion

            return hssfworkbook;
        }

        #region GetToData
        /// <summary>
        /// HSSFWorkbook To DataTable
        /// </summary>
        /// <returns></returns>
        public DataTable GetDataTable(int? userId, string filePath, out string message)
        {
            message = string.Empty;
            bool exists = IsExists(filePath, out message);
            if (!exists)
            {
                return null;
            }

            string fileName = System.IO.Path.GetFileName(filePath);
            try
            {
                IWorkbook workbook = null;
                ISheet sheet = null;
                IRow row = null;

                FileStream streamfile = new FileStream(filePath, FileMode.Open, FileAccess.Read);

                // 2007版本  
                if (filePath.IndexOf(".xlsx") > 0)
                    workbook = new XSSFWorkbook(streamfile);
                // 2003版本  
                else if (filePath.IndexOf(".xls") > 0)
                    workbook = new HSSFWorkbook(streamfile);

                sheet = workbook.GetSheetAt(0);


                DataTable table = new DataTable();
                IRow headerRow = sheet.GetRow(0);//第一行为标题行  
                int cellCount = headerRow.LastCellNum;//总列
                int rowCount = sheet.LastRowNum;//总行

                for (int i = headerRow.FirstCellNum; i < cellCount; i++)
                {
                    DataColumn column = new DataColumn(headerRow.GetCell(i).StringCellValue);
                    table.Columns.Add(column);
                }

                DataColumn colAdder = new DataColumn("Adder");
                DataColumn colAddTime = new DataColumn("AddTime");
                DataColumn colFileName = new DataColumn("FileName");
                table.Columns.Add(colAdder);
                table.Columns.Add(colAddTime);
                table.Columns.Add(colFileName);
                DateTime addTime = DateTime.Now;
                for (int i = (sheet.FirstRowNum + 1); i <= rowCount; i++)
                {
                    row = sheet.GetRow(i);
                    DataRow dataRow = table.NewRow();
                    if (row != null)
                    {
                        dataRow["Adder"] = userId;
                        dataRow["AddTime"] = addTime;
                        dataRow["FileName"] = fileName;
                        for (int j = row.FirstCellNum; j < cellCount; j++)
                        {
                            ICell cell = row.Cells[j];
                            switch (cell.CellType)
                            {
                                case CellType.Blank:
                                    dataRow[j] = string.Empty;
                                    break;
                                case CellType.String:
                                    dataRow[j] = cell.StringCellValue;
                                    break;
                                case CellType.Numeric:
                                    if (HSSFDateUtil.IsCellDateFormatted(cell))
                                    {
                                        dataRow[j] = cell.DateCellValue;
                                    }
                                    else if (HSSFDateUtil.IsCellInternalDateFormatted(cell))
                                    {
                                        dataRow[j] = cell.DateCellValue;
                                    }
                                    else
                                    {
                                        dataRow[j] = cell.NumericCellValue;
                                    }
                                    break;
                                case CellType.Boolean:
                                    dataRow[j] = cell.BooleanCellValue;
                                    break;
                                case CellType.Error:
                                    dataRow[j] = cell.ErrorCellValue.ToString();
                                    break;
                                case CellType.Formula:
                                    HSSFFormulaEvaluator eva = new HSSFFormulaEvaluator(workbook);
                                    dataRow[i] = eva.Evaluate(cell).StringValue;
                                    break;
                                default:
                                    dataRow[j] = cell.ToString();
                                    break;
                            }
                        }
                    }

                    table.Rows.Add(dataRow);
                }

                return table;
            }
            catch (Exception ex)
            {
                message += ex.Message;
                return null;
            }
        }

        /// <summary>
        /// Excel To DataTable
        /// </summary>
        /// <param name="userId"></param>
        /// <param name="filePath"></param>
        /// <param name="message"></param>
        /// <returns></returns>
        public DataTable GetDataTable_Excel(int? userId, string filePath, out string message)
        {
            message = string.Empty;
            string sheetName = "Sheet1$";
            string fileName = System.IO.Path.GetFileName(filePath);
            DataSet ds = new DataSet();
            try
            {
                string extension = System.IO.Path.GetExtension(filePath);
                OleDbConnection objConn = null;
                switch (extension.ToLower())
                {
                    case ".xls":
                        objConn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties=Excel 8.0;");
                        break;
                    case ".xlsx":
                        objConn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties = 'Excel 12.0;HDR=Yes;IMEX=1'");
                        break;
                }

                objConn.Open();

                //DataTable schemaTable = objConn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, null);    // Excel 所有Sheet
                //sheetName = schemaTable.Rows[0][2].ToString().Trim();    // 获取 Excel 的第一个sheet的表名，默认值是sheet1

                string strSql = string.Format("SELECT {1} AS Adder,'{2}' AS AddTime, '{3}' AS FileName, * from [{0}]",
                    sheetName, userId, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), fileName);
                OleDbCommand objCmd = new OleDbCommand(strSql, objConn);
                // OleDbDataReader reader = objCmd.ExecuteReader();                
                OleDbDataAdapter myData = new OleDbDataAdapter(strSql, objConn);
                myData.Fill(ds, sheetName);
                objConn.Close();
                return ds.Tables[0];
            }
            catch (Exception ex)
            {
                message += ex.Message;
                return null;
            }
        }

        /// <summary>
        /// CSV To DataTable
        /// </summary>
        /// <param name="userId"></param>
        /// <param name="filePath"></param>
        /// <param name="message"></param>
        /// <returns></returns>
        public DataTable GetDataTable_CSV(int? userId, string filePath, out string message)
        {
            message = string.Empty;
            bool exists = IsExists(filePath, out message);
            if (!exists)
            {
                return null;
            }

            #region 读取文件
            try
            {
                #region 作废
                //StringBuilder resultAsString = new StringBuilder();
                //long offset = 0x10000000; // 256 megabytes
                //long length = 0x20000000; // 512 megabytes

                //// Create the memory-mapped file.
                //using (var memoryMappedFile = MemoryMappedFile.CreateFromFile(filePath))
                //{
                //    // Create a random access view, from the 256th megabyte (the offset)
                //    // to the 768th megabyte (the offset plus length).
                //    using (var memoryMappedViewStream = memoryMappedFile.CreateViewStream(0, length))
                //    {
                //        for (int i = 0; i < length; i++)
                //        {
                //            //Reads a byte from a stream and advances the position within the stream by one byte, or returns -1 if at the end of the stream.
                //            int result = memoryMappedViewStream.ReadByte();
                //            if (result == -1)
                //            {
                //                break;
                //            }

                //            char letter = (char)result;
                //            resultAsString.Append(letter);
                //        }
                //    }
                //}
                #endregion

                DataTable dt = new DataTable();

                using (FileStream fs = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    BufferedStream bs = new BufferedStream(fs);
                    StreamReader sr = new StreamReader(bs, Encoding.Default);
                    string line;
                    int i = 0;
                    int headLine = 0;
                    while ((line = sr.ReadLine()) != null)
                    {
                        i++;
                        if (line.IndexOf('#') == -1)
                        {
                            List<string> list = line.Split(',').ToList();
                            if (headLine == 0)
                            {
                                #region 头部
                                foreach (var val in list)
                                {
                                    DataColumn column = new DataColumn(val, Type.GetType("System.String"));
                                    dt.Columns.Add(column);
                                }
                                #endregion
                            }
                            else
                            {
                                #region 内容
                                DataRow dr = dt.NewRow();
                                for (int j = 0; j < list.Count; j++)
                                {
                                    dr[j] = list[j].ToString();
                                }

                                dt.Rows.Add(dr);
                                #endregion

                            }
                        }
                    }

                    sr.Close();
                    bs.Close();
                    fs.Close();
                }

                #region read linq
                //var Lines = File.ReadLines(filePath).Select(a => a.Split(';'));
                //var CSV = from line in Lines
                //          select (line).ToArray();
                #endregion
            }
            catch (Exception ex)
            {
            }
            #endregion
            return null;
        }
        #endregion

        #region Odbc GetFile
        /// <summary>
        /// Odbc查询CSV
        /// </summary>
        /// 
        public static void QueryCSVToOdbc(string filePath)
        {
            FileInfo fileinfo = new FileInfo(filePath);
            string tableName = fileinfo.Name;
            string fileFullName = fileinfo.FullName;

            //文件路径
            //string filePath = AppDomain.CurrentDomain.BaseDirectory;
            string pContent = string.Empty;

            OdbcConnection odbcConn = new OdbcConnection();
            OdbcCommand odbcCmd = new OdbcCommand();
            OdbcDataReader dataReader;
            try
            {
                string strConnOledb = "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=";
                strConnOledb += fileFullName;
                strConnOledb += ";Extensions=csv,txt;";

                odbcConn.ConnectionString = strConnOledb;
                odbcConn.Open();
                StringBuilder commandText = new StringBuilder("SELECT ");
                commandText.AppendFormat("* From {0}", tableName);
                odbcCmd.Connection = odbcConn;
                odbcCmd.CommandText = commandText.ToString();
                dataReader = odbcCmd.ExecuteReader();



                while (dataReader.Read())
                {
                    pContent = Convert.ToString(dataReader["content"]);
                }
                dataReader.Close();
            }
            catch (System.Exception ex)
            {
                odbcConn.Close();
            }
            finally
            {
                odbcConn.Close();
            }
        }

        /// <summary>
        /// Oledb查询CSV
        /// </summary>
        public static void QueryCSVToOledb(string filePath)
        {
            FileInfo fileinfo = new FileInfo(filePath);
            string tableName = fileinfo.Name;
            string fileFullName = fileinfo.FullName;

            string pContent = string.Empty;
            OleDbConnection oledbConn = new OleDbConnection();
            OleDbCommand oledbCmd = new OleDbCommand();
            OleDbDataReader dataReader;
            try
            {
                //两种连接方式皆可
                //string strConnOledb = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=";
                string strConnOledb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=";
                strConnOledb += fileFullName;
                strConnOledb += ";Extended Properties='Text;HDR=Yes;IMEX=1;'";

                oledbConn.ConnectionString = strConnOledb;
                oledbConn.Open();
                StringBuilder commandText = new StringBuilder("SELECT ");
                commandText.AppendFormat("* From {0}", tableName);
                oledbCmd.Connection = oledbConn;
                oledbCmd.CommandText = commandText.ToString();
                dataReader = oledbCmd.ExecuteReader();


                while (dataReader.Read())
                {
                    pContent = Convert.ToString(dataReader["content"]);
                }
                dataReader.Close();
            }
            catch (System.Exception ex)
            {
                oledbConn.Close();
            }
            finally
            {
                oledbConn.Close();
            }
        }

        /// <summary>
        /// Odbc查询Excel
        /// </summary>
        public static void QueryExcelToOdbc(string filePath)
        {
            //文件路径
            FileInfo fileinfo = new FileInfo(filePath);
            string tableName = fileinfo.Name;
            filePath = fileinfo.FullName;

            OdbcConnection odbcConn = new OdbcConnection();
            OdbcCommand odbcCmd = new OdbcCommand();
            OdbcDataReader dataReader;
            try
            {
                //连接字符串
                string strConnOledb = "Driver={Microsoft Excel Driver (*.xls)};Dbq=";
                strConnOledb += filePath;
                strConnOledb += ";Extended=xls";

                odbcConn.ConnectionString = strConnOledb;
                odbcConn.Open();
                StringBuilder commandText = new StringBuilder("SELECT ");
                commandText.AppendFormat("* From {0}", "[Sheet1$]");
                odbcCmd.Connection = odbcConn;
                odbcCmd.CommandText = commandText.ToString();
                dataReader = odbcCmd.ExecuteReader();
                while (dataReader.Read())
                {
                    string pContent = Convert.ToString(dataReader["content"]);
                }
                dataReader.Close();
            }
            catch (System.Exception ex)
            {
                odbcConn.Close();
            }
            finally
            {
                odbcConn.Close();
            }
        }

        /// <summary>
        /// Oledb查询Excel
        /// </summary>
        public static void QueryExcelToOledb(string filePath)
        {
            //文件路径
            FileInfo fileinfo = new FileInfo(filePath);
            string tableName = fileinfo.Name;
            filePath = fileinfo.FullName;

            OleDbConnection oledbConn = new OleDbConnection();
            OleDbCommand oledbCmd = new OleDbCommand();
            OleDbDataReader dataReader;
            try
            {
                //一下两种方式皆可
                string strConnOledb = string.Empty;

                if (fileinfo.Extension == ".xlsx")
                {
                    strConnOledb = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=";
                    strConnOledb += filePath;
                    strConnOledb += ";Extended Properties='Excel 12.0;HDR=Yes;IMEX=1;'";
                }
                else
                {
                    strConnOledb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=";
                    strConnOledb += filePath;
                    strConnOledb += ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'";
                }

                oledbConn.ConnectionString = strConnOledb;
                oledbConn.Open();
                StringBuilder commandText = new StringBuilder("SELECT ");
                commandText.AppendFormat("* From {0}", "[Sheet1$]");
                oledbCmd.Connection = oledbConn;
                oledbCmd.CommandText = commandText.ToString();
                dataReader = oledbCmd.ExecuteReader();
                while (dataReader.Read())
                {
                    string pContent = Convert.ToString(dataReader["content"]);
                }
                dataReader.Close();
            }
            catch (System.Exception ex)
            {
                oledbConn.Close();
            }
            finally
            {
                oledbConn.Close();
            }
        }
        #endregion

        /// <summary>
        /// 判读文件是否存在（负载均衡 - 文件同步延迟问题处理）
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="errorMsg"></param>
        /// <returns></returns>
        private bool IsExists(string filePath, out string errorMsg)
        {
            bool flag = false;
            errorMsg = string.Empty;

            #region 负载均衡 - 文件同步延迟问题处理
            int maxSleepTimes = 15;
            for (int i = 0; i < maxSleepTimes; i++)
            {
                if (!File.Exists(filePath))
                {
                    Thread.Sleep(500);
                }
                else
                {
                    #region 异常捕获的方式判断，文件是否有权限；没权限手动赋予权限
                    try
                    {
                        FileStream streamfile = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                        flag = true;
                        streamfile.Dispose();
                        streamfile.Close();
                    }
                    catch (UnauthorizedAccessException ex1)
                    {
                        FileInfo fileInfo = new FileInfo(filePath);
                        FileSecurity fileSecurity = fileInfo.GetAccessControl();
                        fileSecurity.AddAccessRule(new FileSystemAccessRule("Everyone", FileSystemRights.FullControl, AccessControlType.Allow));     //以完全控制为例
                        fileInfo.SetAccessControl(fileSecurity);
                    }
                    catch (Exception ex)
                    {
                        errorMsg += string.Format("{0}", ex.Message);
                    }
                    #endregion

                    break;
                }
            }
            #endregion

            return flag;
        }
    }
}
