using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SimpleOrm
{
    public class ExcelOperator
    {
        /// <summary>
        /// 读取excel文件中指定名字的sheet中的全部数据       
        /// 异常：  
        ///      ArgumentNullException: path不存在或者指定的文件不是 excel文件
        ///      
        ///     System.Data.OleDb.OleDbException:打开连接时出现的连接级别错误。
        /// </summary>
        /// <param name="path"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public static DataTable QueryAll(string path, string sheetName)
        {
            DataTable dt = new DataTable();
            OleDbConnection cn = BuildConnection(path);
            try
            {
                TryOpenConnection(cn);
                var cmd = new OleDbCommand("select * from [" + sheetName + "$]", cn);
                var apt = new OleDbDataAdapter(cmd);
                apt.Fill(dt);
            }
            finally
            {
                cn.Close();
                cn.Dispose();
            }
            if (dt.Rows.Count == 0 || dt.Columns.Count == 0)
            {
                throw new Exception(string.Format("sheet: '{0}' doesn't exist any data", sheetName));
            }
            return dt;
        }

        /// <summary>
        /// 将 table 写入到 path中
        /// 异常：
        ///      InvalidOperationException: path不存在或者指定的文件不是 excel文件       
        ///      OleDbException:打开连接时出现的连接级别错误。
        ///      ArgumentNullException：table is null or empty.
        /// </summary>
        /// <param name="table"></param>
        /// <param name="path"></param>
        public static void WriteOnNewSheet(DataTable table, string path)
        {
            if (table == null|| table.Rows.Count == 0 || table.Columns.Count == 0)
            {
                throw new ArgumentNullException("table is null or empty");
            }
            var cn = BuildConnection(path, 0);
            try
            {
                TryOpenConnection(cn);
                var sql = new StringBuilder().Append("CREATE TABLE [" + table.TableName + "]");
                var values = new StringBuilder();
                sql.Append("(");
                foreach (DataColumn column in table.Columns)
                {
                    sql.Append("[" + column.ColumnName + "] text,");
                    values.Append("["+column.ColumnName + "],");
                }
                sql.Remove(sql.Length - 1, 1);
                values.Remove(values.Length - 1, 1);
                sql.Append(")");
                var cmd = new OleDbCommand(sql.ToString(), cn);
                cmd.ExecuteNonQuery();
                StringBuilder sqlValue = new StringBuilder();
                foreach (DataRow row in table.Rows)
                {
                    sql.Clear();
                    sql.Append("INSERT INTO [" + table.TableName + "] (").Append(values).Append(") values (");
                    foreach (DataColumn column in table.Columns)
                    {
                        sql.Append("'"+row[column] + "',");
                    }
                    sql.Remove(sql.Length - 1, 1);
                    sql.Append(")");
                    cmd.CommandText = sql.ToString();
                    cmd.ExecuteNonQuery();
                }
            }
            finally
            {
                cn.Close();
                cn.Dispose();
            }
        }

        /// <summary>
        /// 使用path创建一个 Excel连接
        /// 异常：
        ///     ArgumentNullException:path 不存在或者指定的文件不是 excel文件， 
        ///     
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        private static OleDbConnection BuildConnection(string path, int readOnly = 1)
        {
            if (string.IsNullOrEmpty(path))
            {
                throw new ArgumentNullException("path is not valid");
            }
            if (!File.Exists(path))
            {
                throw new InvalidOperationException(string.Format("the path:{0} is not exist", path));
            }
            var extension = Path.GetExtension(path);
            if (extension != ".xls" && extension != ".xlsx")
            {
                throw new InvalidOperationException(string.Format("path:{0}, the extension,{1}, is  not .xls or .xlsx, invalid operation", path, extension));
            }
            string cnString;
            if (extension == ".xls")
            {
                cnString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR=yes;IMEX={1};'", path, readOnly);
            }
            else
            {
                cnString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0;HDR=yes;IMEX={1};'", path, readOnly);
            }
            return new OleDbConnection(cnString);
        }

        /// <summary>
        /// 尝试打开连接，如果连接已经打开，不会抛出 InvalidOperationException
        /// </summary>
        /// <param name="cn"></param>
        private static void TryOpenConnection(OleDbConnection cn)
        {
            try
            {
                cn.Open();
            }
            catch (InvalidOperationException)
            {

            }
        }        
    }    
}
