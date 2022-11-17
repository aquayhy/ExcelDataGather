using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Collections;
using System.Data.OleDb;
using System.Collections.Generic;

namespace ExcelDataGather
{
    public abstract class ExcelHelper
    {
        ///用于读取EXCEL文件的类，从SQLHelper修改

        //Database connection strings
        public static string ConnectionString;
        public static void SetConnectionString(String filename)
        {
            string OleDbConnStringXLSX = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\'{0}\';Extended Properties=\'Excel 12.0;HDR=yes\'";
            string OleDbConnStringXLS = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\'{0}\';Extended Properties=\'Excel 8.0;HDR=yes\'";
            if (System.IO.Path.GetExtension(filename) == ".xls")
                ConnectionString = String.Format(OleDbConnStringXLS, filename);
            else
                ConnectionString = String.Format(OleDbConnStringXLSX, filename);
        }

        // Hashtable to store cached parameters
        private static Hashtable parmCache = Hashtable.Synchronized(new Hashtable());

        /// <summary>
        /// Execute a SqlCommand (that returns no resultset) against the database specified in the connection string 
        /// using the provided parameters.
        /// </summary>
        /// <remarks>
        /// e.g.:  
        ///  int result = ExecuteNonQuery(connString, CommandType.StoredProcedure, "PublishOrders", new SqlParameter("@prodid", 24));
        /// </remarks>
        /// <param name="connectionString">a valid connection string for a SqlConnection</param>
        /// <param name="commandType">the CommandType (stored procedure, text, etc.)</param>
        /// <param name="commandText">the stored procedure name or T-SQL command</param>
        /// <param name="commandParameters">an array of SqlParamters used to execute the command</param>
        /// <returns>an int representing the number of rows affected by the command</returns>
        public static int ExecuteNonQuery(string connectionString, CommandType cmdType, string cmdText, params OleDbParameter[] commandParameters)
        {

            OleDbCommand cmd = new OleDbCommand();

            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                PrepareCommand(cmd, conn, null, cmdType, cmdText, commandParameters);
                int val = cmd.ExecuteNonQuery();
                cmd.Parameters.Clear();
                return val;
            }
        }

        /// <summary>
        /// Execute a SqlCommand (that returns no resultset) against an existing database connection 
        /// using the provided parameters.
        /// </summary>
        /// <remarks>
        /// e.g.:  
        ///  int result = ExecuteNonQuery(connString, CommandType.StoredProcedure, "PublishOrders", new SqlParameter("@prodid", 24));
        /// </remarks>
        /// <param name="conn">an existing database connection</param>
        /// <param name="commandType">the CommandType (stored procedure, text, etc.)</param>
        /// <param name="commandText">the stored procedure name or T-SQL command</param>
        /// <param name="commandParameters">an array of SqlParamters used to execute the command</param>
        /// <returns>an int representing the number of rows affected by the command</returns>
        public static int ExecuteNonQuery(OleDbConnection connection, CommandType cmdType, string cmdText, params OleDbParameter[] commandParameters)
        {

            OleDbCommand cmd = new OleDbCommand();

            PrepareCommand(cmd, connection, null, cmdType, cmdText, commandParameters);
            int val = cmd.ExecuteNonQuery();
            cmd.Parameters.Clear();
            return val;
        }

        /// <summary>
        /// Execute a SqlCommand (that returns no resultset) using an existing SQL Transaction 
        /// using the provided parameters.
        /// </summary>
        /// <remarks>
        /// e.g.:  
        ///  int result = ExecuteNonQuery(connString, CommandType.StoredProcedure, "PublishOrders", new SqlParameter("@prodid", 24));
        /// </remarks>
        /// <param name="trans">an existing sql transaction</param>
        /// <param name="commandType">the CommandType (stored procedure, text, etc.)</param>
        /// <param name="commandText">the stored procedure name or T-SQL command</param>
        /// <param name="commandParameters">an array of SqlParamters used to execute the command</param>
        /// <returns>an int representing the number of rows affected by the command</returns>
        public static int ExecuteNonQuery(OleDbTransaction trans, CommandType cmdType, string cmdText, params OleDbParameter[] commandParameters)
        {
            OleDbCommand cmd = new OleDbCommand();
            PrepareCommand(cmd, trans.Connection, trans, cmdType, cmdText, commandParameters);
            int val = cmd.ExecuteNonQuery();
            cmd.Parameters.Clear();
            return val;
        }

        /// <summary>
        /// Execute a SqlCommand that returns a resultset against the database specified in the connection string 
        /// using the provided parameters.
        /// </summary>
        /// <remarks>
        /// e.g.:  
        ///  SqlDataReader r = ExecuteReader(connString, CommandType.StoredProcedure, "PublishOrders", new SqlParameter("@prodid", 24));
        /// </remarks>
        /// <param name="connectionString">a valid connection string for a SqlConnection</param>
        /// <param name="commandType">the CommandType (stored procedure, text, etc.)</param>
        /// <param name="commandText">the stored procedure name or T-SQL command</param>
        /// <param name="commandParameters">an array of SqlParamters used to execute the command</param>
        /// <returns>A SqlDataReader containing the results</returns>
        public static OleDbDataReader ExecuteReader(string connectionString, CommandType cmdType, string cmdText, params OleDbParameter[] commandParameters)
        {
            OleDbCommand cmd = new OleDbCommand();
            OleDbConnection conn = new OleDbConnection(connectionString);

            // we use a try/catch here because if the method throws an exception we want to 
            // close the connection throw code, because no datareader will exist, hence the 
            // commandBehaviour.CloseConnection will not work
            try
            {
                PrepareCommand(cmd, conn, null, cmdType, cmdText, commandParameters);
                OleDbDataReader rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection);
                cmd.Parameters.Clear();
                return rdr;
            }
            catch
            {
                conn.Close();
                throw;
            }
        }

        /// <summary>
        /// Execute a SqlCommand that returns the first column of the first record against the database specified in the connection string 
        /// using the provided parameters.
        /// </summary>
        /// <remarks>
        /// e.g.:  
        ///  Object obj = ExecuteScalar(connString, CommandType.StoredProcedure, "PublishOrders", new SqlParameter("@prodid", 24));
        /// </remarks>
        /// <param name="connectionString">a valid connection string for a SqlConnection</param>
        /// <param name="commandType">the CommandType (stored procedure, text, etc.)</param>
        /// <param name="commandText">the stored procedure name or T-SQL command</param>
        /// <param name="commandParameters">an array of SqlParamters used to execute the command</param>
        /// <returns>An object that should be converted to the expected type using Convert.To{Type}</returns>
        public static object ExecuteScalar(string connectionString, CommandType cmdType, string cmdText, params OleDbParameter[] commandParameters)
        {
            OleDbCommand cmd = new OleDbCommand();

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                PrepareCommand(cmd, connection, null, cmdType, cmdText, commandParameters);
                object val = cmd.ExecuteScalar();
                cmd.Parameters.Clear();
                return val;
            }
        }

        /// <summary>
        /// Execute a SqlCommand that returns the first column of the first record against an existing database connection 
        /// using the provided parameters.
        /// </summary>
        /// <remarks>
        /// e.g.:  
        ///  Object obj = ExecuteScalar(connString, CommandType.StoredProcedure, "PublishOrders", new SqlParameter("@prodid", 24));
        /// </remarks>
        /// <param name="conn">an existing database connection</param>
        /// <param name="commandType">the CommandType (stored procedure, text, etc.)</param>
        /// <param name="commandText">the stored procedure name or T-SQL command</param>
        /// <param name="commandParameters">an array of SqlParamters used to execute the command</param>
        /// <returns>An object that should be converted to the expected type using Convert.To{Type}</returns>
        public static object ExecuteScalar(OleDbConnection connection, CommandType cmdType, string cmdText, params OleDbParameter[] commandParameters)
        {

            OleDbCommand cmd = new OleDbCommand();

            PrepareCommand(cmd, connection, null, cmdType, cmdText, commandParameters);
            object val = cmd.ExecuteScalar();
            cmd.Parameters.Clear();
            return val;
        }

        /// <summary>
        /// add parameter array to the cache
        /// </summary>
        /// <param name="cacheKey">Key to the parameter cache</param>
        /// <param name="cmdParms">an array of SqlParamters to be cached</param>
        public static void CacheParameters(string cacheKey, params OleDbParameter[] commandParameters)
        {
            parmCache[cacheKey] = commandParameters;
        }

        /// <summary>
        /// Retrieve cached parameters
        /// </summary>
        /// <param name="cacheKey">key used to lookup parameters</param>
        /// <returns>Cached SqlParamters array</returns>
        public static OleDbParameter[] GetCachedParameters(string cacheKey)
        {
            OleDbParameter[] cachedParms = (OleDbParameter[])parmCache[cacheKey];

            if (cachedParms == null)
                return null;

            OleDbParameter[] clonedParms = new OleDbParameter[cachedParms.Length];

            for (int i = 0, j = cachedParms.Length; i < j; i++)
                clonedParms[i] = (OleDbParameter)((ICloneable)cachedParms[i]).Clone();

            return clonedParms;
        }

        /// <summary>
        /// Prepare a command for execution
        /// </summary>
        /// <param name="cmd">SqlCommand object</param>
        /// <param name="conn">SqlConnection object</param>
        /// <param name="trans">SqlTransaction object</param>
        /// <param name="cmdType">Cmd type e.g. stored procedure or text</param>
        /// <param name="cmdText">Command text, e.g. Select * from Products</param>
        /// <param name="cmdParms">SqlParameters to use in the command</param>
        private static void PrepareCommand(OleDbCommand cmd, OleDbConnection conn, OleDbTransaction trans, CommandType cmdType, string cmdText, OleDbParameter[] cmdParms)
        {

            if (conn.State != ConnectionState.Open)
                conn.Open();

            cmd.Connection = conn;
            cmd.CommandText = cmdText;

            if (trans != null)
                cmd.Transaction = trans;

            cmd.CommandType = cmdType;

            if (cmdParms != null)
            {
                foreach (OleDbParameter parm in cmdParms)
                    cmd.Parameters.Add(parm);
            }
        }

        private static void PrepareInsertCommand(OleDbCommand cmd, OleDbConnection conn, OleDbTransaction trans, CommandType cmdType, string cmdText, OleDbParameter[] cmdParms)
        {

            if (conn.State != ConnectionState.Open)
                conn.Open();

            cmd.Connection = conn;
            cmd.CommandText = cmdText;

            if (trans != null)
                cmd.Transaction = trans;

            cmd.CommandType = cmdType;

            if (cmdParms != null)
            {
                OleDbParameter tempparm = new OleDbParameter("ReferenceID", OleDbType.Integer, 4, ParameterDirection.ReturnValue, false, 0, 0, string.Empty, DataRowVersion.Default, null);
                cmd.Parameters.Add(tempparm);
                foreach (OleDbParameter parm in cmdParms)
                    cmd.Parameters.Add(parm);
            }
        }

        public static int Insert(string connectionString, CommandType cmdType, string cmdText, params OleDbParameter[] commandParameters)
        {

            OleDbCommand cmd = new OleDbCommand();

            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                PrepareInsertCommand(cmd, conn, null, cmdType, cmdText, commandParameters);
                int val = cmd.ExecuteNonQuery();
                if (val == 1)
                {
                    val = Convert.ToInt32(cmd.Parameters[0].Value);
                }
                else
                {
                    val = 0;
                }
                cmd.Parameters.Clear();
                return val;
            }
        }

        public static DataSet TransExcelToDataSet(string fileName)
        {
            List<string> sheetNames = GetAllSheetNameOfExcel(fileName);
            return (TransExcelToDataSet(fileName, sheetNames));
        }

        public static List<string> GetAllSheetNameOfExcel(string fileName)
        {
            List<String> allTableName = new List<string>();
            string strConn;
            string OleDbConnStringXLSX = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\'{0}\';Extended Properties=\'Excel 12.0;HDR=yes;IMEX=1\'";
            string OleDbConnStringXLS = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\'{0}\';Extended Properties=\'Excel 8.0;HDR=yes;IMEX=1\'";
            if (System.IO.Path.GetExtension(fileName) == ".xls")
                strConn = String.Format(OleDbConnStringXLS, fileName);
            else
                strConn = String.Format(OleDbConnStringXLSX, fileName);
            using (OleDbConnection objConn = new OleDbConnection(strConn))
            {
                if (objConn.State != ConnectionState.Open)
                    objConn.Open();
                DataTable sheetNames = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                foreach (DataRow dr in sheetNames.Rows)
                {
                    allTableName.Add(dr[2].ToString().Replace("$", ""));
                }
            }
            return allTableName;
        }

        public static DataTable TransExcelToDataTable(string fileName)
        {
            return TransExcelToDataTable(fileName, "Sheet1");
        }

        public static DataTable TransExcelToDataTable(string fileName, string sheetName)
        {
            OleDbConnection objConn = null;
            DataSet data = new DataSet();
            string strConn;
            string OleDbConnStringXLSX = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\'{0}\';Extended Properties=\'Excel 12.0;HDR=yes;IMEX=1\'";
            string OleDbConnStringXLS = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\'{0}\';Extended Properties=\'Excel 8.0;HDR=yes;IMEX=1\'";
            if (System.IO.Path.GetExtension(fileName) == ".xls")
                strConn = String.Format(OleDbConnStringXLS, fileName);
            else
                strConn = String.Format(OleDbConnStringXLSX, fileName);

            try
            {

                objConn = new OleDbConnection(strConn);
                using (objConn)
                {
                    OleDbDataAdapter sqlada = null;
                    //遍历从配置文件中读取的sheet名称
                    string strSql = "select * From [" + sheetName.Trim() + "$]";
                    sqlada = new OleDbDataAdapter(strSql, objConn);
                    //填充dataset
                    sqlada.Fill(data, sheetName);
                }
            }
            catch (Exception e)
            {
                throw new Exception("将excel中指定sheet内容读入dataset出错！" + e.Message + " strConn: " + strConn + " ; fileName:" + fileName);
                //throw e;
            }
            finally
            {
                if (objConn != null)
                {
                    objConn.Dispose();
                    objConn.Close();
                }

                GC.Collect();
            }
            return data.Tables[0];
        }

        public static DataSet TransExcelToDataSet(string fileName, List<string> sheetNames)
        {
            OleDbConnection objConn = null;
            DataSet data = new DataSet();
            string strConn;
            string OleDbConnStringXLSX = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\'{0}\';Extended Properties=\'Excel 12.0;HDR=yes;IMEX=1\'";
            string OleDbConnStringXLS = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\'{0}\';Extended Properties=\'Excel 8.0;HDR=yes;IMEX=1\'";
            if (System.IO.Path.GetExtension(fileName) == ".xls")
                strConn = String.Format(OleDbConnStringXLS, fileName);
            else
                strConn = String.Format(OleDbConnStringXLSX, fileName);

            try
            {

                objConn = new OleDbConnection(strConn);
                using (objConn)
                {
                    OleDbDataAdapter sqlada = null;
                    //遍历从配置文件中读取的sheet名称
                    foreach (string sheetName in sheetNames)
                    {
                        if (!string.IsNullOrEmpty(sheetName))
                        {
                            string strSql = "select * From [" + sheetName.Trim() + "$]";
                            sqlada = new OleDbDataAdapter(strSql, objConn);
                            //填充dataset
                            sqlada.Fill(data, sheetName);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                throw new Exception("将excel中指定sheet内容读入dataset出错！" + e.Message + " strConn: " + strConn + " ; fileName:" + fileName);
                //throw e;
            }
            finally
            {
                if (objConn != null)
                {
                    objConn.Dispose();
                    objConn.Close();
                }

                GC.Collect();
            }
            return data;
        }
    }
}
