using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ModelConverter
{
    class DatabaseHelper
    {
        string CONNECTION_STRING = "Data Source=192.168.1.230;Initial Catalog=Demo;Persist Security Info=True;User id=sa;Password=admin@123;";
        SqlConnection sqlConnect = new SqlConnection();
        SqlCommand cmd = new SqlCommand();

        public DatabaseHelper()
        {
            CONNECTION_STRING = string.Format("Data Source = {0}; Initial Catalog ={1};Persist Security Info=True;User id ={2}; Password={3}; ",
                                              ConfigurationManager.AppSettings["Server"],
                                              ConfigurationManager.AppSettings["Database"],
                                              ConfigurationManager.AppSettings["UserName"],
                                              ConfigurationManager.AppSettings["Password"]);
        }

        public bool TestConnection()
        {
            try
            {
                using (sqlConnect = new SqlConnection(CONNECTION_STRING))
                {
                    sqlConnect.Open();
                    Console.WriteLine("connect success");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (sqlConnect.State != ConnectionState.Closed)
                {
                    sqlConnect.Close();
                }
            }
            return true;
        }

        /// <summary>
        /// Get all table names
        /// </summary>
        /// <returns></returns>
        public DataTable GetAllTable()
        {
            //Only table
            //"SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE='BASE TABLE' ORDER BY TABLE_NAME"
            //Table and view
            //"SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES ORDER BY TABLE_NAME"
            string commandText = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES ORDER BY TABLE_NAME";
            DataTable returnTable = new DataTable();
            returnTable.Columns.Add(TableName.NAME, typeof(string));
            try
            {
                using (sqlConnect = new SqlConnection(CONNECTION_STRING))
                {
                    sqlConnect.Open();
                    cmd = new SqlCommand();
                    cmd.Connection = sqlConnect;
                    cmd.CommandText = commandText;
                    cmd.CommandType = CommandType.Text;
                    SqlDataReader reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        DataRow r = returnTable.NewRow();
                        r[0] = reader[0];
                        returnTable.Rows.Add(r);
                    }
                    Console.WriteLine("Get all tables success");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (sqlConnect.State != ConnectionState.Closed)
                {
                    sqlConnect.Close();
                }
            }
            return returnTable;
        }

        /// <summary>
        /// Get all columns in all tables
        /// </summary>
        /// <returns></returns>
        public DataTable GetAllColumn()
        {
            string commandText = "select TABLE_NAME, COLUMN_NAME, ORDINAL_POSITION, COLUMN_DEFAULT, IS_NULLABLE, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH from information_schema.columns";
            DataTable returnTable = new DataTable();
            returnTable.Columns.Add(ColumnName.TableName, typeof(string));
            returnTable.Columns.Add(ColumnName.ColName, typeof(string));
            returnTable.Columns.Add(ColumnName.Position, typeof(int));
            returnTable.Columns.Add(ColumnName.Default, typeof(string));
            returnTable.Columns.Add(ColumnName.IsNull, typeof(string));
            returnTable.Columns.Add(ColumnName.DataType, typeof(string));
            returnTable.Columns.Add(ColumnName.MaxLength, typeof(string));
            try
            {
                using (sqlConnect = new SqlConnection(CONNECTION_STRING))
                {
                    sqlConnect.Open();
                    cmd = new SqlCommand();
                    cmd.Connection = sqlConnect;
                    cmd.CommandText = commandText;
                    cmd.CommandType = CommandType.Text;
                    SqlDataReader reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        DataRow r = returnTable.NewRow();
                        r[ColumnName.TableName] = reader[0];
                        r[ColumnName.ColName] = reader[1];
                        r[ColumnName.Position] = (int)reader[2];
                        r[ColumnName.Default] = reader[3];
                        r[ColumnName.IsNull] = reader[4];
                        r[ColumnName.DataType] = reader[5];
                        r[ColumnName.MaxLength] = reader[6];
                        returnTable.Rows.Add(r);
                    }
                    Console.WriteLine("Get all columns success");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (sqlConnect.State != ConnectionState.Closed)
                {
                    sqlConnect.Close();
                }
            }
            return returnTable;
        }

        /// <summary>
        /// Get list of infor column by table
        /// </summary>
        /// <param name="tableName">Table name</param>
        /// <returns></returns>
        public DataTable GetColumnInfor(string tableName)
        {
            StringBuilder commandText = new StringBuilder();
            commandText.AppendLine("SELECT  ISNULL(a.[COLUMN_NAME], '') AS [COLUMN_NAME], ");
            commandText.AppendLine("        ( ");
            commandText.AppendLine("            CASE ");
            commandText.AppendLine("                WHEN ");
            commandText.AppendLine("                    ( ");
            commandText.AppendLine("                        SELECT  Count(*) ");
            commandText.AppendLine("                        FROM    INFORMATION_SCHEMA.KEY_COLUMN_USAGE ");
            commandText.AppendLine("                        WHERE   [TABLE_NAME] = '" + tableName + "' AND ");
            commandText.AppendLine("                                [COLUMN_NAME] = a.[COLUMN_NAME] ");
            commandText.AppendLine("                    ) > 0 ");
            commandText.AppendLine("                THEN 'Y' ");
            commandText.AppendLine("                ELSE '' ");
            commandText.AppendLine("            END ");
            commandText.AppendLine("        ) AS [PRIMARY_KEY], ");
            commandText.AppendLine("        ( ");
            commandText.AppendLine("            CASE ");
            commandText.AppendLine("                WHEN a.[IS_NULLABLE] = 'NO' ");
            commandText.AppendLine("                    THEN 'N' ");
            commandText.AppendLine("                ELSE 'Y' ");
            commandText.AppendLine("            END ");
            commandText.AppendLine("        ) AS [IS_NULLABLE], ");
            commandText.AppendLine("        ISNULL(a.[DATA_TYPE], '') AS [DATA_TYPE], ");
            commandText.AppendLine("        ISNULL(CONVERT(NCHAR(6),a.[CHARACTER_MAXIMUM_LENGTH]), '') AS [CHARACTER_MAXIMUM_LENGTH], ");
            commandText.AppendLine("        ( ");
            commandText.AppendLine("            CASE ");
            commandText.AppendLine("                WHEN a.[COLUMN_DEFAULT] IS NOT NULL ");
            commandText.AppendLine("                    THEN REPLACE(REPLACE(REPLACE(REPLACE(a.[COLUMN_DEFAULT], N'(', ''), ')',''), 'N''', ''), '''', '') ");
            commandText.AppendLine("                ELSE '' ");
            commandText.AppendLine("            END ");
            commandText.AppendLine("        ) AS [COLUMN_DEFAULT], ");
            commandText.AppendLine("        ( ");
            commandText.AppendLine("            SELECT TOP 1 ISNULL(sep.[value], '') ");
            commandText.AppendLine("            FROM SYS.TABLES st ");
            commandText.AppendLine("            INNER JOIN SYS.COLUMNS sc ON st.[object_id] = sc.[object_id] ");
            commandText.AppendLine("            LEFT JOIN SYS.EXTENDED_PROPERTIES sep ON st.[object_id] = sep.[major_id] ");
            commandText.AppendLine("                                                     AND sc.[column_id] = sep.[minor_id] ");
            commandText.AppendLine("                                                     AND sep.[name] = 'MS_Description' ");
            commandText.AppendLine("            WHERE	st.[name] = '" + tableName + "' AND ");
            commandText.AppendLine("                    sc.[name] = a.[COLLATION_NAME] ");
            commandText.AppendLine("        ) AS [DESCRIPTION], ");
            commandText.AppendLine("        '' AS [IDENTITY], ");
            commandText.AppendLine("        '' AS [UNIQUE], ");
            commandText.AppendLine("        '' AS [FOREIGN_KEY] ");
            commandText.AppendLine("FROM	INFORMATION_SCHEMA.COLUMNS a ");
            commandText.AppendLine("WHERE	a.[TABLE_NAME] = N'" + tableName + "' ");
            commandText.AppendLine("ORDER BY a.[ORDINAL_POSITION] ");
            DataTable returnTable = new DataTable();
            returnTable.Columns.Add(ColumnName.ColName, typeof(string));
            returnTable.Columns.Add(ColumnName.PrimaryKey, typeof(string));
            returnTable.Columns.Add(ColumnName.IsNull, typeof(string));
            returnTable.Columns.Add(ColumnName.DataType, typeof(string));
            returnTable.Columns.Add(ColumnName.MaxLength, typeof(string));
            returnTable.Columns.Add(ColumnName.Default, typeof(string));
            returnTable.Columns.Add(ColumnName.Description, typeof(string));
            returnTable.Columns.Add(ColumnName.Identity, typeof(string));
            returnTable.Columns.Add(ColumnName.Unique, typeof(string));
            returnTable.Columns.Add(ColumnName.ForeignKey, typeof(string));
            try
            {
                using (sqlConnect = new SqlConnection(CONNECTION_STRING))
                {
                    sqlConnect.Open();
                    cmd = new SqlCommand();
                    cmd.Connection = sqlConnect;
                    cmd.CommandText = commandText.ToString();
                    cmd.CommandType = CommandType.Text;
                    SqlDataReader reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        DataRow r = returnTable.NewRow();
                        r[0] = reader[0];
                        r[1] = reader[1];
                        r[2] = reader[2];
                        r[3] = reader[3];
                        r[4] = reader[4];
                        r[5] = reader[5];
                        r[6] = reader[6];
                        r[7] = reader[7];
                        r[8] = reader[8];
                        r[9] = reader[9];
                        returnTable.Rows.Add(r);
                    }
                    Console.WriteLine("Get list of column in table '" + tableName + "' success");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (sqlConnect.State != ConnectionState.Closed)
                {
                    sqlConnect.Close();
                }
            }
            return returnTable;
        }

    }
}
