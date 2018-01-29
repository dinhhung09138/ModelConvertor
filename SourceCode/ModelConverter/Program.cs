﻿using System;
using System.Configuration;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.IO;

namespace ModelConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            DatabaseHelper db = new DatabaseHelper();
            db.TestConnection();
            DataTable tbTable = db.GetAllTable();
            DataTable tbColumn = db.GetAllColumn();
            FileHelper file = new FileHelper();
            for (int i = 0; i < tbTable.Rows.Count; i++)
            {
                file.Execute(tbTable.Rows[i][TableName.NAME].ToString().ToUpper(), tbColumn.Select(ColumnName.TableName + " = '" + tbTable.Rows[i][TableName.NAME].ToString() + "'"));
            }
            //file.Execute("TEST", tbColumn);
            Console.WriteLine("Success");
        }
    }

    class FileHelper
    {
        string CURRENT_FOLDER = Directory.GetCurrentDirectory();
        string OUTPUT_FOLDER_NAME = "";
        string FORMAT_MONEY = "#,##0";
        string FORMAT_NUMBER = "#,##0";
        string FORMAT_DATE = "dd/MM/yyyy";
        string FORMAT_TIME = "hh:mm";
        string FORMAT_FULLTIME = "dd/MM/yyyy hh:mm";

        public FileHelper()
        {
            OUTPUT_FOLDER_NAME = string.Format("{0}\\{1}", CURRENT_FOLDER, "Output");
            if (Directory.Exists(OUTPUT_FOLDER_NAME))
            {
                Directory.Delete(OUTPUT_FOLDER_NAME, true);
            }
            Directory.CreateDirectory(OUTPUT_FOLDER_NAME);
        }

        public void Execute(string tableName, DataRow[] columnData)
        {
            StringBuilder sb = new StringBuilder();
            string fileName = string.Format("{0}\\{1}", OUTPUT_FOLDER_NAME, tableName + ".cs");
            //Write Comment
            sb = this.WriteCommand(sb, tableName);
            //
            sb.AppendLine("namespace Models");
            sb.AppendLine("{");
            //Write lib
            sb = this.WriteUsingLibrary(sb);
            //WriteClass
            sb = this.WriteClass(sb, tableName, columnData);
            sb.AppendLine("}");
            //Write file
            this.WriteFile(fileName, sb);

            Console.WriteLine("Execute file success");

        }

        /// <summary>
        /// Write command for model
        /// </summary>
        /// <param name="sb"></param>
        /// <param name="tableName"></param>
        /// <returns></returns>
        private StringBuilder WriteCommand(StringBuilder sb, string tableName)
        {
            sb.AppendLine("//--------------------");
            sb.AppendLine("// <auto-generated>");
            sb.AppendLine("//\t");
            sb.AppendLine("//\t");
            sb.AppendLine("//\tTable: " + tableName);
            sb.AppendLine("// </auto-generated>");
            sb.AppendLine("//--------------------");
            return sb;
        }

        /// <summary>
        /// Write using libiray in model
        /// </summary>
        /// <param name="sb"></param>
        /// <returns></returns>
        private StringBuilder WriteUsingLibrary(StringBuilder sb)
        {
            sb.AppendLine("\tusing System;");
            sb.AppendLine("\tusing System.Collections.Generic;");
            sb.AppendLine("\tusing System.ComponentModel.DataAnnotations;");
            sb.AppendLine("\t");
            return sb;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sb"></param>
        /// <param name="tableName">Table name</param>
        /// <param name="columnData">Data of column name</param>
        /// <returns></returns>
        private StringBuilder WriteClass(StringBuilder sb, string tableName, DataRow[] columnData)
        {
            sb.AppendLine("\tpublic class " + tableName);
            sb.AppendLine("\t{");
            sb.AppendLine("\t\t#region");
            sb.AppendLine("\t");
            for (int i = 0; i < columnData.Length; i++)
            {
                switch (columnData[i][ColumnName.DataType].ToString().ToLower())
                {
                    case "nvarchar":
                    case "nchar":
                        sb = this.WriteStringColumn(sb, columnData[i]);
                        break;
                    case "int":
                        sb = this.WriteIntColumn(sb, columnData[i]);
                        break;
                    case "money":
                    case "decimal":
                        sb = this.WriteMoneyColumn(sb, columnData[i]);
                        break;
                    case "float":
                        sb = this.WriteFloatColumn(sb, columnData[i]);
                        break;
                    case "bigint":
                        sb = this.WriteLongColumn(sb, columnData[i]);
                        break;
                    case "smallint":
                        break;
                    case "datetime":
                    case "timestamp":
                        sb = this.WriteDatetimeColumn(sb, columnData[i]);
                        break;
                    case "uniqueidentifier":
                        sb = this.WriteGuidColumn(sb, columnData[i]);
                        break;
                    case "bit":
                        sb = this.WriteBooleanColumn(sb, columnData[i]);
                        break;
                    default:
                        break;

                }
            }
            sb.AppendLine("\t");
            sb.AppendLine("\t\t#endregion");
            sb.AppendLine("\t}");
            return sb;
        }

        /// <summary>
        /// Write string column
        /// </summary>
        /// <param name="sb"></param>
        /// <param name="row">row data</param>
        /// <returns></returns>
        private StringBuilder WriteStringColumn(StringBuilder sb, DataRow row)
        {
            sb.AppendLine("\t\t");
            sb.AppendLine("\t\t//<summary>");
            sb.AppendLine("\t\t//");
            sb.AppendLine("\t\t//<summary>");
            sb.AppendLine("\t\t[Display(Name =\"\")]");
            if (row[ColumnName.IsNull].ToString() == "NO")
            {
                sb.AppendLine("\t\t[Required(ErrorMessage = \"\")]");
            }
            sb.AppendLine("\t\t[StringLength(" + row[ColumnName.MaxLength] + ",ErrorMessage = \"\")]");
            sb.AppendLine("\t\tpublic string " + Utils.UppercaseWords(row[ColumnName.ColName].ToString(), '_') + " { get; set; } = \"" + Utils.GetDefaultStringValue(row[ColumnName.Default].ToString()) + "\";");
            sb.AppendLine("\t\t");
            return sb;
        }

        /// <summary>
        /// Write guid column
        /// </summary>
        /// <param name="sb"></param>
        /// <param name="row">row data</param>
        /// <returns></returns>
        private StringBuilder WriteGuidColumn(StringBuilder sb, DataRow row)
        {
            sb.AppendLine("\t\t");
            sb.AppendLine("\t\t//<summary>");
            sb.AppendLine("\t\t//");
            sb.AppendLine("\t\t//<summary>");
            sb.AppendLine("\t\t[Display(Name =\"\")]");
            if (row[ColumnName.IsNull].ToString() == "NO")
            {
                sb.AppendLine("\t\t[Required(ErrorMessage = \"\")]");
            }
            sb.AppendLine("\t\tpublic Guid " + Utils.UppercaseWords(row[ColumnName.ColName].ToString(), '_') + " { get; set; } = Guid.NewGuid();");
            sb.AppendLine("\t\t");
            return sb;
        }

        /// <summary>
        /// Write int column
        /// </summary>
        /// <param name="sb"></param>
        /// <param name="row">row data</param>
        /// <returns></returns>
        private StringBuilder WriteIntColumn(StringBuilder sb, DataRow row)
        {
            string valueTmp = Utils.GetDefaultStringValue(row[ColumnName.Default].ToString());
            sb.AppendLine("\t\t");
            sb.AppendLine("\t\t//<summary>");
            sb.AppendLine("\t\t//");
            sb.AppendLine("\t\t//<summary>");
            sb.AppendLine("\t\t[Display(Name =\"\")]");
            if (row[ColumnName.IsNull].ToString() == "NO")
            {
                sb.AppendLine("\t\t[Required(ErrorMessage = \"\")]");
                sb.AppendLine("\t\t[DisplayFormat(DataFormatString = \"" + FORMAT_NUMBER + "\", ApplyFormatInEditMode  = true, NullDisplayText = \"0\")]");
                sb.AppendLine("\t\tpublic int " + Utils.UppercaseWords(row[ColumnName.ColName].ToString(), '_') + " { get; set; } " + (valueTmp.Length > 0 ? " = " + valueTmp : "") + ";");
            }
            else
            {
                sb.AppendLine("\t\t[DisplayFormat(DataFormatString = \"" + FORMAT_NUMBER + "\", ApplyFormatInEditMode  = true, NullDisplayText = \"0\")]");
                sb.AppendLine("\t\tpublic Nullable<int> " + Utils.UppercaseWords(row[ColumnName.ColName].ToString(), '_') + " { get; set; } " + (valueTmp.Length > 0 ? " = " + valueTmp : "") + ";");
            }
            
            sb.AppendLine("\t\t");
            return sb;
        }

        /// <summary>
        /// Write money column
        /// </summary>
        /// <param name="sb"></param>
        /// <param name="row">row data</param>
        /// <returns></returns>
        private StringBuilder WriteMoneyColumn(StringBuilder sb, DataRow row)
        {
            sb.AppendLine("\t\t");
            sb.AppendLine("\t\t//<summary>");
            sb.AppendLine("\t\t//");
            sb.AppendLine("\t\t//<summary>");
            sb.AppendLine("\t\t[Display(Name =\"\")]");
            if (row[ColumnName.IsNull].ToString() == "NO")
            {
                sb.AppendLine("\t\t[Required(ErrorMessage = \"\")]");
            }
            sb.AppendLine("\t\t[Range(0.1, 99999999, ErrorMessage = \"\")]");
            sb.AppendLine("\t\t[DisplayFormat(DataFormatString = \"" + FORMAT_MONEY + "\", ApplyFormatInEditMode  = true, NullDisplayText = \"0\")]");
            sb.AppendLine("\t\tpublic decimal " + Utils.UppercaseWords(row[ColumnName.ColName].ToString(), '_') + " { get; set; } = " + Utils.GetDefaultNumberValue(row[ColumnName.Default].ToString()) + ";");

            sb.AppendLine("\t\t");
            return sb;
        }

        /// <summary>
        /// Write long column
        /// </summary>
        /// <param name="sb"></param>
        /// <param name="row">row data</param>
        /// <returns></returns>
        private StringBuilder WriteLongColumn(StringBuilder sb, DataRow row)
        {
            string valueTmp = Utils.GetDefaultStringValue(row[ColumnName.Default].ToString());
            sb.AppendLine("\t\t");
            sb.AppendLine("\t\t//<summary>");
            sb.AppendLine("\t\t//");
            sb.AppendLine("\t\t//<summary>");
            sb.AppendLine("\t\t[Display(Name =\"\")]");
            if (row[ColumnName.IsNull].ToString() == "NO")
            {
                sb.AppendLine("\t\t[Required(ErrorMessage = \"\")]");
                sb.AppendLine("\t\t[DisplayFormat(DataFormatString = \"" + FORMAT_NUMBER + "\", ApplyFormatInEditMode  = true, NullDisplayText = \"0\")]");
                sb.AppendLine("\t\tpublic long " + Utils.UppercaseWords(row[ColumnName.ColName].ToString(), '_') + " { get; set; } " + (valueTmp.Length > 0 ? " = " + valueTmp : "") + ";");
            }
            else
            {
                sb.AppendLine("\t\t[DisplayFormat(DataFormatString = \"" + FORMAT_NUMBER + "\", ApplyFormatInEditMode  = true, NullDisplayText = \"0\")]");
                sb.AppendLine("\t\tpublic Nullable<long> " + Utils.UppercaseWords(row[ColumnName.ColName].ToString(), '_') + " { get; set; } " + (valueTmp.Length > 0 ? " = " + valueTmp : "") + ";");
            }

            sb.AppendLine("\t\t");
            return sb;
        }

        /// <summary>
        /// Write float column
        /// </summary>
        /// <param name="sb"></param>
        /// <param name="row">row data</param>
        /// <returns></returns>
        private StringBuilder WriteFloatColumn(StringBuilder sb, DataRow row)
        {
            string valueTmp = Utils.GetDefaultStringValue(row[ColumnName.Default].ToString());
            sb.AppendLine("\t\t");
            sb.AppendLine("\t\t//<summary>");
            sb.AppendLine("\t\t//");
            sb.AppendLine("\t\t//<summary>");
            sb.AppendLine("\t\t[Display(Name =\"\")]");
            if (row[ColumnName.IsNull].ToString() == "NO")
            {
                sb.AppendLine("\t\t[Required(ErrorMessage = \"\")]");
                sb.AppendLine("\t\t[DisplayFormat(DataFormatString = \"" + FORMAT_NUMBER + "\", ApplyFormatInEditMode  = true, NullDisplayText = \"0\")]");
                sb.AppendLine("\t\tpublic float " + Utils.UppercaseWords(row[ColumnName.ColName].ToString(), '_') + " { get; set; } " + (valueTmp.Length > 0 ? " = " + valueTmp : "") + ";");
            }
            else
            {
                sb.AppendLine("\t\t[DisplayFormat(DataFormatString = \"" + FORMAT_NUMBER + "\", ApplyFormatInEditMode  = true, NullDisplayText = \"0\")]");
                sb.AppendLine("\t\tpublic Nullable<float> " + Utils.UppercaseWords(row[ColumnName.ColName].ToString(), '_') + " { get; set; } " + (valueTmp.Length > 0 ? " = " + valueTmp : "") + ";");
            }

            sb.AppendLine("\t\t");
            return sb;
        }

        /// <summary>
        /// Write money column
        /// </summary>
        /// <param name="sb"></param>
        /// <param name="row">row data</param>
        /// <returns></returns>
        private StringBuilder WriteDatetimeColumn(StringBuilder sb, DataRow row)
        {
            sb.AppendLine("\t\t");
            sb.AppendLine("\t\t//<summary>");
            sb.AppendLine("\t\t//");
            sb.AppendLine("\t\t//<summary>");
            sb.AppendLine("\t\t[Display(Name =\"\")]");
            if (row[ColumnName.IsNull].ToString() == "NO")
            {
                sb.AppendLine("\t\t[Required(ErrorMessage = \"\")]");
                sb.AppendLine("\t\t[DisplayFormat(DataFormatString = \"" + FORMAT_DATE + "\", ApplyFormatInEditMode  = true)]");
                sb.AppendLine("\t\tpublic DateTime " + Utils.UppercaseWords(row[ColumnName.ColName].ToString(), '_') + " { get; set; } = DateTime.Now;");
            }
            else
            {
                sb.AppendLine("\t\t[DisplayFormat(DataFormatString = \"" + FORMAT_DATE + "\", ApplyFormatInEditMode  = true)]");
                sb.AppendLine("\t\tpublic Nullable<DateTime> " + Utils.UppercaseWords(row[ColumnName.ColName].ToString(), '_') + " { get; set; } = DateTime.Now;");
            }
            sb.AppendLine("\t\t");
            return sb;
        }

        /// <summary>
        /// Write boolean column
        /// </summary>
        /// <param name="sb"></param>
        /// <param name="row">row data</param>
        /// <returns></returns>
        private StringBuilder WriteBooleanColumn(StringBuilder sb, DataRow row)
        {
            string tmpValueDefault = Utils.GetDefaultStringValue(row[ColumnName.Default].ToString());
            sb.AppendLine("\t\t");
            sb.AppendLine("\t\t//<summary>");
            sb.AppendLine("\t\t//");
            sb.AppendLine("\t\t//<summary>");
            sb.AppendLine("\t\t[Display(Name =\"\")]");
            if (row[ColumnName.IsNull].ToString() == "NO")
            {
                sb.AppendLine("\t\t[Required(ErrorMessage = \"\")]");
            }
            sb.AppendLine("\t\tpublic Boolean " + Utils.UppercaseWords(row[ColumnName.ColName].ToString(), '_') + " { get; set; } = \"" + (tmpValueDefault.Length > 0 && tmpValueDefault == "1" ? "true" : "false") + "\";");
            sb.AppendLine("\t\t");
            return sb;
        }



        /// <summary>
        /// Write file to disk
        /// </summary>
        /// <param name="fileName">filepath</param>
        /// <param name="sb"></param>
        public void WriteFile(string fileName, StringBuilder sb)
        {
            try
            {
                StreamWriter sw = new StreamWriter(fileName, true);
                sw.Write(sb.ToString());
                sw.Close();
                Console.WriteLine("Write file succes!!!");
            }
            catch(Exception ex)
            {
                Console.WriteLine("Write file is error");
                Console.WriteLine(ex.Message);
                Console.WriteLine("---------");
            }
        }
    }

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
                using(sqlConnect = new SqlConnection(CONNECTION_STRING))
                {
                    sqlConnect.Open();
                    Console.WriteLine("connect success");
                }
            }
            catch(Exception ex)
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
                if(sqlConnect.State != ConnectionState.Closed)
                {
                    sqlConnect.Close();
                }
            }
            return returnTable;
        }

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

    }

    class Utils
    {
        public static string UppercaseWords(string text, char sperateCode)
        {
            string[] tmpString = text.Split(sperateCode);
            string returnString = "";
            foreach (var s in tmpString)
            {
                if(s.Length > 0)
                returnString += Char.ToUpper(s[0]) + s.Substring(1).ToLower();
            }
            return returnString;
        }

        public static string GetDefaultStringValue(string text)
        {
            if (text.Length == 0)
                return "";
            string returnString = text.Replace("(N'", "").Replace("')","").Replace("((","").Replace("))","");
            return returnString;
        }

        public static string GetDefaultNumberValue(string text)
        {
            if (text.Length == 0)
                return "0";
            string returnString = text.Replace("(N'", "").Replace("')", "").Replace("((", "").Replace("))", "");
            return returnString;
        }
    }

    public class TableName
    {
        public const string NAME = "NAME";
    }

    public class ColumnName
    {
        public const string TableName = "TABLE_NAME";
        public const string ColName = "COLUMN_NAME";
        public const string Position = "POSITION";
        public const string Default = "DEFAULT";
        public const string IsNull = "IS_NULL";
        public const string DataType = "DATA_TYPE";
        public const string MaxLength = "MAXLENGTH";
    }
}
