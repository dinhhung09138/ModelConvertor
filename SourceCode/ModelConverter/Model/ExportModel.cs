using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ModelConverter.Model
{
    class ExportModel
    {
        public ExportModel()
        {

        }

        public void Execute()
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
}
