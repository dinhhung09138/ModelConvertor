using System;
using System.Configuration;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using ModelConverter.Model;
using ModelConverter.Excel;
using ModelConverter.ObjectToXml;


namespace ModelConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            //ExportModel model = new ExportModel();
            //model.Execute();
            //
            ExportExcel excel = new ExportExcel();
            excel.Execute();
            //
            //Execute exe = new Execute();
            //exe.Mail();
        }
    }
    
    
}
