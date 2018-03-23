using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.IO;

namespace ModelConverter.Excel
{
    public class ExportExcel
    {
        Application xlApp;
        Workbook xlWorkBook;
        object missValue = Type.Missing;
        int sheetCount = 0;

        private readonly string FILE_PATH = Path.Combine(Environment.CurrentDirectory, "Data.xls");

        public ExportExcel()
        {
            if (File.Exists(FILE_PATH))
            {
                File.Delete(FILE_PATH);
            }
            xlApp = new Application();
        }

        public void Execute()
        {
            if (!CheckExcelApp())
            {
                ReleaseMemory();
                return;
            }
            DatabaseHelper db = new DatabaseHelper();
            db.TestConnection();
            System.Data.DataTable tbTable = db.GetAllTable();
            System.Data.DataTable tbColumn = db.GetAllColumn();
            //
            CreateSheetSummary(tbTable);
            //
            for (int i = 0; i < tbTable.Rows.Count; i++)
            {
                var columnLits = db.GetColumnInfor(tbTable.Rows[i][TableName.NAME].ToString().ToUpper());
                CreatSheetData(columnLits, tbTable.Rows[i][TableName.NAME].ToString().ToUpper(), i + 1);
            }
            SaveExcelFile();
            Console.WriteLine("Finish");
        }

        private bool CheckExcelApp()
        {
            if (xlApp == null)
            {
                Console.WriteLine("Excel is not properly installed!!");
                return false;
            }
            xlWorkBook = xlApp.Workbooks.Add(missValue);
            
            Console.WriteLine("Check excel app success!!");
            return true;
        }

        private void CreateSheetSummary(System.Data.DataTable data)
        {
            try
            {
                Worksheet xlSheetPaper = CreateSheet("Tổng quan");
                Worksheet xlSheetSummary = CreateSheet("Summary");
                // xlWorkSheet.Cells[1, 1] = "Something";
            }
            catch (Exception ex)
            {
                // Marshal.ReleaseComObject(xlWorkSheet);
                Console.WriteLine("Create sheet CreateSheetSummary");
                Console.WriteLine(ex.Message);

            }
            finally
            {
                //ReleaseMemory();
            }
        }

        private void CreatSheetData(System.Data.DataTable data, string tableName, int sheetCount)
        {
            Worksheet xlWorkSheet = null;
            try
            {
                sheetCount = xlWorkBook.Sheets.Count;
                xlWorkSheet = xlWorkBook.Worksheets.Add(missValue, xlWorkBook.Worksheets[sheetCount], missValue, missValue);
                //xlWorkSheet = lSheet.Add(lSheet[sheetCount], missValue, missValue, missValue);
                xlWorkSheet.Name = tableName;
               // xlWorkSheet.Cells[1, 1] = "Something";
            }
            catch (Exception ex)
            {
               // Marshal.ReleaseComObject(xlWorkSheet);
                Console.WriteLine("Create sheet '" + tableName + "' Error");
                Console.WriteLine(ex.Message);

            }
            finally
            {
                //ReleaseMemory();
            }
        }

        private void SaveExcelFile()
        {
            try
            {

                xlWorkBook.SaveAs(FILE_PATH, XlFileFormat.xlWorkbookNormal, missValue, missValue, missValue, missValue,
                    XlSaveAsAccessMode.xlExclusive, missValue, missValue, missValue, missValue, missValue);
                xlWorkBook.Close(true, missValue, missValue);
                xlApp.Quit();
                

                Console.WriteLine("Write file success");
            }
            catch(Exception ex)
            {
                Console.WriteLine("Error when save file");
                Console.WriteLine(ex.Message);
            }
            finally
            {
                ReleaseMemory();
            }

        }

        private void ReleaseMemory()
        {
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }

        private Worksheet CreateSheet(string sheetName)
        {
            sheetCount = xlWorkBook.Sheets.Count;
            Worksheet xlSheet = xlWorkBook.Worksheets.Add(missValue, xlWorkBook.Worksheets[sheetCount], missValue, missValue);
            xlSheet.Name = sheetName;
            xlSheet.Columns.AutoFit();

            return xlSheet;
        }
    }
}
