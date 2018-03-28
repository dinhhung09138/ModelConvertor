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
                var _columnList = db.GetColumnInfor(tbTable.Rows[i][TableName.NAME].ToString().ToUpper());
                CreatSheetData(_columnList, tbTable.Rows[i][TableName.NAME].ToString().ToUpper(), i + 1);
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
            xlApp.StandardFont = "Arial";
            xlApp.StandardFontSize = 9;
            
            xlWorkBook = xlApp.Workbooks.Add(missValue);
            xlWorkBook.Windows[1].Zoom = 80;
            Console.WriteLine("Check excel app success!!");
            return true;
        }

        private void CreateSheetSummary(System.Data.DataTable data)
        {
            try
            {
                Worksheet xlSheetPaper = CreateSheet("Trang bìa");
                Worksheet _xlSheetSummary = CreateSheet("Summary");

                #region " [ Summary ] "

                _xlSheetSummary.Cells[3, 2] = "Danh sách bảng";
                _xlSheetSummary.Range[_xlSheetSummary.Cells[3, 2], _xlSheetSummary.Cells[3, 5]].Merge();

                int _row = 5;
                _xlSheetSummary.Cells[_row, 2] = "Loại";
                _xlSheetSummary.Columns[2].ColumnWidth = 20;
                _xlSheetSummary.Cells[_row, 3] = "Tên logic";
                _xlSheetSummary.Columns[3].ColumnWidth = 40;
                _xlSheetSummary.Cells[_row, 4] = "Tên vật lý";
                _xlSheetSummary.Columns[4].ColumnWidth = 30;
                _xlSheetSummary.Cells[_row, 5] = "Mục đích";
                _xlSheetSummary.Columns[5].ColumnWidth = 70;
                //
                ((Range)_xlSheetSummary.get_Range(string.Format("B{0}:E{1}", _row, _row + data.Rows.Count))).Cells.Borders.LineStyle = XlLineStyle.xlContinuous;

                for (int i = 0; i < data.Rows.Count; i++)
                {
                    _xlSheetSummary.Cells[_row + i + 1, 2] = "";
                    _xlSheetSummary.Cells[_row + i + 1, 3] = "";
                    _xlSheetSummary.Cells[_row + i + 1, 4] = data.Rows[i][TableName.NAME];
                    _xlSheetSummary.Cells[_row + i + 1, 5] = "";
                }

                #endregion

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
            Worksheet _xlSheet = null;
            try
            {
                sheetCount = xlWorkBook.Sheets.Count;
                _xlSheet = xlWorkBook.Worksheets.Add(missValue, xlWorkBook.Worksheets[sheetCount], missValue, missValue);
                _xlSheet.Name = tableName;
                
                //
                _xlSheet.Columns[2].ColumnWidth = 10;
                _xlSheet.Columns[3].ColumnWidth = 20;
                _xlSheet.Columns[4].ColumnWidth = 20;
                _xlSheet.Columns[5].ColumnWidth = 12;
                _xlSheet.Columns[6].ColumnWidth = 18;
                _xlSheet.Columns[7].ColumnWidth = 12;
                _xlSheet.Columns[8].ColumnWidth = 12;
                _xlSheet.Columns[9].ColumnWidth = 12;
                _xlSheet.Columns[10].ColumnWidth = 12;
                _xlSheet.Columns[11].ColumnWidth = 20;
                _xlSheet.Columns[12].ColumnWidth = 12;
                _xlSheet.Columns[13].ColumnWidth = 30;
                _xlSheet.Columns[14].ColumnWidth = 50;
                //
                _xlSheet.Cells[4, 2] = "Nhóm";
                _xlSheet.Cells[5, 2] = "Tên vật lý";
                _xlSheet.Cells[6, 2] = "Tên logic";
                //
                _xlSheet.Cells[4, 3] = "";
                _xlSheet.Range[_xlSheet.Cells[4, 3], _xlSheet.Cells[4, 6]].Merge();
                _xlSheet.Cells[5, 3] = tableName;
                _xlSheet.Range[_xlSheet.Cells[5, 3], _xlSheet.Cells[5, 6]].Merge();
                _xlSheet.Cells[6, 3] = "";
                _xlSheet.Range[_xlSheet.Cells[6, 3], _xlSheet.Cells[6, 6]].Merge();
                //
                ((Range)_xlSheet.get_Range("B4:F6")).Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
                //
                int _row = 8;
                //
                //((Range)_xlSheet.get_Range(string.Format("B{0}:N{1}", _row, _row))).Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
                // _xlSheet.get_Range(string.Format("B{0}:N{1}", _row, _row)).Style.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                ((Range)_xlSheet.get_Range(string.Format("B{0}:N{1}", _row, _row + data.Rows.Count))).Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
                _xlSheet.Cells[_row, 2] = "STT";
                //_xlSheet.Cells[_row, 2].Style.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                _xlSheet.Cells[_row, 3] = "Column name";
                //_xlSheet.Cells[_row, 3].Style.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                _xlSheet.Cells[_row, 4] = "Physical name";
                //_xlSheet.Cells[_row, 4].Style.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                _xlSheet.Cells[_row, 5] = "Primary key";
                //_xlSheet.Cells[_row, 5].Style.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                _xlSheet.Cells[_row, 6] = "Data type";
                //_xlSheet.Cells[_row, 6].Style.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                _xlSheet.Cells[_row, 7] = "Data length";
                //_xlSheet.Cells[_row, 7].Style.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                _xlSheet.Cells[_row, 8] = "Allow null";
                //_xlSheet.Cells[_row, 8].Style.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                _xlSheet.Cells[_row, 9] = "Index";
                //_xlSheet.Cells[_row, 9].Style.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                _xlSheet.Cells[_row, 10] = "Indentity";
                //_xlSheet.Cells[_row, 10].Style.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                _xlSheet.Cells[_row, 11] = "Init value";
                //_xlSheet.Cells[_row, 11].Style.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                _xlSheet.Cells[_row, 12] = "Unique";
                //_xlSheet.Cells[_row, 12].Style.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                _xlSheet.Cells[_row, 13] = "Foreign key";
                //_xlSheet.Cells[_row, 13].Style.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                _xlSheet.Cells[_row, 14] = "Memo";
                //_xlSheet.Cells[_row, 14].Style.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                for (int i = 0; i < data.Rows.Count; i++)
                {
                    _xlSheet.Cells[_row + i + 1, 2] = (i + 1);
                    _xlSheet.Cells[_row + i + 1, 3] = "";
                    _xlSheet.Cells[_row + i + 1, 4] = data.Rows[i][ColumnName.ColName];
                    _xlSheet.Cells[_row + i + 1, 5] = data.Rows[i][ColumnName.PrimaryKey];
                    _xlSheet.Cells[_row + i + 1, 6] = data.Rows[i][ColumnName.DataType];
                    _xlSheet.Cells[_row + i + 1, 7] = data.Rows[i][ColumnName.MaxLength];
                    _xlSheet.Cells[_row + i + 1, 8] = data.Rows[i][ColumnName.IsNull];
                    _xlSheet.Cells[_row + i + 1, 9] = ""; 
                    _xlSheet.Cells[_row + i + 1, 10] = data.Rows[i][ColumnName.Identity];
                    _xlSheet.Cells[_row + i + 1, 11] = data.Rows[i][ColumnName.Default];
                    _xlSheet.Cells[_row + i + 1, 12] = data.Rows[i][ColumnName.Unique];
                    _xlSheet.Cells[_row + i + 1, 13] = data.Rows[i][ColumnName.ForeignKey];
                    _xlSheet.Cells[_row + i + 1, 14] = "";
                }
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
                OpenExcelFile();

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

        private void OpenExcelFile()
        {
            try
            {
                FileInfo fi = new FileInfo(FILE_PATH);
                if (fi.Exists)
                {
                    System.Diagnostics.Process.Start(FILE_PATH);
                }
                else
                {
                    //file doesn't exist
                }
            }
            catch(Exception ex)
            {
                Console.Write(ex.Message);
            }
        }
    }
}
