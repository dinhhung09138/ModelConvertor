using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.IO;
using OfficeOpenXml;

namespace ModelConverter.Excel
{
    public class ExportExcel
    {

        private readonly string FILE_PATH = Path.Combine(Environment.CurrentDirectory, "Data.xlsx");

        public ExportExcel()
        {
            if (File.Exists(FILE_PATH))
            {
                File.Delete(FILE_PATH);
            }
        }

        public void Execute()
        {
            //DatabaseHelper db = new DatabaseHelper();
            //db.TestConnection();
            //DataTable tbTable = db.GetAllTable();
            //DataTable tbColumn = db.GetAllColumn();
            ////
            //ExcelPackage ExcelPkg = new ExcelPackage();
            ////
            //CreateSheetSummary(tbTable);
            ////
            //for (int i = 0; i < tbTable.Rows.Count; i++)
            //{
            //    var _columnList = db.GetColumnInfor(tbTable.Rows[i][TableName.NAME].ToString().ToUpper());
            //    CreateSheetData(ExcelPkg, _columnList, tbTable.Rows[i][TableName.NAME].ToString().ToUpper());
            //}
            //ExcelPkg.SaveAs(new FileInfo(FILE_PATH));
            ////

            CreateExcelFile(FILE_PATH);

            Console.WriteLine("Finish");
            OpenExcelFile();
            //CreateExcelFile(Path.Combine(Environment.CurrentDirectory, "Data1.xlsx"));
        }

        private void CreateSheetSummary(DataTable data)
        {
            return;
            try
            {
                //Worksheet xlSheetPaper = CreateSheet("Trang bìa");
                //Worksheet _xlSheetSummary = CreateSheet("Summary");

                //#region " [ Summary ] "

                //_xlSheetSummary.Cells[3, 2] = "Danh sách bảng";
                //_xlSheetSummary.Range[_xlSheetSummary.Cells[3, 2], _xlSheetSummary.Cells[3, 5]].Merge();

                //int _row = 5;
                //_xlSheetSummary.Cells[_row, 2] = "Loại";
                //_xlSheetSummary.Columns[2].ColumnWidth = 20;
                //_xlSheetSummary.Cells[_row, 3] = "Tên logic";
                //_xlSheetSummary.Columns[3].ColumnWidth = 40;
                //_xlSheetSummary.Cells[_row, 4] = "Tên vật lý";
                //_xlSheetSummary.Columns[4].ColumnWidth = 30;
                //_xlSheetSummary.Cells[_row, 5] = "Mục đích";
                //_xlSheetSummary.Columns[5].ColumnWidth = 70;
                ////
                //((Range)_xlSheetSummary.get_Range(string.Format("B{0}:E{1}", _row, _row + data.Rows.Count))).Cells.Borders.LineStyle = XlLineStyle.xlContinuous;

                //for (int i = 0; i < data.Rows.Count; i++)
                //{
                //    _xlSheetSummary.Cells[_row + i + 1, 2] = "";
                //    _xlSheetSummary.Cells[_row + i + 1, 3] = "";
                //    _xlSheetSummary.Cells[_row + i + 1, 4] = data.Rows[i][TableName.NAME];
                //    _xlSheetSummary.Cells[_row + i + 1, 5] = "";
                //}

                //#endregion

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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="data"></param>
        /// <param name="tableName"></param>
        /// <returns></returns>
        static ExcelWorksheet CreateSheetData(ExcelPackage excelPackage, DataTable data, string tableName)
        {
            ExcelWorksheet _xlSheet = excelPackage.Workbook.Worksheets.Add(tableName);
            //
            _xlSheet.Cells[3, 2].Value = "Nhóm";
            _xlSheet.Cells[4, 2].Value = "Tên vật lý";
            _xlSheet.Cells[5, 2].Value = "Tên logic";
            using (ExcelRange range = _xlSheet.Cells[3, 2, 5, 2])
            {
                range.Style.Font.Bold = true;
                range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
            }
            using (ExcelRange range = _xlSheet.Cells[3, 3, 3, 5])
            {
                range.Merge = true;
            }
            using (ExcelRange range = _xlSheet.Cells[4, 3, 4, 5])
            {
                range.Merge = true;
            }
            using (ExcelRange range = _xlSheet.Cells[5, 3, 5, 5])
            {
                range.Merge = true;
            }
            using (ExcelRange range = _xlSheet.Cells[3, 2, 5, 5])
            {
                range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }
            //
            int _row = 7;
            //
            _xlSheet.Cells[_row, 2].Value = "STT";
            _xlSheet.Cells[_row, 3].Value = "Column name";
            _xlSheet.Cells[_row, 4].Value = "Physical name";
            _xlSheet.Cells[_row, 5].Value = "Primary key";
            _xlSheet.Cells[_row, 6].Value = "Data type";
            _xlSheet.Cells[_row, 7].Value = "Data length";
            _xlSheet.Cells[_row, 8].Value = "Allow null";
            _xlSheet.Cells[_row, 9].Value = "Index";
            _xlSheet.Cells[_row, 10].Value = "Indentity";
            _xlSheet.Cells[_row, 11].Value = "Init value";
            _xlSheet.Cells[_row, 12].Value = "Unique";
            _xlSheet.Cells[_row, 13].Value = "Foreign key";
            _xlSheet.Cells[_row, 14].Value = "Memo";
            using (ExcelRange range = _xlSheet.Cells[_row, 2, _row, 14])
            {
                range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange range = _xlSheet.Cells[_row, 2, _row + data.Rows.Count + 1, 14])
            {
                range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }
            //
            for (int i = 0; i < data.Rows.Count; i++)
            {
                _xlSheet.Cells[_row + i + 1, 2].Value = (i + 1);
                _xlSheet.Cells[_row + i + 1, 3].Value = "";
                _xlSheet.Cells[_row + i + 1, 4].Value = data.Rows[i][ColumnName.ColName];
                _xlSheet.Cells[_row + i + 1, 5].Value = data.Rows[i][ColumnName.PrimaryKey];
                _xlSheet.Cells[_row + i + 1, 6].Value = data.Rows[i][ColumnName.DataType];
                _xlSheet.Cells[_row + i + 1, 7].Value = data.Rows[i][ColumnName.MaxLength];
                _xlSheet.Cells[_row + i + 1, 8].Value = data.Rows[i][ColumnName.IsNull];
                _xlSheet.Cells[_row + i + 1, 9].Value = "";
                _xlSheet.Cells[_row + i + 1, 10].Value = data.Rows[i][ColumnName.Identity];
                _xlSheet.Cells[_row + i + 1, 11].Value = data.Rows[i][ColumnName.Default];
                _xlSheet.Cells[_row + i + 1, 12].Value = data.Rows[i][ColumnName.Unique];
                _xlSheet.Cells[_row + i + 1, 13].Value = data.Rows[i][ColumnName.ForeignKey];
                _xlSheet.Cells[_row + i + 1, 14].Value = "";
            }
            //
            _xlSheet.Cells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            _xlSheet.Cells.AutoFitColumns();
            _xlSheet.View.ShowGridLines = false;
            _xlSheet.Protection.IsProtected = false;
            _xlSheet.Protection.AllowSelectLockedCells = false;
            //

            return _xlSheet;
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
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
        }

        public void CreateExcelFile(string filePath)
        {
            DatabaseHelper db = new DatabaseHelper();
            db.TestConnection();
            DataTable tbTable = db.GetAllTable();
            DataTable tbColumn = db.GetAllColumn();
            //
            ExcelPackage _excelPkg = new ExcelPackage();
            for (int i = 0; i < tbTable.Rows.Count; i++)
            {
                var _columnList = db.GetColumnInfor(tbTable.Rows[i][TableName.NAME].ToString().ToUpper());
                ExcelContentModel _excelModel = CreateExcelModel(tbTable.Rows[i][TableName.NAME].ToString().ToUpper(), _columnList);
                CreateSheet(_excelPkg, _excelModel);
            }
            //
            _excelPkg.SaveAs(new FileInfo(filePath));
            //
            Console.WriteLine("Finish");
            OpenExcelFile();
        }

        private ExcelWorksheet CreateSheet(ExcelPackage excelPackage, ExcelContentModel excelModel)
        {
            ExcelWorksheet _sheet = excelPackage.Workbook.Worksheets.Add(excelModel.SheetName);
            //Data content
            foreach (var item in excelModel.Data)
            {
                _sheet.Cells[item.Row, item.Col].Value = item.Value;
                if (item.Bold)
                {
                    _sheet.Cells[item.Row, item.Col].Style.Font.Bold = item.Bold;
                }
            }
            //Horizontal formating
            foreach (var item in excelModel.Horizontals)
            {
                using (ExcelRange range = _sheet.Cells[item.FromRow, item.FromColumn, item.ToRow, item.ToColumn])
                {
                    range.Style.HorizontalAlignment = item.Horizontal;
                }
            }
            //Vertical formating
            foreach (var item in excelModel.Verticals)
            {
                using (ExcelRange range = _sheet.Cells[item.FromRow, item.FromColumn, item.ToRow, item.ToColumn])
                {
                    range.Style.VerticalAlignment = item.Vertical;
                }
            }
            //Merge
            foreach (var item in excelModel.Merges)
            {
                using (ExcelRange range = _sheet.Cells[item.FromRow, item.FromColumn, item.ToRow, item.ToColumn])
                {
                    range.Merge = true;
                }
            }
            //Border
            foreach (var item in excelModel.Borders)
            {
                using (ExcelRange range = _sheet.Cells[item.FromRow, item.FromColumn, item.ToRow, item.ToColumn])
                {
                    range.Style.Border.Top.Style = item.Border;
                    range.Style.Border.Right.Style = item.Border;
                    range.Style.Border.Bottom.Style = item.Border;
                    range.Style.Border.Left.Style = item.Border;
                }
            }
            //Set format text for all sheet
//            _sheet.Cells.Style.Numberformat.Format = "@";
            //Format specific cell
            foreach (var item in excelModel.Formatings)
            {
                _sheet.Cells[item.FromRow, item.FromColumn, item.ToRow, item.ToColumn].Style.Numberformat.Format = item.Format;
            }
            //
            _sheet.Cells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            _sheet.Cells.AutoFitColumns();
            _sheet.View.ShowGridLines = excelModel.ShowGridLines;
            _sheet.Protection.IsProtected = excelModel.IsProtected;
            _sheet.Protection.AllowSelectLockedCells = excelModel.AllowSelectLockedCells;
            return _sheet;
        }

        private ExcelContentModel CreateExcelModel(string sheetName, DataTable data)
        {
            ExcelContentModel _return = new ExcelContentModel();
            _return.SheetName = sheetName;
            _return.ShowGridLines = false;
            _return.IsProtected = false;
            _return.AllowSelectLockedCells = false;

            int _row = 7;

            #region " [ Data ] "

            _return.Data.Add(new ExcelDataDetailModel() {
                Row = 3,
                Col = 2,
                Value = "Nhóm",
                Bold = false
            });
            _return.Data.Add(new ExcelDataDetailModel()
            {
                Row = 4,
                Col = 2,
                Value = "Tên vật lý",
                Bold = false
            });
            _return.Data.Add(new ExcelDataDetailModel()
            {
                Row = 5,
                Col = 2,
                Value = "Tên logic",
                Bold = false
            });
            _return.Data.Add(new ExcelDataDetailModel()
            {
                Row = _row,
                Col = 2,
                Value = "STT",
                Bold = false
            });
            _return.Data.Add(new ExcelDataDetailModel()
            {
                Row = _row,
                Col = 3,
                Value = "Column name",
                Bold = false
            });
            _return.Data.Add(new ExcelDataDetailModel()
            {
                Row = _row,
                Col = 4,
                Value = "Physical name",
                Bold = false
            });
            _return.Data.Add(new ExcelDataDetailModel()
            {
                Row = _row,
                Col = 5,
                Value = "Primary key",
                Bold = false
            });
            _return.Data.Add(new ExcelDataDetailModel()
            {
                Row = _row,
                Col = 6,
                Value = "Data type",
                Bold = false
            });
            _return.Data.Add(new ExcelDataDetailModel()
            {
                Row = _row,
                Col = 7,
                Value = "Data length",
                Bold = false
            });
            _return.Data.Add(new ExcelDataDetailModel()
            {
                Row = _row,
                Col = 8,
                Value = "Allow null",
                Bold = false
            });
            _return.Data.Add(new ExcelDataDetailModel()
            {
                Row = _row,
                Col = 9,
                Value = "Index",
                Bold = false
            });
            _return.Data.Add(new ExcelDataDetailModel()
            {
                Row = _row,
                Col = 10,
                Value = "Indentity",
                Bold = false
            });
            _return.Data.Add(new ExcelDataDetailModel()
            {
                Row = _row,
                Col = 11,
                Value = "Init value",
                Bold = false
            });
            _return.Data.Add(new ExcelDataDetailModel()
            {
                Row = _row,
                Col = 12,
                Value = "Unique",
                Bold = false
            });
            _return.Data.Add(new ExcelDataDetailModel()
            {
                Row = _row,
                Col = 13,
                Value = "Foreign key",
                Bold = false
            });
            _return.Data.Add(new ExcelDataDetailModel()
            {
                Row = _row,
                Col = 14,
                Value = "Memo",
                Bold = false
            });

            for (int i = 0; i < data.Rows.Count; i++)
            {
                _return.Data.Add(new ExcelDataDetailModel() {
                    Row  = _row + i + 1,
                    Col = 2,
                    Value = (i + 1 + 1000)
                });
                _return.Data.Add(new ExcelDataDetailModel()
                {
                    Row = _row + i + 1,
                    Col = 4,
                    Value = data.Rows[i][ColumnName.ColName].ToString()
                });
                _return.Data.Add(new ExcelDataDetailModel()
                {
                    Row = _row + i + 1,
                    Col = 5,
                    Value = data.Rows[i][ColumnName.PrimaryKey].ToString()
                });
                _return.Data.Add(new ExcelDataDetailModel()
                {
                    Row = _row + i + 1,
                    Col = 6,
                    Value = data.Rows[i][ColumnName.DataType].ToString()
                });
                _return.Data.Add(new ExcelDataDetailModel()
                {
                    Row = _row + i + 1,
                    Col = 7,
                    Value = data.Rows[i][ColumnName.MaxLength]
                });
                _return.Data.Add(new ExcelDataDetailModel()
                {
                    Row = _row + i + 1,
                    Col = 8,
                    Value = data.Rows[i][ColumnName.IsNull].ToString()
                });
                _return.Data.Add(new ExcelDataDetailModel()
                {
                    Row = _row + i + 1,
                    Col = 10,
                    Value = data.Rows[i][ColumnName.Identity].ToString()
                });
                _return.Data.Add(new ExcelDataDetailModel()
                {
                    Row = _row + i + 1,
                    Col = 11,
                    Value = data.Rows[i][ColumnName.Default].ToString()
                });
                _return.Data.Add(new ExcelDataDetailModel()
                {
                    Row = _row + i + 1,
                    Col = 12,
                    Value = data.Rows[i][ColumnName.Unique].ToString()
                });
                _return.Data.Add(new ExcelDataDetailModel()
                {
                    Row = _row + i + 1,
                    Col = 13,
                    Value = data.Rows[i][ColumnName.ForeignKey].ToString()
                });
            }

            #endregion

            #region " [ Merge ] "

            _return.Merges.Add(new ExcelMergeDetailModel() {
                FromRow = 3,
                FromColumn = 3, 
                ToRow = 3,
                ToColumn = 5
            });
            _return.Merges.Add(new ExcelMergeDetailModel()
            {
                FromRow = 4,
                FromColumn = 3,
                ToRow = 4,
                ToColumn = 5
            });
            _return.Merges.Add(new ExcelMergeDetailModel()
            {
                FromRow = 5,
                FromColumn = 3,
                ToRow = 5,
                ToColumn = 5
            });

            #endregion

            #region " [ Horizontal ] "

            _return.Horizontals.Add(new ExcelHorizontalDetailModel() {
                FromRow = 3,
                FromColumn = 2,
                ToRow = 5,
                ToColumn = 5,
                Horizontal = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left
            });
            _return.Horizontals.Add(new ExcelHorizontalDetailModel()
            {
                FromRow = _row,
                FromColumn = 2,
                ToRow = _row,
                ToColumn = 14,
                Horizontal = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center
            });
            _return.Horizontals.Add(new ExcelHorizontalDetailModel()
            {
                FromRow = _row + 1,
                FromColumn = 2,
                ToRow = _row + data.Rows.Count + 1,
                ToColumn = 2,
                Horizontal = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center
            });
            _return.Horizontals.Add(new ExcelHorizontalDetailModel()
            {
                FromRow = _row + 1,
                FromColumn = 5,
                ToRow = _row + data.Rows.Count + 1,
                ToColumn = 5,
                Horizontal = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center
            });
            _return.Horizontals.Add(new ExcelHorizontalDetailModel()
            {
                FromRow = _row + 1,
                FromColumn = 7,
                ToRow = _row + data.Rows.Count + 1,
                ToColumn = 7,
                Horizontal = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center
            });
            _return.Horizontals.Add(new ExcelHorizontalDetailModel()
            {
                FromRow = _row + 1,
                FromColumn = 8,
                ToRow = _row + data.Rows.Count + 1,
                ToColumn = 8,
                Horizontal = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center
            });
            _return.Horizontals.Add(new ExcelHorizontalDetailModel()
            {
                FromRow = _row + 1,
                FromColumn = 9,
                ToRow = _row + data.Rows.Count + 1,
                ToColumn = 9,
                Horizontal = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center
            });
            _return.Horizontals.Add(new ExcelHorizontalDetailModel()
            {
                FromRow = _row + 1,
                FromColumn = 10,
                ToRow = _row + data.Rows.Count + 1,
                ToColumn = 10,
                Horizontal = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center
            });
            _return.Horizontals.Add(new ExcelHorizontalDetailModel()
            {
                FromRow = _row + 1,
                FromColumn = 11,
                ToRow = _row + data.Rows.Count + 1,
                ToColumn = 11,
                Horizontal = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center
            });
            _return.Horizontals.Add(new ExcelHorizontalDetailModel()
            {
                FromRow = _row + 1,
                FromColumn = 12,
                ToRow = _row + data.Rows.Count + 1,
                ToColumn = 12,
                Horizontal = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center
            });

            #endregion

            #region " [ Vertical ] "

            #endregion

            #region " [ Border ] "

            _return.Borders.Add(new ExcelBorderDetailModel() {
                FromRow = 3,
                FromColumn = 2,
                ToRow = 5, 
                ToColumn = 5,
                Border = OfficeOpenXml.Style.ExcelBorderStyle.Thin
            });

            _return.Borders.Add(new ExcelBorderDetailModel()
            {
                FromRow = _row,
                FromColumn = 2,
                ToRow = _row + data.Rows.Count + 1,
                ToColumn = 14,
                Border = OfficeOpenXml.Style.ExcelBorderStyle.Thin
            });

            #endregion

            #region " [ Formatting ] "

            _return.Formatings.Add(new ExcelFromatingDetailModel() {
                FromRow = _row + 1,
                FromColumn = 2,
                ToRow = _row + data.Rows.Count + 1,
                ToColumn = 2,
                Format = "#,##0"
            });

            #endregion

            return _return;
        }

    }

    /// <summary>
    /// Excel content model
    /// </summary>
    public class ExcelContentModel
    {
        /// <summary>
        /// Sheet name
        /// </summary>
        public string SheetName { get; set; } = "";

        /// <summary>
        /// Show gridline. Default is false
        /// </summary>
        public bool ShowGridLines { get; set; } = false;

        /// <summary>
        /// Is protection sheet
        /// Default is false
        /// </summary>
        public bool IsProtected { get; set; } = false;

        /// <summary>
        /// Allow select locked cells
        /// Default false
        /// </summary>
        public bool AllowSelectLockedCells { get; set; } = false;

        /// <summary>
        /// List of data content in sheet
        /// </summary>
        public List<ExcelDataDetailModel> Data { get; set; } = new List<ExcelDataDetailModel>();

        /// <summary>
        /// Horizontal formating
        /// </summary>
        public List<ExcelHorizontalDetailModel> Horizontals { get; set; } = new List<ExcelHorizontalDetailModel>();

        /// <summary>
        /// Vertical formating
        /// </summary>
        public List<ExcelVerticalDetailModel> Verticals { get; set; } = new List<ExcelVerticalDetailModel>();

        /// <summary>
        /// Merge cells list
        /// </summary>
        public List<ExcelMergeDetailModel> Merges { get; set; } = new List<ExcelMergeDetailModel>();

        /// <summary>
        /// Border cells list
        /// </summary>
        public List<ExcelBorderDetailModel> Borders { get; set; } = new List<ExcelBorderDetailModel>();

        /// <summary>
        /// Format data content
        /// </summary>
        public List<ExcelFromatingDetailModel> Formatings { get; set; } = new List<ExcelFromatingDetailModel>();

    }

    /// <summary>
    /// Value data model
    /// </summary>
    public class ExcelDataDetailModel
    {
        /// <summary>
        /// Row of cell
        /// </summary>
        public int Row { get; set; } = 0;

        /// <summary>
        /// Column of cell
        /// </summary>
        public int Col { get; set; } = 0;

        /// <summary>
        /// Value in a cell
        /// </summary>
        public object Value { get; set; } = "";

        /// <summary>
        /// Bold
        /// </summary>
        public bool Bold { get; set; } = false;
    }

    /// <summary>
    /// Range horizontal formating
    /// </summary>
    public class ExcelHorizontalDetailModel
    {
        /// <summary>
        /// Column of begin cell
        /// </summary>
        public int FromColumn { get; set; } = 0;

        /// <summary>
        /// Row of begin cell
        /// </summary>
        public int FromRow { get; set; } = 0;

        /// <summary>
        /// Column of end cell
        /// </summary>
        public int ToColumn { get; set; } = 0;

        /// <summary>
        /// Row of end cell
        /// </summary>
        public int ToRow { get; set; } = 0;

        /// <summary>
        /// Horizontal alignment. Default is left
        /// </summary>
        public OfficeOpenXml.Style.ExcelHorizontalAlignment Horizontal { get; set; } = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
    }

    /// <summary>
    /// Range vertical formating
    /// </summary>
    public class ExcelVerticalDetailModel
    {
        /// <summary>
        /// Column of begin cell
        /// </summary>
        public int FromColumn { get; set; } = 0;

        /// <summary>
        /// Row of begin cell
        /// </summary>
        public int FromRow { get; set; } = 0;

        /// <summary>
        /// Column of end cell
        /// </summary>
        public int ToColumn { get; set; } = 0;

        /// <summary>
        /// Row of end cell
        /// </summary>
        public int ToRow { get; set; } = 0;

        /// <summary>
        /// Vertical alignment. Default center
        /// </summary>
        public OfficeOpenXml.Style.ExcelVerticalAlignment Vertical { get; set; } = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
    }

    /// <summary>
    /// merge formating
    /// </summary>
    public class ExcelMergeDetailModel
    {
        /// <summary>
        /// Column of begin cell
        /// </summary>
        public int FromColumn { get; set; } = 0;

        /// <summary>
        /// Row of begin cell
        /// </summary>
        public int FromRow { get; set; } = 0;

        /// <summary>
        /// Column of end cell
        /// </summary>
        public int ToColumn { get; set; } = 0;

        /// <summary>
        /// Row of end cell
        /// </summary>
        public int ToRow { get; set; } = 0;
        
    }

    /// <summary>
    /// Range border formating
    /// </summary>
    public class ExcelBorderDetailModel
    {
        /// <summary>
        /// Column of begin cell
        /// </summary>
        public int FromColumn { get; set; } = 0;

        /// <summary>
        /// Row of begin cell
        /// </summary>
        public int FromRow { get; set; } = 0;

        /// <summary>
        /// Column of end cell
        /// </summary>
        public int ToColumn { get; set; } = 0;

        /// <summary>
        /// Row of end cell
        /// </summary>
        public int ToRow { get; set; } = 0;

        /// <summary>
        /// Border value. Default thin
        /// </summary>
        public OfficeOpenXml.Style.ExcelBorderStyle Border { get; set; } = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
    }

    /// <summary>
    /// Data formating
    /// </summary>
    public class ExcelFromatingDetailModel
    {
        /// <summary>
        /// Column of begin cell
        /// </summary>
        public int FromColumn { get; set; } = 0;

        /// <summary>
        /// Row of begin cell
        /// </summary>
        public int FromRow { get; set; } = 0;

        /// <summary>
        /// Column of end cell
        /// </summary>
        public int ToColumn { get; set; } = 0;

        /// <summary>
        /// Row of end cell
        /// </summary>
        public int ToRow { get; set; } = 0;

        /// <summary>
        /// Format string. Default @
        /// </summary>
        public string Format { get; set; } = "@";
    }

}
