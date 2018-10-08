using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Collections;
using DocumentFormat.OpenXml;
using System.Text.RegularExpressions;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelParser
{
    /// <summary>
    /// Represents a worksheet in an Excel workbook
    /// </summary>
    public class ExcelSpreadsheet
    {
        DataTable table = new DataTable();
        string name;

        /// <summary>
        /// The underlying data table which constitutes the data in the spreadsheet.
        /// This includes the rows, columns, and overall schema of the worksheet.
        /// </summary>
        public DataTable Table
        {
            get
            {
                return table;
            }
        }

        /// <summary>
        /// The name of the worksheet
        /// </summary>
        public string Name
        {
            get
            {
                return name;
            }
            set
            {
                if (value == "")
                    throw new Exception("Worksheet name cannot be blank.");
                table.TableName = value;
                name = value;
            }
        }

        /// <summary>
        /// Create a new instance of the excel spreadsheet class.
        /// </summary>
        /// <param name="_name">Name of the spreadsheet.</param>
        public ExcelSpreadsheet(string _name)
        {
            table = new DataTable();
            name = _name;
        }
    }

    /// <summary>
    /// Represents a full Excel workbook.
    /// </summary>
    public class ExcelWorkbook : IEnumerable<ExcelSpreadsheet>
    {
        private List<ExcelSpreadsheet> spreadsheets;

        /// <summary>
        /// Initialize new instance of an Excel workbook.
        /// </summary>
        public ExcelWorkbook()
        {
            spreadsheets = new List<ExcelSpreadsheet>();
        }

        /// <summary>
        /// Get a listing of spreadsheets.
        /// </summary>
        public List<ExcelSpreadsheet> Spreadsheets
        {
            get
            {
                return spreadsheets;
            }
        }

        /// <summary>
        /// Get a particular worksheet by name.
        /// </summary>
        /// <param name="worksheet">Name of the spreadsheet in the workbook.</param>
        /// <returns>Excel spreadsheet with name "spreadsheet".</returns>
        public ExcelSpreadsheet this[string spreadsheet]
        {
            get
            {
                return spreadsheets.First(s => s.Name == spreadsheet);
            }
        }

        /// <summary>
        /// Returns true or false as to whether the specified spreadsheet idenitifed by "name" is in the workbook.
        /// </summary>
        /// <param name="spreadsheet">Name of the spreadsheet to check against spreadsheets currently in the workbook.</param>
        /// <returns>Returns true or false depending on if the spreadsheet with name "spreadsheet" is in the workbook.</returns>
        public bool Contains(string spreadsheet)
        {
            return spreadsheets.Exists(s => s.Name == spreadsheet);
        }

        public static ExcelWorkbook FromFile(string path)
        {
            ExcelWorkbook workbook = new ExcelWorkbook();
            using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (SpreadsheetDocument xlsx = SpreadsheetDocument.Open(fs, false))
                {
                    WorkbookPart workbookpart = xlsx.WorkbookPart;
                    SharedStringTable sharedStrings = null;
                    if (workbookpart.SharedStringTablePart != null)
                    {
                            sharedStrings = workbookpart.SharedStringTablePart.SharedStringTable;
                    }
                    List<WorksheetPart> worksheetParts = workbookpart.WorksheetParts.ToList<WorksheetPart>();
                    foreach (WorksheetPart worksheetPart in worksheetParts)
                    {
                        Worksheet worksheet = worksheetPart.Worksheet;
                        IEnumerable<Row> rows = worksheet.Descendants<Row>();
                        int i = 0;
                        ExcelSpreadsheet spreadsheet = new ExcelSpreadsheet(workbookpart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault(s => s.Id.HasValue && s.Id.Value == workbookpart.GetIdOfPart(worksheetPart)).Name);
                        DataTable table = spreadsheet.Table;
                        foreach (Row row in rows)
                        {
                            if (i == 0)
                            {
                                foreach (Cell cell in row.Elements<Cell>())
                                    table.Columns.Add();
                            }
                            DataRow datarow = table.NewRow();
                            int j = 0;
                            foreach (Cell cell in row.Elements<Cell>())
                            {
                                if (cell.DataType != null && cell.DataType == CellValues.SharedString)
                                    datarow[j++] = sharedStrings.ChildElements[int.Parse(cell.CellValue.Text)].InnerText;
                                else if (cell.CellValue != null)
                                    datarow[j++] = cell.CellValue.Text;
                                else
                                    datarow[j++] = cell.InnerText;
                            }
                            table.Rows.Add(datarow);
                            i++;
                        }
                        if (spreadsheet.Table.Rows.Count > 0)
                        {
                            Regex regex = new Regex(@"^(.+\\)*(.*?)\.xlsx$");
                            spreadsheet.Name = (regex.Match(path).Groups[2].ToString() + "_" + spreadsheet.Name);
                            workbook.Spreadsheets.Add(spreadsheet);
                        }
                    }
                }
            }
            return workbook;
        }

        /// <summary>
        /// Returns the enumerator for the list of spreadsheets.
        /// </summary>
        /// <returns>Enumerator for the list.</returns>
        public IEnumerator<ExcelSpreadsheet> GetEnumerator()
        {
            return spreadsheets.GetEnumerator();
        }

        /// <summary>
        /// Returns the enumerator for the list of spreadsheets.
        /// </summary>
        /// <returns>Enumerator for the list.</returns>
        IEnumerator IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }
    }

    /// <summary>
    /// File management utilities for Excel spreadsheets
    /// </summary>
    public class ExcelWorkbookFinder
    {
        /// <summary>
        /// Returns all Excel workbooks in the directory specified by dirpath.
        /// </summary>
        /// <param name="dirpath">A path to a directory.</param>
        /// <returns></returns>
        public static ExcelWorkbook[] GetAllWorkbooks(string dirpath)
        {
            return GetAllWorkbooks(dirpath, ".*", true);
        }

        /// <summary>
        /// Returns all workbooks matching a specified set of conditions.
        /// </summary>
        /// <param name="dirpath">The directory under which to search. This search is NOT recursive.</param>
        /// <param name="filename">File name for which to search.</param>
        /// <param name="filenameIsRegEx">True or false as to whether "filename" is a regular expression to use on files in "dirpath"</param>
        /// <returns>Array of Excel workbooks.</returns>
        public static ExcelWorkbook[] GetAllWorkbooks(string dirpath, string filename, bool filenameIsRegEx = false)
        {
            List<ExcelWorkbook> workbooks = new List<ExcelWorkbook>();
            foreach (string file in Directory.EnumerateFiles(dirpath, "*.xlsx", SearchOption.TopDirectoryOnly))
            {
                FileInfo info = new FileInfo(file);
                if ((Path.GetFileName(file) == filename || (filenameIsRegEx && (new Regex(filename)).IsMatch(Path.GetFileName(file)))) && Path.GetFileName(file).Trim().Substring(Path.GetFileName(file).Trim().Length - 5 > 0 ? Path.GetFileName(file).Trim().Length - 5 : 0, 5).ToLower() == ".xlsx" && !info.Attributes.HasFlag(FileAttributes.Hidden))
                    workbooks.Add(ExcelWorkbook.FromFile(file));
            }
            return workbooks.ToArray();
        }
    }
}
