using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using DocumentFormat.OpenXml;
using System.IO;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ExcelParser
{
    namespace Reports
    {
        /// <summary>
        /// Represents a report. This class is used for exporting reports.
        /// The data dictionary associates tokens in the report template with data tables.
        /// Note that if a data table contains only one value (i.e. one row and one column),
        /// the data is populated in the generated report as a string value rather than a table.
        /// </summary>
        public class Report
        {
            private string outputPath;
            private string reportName;
            private Dictionary<string, DataTable> dataDictionary;
            private ReportStyle style;
            private string templatepath;

            /// <summary>
            /// The full path, including filename, of the report.
            /// </summary>
            public string OutputPath { get { return outputPath; } set { outputPath = value; } }
            /// <summary>
            /// The name of the report (.docx filename).
            /// </summary>
            public string ReportName { get { return reportName; } set { reportName = value; } }
            /// <summary>
            /// The data dictionary which represents a mapping between data tables and tokens in the report.
            /// When the report is generated from the template, {Token} will be replaced with the data table.
            /// </summary>
            public Dictionary<string, DataTable> DataDictionary { get { return dataDictionary; } }
            /// <summary>
            /// The full path to the template report.
            /// </summary>
            public string TemplatePath { get { return templatepath; } set { templatepath = value; } }

            /// <summary>
            /// Initializes a new instance of a report. Reports are used to generate the actual outputted document.
            /// </summary>
            /// <param name="_reportName">The output filename of the report (.docx file).</param>
            /// <param name="_templatepath">The path to the template which contains tokens to generate the actual report.</param>
            /// <param name="_style">Optional style. This may or may not be implemented in the future.</param>
            public Report(string _reportName, string _templatepath, Dictionary<string, DataTable> _dataDictionary = null, ReportStyle _style = null)
            {
                outputPath = Environment.GetEnvironmentVariable("HOMEPATH") + @"\Documents\ReportOutputs\" + _reportName;
                if (!Directory.Exists(Environment.GetEnvironmentVariable("HOMEPATH") + @"\Documents\ReportOutputs"))
                    Directory.CreateDirectory(Environment.GetEnvironmentVariable("HOMEPATH") + @"\Documents\ReportOutputs");
                reportName = _reportName;
                style = _style;
                templatepath = _templatepath;
                dataDictionary = (dataDictionary == null ? new Dictionary<string, DataTable>() : _dataDictionary);
            }

            /// <summary>
            /// Builds the OpenXML-based table using the data table provided.
            /// </summary>
            /// <returns>OpenXML Table</returns>
            private static Table BuildTable(DataTable table)
            {
                Table openxmlTable = new Table();
                TableProperties properties = new TableProperties(
                    new TableBorders(
                        new TopBorder()
                        {
                            Val = BorderValues.Single,
                            Color = "auto",
                            Size = (UInt32Value)4U,
                            Space = (UInt32Value)0U
                        },
                        new BottomBorder()
                        {
                            Val = BorderValues.Single,
                            Color = "auto",
                            Size = (UInt32Value)4U,
                            Space = (UInt32Value)0U
                        },
                        new LeftBorder()
                        {
                            Val = BorderValues.Single,
                            Color = "auto",
                            Size = (UInt32Value)4U,
                            Space = (UInt32Value)0U
                        },
                        new RightBorder()
                        {
                            Val = BorderValues.Single,
                            Color = "auto",
                            Size = (UInt32Value)4U,
                            Space = (UInt32Value)0U
                        },
                        new InsideHorizontalBorder()
                        {
                            Val = BorderValues.Single,
                            Color = "auto",
                            Size = (UInt32Value)4U,
                            Space = (UInt32Value)0U
                        },
                        new InsideVerticalBorder()
                        {
                            Val = BorderValues.Single,
                            Color = "auto",
                            Size = (UInt32Value)4U,
                            Space = (UInt32Value)0U
                        }
                    ),
                    new TableJustification() { Val = TableRowAlignmentValues.Center },
                    new TableStyle() { Val = "TableGrid" }
                );
                openxmlTable.AppendChild<TableProperties>(properties);
                foreach (DataRow row in table.Rows)
                {
                    TableRow tr = new TableRow();
                    foreach (DataColumn column in table.Columns)
                    {
                        TableCell tc = new TableCell();
                        tc.Append(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Auto }));
                        tc.Append(new Paragraph(new Run(new Text(row[column].ToString()))));
                        tr.Append(tc);
                    }
                    openxmlTable.Append(tr);
                }
                return openxmlTable;
            }

            /// <summary>
            /// Generates the report.
            /// </summary>
            /// <param name="report">The report to generate.</param>
            public static void GenerateReport(Report report)
            {
                Dictionary<string, DataTable> inspectDictionary = new Dictionary<string, DataTable>(report.DataDictionary);
                report.DataDictionary.Clear();
                foreach (KeyValuePair<string, DataTable> item in inspectDictionary)
                {
                    string correctedKey = item.Key;
                    if (correctedKey[0] != '{' || correctedKey[correctedKey.Length - 1] != '}')
                        correctedKey = "{" + correctedKey + "}";
                    report.DataDictionary.Add(correctedKey, item.Value);
                }
                byte[] rawdata;
                using (FileStream fs = new FileStream(report.TemplatePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (MemoryStream ms = new MemoryStream())
                    {
                        fs.CopyTo(ms);
                        rawdata = ms.ToArray();
                    }
                }
                using (MemoryStream ms = new MemoryStream())
                {
                    ms.Write(rawdata, 0, rawdata.Length);
                    ms.Seek(0L, SeekOrigin.Begin);
                    using (WordprocessingDocument docx = WordprocessingDocument.Open(ms, true))
                    {
                        Dictionary<int, Text[]> tokens = new Dictionary<int, Text[]>();
                        Dictionary<int, string> tokenStrings = new Dictionary<int, string>();
                        Text[] textItems = null;
                        Text previous = null;
                        Text previous2 = null;
                        int index = 0;
                        string token = "";
                        foreach (Text item in docx.MainDocumentPart.Document.Body.Descendants<Text>())
                        {
                            // case "... {token} ..."
                            if ((token = ContainsToken(item.Text, report.DataDictionary)) != "")
                            {
                                textItems = new Text[1];
                                textItems[0] = item;
                                tokens.Add(index, textItems);
                                tokenStrings.Add(index++, token);
                            }
                            // case "... {" + "token} ..."
                            else if ((token = ContainsToken("{" + item.Text, report.DataDictionary)) != "" && previous != null && previous.Text[previous.Text.Length - 1] == '{')
                            {
                                textItems = new Text[2];
                                textItems[0] = previous;
                                textItems[1] = item;
                                // use even index
                                if (index % 2 == 1)
                                    index++;
                                tokens.Add(index, textItems);
                                tokenStrings.Add(index++, token);
                            }
                            // case "... {token" + "} ..."
                            else if (previous != null && (token = ContainsToken(previous.Text + "}", report.DataDictionary)) != "" && item.Text[0] == '}')
                            {
                                textItems = new Text[2];
                                textItems[0] = previous;
                                textItems[1] = item;
                                // use odd index
                                if (index % 2 == 0)
                                    index++;
                                tokens.Add(index, textItems);
                                tokenStrings.Add(index++, token);
                            }
                            // case " ... {" + "token" + "} ..."
                            else if (previous2 != null && previous2.Text[previous2.Text.Length - 1] == '{' && (token = ContainsToken("{" + previous.Text + "}", report.DataDictionary)) != "" && item.Text[0] == '}')
                            {
                                textItems = new Text[3];
                                textItems[0] = previous2;
                                textItems[1] = previous;
                                textItems[2] = item;
                                tokens.Add(index, textItems);
                                tokenStrings.Add(index++, token);
                            }
                            previous2 = previous;
                            previous = item;
                        }
                        OpenXmlElement element = null;
                        foreach (KeyValuePair<int, Text[]> item in tokens)
                        {
                            if (item.Value.Count() == 3)
                            {
                                item.Value[0].Text = item.Value[0].Text.Remove(item.Value[0].Text.Length - 1, 1);
                                item.Value[0].Space = SpaceProcessingModeValues.Preserve;
                                item.Value[1].Text = "";
                                item.Value[2].Text = item.Value[2].Text.Remove(0, 1);
                                item.Value[2].Space = SpaceProcessingModeValues.Preserve;
                                element = item.Value[2];
                            }
                            else if (item.Value.Count() == 2)
                            {
                                if (item.Key % 2 == 0)
                                {
                                    item.Value[0].Text = item.Value[0].Text.Remove(0, 1);
                                    item.Value[0].Space = SpaceProcessingModeValues.Preserve;
                                    item.Value[1].Text = item.Value[1].Text.Replace(tokenStrings[item.Key].Replace("{", ""), "");
                                    item.Value[1].Space = SpaceProcessingModeValues.Preserve;
                                }
                                else
                                {
                                    item.Value[0].Text = item.Value[0].Text.Replace(tokenStrings[item.Key].Replace("}", ""), "");
                                    item.Value[0].Space = SpaceProcessingModeValues.Preserve;
                                    item.Value[1].Text = item.Value[1].Text.Remove(item.Value[1].Text.Length - 1, 1);
                                    item.Value[1].Space = SpaceProcessingModeValues.Preserve;
                                }
                                element = item.Value[1];
                            }
                            else
                            {
                                item.Value[0].Text = item.Value[0].Text.Replace(tokenStrings[item.Key], "");
                                item.Value[0].Space = SpaceProcessingModeValues.Preserve;
                                element = item.Value[0];
                            }
                            while (element.GetType() != typeof(Paragraph))
                            {
                                element = element.Parent;
                            }
                            Paragraph paragraph = new Paragraph();
                            Run run = new Run();
                            run.Append(BuildTable(report.DataDictionary[tokenStrings[item.Key]]));
                            paragraph.Append(run);
                            element.InsertAfterSelf(paragraph);
                        }
                        docx.MainDocumentPart.Document.Save();
                    }
                    ms.Seek(0L, SeekOrigin.Begin);
                    using (FileStream fs = new FileStream(report.OutputPath, FileMode.Create))
                    {
                        ms.CopyTo(fs);
                    }
                }
            }

            /// <summary>
            /// Searches the string text for any occurence in the token dictionary.
            /// </summary>
            /// <param name="text">The larger string to search.</param>
            /// <param name="dataDictionary">The dictionary with keys as tokens</param>
            /// <returns>Returns "" if no match; else, returns to the token with braces included.</returns>
            private static string ContainsToken(string text, Dictionary<string, DataTable> dataDictionary)
            {
                foreach (string key in dataDictionary.Keys)
                {
                    if (text.Contains(key))
                        return key;
                }
                return "";
            }

            /// <summary>
            /// Generate a series of reports.
            /// </summary>
            /// <param name="reports">The array of reports to generate.</param>
            public static void GenerateReports(Report[] reports)
            {
                foreach (Report r in reports)
                {
                    GenerateReport(r);
                }
            }
        }
    }

    /// <summary>
    /// Style for the tables which populate the tokens. This may or may not be implemented in the future.
    /// </summary>
    public class ReportStyle
    {
        /// <summary>
        /// Initializes a new instance of the ReportStyle class for generated reports.
        /// </summary>
        public ReportStyle()
        {

        }
    }
}
