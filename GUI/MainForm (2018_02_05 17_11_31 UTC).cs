using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DBInterface;
using ExcelParser;
using ExcelParser.Reports;
using Microsoft.VisualBasic;

namespace GUI
{
    public partial class MainForm : Form
    {
        public DatabaseObject dbObject1 { get; private set; }
        public String currentTableName { get; private set; }
        public List<String> columnNameList = new List<string>();
        public bool rowFirst = false;

        public MainForm()
        {
            try{
                InitializeComponent();
            }
            catch(Exception ex)
            {
                MessageBox.Show("An Error Occurred: " + ex.Message.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void splitContainer1_SplitterMoved(object sender, SplitterEventArgs e)
        {

        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            /*
            //workSheetList.SelectionMode = SelectionMode.MultiExtended;
            DataTable table = new DataTable();

            table.Columns.Add("Column 1",typeof(int));
            table.Columns.Add("Column 2", typeof(string));
            table.Columns.Add("Column 3", typeof(string));

            table.Rows.Add(1, "item 1", "string 1");
            table.Rows.Add(2, "item 2", "string 2");
            table.Rows.Add(3, "item 3", "string 3");
            table.Rows.Add(4, "item 4", "string 4");
            table.Rows.Add(5, "item 5", "string 5");

            dataGridView1.DataSource = table;

            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            }*/

            
        }

        private void fileToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void removeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Are you sure you want to delete this column/row","", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // The user wants to exit the application. Close everything down.
            Application.Exit();
        }

        private void openWorkbookToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                String pubdirect = "";
                StringBuilder regx2 = new StringBuilder();

                openFileDialog1.InitialDirectory = "c:\\";
                openFileDialog1.Filter = "All files (*.*)|*.*|Excel files (*.xlsx)|*.xlsx";
                openFileDialog1.FilterIndex = 2;
                openFileDialog1.RestoreDirectory = true;
                openFileDialog1.Multiselect = true;
                String fname = "";

                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    int x = 0;
                    int y = 0;
                    this.Cursor = Cursors.WaitCursor;
                    foreach (string fileName in openFileDialog1.FileNames) {

                        if (y == 1)
                        {
                            regx2.Append("|");
                        }
                        regx2.Append(Path.GetFileNameWithoutExtension(fileName));
                        regx2.Append(@"\.xlsx");
                        y = 1;
                        //string strfilename = openFileDialog1.InitialDirectory + fileName;
                        string strdirect = Path.GetDirectoryName(fileName);
                        pubdirect = strdirect;
                        string justName = Path.GetFileName(fileName);

                        //MessageBox.Show(strdirect);
                        //MessageBox.Show(justName);

                        ExcelWorkbook[] workbooks = ExcelWorkbookFinder.GetAllWorkbooks(strdirect, justName, false);

                        //ExcelParser.ExcelSpreadsheet spreadsheet = workbooks[0].First(s => s.Name == "test_Sheet1");

                        foreach (ExcelWorkbook fwork in workbooks)
                        {
                            foreach (ExcelParser.ExcelSpreadsheet fsheet in fwork.Spreadsheets)
                            {
                                dbObject1.CreateTable(fsheet);
                            }
                        }

                        if (x == 0)
                            fname = workbooks[0].Spreadsheets[0].Name;

                        //toSend = displayfirst +  + "]";
                        currentTableName = fname;
                        dataGridView1.DataSource = dbObject1.ReadTable(workbooks[0].Spreadsheets[0].Name);
                        workSheetList.Items.Add(workbooks[0].Spreadsheets[0].Name);
                        x++;
                    }
                }
                // MessageBox.Show(regx2.ToString());
                // MessageBox.Show(pubdirect);
                dbObject1.ExecuteQuery("INSERT INTO sqlregularexp (regexp, directorypath) VALUES ('" + regx2.ToString() + "','" + pubdirect + "');");
                queryBox.Text = "Select * From [" + fname + "]";
                this.Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                MessageBox.Show("An Error Occurred: " + ex.Message.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void openDirectoryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try {
                String fname = "";
                String regx1 = ".*";
                string folderpath = "";
                this.Cursor = Cursors.WaitCursor;
                FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();

                if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
                {
                    int x = 0;
                    folderpath = folderBrowserDialog1.SelectedPath;
                    //MessageBox.Show(folderpath);

                    ExcelWorkbook[] workbooks = ExcelWorkbookFinder.GetAllWorkbooks(folderpath);

                    foreach (ExcelWorkbook fwork in workbooks)
                    {
                        foreach (ExcelParser.ExcelSpreadsheet fsheet in fwork.Spreadsheets)
                        {
                            dbObject1.CreateTable(fsheet);
                        }
                    }

                    for (int i = 0; i < workbooks.Length; i++)
                    {
                        for (int j = 0; j < workbooks[i].Spreadsheets.Count; j++) {
                            workSheetList.Items.Add(workbooks[i].Spreadsheets[j].Name);
                            if (i == 0 && j == 0 && x == 0)
                            {
                                fname = workbooks[i].Spreadsheets[j].Name;
                            }
                            x++;
                        }
                    }
                    currentTableName = fname;
                    dataGridView1.DataSource = dbObject1.ReadTable(workbooks[0].Spreadsheets[0].Name);

                }
                //MessageBox.Show(regx1);
                //MessageBox.Show(folderpath);
                dbObject1.ExecuteQuery("INSERT INTO sqlregularexp (regexp, directorypath) VALUES ('" + regx1 + "','" + folderpath + "');");
                queryBox.Text = "Select * From [" + fname + "]";
                this.Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                MessageBox.Show("An Error Occurred: " + ex.Message.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void tableLayoutPanel4_Paint(object sender, PaintEventArgs e)
        {

        }

        private void openDatabaseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try {
                OpenFileDialog openFileDialog2 = new OpenFileDialog();
                String dbPath = "";

                openFileDialog2.InitialDirectory = "c:\\";
                openFileDialog2.Filter = "All files (*.*)|*.*|SQLite files (*.sqlite)|*.sqlite";
                openFileDialog2.FilterIndex = 2;
                openFileDialog2.RestoreDirectory = true;

                if (openFileDialog2.ShowDialog() == DialogResult.OK)
                {
                    workSheetList.Items.Clear();
                    savedQueriesList.Items.Clear();
                    dbPath = openFileDialog2.FileName;
                    //MessageBox.Show(dbPath);
                    dbObject1 = new DatabaseObject(dbPath);
                    // Get existing tables
                }

                DataTable allTables = dbObject1.GetAllTables();

                foreach (DataRow dr in allTables.Rows)
                {
                    int x = 0;
                    //string tableName = (string)dr[0];
                    workSheetList.Items.Add((string)dr[0]);
                    if (x == 0)
                    {
                        dataGridView1.DataSource = dbObject1.ReadTable((string)dr[0]);
                        currentTableName = (string)dr[0];
                        queryBox.Text = "Select * From [" + (string)dr[0] + "]";
                    }
                    x++;
                }

                DataTable allQueries = dbObject1.GetSavedQueries();

                foreach (DataRow dr in allQueries.Rows)
                {
                    //string token = (string)dr[0];
                    savedQueriesList.Items.Add((string)dr[0]);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An Error Occurred: " + ex.Message.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView1_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
           
        }

        private void tableLayoutPanel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void hideToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try {
                int rowId = 0;
                int columnId = 0;

                queryBox.Text = "";

                DataTable columnNames = dbObject1.GetColumnNames(currentTableName);

                foreach (DataRow dr in columnNames.Rows)
                {
                    columnNameList.Add(dr[1].ToString());
                }

                columnNameList.RemoveAt(0);

                String cList = "";

                foreach (String s in columnNameList)
                {
                    cList += ("[" + s + "],");
                }

                String hideThese1 = "SELECT " + cList.Trim(',') + " FROM [";
                String hideThese2 = "] Where rowid not in (";

                if (MessageBox.Show("Are you sure you want to hide this column/row", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    this.Cursor = Cursors.WaitCursor;
                    int selectedCellCount = dataGridView1.GetCellCount(DataGridViewElementStates.Selected);
                    if (selectedCellCount > 0)
                    {
                        int rcount = dataGridView1.Rows.Count;
                        int ccount = dataGridView1.Columns.Count;

                        if (dataGridView1.AreAllCellsSelected(true))
                        {
                            MessageBox.Show("Hiding all cells is not recommended, please delete", "Selected Cells");
                        }
                        else
                        {
                            StringBuilder nowinfo = new StringBuilder();
                            int selectedRowCount = dataGridView1.Rows.GetRowCount(DataGridViewElementStates.Selected);
                            if (selectedRowCount > 0)
                            {
                                StringBuilder sb = new StringBuilder();
                                sb.Append(hideThese1);
                                sb.Append(currentTableName);
                                sb.Append(hideThese2);

                                for (int i = 0; i < selectedRowCount; i++)
                                {
                                    //sb.Append("Row: ");
                                    rowId = dataGridView1.SelectedRows[i].Index + 1;
                                    if (i == 0)
                                    {
                                        sb.Append(rowId.ToString());
                                    }
                                    else
                                    {
                                        sb.Append(",");
                                        sb.Append(rowId.ToString());
                                    }
                                    //sb.Append(Environment.NewLine);
                                }
                                sb.Append(")");
                                nowinfo.Append(sb.ToString());
                                rowFirst = true;
                                ///nowinfo.Append(Environment.NewLine);
                                //sb.Append("Total: " + selectedRowCount.ToString()); 
                                //MessageBox.Show(sb.ToString(), "Selected Rows");   
                            }
                            else
                            {
                                int selectedColumnCount = dataGridView1.Columns.GetColumnCount(DataGridViewElementStates.Selected);
                                if (selectedCellCount > 0)
                                {
                                    StringBuilder sb2 = new StringBuilder(); ;
                                    int lcol = -1;

                                    if (rcount == selectedCellCount)
                                    {
                                        lcol = dataGridView1.SelectedCells[4].ColumnIndex;
                                        //lcol = lcol;

                                        //queryBox.Text = "";

                                    }

                                    DataTable dt = (DataTable)dataGridView1.DataSource;

                                    //MessageBox.Show(dt.Columns[lcol].ColumnName);
                                    //MessageBox.Show(lcol.ToString());

                                    //columnNameList = columnNameList.FindAll(s => s != dt.Columns[lcol].ColumnName || s == "rowid");

                                    columnNameList.Clear();

                                    // DataTable columnNames2 = dbObject1.GetColumnNames(currentTableName);

                                    foreach (DataRow dr in columnNames.Rows)
                                    {
                                        columnNameList.Add(dr[1].ToString());
                                    }
                                    if (rowFirst == true)
                                    {
                                        columnNameList.RemoveAt(lcol + 1);
                                        columnNameList.RemoveAt(0);
                                    }
                                    else
                                    {
                                        columnNameList.RemoveAt(lcol);
                                        columnNameList.RemoveAt(0);
                                    }
                                    cList = "";

                                    foreach (String s in columnNameList)
                                    {
                                        cList += ("[" + s + "],");
                                    }

                                    hideThese1 = "SELECT " + cList.Trim(',') + " FROM [";
                                    hideThese2 = "]";

                                    //StringBuilder sb2 = new StringBuilder();
                                    sb2.Append(hideThese1);
                                    sb2.Append(currentTableName);
                                    sb2.Append(hideThese2);

                                    //MessageBox.Show(sb2.ToString());

                                    if (lcol != -1)
                                    {
                                        nowinfo.Append(sb2.ToString());
                                    }
                                }
                            }

                            if (nowinfo.ToString() != "") {
                                queryBox.Text = nowinfo.ToString();
                                //MessageBox.Show(nowinfo.ToString(), "Selected Columns");
                                dataGridView1.DataSource = dbObject1.ExecuteQuery(nowinfo.ToString());
                                //Send column or row number here
                            }
                            // else { MessageBox.Show("Improper selection"); }
                        }
                    }

                    this.Cursor = Cursors.Default;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An Error Occurred: " + ex.Message.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void runQueryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try {
                this.Cursor = Cursors.WaitCursor;
                string queryString = queryBox.Text;
                // MessageBox.Show(queryString);
                dataGridView1.DataSource = dbObject1.ExecuteQuery(queryString);
                //Send queryString here            
                this.Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                MessageBox.Show("An Error Occurred: " + ex.Message.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void removeSpreadsheetToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try {
                String dropTable = "Drop Table ";
                for (int i = workSheetList.SelectedIndices.Count - 1; i >= 0; i--)
                {
                    //MessageBox.Show(dropTable + workSheetList.SelectedItem.ToString());
                    dbObject1.ExecuteNonQuery(dropTable + workSheetList.SelectedItem.ToString());
                    workSheetList.Items.RemoveAt(workSheetList.SelectedIndices[i]);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An Error Occurred: " + ex.Message.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void workSheetList_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void openSpreadsheetToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try {
                queryBox.Text = "";

                dataGridView1.DataSource = dbObject1.ReadTable(workSheetList.SelectedItem.ToString());

                DataTable columnNames = dbObject1.GetColumnNames(workSheetList.SelectedItem.ToString());

                columnNameList.Clear();

                foreach (DataRow dr in columnNames.Rows)
                {
                    columnNameList.Add(dr[1].ToString());
                }

                columnNameList.RemoveAt(0);

                String cList = "";

                foreach (String s in columnNameList)
                {
                    cList += ("[" + s + "],");
                }

                queryBox.Text = "Select " + cList.Trim(',') + " From [" + workSheetList.SelectedItem.ToString() + "]";
                dataGridView1.DataSource = dbObject1.ExecuteQuery(queryBox.Text);

                rowFirst = true;

                currentTableName = workSheetList.SelectedItem.ToString();

                /* for (int i = workSheetList.SelectedIndices.Count - 1; i >= 0; i--)
                 {
                     MessageBox.Show(workSheetList.SelectedItem.ToString());
                 }*/
            }
            catch (Exception ex)
            {
                MessageBox.Show("An Error Occurred: " + ex.Message.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try {
                String queryString = "";
                String tokenName = savedQueriesList.SelectedItem.ToString();

                DataTable allQueries = dbObject1.GetSavedQueries();

                foreach (DataRow dr in allQueries.Rows)
                {
                    string token = (string)dr[0];
                    string query = (string)dr[1];

                    if (tokenName == token)
                    {
                        queryString = query;
                    }
                }
                queryBox.Text = queryString;
            }
            catch (Exception ex)
            {
                MessageBox.Show("An Error Occurred: " + ex.Message.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void runToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try {
                String queryString = "";
                String tokenName = savedQueriesList.SelectedItem.ToString();

                DataTable allQueries = dbObject1.GetSavedQueries();

                foreach (DataRow dr in allQueries.Rows)
                {
                    string token = (string)dr[0];
                    string query = (string)dr[1];

                    if (tokenName == token)
                    {
                        queryString = query;
                    }
                }
                //MessageBox.Show(queryString);
                dataGridView1.DataSource = dbObject1.ExecuteQuery(queryString);
                queryBox.Text = queryString;
            }
            catch (Exception ex)
            {
                MessageBox.Show("An Error Occurred: " + ex.Message.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try {
                //String dropRow1 = "Delete From ";
                //String dropRow2 = "Where ";
                for (int i = savedQueriesList.SelectedIndices.Count - 1; i >= 0; i--)
                {
                    // MessageBox.Show(savedQueriesList.SelectedItem.ToString());
                    //dbObject1.ExecuteQuery(dropTable + workSheetList.SelectedItem.ToString());
                    savedQueriesList.Items.RemoveAt(savedQueriesList.SelectedIndices[i]);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An Error Occurred: " + ex.Message.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }



        private void newDatabaseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try {

                workSheetList.Items.Clear();
                savedQueriesList.Items.Clear();
                queryBox.Text = "Select * From";
                //dataGridView1.Rows.Clear();
                //dataGridView1.Refresh();

                string filepath = "";
                SaveFileDialog saveFileDialog1 = new SaveFileDialog();

                saveFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
                saveFileDialog1.FilterIndex = 2;
                saveFileDialog1.RestoreDirectory = true;

                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    filepath = saveFileDialog1.InitialDirectory + saveFileDialog1.FileName + ".sqlite";
                    //MessageBox.Show(filepath);
                    dbObject1 = new DatabaseObject(filepath);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An Error Occurred: " + ex.Message.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            geenerateReportToolStripMenuItem_Click(sender, e);
        }

        private void RefreshButton_Click(object sender, EventArgs e)
        {
            try {
                String fname = "";
                dbObject1.DropTables();
                workSheetList.Items.Clear();
                DataTable repopulate = dbObject1.ReadTable("sqlregularexp");

                foreach (DataRow dr in repopulate.Rows)
                {
                    int x = 0;
                    string regex = (string)dr[0];
                    string dpath = (string)dr[1];

                    // MessageBox.Show(regex);
                    // MessageBox.Show(dpath);

                    ExcelWorkbook[] workbooks = ExcelWorkbookFinder.GetAllWorkbooks(dpath, regex, true);

                    //ExcelParser.ExcelSpreadsheet spreadsheet = workbooks[0].First(s => s.Name == "test_Sheet1");

                    foreach (ExcelWorkbook fwork in workbooks)
                    {
                        foreach (ExcelParser.ExcelSpreadsheet fsheet in fwork.Spreadsheets)
                        {
                            dbObject1.CreateTable(fsheet);
                            workSheetList.Items.Add(fsheet.Name);
                            if (x == 0)
                            {
                                fname = workbooks[0].Spreadsheets[0].Name;
                                currentTableName = fname;
                            }
                            x++;
                        }
                    }

                }
                dataGridView1.DataSource = dbObject1.ReadTable(fname);
            }
            catch (Exception ex)
            {
                MessageBox.Show("An Error Occurred: " + ex.Message.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ExceltoDBbutton_Click(object sender, EventArgs e)
        {
            openWorkbookToolStripMenuItem_Click(sender, e);
        }

        public void geenerateReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try {
                DataTable allQueries = dbObject1.GetSavedQueries();

                String reportName = Interaction.InputBox("What would you like as the name of you report? Afterwards select your template file.", "Report Name/Template Window", "Report.docx");

                OpenFileDialog openFileDialog3 = new OpenFileDialog();
                String templatePath = "";

                openFileDialog3.InitialDirectory = "c:\\";
                openFileDialog3.Filter = "All files (*.*)|*.*|Word files (*.docx)|*.docx";
                openFileDialog3.FilterIndex = 2;
                openFileDialog3.RestoreDirectory = true;

                if (openFileDialog3.ShowDialog() == DialogResult.OK)
                {
                    templatePath = openFileDialog3.FileName;

                    Report report = new Report(reportName, templatePath);

                    foreach (DataRow dr in allQueries.Rows)
                    {
                        string token = (string)dr[0];
                        string query = (string)dr[1];
                        DataTable result;
                        DataTable reportTable = new DataTable();
                        DataRow columnData;
                        result = dbObject1.ExecuteQuery(query);
                        columnData = result.NewRow();

                        foreach (DataColumn dc in result.Columns)
                        {
                            reportTable.Columns.Add();
                        }

                        DataRow headerRow = reportTable.NewRow();
                        int k = 0;

                        foreach (DataColumn dc in result.Columns)
                        {
                            headerRow[k++] = dc.ColumnName;
                        }

                        reportTable.Rows.Add(headerRow);

                        foreach (DataRow row in result.Rows)
                        {
                            k = 0;
                            DataRow newRow = reportTable.NewRow();
                            foreach (DataColumn dc in result.Columns)
                            {
                                newRow[k++] = row[dc].ToString();
                            }
                            reportTable.Rows.Add(newRow);
                        }
                        
                        report.DataDictionary.Add("{" + token + "}", reportTable);
                    }
                    Report.GenerateReport(report);

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An Error Occurred: " + ex.Message.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void saveQueryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try {
                string queryString = queryBox.Text;
                String tokenName = Interaction.InputBox("What would you like query to be called?", "Query Name Window", "Query1");

                dbObject1.ExecuteQuery("INSERT INTO sqlqueries (queryname, query) VALUES ('" + tokenName + "','" + queryString.Replace("'", "''") + "');");

                savedQueriesList.Items.Add(tokenName);
            }
            catch (Exception ex)
            {
                MessageBox.Show("An Error Occurred: " + ex.Message.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void queryBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void worksheetMenuStrip1_Opening(object sender, CancelEventArgs e)
        {

        }
    }
}
