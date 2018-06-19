using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Data.OleDb;
using System.IO;
using NG2LoadExcelSheets.CommonClass;
using System.Data;
using NG2LoadExcelSheets.ExcelInputSheets;

namespace NG2LoadExcelSheets
{
    public partial class UploadFile : Form
    {

        #region "Public Property"

        public ExcelSheetRecords excelSheetRecords = new ExcelSheetRecords();

        public DBConnection _dbConnection = new DBConnection();

        public S1AllTicketsCreated s1AllTicketsCreated = new S1AllTicketsCreated();

        #endregion

        #region "Constructor"

        public UploadFile()
        {
            InitializeComponent();
        }

        #endregion

        #region "Event- Browse the file"

        /// <summary>
        /// Browse the file
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnBrowseFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog _openFileDialog = new OpenFileDialog();
            _openFileDialog.InitialDirectory = "D:/";
            _openFileDialog.Title = "Select file to be upload";
            _openFileDialog.Filter = "Select Valid Document(*.pdf; *.doc; *.xlsx; *.html)|*.pdf; *.docx; *.xlsx; *.html"; ;
            _openFileDialog.FilterIndex = 1;
            try
            {
                if (_openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    if (_openFileDialog.CheckFileExists)
                    {
                        string path = System.IO.Path.GetFullPath(_openFileDialog.FileName);
                        label2.Text = _openFileDialog.FileName;
                        lblDisplayFileLocation.Text = path;
                    }
                }
                else
                {
                    MessageBox.Show("Please Upload document.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        #endregion
        
        #region "Event- Upload excel sheet"

        private void btnUpload_Click(object sender, EventArgs e)
        {
            try
            {
                //GenerateSheets();
                Save();
                //UpDataDB();
            }
            catch (Exception ex)
            {

            }
        }

        #endregion

        /// <summary>
        /// Bulk upload (Moved data from excel sheets to database)
        /// </summary>
        public void Save()
        {
            bool Result = false;
            object readOnly = false;
            Microsoft.Office.Interop.Excel.Application oExcelApp = new Microsoft.Office.Interop.Excel.Application();
            object missing = System.Reflection.Missing.Value;
            Workbook oExcelWorkBook = oExcelApp.Workbooks.Open(lblDisplayFileLocation.Text.ToString().Trim());
            int numSheets = oExcelWorkBook.Sheets.Count;
            excelSheetRecords.FileName = label2.Text.ToString().Trim(); //get the path of the file 
            excelSheetRecords.FileExt = Path.GetExtension(label2.Text.ToString().Trim()); //get the file extension
            System.Data.DataTable dtColumnTitle = excelSheetRecords.GetColumnTitles();
            foreach (Worksheet worksheet in oExcelWorkBook.Worksheets)  // Get total number of sheets
            {
                Range excelRange = worksheet.UsedRange;
                excelSheetRecords.SheetName = worksheet.Name;
                //int RowCount = excelRange.Rows.Count;
                //int ColumnCount = excelRange.Columns.Count;
                
                try
                {
                    if (excelSheetRecords.SheetName.ToString().Trim().Equals(dtColumnTitle.Select("SheetName = '" + _dbConnection.AllTicketsCreated() + "'")[0]["SheetName"]))
                    {
                        s1AllTicketsCreated.AssignColumnName(dtColumnTitle.Select("SheetName = '" + _dbConnection.AllTicketsCreated() + "'"));
                        Result = s1AllTicketsCreated.Save(excelSheetRecords);
                        MessageBox.Show("Records stored successfully");
                    }

                    //...



                    // Filter
                    // To generate 6 files
                }
                catch (Exception ex)
                {

                }
            }
        }

        #region "Backup"

        public void UpDataDB()
        {
            // Prerequisite: The data to be inserted is available in a DataTable/DataSet.
            System.Data.DataTable data = new System.Data.DataTable();
            data.Columns.Add("RegNo", typeof(int));
            data.Columns.Add("Name", typeof(string));
            data.Rows.Add(122, "SK");
            data.Rows.Add(123, "Test");

            // Now, open a database connection using the Microsoft.Jet.OLEDB provider.
            // The "using" statement ensures that the connection is closed no matter what.
            using (var connection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data source=" + _dbConnection.ConnectionString()))
            {
                connection.Open();

                // Create an OleDbDataAdapter and provide it with an INSERT command.
                var adapter = new OleDbDataAdapter();
                adapter.InsertCommand = new OleDbCommand("INSERT INTO SampleSheet (RegNo, Name) VALUES (@RegNo , @Name)", connection);
                adapter.InsertCommand.Parameters.Add("@RegNo", OleDbType.VarChar, 40, "RegNo");
                adapter.InsertCommand.Parameters.Add("@Name", OleDbType.VarChar, 24, "Name");

                // Hit the big red button!
                adapter.Update(data);
            }
        }

        public void GenerateSheets()
        {
            Microsoft.Office.Interop.Excel.Application oExcelApp = new Microsoft.Office.Interop.Excel.Application();

            object readOnly = false;

            object missing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Excel.Workbook oExcelWorkBook = oExcelApp.Workbooks.Open(lblDisplayFileLocation.Text.ToString().Trim()
                                //,missing, readOnly, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing
                                );

            int numSheets = oExcelWorkBook.Sheets.Count;

            excelSheetRecords.FileName = label2.Text.ToString().Trim(); //get the path of the file 
            excelSheetRecords.FileExt = Path.GetExtension(label2.Text.ToString().Trim()); //get the file extension
            foreach (Worksheet worksheet in oExcelWorkBook.Worksheets)
            {
                Range excelRange = worksheet.UsedRange;
                excelSheetRecords.SheetName = worksheet.Name;
                int RowCount = excelRange.Rows.Count;
                int ColumnCount = excelRange.Columns.Count;

                if (excelSheetRecords.FileExt.CompareTo(".xls") == 0 || excelSheetRecords.FileExt.CompareTo(".xlsx") == 0)
                {
                    try
                    {
                        System.Data.DataTable dtExcelSheetData = new System.Data.DataTable();
                        dtExcelSheetData = excelSheetRecords.ReadExcel();
                        gridMatricsReport.Visible = true;
                        gridMatricsReport.DataSource = dtExcelSheetData;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                    }
                }
                else
                {
                    MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error  
                }

            }




            /*
            foreach (Worksheet worksheet in oExcelWorkBook.Worksheets)
            {
                Range excelRange = worksheet.UsedRange;
                string ExcelSheetName = worksheet.Name;
                int RowCount = excelRange.Rows.Count;
                int ColumnCount = excelRange.Columns.Count;
                for (int r = 1; r <= RowCount; r++)
                {
                    for (int c = 1; c <= ColumnCount; c++)
                    {
                        dynamic cell = excelRange.Cells[r, c];
                        try
                        {
                            if (cell.Locked == false)
                            {
                                string content = cell.Value2;
                                if (content != null && !content.Trim().Equals(""))
                                {
                                    content = content.Trim();
                                    cell.Value2 = cell.Value2 + " - This is a test";
                                }
                            }
                        }
                        catch (Exception)
                        {
                            // we are using dynamic type for cell variable so
                            // the variable might not have all the properties we used in our code
                        }

                    }
                }
            }
            */

            //oExcelWorkBook.Save();
            //oExcelApp.Application.Quit();
        }

     

        public void GenerateSheets1()
        {
            string filePath = string.Empty;
            string fileExt = string.Empty;

            filePath = label2.Text; //get the path of the file  
            fileExt = Path.GetExtension(filePath); //get the file extension  
            if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
            {
                try
                {
                    System.Data.DataTable dtExcel = new System.Data.DataTable();
                    //dtExcel = excelSheetRecords.ReadExcel(filePath, fileExt); //read excel file  
                    gridMatricsReport.Visible = true;
                    gridMatricsReport.DataSource = dtExcel;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
            }
            else
            {
                MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error  
            }


            Microsoft.Office.Interop.Excel.Application oExcelApp = new Microsoft.Office.Interop.Excel.Application();

            object readOnly = false;

            object missing = System.Reflection.Missing.Value;

            Microsoft.Office.Interop.Excel.Workbook oExcelWorkBook = oExcelApp.Workbooks.Open(lblDisplayFileLocation.Text.ToString().Trim(),
                                missing, readOnly, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);

            int numSheets = oExcelWorkBook.Sheets.Count;

            foreach (Worksheet worksheet in oExcelWorkBook.Worksheets)
            {
                Range excelRange = worksheet.UsedRange;
                string ExcelSheetName = worksheet.Name;
                int RowCount = excelRange.Rows.Count;
                int ColumnCount = excelRange.Columns.Count;
                for (int r = 1; r <= RowCount; r++)
                {
                    for (int c = 1; c <= ColumnCount; c++)
                    {
                        dynamic cell = excelRange.Cells[r, c];
                        try
                        {
                            if (cell.Locked == false)
                            {
                                string content = cell.Value2;
                                if (content != null && !content.Trim().Equals(""))
                                {
                                    content = content.Trim();
                                    cell.Value2 = cell.Value2 + " - This is a test";
                                }
                            }
                        }
                        catch (Exception)
                        {
                            // we are using dynamic type for cell variable so
                            // the variable might not have all the properties we used in our code
                        }

                    }
                }
            }


            oExcelWorkBook.Save();
            oExcelApp.Application.Quit();
        }

        #endregion

        #region "Working- Bind excel records into grid view"

        /// <summary>
        /// Bind data to grid
        /// </summary>
        public void GenerateExcelData()
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                Workbook xlWorkbook = xlApp.Workbooks.Open(lblDisplayFileLocation.Text.ToString().Trim());
                _Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;

                // dt.Column = colCount;  
                gridMatricsReport.ColumnCount = colCount;
                gridMatricsReport.RowCount = rowCount;

                for (int i = 1; i <= rowCount; i++)
                {
                    for (int j = 1; j <= colCount; j++)
                    {
                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        {
                            gridMatricsReport.Rows[i - 1].Cells[j - 1].Value = xlRange.Cells[i, j].Value2.ToString();
                        }
                    }
                }

                //cleanup  
                GC.Collect();
                GC.WaitForPendingFinalizers();

                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //close and release  
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);

                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
            }
            catch (Exception ex)
            {

            }
        }

        #endregion

        
    }
}
