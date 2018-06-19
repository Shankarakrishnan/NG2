using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.Linq;

namespace NG2LoadExcelSheets.CommonClass
{
    public class ExcelSheetRecords
    {

        #region "Public Property"

        /// <summary>
        /// get or set the File Name
        /// </summary>
        public string FileName { get; set; }

        /// <summary>
        /// get or set the File Extension
        /// </summary>
        public string FileExt { get; set; }

        /// <summary>
        /// get or set the Sheet Name
        /// </summary>
        public string SheetName { get; set; }

        public DBConnection _dbConnection = new DBConnection();

        #endregion

        #region "Public Methods"

        /// <summary>
        /// Read excel sheet data
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="fileExt"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public DataTable ReadExcel()
        {
            string conn = string.Empty;
            DataTable dtexcel = new DataTable();
            if (FileExt.CompareTo(".xls") == 0)
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + FileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
            else
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileName + ";Extended Properties='Excel 12.0;HDR=NO';"; //for above excel 2007  
            using (OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from [" + SheetName + "$]", con); //here we read data from sheet1  
                    oleAdpt.Fill(dtexcel); //fill excel data into dataTable  
                    dtexcel.Rows.Remove(dtexcel.Rows[0]);
                }
                catch (Exception ex)
                {

                }
            }
            return dtexcel;
        }

        public DataTable GetColumnTitles()
        {
            OleDbConnection con = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data source=" + _dbConnection.ConnectionString());
            DataTable dt = new DataTable();
            using (dt = new DataTable())
            {
                con.Open();
                using (OleDbDataAdapter adapter = new OleDbDataAdapter("SELECT * FROM Title", con))
                {
                    adapter.Fill(dt);
                }
               
            }
            return dt;
        }

        #endregion

    }
}
