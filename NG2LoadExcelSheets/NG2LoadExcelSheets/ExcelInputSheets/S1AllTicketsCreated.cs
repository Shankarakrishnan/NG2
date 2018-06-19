using NG2LoadExcelSheets.CommonClass;
using System;
using System.Data;
using System.Data.OleDb;
using System.Linq;

namespace NG2LoadExcelSheets.ExcelInputSheets
{

    /// <summary>
    /// Manipulating Sheet-1 
    /// </summary>
    public class S1AllTicketsCreated
    {

        #region "Public Property"

        public string SheetID { get; set; }
        public string SheetName { get; set; }
        public string AccountNumber { get; set; }
        public string TicketNumber { get; set; }
        public string TicketAssignedTo { get; set; }
        public string CreateDateTime { get; set; }
        public string CompleteDateTime { get; set; }
        public string WorkTypeName { get; set; }
        public string NumberOfWorkEntries { get; set; }
        public string IssueType { get; set; }
        public string SubIssueType { get; set; }
        public string Source { get; set; }
        public string UDFDeferredForResolution { get; set; }
        public string Priority { get; set; }
        public string TicketStatus { get; set; }
        public string TotalWorkedHours { get; set; }
        public string TicketDurationDays { get; set; }
        public string TicketDurationHours { get; set; }
        public string TicketQueue { get; set; }

        public DBConnection _dbConnection = new DBConnection();

        #endregion

        #region "Costructor"

        /// <summary>
        /// Default constructor
        /// </summary>
        public S1AllTicketsCreated()
        {

        }

        #endregion

        #region "Public Methods"

        #region "Bulk insert- insert bulk data into MS Access"

        /// <summary>
        /// Convert Excel- Sheet 1
        /// </summary>
        /// <param name="excelSheetRecords"></param>
        /// <returns></returns>
        public bool Save(ExcelSheetRecords excelSheetRecords)
        {
            try
            {
                DataTable data = new DataTable();
                data.Columns.Add(AccountNumber, typeof(string));
                data.Columns.Add(TicketNumber, typeof(string));
                data.Columns.Add(TicketAssignedTo, typeof(string));
                data.Columns.Add(CreateDateTime, typeof(string));
                data.Columns.Add(CompleteDateTime, typeof(string));
                data.Columns.Add(WorkTypeName, typeof(string));
                data.Columns.Add(NumberOfWorkEntries, typeof(string));
                data.Columns.Add(IssueType, typeof(string));
                data.Columns.Add(SubIssueType, typeof(string));
                data.Columns.Add(Source, typeof(string));
                data.Columns.Add(UDFDeferredForResolution, typeof(string));
                data.Columns.Add(Priority, typeof(string));
                data.Columns.Add(TicketStatus, typeof(string));
                data.Columns.Add(TotalWorkedHours, typeof(string));
                data.Columns.Add(TicketDurationDays, typeof(string));
                data.Columns.Add(TicketDurationHours, typeof(string));
                data.Columns.Add(TicketQueue, typeof(string));

                DataTable dtData = excelSheetRecords.ReadExcel();
                dtData.Columns["F1"].ColumnName = AccountNumber;
                dtData.Columns["F2"].ColumnName = TicketNumber;
                dtData.Columns["F3"].ColumnName = TicketAssignedTo;
                dtData.Columns["F4"].ColumnName = CreateDateTime;
                dtData.Columns["F5"].ColumnName = CompleteDateTime;
                dtData.Columns["F6"].ColumnName = WorkTypeName;
                dtData.Columns["F7"].ColumnName = NumberOfWorkEntries;
                dtData.Columns["F8"].ColumnName = IssueType;
                dtData.Columns["F9"].ColumnName = SubIssueType;
                dtData.Columns["F10"].ColumnName = Source;
                dtData.Columns["F11"].ColumnName = UDFDeferredForResolution;
                dtData.Columns["F12"].ColumnName = Priority;
                dtData.Columns["F13"].ColumnName = TicketStatus;
                dtData.Columns["F14"].ColumnName = TotalWorkedHours;
                dtData.Columns["F15"].ColumnName = TicketDurationDays;
                dtData.Columns["F16"].ColumnName = TicketDurationHours;
                dtData.Columns["F17"].ColumnName = TicketQueue;

                //remove empty row, where whole columns (row) contains null in excel document
                dtData = dtData.Rows.Cast<DataRow>().Where(row => !row.ItemArray.All(field => field is DBNull || string.IsNullOrWhiteSpace(field as string))).CopyToDataTable();

                //We should merge datatype as string for bulk upload
                //Other datatypes to string 
                foreach (DataRow iData in dtData.Rows)
                {
                    data.Rows.Add(iData[AccountNumber], iData[TicketNumber], iData[TicketAssignedTo], iData[CreateDateTime],
                     iData[CompleteDateTime], iData[WorkTypeName], iData[NumberOfWorkEntries], iData[IssueType],
                     iData[SubIssueType], iData[Source], iData[UDFDeferredForResolution], iData[Priority],
                     iData[TicketStatus], iData[TotalWorkedHours], iData[TicketDurationDays], iData[TicketDurationHours],
                     iData[TicketQueue]);
                }

                string InsertQuery = "INSERT INTO AllTicketsCreated " +
                                            "(" + AccountNumber + ", " + TicketNumber + ", " + TicketAssignedTo + "," +
                                                CreateDateTime + ", " + CompleteDateTime + ", " + WorkTypeName + "," +
                                                NumberOfWorkEntries + ", " + IssueType + ", " + SubIssueType + "," +
                                                Source + ", " + UDFDeferredForResolution + ", " + Priority + "," +
                                                TicketStatus + ", " + TotalWorkedHours + ", " + TicketDurationDays + "," +
                                                TicketDurationHours + ", " + TicketQueue + ")" +

                                    " VALUES (@" + AccountNumber + " , @" + TicketNumber + ", @" + TicketAssignedTo +
                                    ", @" + CreateDateTime + " , @" + CompleteDateTime + ", @" + WorkTypeName +
                                    ", @" + NumberOfWorkEntries + " , @" + IssueType + ", @" + SubIssueType +
                                    ", @" + Source + " , @" + UDFDeferredForResolution + ", @" + Priority +
                                    ", @" + TicketStatus + " , @" + TotalWorkedHours + ", @" + TicketDurationDays +
                                    ", @" + TicketDurationHours + " , @" + TicketQueue + " )";
                using (var connection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data source=" + _dbConnection.ConnectionString() + ""))
                {
                    connection.Open();
                    var adapter = new OleDbDataAdapter();
                    adapter.InsertCommand = new OleDbCommand(InsertQuery, connection);
                    adapter.InsertCommand.Parameters.Add("@" + AccountNumber, OleDbType.VarChar, 255, AccountNumber);
                    adapter.InsertCommand.Parameters.Add("@" + TicketNumber, OleDbType.VarChar, 255, TicketNumber);
                    adapter.InsertCommand.Parameters.Add("@" + TicketAssignedTo, OleDbType.VarChar, 255, TicketAssignedTo);

                    adapter.InsertCommand.Parameters.Add("@" + CreateDateTime, OleDbType.VarChar, 25, CreateDateTime);
                    adapter.InsertCommand.Parameters.Add("@" + CompleteDateTime, OleDbType.VarChar, 25, CompleteDateTime);
                    adapter.InsertCommand.Parameters.Add("@" + WorkTypeName, OleDbType.VarChar, 255, WorkTypeName);

                    adapter.InsertCommand.Parameters.Add("@" + NumberOfWorkEntries, OleDbType.VarChar, 10, NumberOfWorkEntries);
                    adapter.InsertCommand.Parameters.Add("@" + IssueType, OleDbType.VarChar, 100, IssueType);
                    adapter.InsertCommand.Parameters.Add("@" + SubIssueType, OleDbType.VarChar, 255, SubIssueType);

                    adapter.InsertCommand.Parameters.Add("@" + Source, OleDbType.VarChar, 100, Source);
                    adapter.InsertCommand.Parameters.Add("@" + UDFDeferredForResolution, OleDbType.VarChar, 10, UDFDeferredForResolution);
                    adapter.InsertCommand.Parameters.Add("@" + Priority, OleDbType.VarChar, 100, Priority);

                    adapter.InsertCommand.Parameters.Add("@" + TicketStatus, OleDbType.VarChar, 100, TicketStatus);
                    adapter.InsertCommand.Parameters.Add("@" + TotalWorkedHours, OleDbType.VarChar, 10, TotalWorkedHours);
                    adapter.InsertCommand.Parameters.Add("@" + TicketDurationDays, OleDbType.VarChar, 10, TicketDurationDays);

                    adapter.InsertCommand.Parameters.Add("@" + TicketDurationHours, OleDbType.VarChar, 10, TicketDurationHours);
                    adapter.InsertCommand.Parameters.Add("@" + TicketQueue, OleDbType.VarChar, 100, TicketQueue);

                    //Bulk upload
                    adapter.Update(data);
                }
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        #endregion


        #region "AssignColumnName"

        /// <summary>
        /// Assign column name to the property dynamically
        /// </summary>
        /// <param name="dtColumn"></param>
        public void AssignColumnName(DataRow[] dtColumn)
        {
            this.AccountNumber = dtColumn[0]["AccountNumber"].ToString().Trim();
            this.TicketNumber = dtColumn[0]["TicketNumber"].ToString().Trim();
            this.TicketAssignedTo = dtColumn[0]["TicketAssignedTo"].ToString().Trim();
            this.CreateDateTime = dtColumn[0]["CreateDateTime"].ToString().Trim();
            this.CompleteDateTime = dtColumn[0]["CompleteDateTime"].ToString().Trim();
            this.WorkTypeName = dtColumn[0]["WorkTypeName"].ToString().Trim();
            this.NumberOfWorkEntries = dtColumn[0]["NumberOfWorkEntries"].ToString().Trim();
            this.IssueType = dtColumn[0]["IssueType"].ToString().Trim();
            this.SubIssueType = dtColumn[0]["SubIssueType"].ToString().Trim();
            this.Source = dtColumn[0]["Source"].ToString().Trim();
            this.UDFDeferredForResolution = dtColumn[0]["UDFDeferredForResolution"].ToString().Trim();
            this.Priority = dtColumn[0]["Priority"].ToString().Trim();
            this.TicketStatus = dtColumn[0]["TicketStatus"].ToString().Trim();
            this.TotalWorkedHours = dtColumn[0]["TotalWorkedHours"].ToString().Trim();
            this.TicketDurationDays = dtColumn[0]["TicketDurationDays"].ToString().Trim();
            this.TicketDurationHours = dtColumn[0]["TicketDurationHours"].ToString().Trim();
            this.TicketQueue = dtColumn[0]["TicketQueue"].ToString().Trim();
        }

        #endregion

        #endregion

    }
}
