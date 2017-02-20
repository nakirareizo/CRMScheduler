using BOL.Wizard;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using BOL.AppConst;
using BOL.AuditLog.Modules;
using System.Configuration;
using System.Data.OleDb;
using System.IO;
using System.Collections;

namespace CRMSync.Classes.MainProcess
{
    class ApprovalLetterDate
    {

        private struct RecordsCount
        {
            public int InsertCount;
            public int UpdateCount;
            public int NonMSCCount;
            public int SameDateCount;
            public int NoAccountCount;
        }

        internal void StartSync()
        {
            LogFileHelper.logList = new ArrayList();
            Console.WriteLine(string.Format("[{0}] : Start Sync Approval Letter", DateTime.Now.ToString()));
            LogFileHelper.logList.Add(string.Format("[{0}] : Start Sync Approval Letter", DateTime.Now.ToString()));
            RecordsCount SyncCount = default(RecordsCount);

            // Initialising the variables
            SyncCount.InsertCount = 0;
            SyncCount.NoAccountCount = 0;
            SyncCount.NonMSCCount = 0;
            SyncCount.SameDateCount = 0;
            SyncCount.UpdateCount = 0;
            string FileID = "";
            try
            {
                //System.Data.DataTable wizardData = this.GetWizardApprovedDate();
                //UPDATED 10062016
                DataTable wizardData = this.GetFromWizardApprovedDate();
                //statusdate
                //string Filename0 = getFileName();
                //DataTable wizardData = GetWizardApprovedDateFromExcelFile(Filename0, "ApprovalDates_10062016");
                int totalRecord = wizardData.Rows.Count;
                Console.WriteLine(string.Format("[{0}] : Total record for Approval Letter Dates is {1}", DateTime.Now.ToString(), totalRecord));
                LogFileHelper.logList.Add(string.Format("[{0}] : Total record for Approval Letter Dates is {1}", DateTime.Now.ToString(), totalRecord));
                int count = 0;
                //if (totalRecord > 0)
                //{
                //    ExcelFileHelper.GenerateExcelFileApprovalDates(wizardData, DateTime.Now.ToString("dd-MM-yyyy"));
                //}
                Guid? MOFApprovalStatus = SyncHelper.GetCodeMasterID("Approval Letter", CodeType.MSCApprovalStatus, true);
                Guid MOFApprovalStatusCID = new Guid(MOFApprovalStatus.Value.ToString());
                foreach (System.Data.DataRow row in wizardData.Rows)
                {
                    count += 1;

                    FileID = row["RefNumber"].ToString();
                    FileID = "CS/3/" + FileID;
                    // Field below must be change ****************************************
                    string ApprovalLetterDate = SyncHelper.ConvertStringToDateTime(row["StatusDate"].ToString().Trim(), false).ToString();
                    //Dim StatusDescription As String = SyncHelper.GetMappingValue(AppConst.CodeType.MSCApprovalStatus, "MOF Approved")
                    if (!string.IsNullOrEmpty(ApprovalLetterDate))
                    {

                        Guid? accountID = GetAccountIDByFileID(FileID);

                        if (accountID != null)
                        {
                            if (IsNonMSC(accountID.Value))
                            {
                                SyncCount.NonMSCCount += 1;
                                Console.WriteLine(string.Format("[{0}] {2}/{3} : Skip Non MSC record FileID {1}", DateTime.Now.ToString(), FileID, count, totalRecord));
                                LogFileHelper.logList.Add(string.Format("[{0}] {2}/{3} : Skip Non MSC record FileID {1}", DateTime.Now.ToString(), FileID, count, totalRecord));
                                continue;
                            }

                            Guid? MSCStatusHistoryID = GetMSCStatusHistoryID(accountID.Value, MOFApprovalStatusCID);

                            if (MSCStatusHistoryID != null)
                            {
                                if (IsDifferentApprovalDate(MSCStatusHistoryID, ApprovalLetterDate))
                                {
                                    UpdateApprovalLetterDate(accountID.Value, MSCStatusHistoryID.Value, Convert.ToDateTime(ApprovalLetterDate));
                                    SyncCount.UpdateCount += 1;
                                    Console.WriteLine(string.Format("[{0}] {2}/{3} : Update MSCStatusHistory FileID {1}", DateTime.Now.ToString(), FileID, count, totalRecord));
                                    LogFileHelper.logList.Add(string.Format("[{0}] {2}/{3} : Update MSCStatusHistory FileID {1}", DateTime.Now.ToString(), FileID, count, totalRecord));
                                }
                                else
                                {
                                    SyncCount.SameDateCount += 1;
                                    Console.WriteLine(string.Format("[{0}] {2}/{3} : Skip same approval letter date FileID {1}", DateTime.Now.ToString(), FileID, count, totalRecord));
                                    LogFileHelper.logList.Add(string.Format("[{0}] {2}/{3} : Skip same approval letter date FileID {1}", DateTime.Now.ToString(), FileID, count, totalRecord));                            // Skip to update same date
                                }
                            }
                            else
                            {
                                SyncCount.InsertCount += 1;
                                CreateApprovalLetterDate(accountID.Value, MOFApprovalStatusCID, Convert.ToDateTime(ApprovalLetterDate));
                                Console.WriteLine(string.Format("[{0}] {2}/{3} : Create MSCStatusHistory FileID {1}", DateTime.Now.ToString(), FileID, count, totalRecord));
                                LogFileHelper.logList.Add(string.Format("[{0}] {2}/{3} : Create MSCStatusHistory FileID {1}", DateTime.Now.ToString(), FileID, count, totalRecord));
                            }
                        }
                        else
                        {
                            SyncCount.NonMSCCount += 1;
                            Console.WriteLine(string.Format("[{0}] {2}/{3} : Record FileID {1} not found", DateTime.Now.ToString(), FileID, count, totalRecord));
                            LogFileHelper.logList.Add(string.Format("[{0}] {2}/{3} : Record FileID {1} not found", DateTime.Now.ToString(), FileID, count, totalRecord));
                        }
                    }
                    Console.WriteLine(string.Format("[{0}] : End Sync Approval Letter", DateTime.Now.ToString()));
                    LogFileHelper.logList.Add(string.Format("[{0}] : End Sync Approval Letter", DateTime.Now.ToString()));
                }
            }
            catch (Exception ex)
            {
                LogFileHelper.logList.Add("Error for MSCFileID" + FileID + ", Error : " + ex.Message);
                List<string> TOs = new List<string>();
                //TOs.AddRange(BOL.Common.Modules.Parameter.WIZARD_RCPNT.Split(','));
                TOs.Add("appanalyst@mdec.com.my");
                bool SendSuccess = BOL.Utils.Email.SendMail(TOs.ToArray(), null, null, BOL.Common.Modules.Parameter.WIZARD_SUBJ, string.Format("{0} SyncApprovalLetterDates {1}", BOL.Common.Modules.Parameter.WIZARD_DESC, ex.Message), null);
            }
            //if (LogFileHelper.logList.Count > 0)
            //{
            //    string ModeSync = "ApprovalLetterDatesSyncLog_";
            //    LogFileHelper.WriteLog(LogFileHelper.logList, ModeSync);
            //}
        }
        public static string getFileName()
        {
            //LogList = new List<string>();
            string Directory = @"C:\CRMSync\ExcelRecord";
            var directory0 = new DirectoryInfo(Directory);
            var myFile0 = (from f in directory0.GetFiles()
                           orderby f.LastWriteTime descending
                           select f).First();
            string Filename0 = Directory + @"\" + myFile0;
            return Filename0;
        }
        private DataTable GetWizardApprovedDateFromExcelFile(string fullFilename, string tableName)
        {
            string conStr = "";
            string Extension = ".xlsx";
            switch (Extension)
            {
                case ".xls": //Excel 97-03
                    conStr = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};
                         Extended Properties = 'Excel 8.0;HDR={1}'";
                    break;
                case ".xlsx": //Excel 07
                    conStr = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};
                         Extended Properties = 'Excel 8.0;HDR={1}'";
                    break;
            }
            conStr = String.Format(conStr, fullFilename, true);
            OleDbConnection connExcel = new OleDbConnection(conStr);
            OleDbCommand cmdExcel = new OleDbCommand();
            OleDbDataAdapter oda = new OleDbDataAdapter();
            DataTable dt = new DataTable();
            cmdExcel.Connection = connExcel;
            //Get the name of First Sheet
            connExcel.Open();
            DataTable dtExcelSchema;
            dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            string SheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
            //Read Data from First Sheet
            cmdExcel.CommandText = "SELECT * From [" + SheetName + "]";
            oda.SelectCommand = cmdExcel;
            oda.Fill(dt);
            connExcel.Close();

            return dt;
        }

        private DataTable GetFromWizardApprovedDate()
        {
            SqlConnection con = SyncHelper.WizardProductionConnection();
            SqlCommand com = new SqlCommand();
            SqlDataAdapter ad = new SqlDataAdapter(com);
            System.Text.StringBuilder sql = new System.Text.StringBuilder();
            sql.AppendLine("SELECT [RefNumber],[Status],[StatusDate]");
            sql.AppendLine("FROM [Production].[dbo].[tbRefStatusHist]");
            sql.AppendLine("where  YEAR(StatusDate) =" + DateTime.Now.Year + "  AND Status = 'PMA'");
            //sql.AppendLine("where Status = 'PMA'");

            com.CommandText = sql.ToString();
            com.CommandType = CommandType.Text;
            com.Connection = con;
            com.CommandTimeout = int.MaxValue;

            con.Open();
            try
            {
                DataTable dt = new DataTable();
                ad.Fill(dt);
                return dt;
            }
            catch (Exception ex)
            {
                throw;
            }
            finally
            {
                con.Close();
            }
        }

        public void UpdateApprovalLetterDate(Guid AccountID, Guid MSCStatusHistoryID, DateTime MOFApprovalDate)
        {
            SqlConnection con = SyncHelper.NewCRMConnection();
            SqlCommand com = new SqlCommand();

            System.Text.StringBuilder sql = new System.Text.StringBuilder();


            sql.AppendLine("UPDATE MSCStatusHistory SET");
            sql.AppendLine("MSCApprovalDate = @MSCApprovalDate,");
            sql.AppendLine("DataSource = @DataSource,");
            sql.AppendLine("ModifiedBy = @ModifiedBy,");
            sql.AppendLine("ModifiedByName = @ModifiedByName,");
            sql.AppendLine("ModifiedDate = @ModifiedDate");
            sql.AppendLine("WHERE MSCStatusHistoryID = @MSCStatusHistoryID");

            com.CommandText = sql.ToString();
            com.CommandType = CommandType.Text;
            com.Connection = con;
            com.CommandTimeout = int.MaxValue;

            com.Parameters.Add(new SqlParameter("@MSCStatusHistoryID", MSCStatusHistoryID));
            com.Parameters.Add(new SqlParameter("@MSCApprovalDate", MOFApprovalDate));
            com.Parameters.Add(new SqlParameter("@DataSource", BOL.Wizard.SyncHelper.DataSource));
            com.Parameters.Add(new SqlParameter("@ModifiedBy", SyncHelper.AdminID));
            com.Parameters.Add(new SqlParameter("@ModifiedByName", SyncHelper.AdminName));
            com.Parameters.Add(new SqlParameter("@ModifiedDate", DateTime.Now));

            try
            {
                BOL.AuditLog.Modules.AccountLog alMgr = new BOL.AuditLog.Modules.AccountLog();
                DataRow oldLogData = alMgr.SelectAccountForLog_MSCStatus(MSCStatusHistoryID);

                con.Open();
                com.ExecuteNonQuery();

                //Company Change Log - MSC History
                DataRow newLogData = alMgr.SelectAccountForLog_MSCStatus(MSCStatusHistoryID);
                //alMgr.CreateAccountLog_MSCStatusHistory(MSCStatusHistoryID, AccountID, Nothing, newLogData, New Guid(SyncHelper.AdminID), SyncHelper.AdminName)
                alMgr.CreateAccountLog_MSCStatusHistory(MSCStatusHistoryID, AccountID, oldLogData, newLogData, new Guid(SyncHelper.AdminID), SyncHelper.AdminName);

            }
            catch (Exception ex)
            {
                throw;
            }
            finally
            {
                con.Close();
            }
        }
        public Guid? GetMSCStatusHistoryID(Guid AccountID, Guid MSCApprovalStatusCID)
        {
            Guid output = new Guid();
            SqlConnection con = SyncHelper.NewCRMConnection();
            SqlCommand com = new SqlCommand();
            SqlDataAdapter ad = new SqlDataAdapter(com);

            System.Text.StringBuilder sql = new System.Text.StringBuilder();
            sql.AppendLine("SELECT MSCStatusHistoryID");
            sql.AppendLine("FROM MSCStatusHistory");
            sql.AppendLine("WHERE AccountID = @AccountID");
            sql.AppendLine("AND MSCApprovalStatusCID = @MSCApprovalStatusCID");

            com.CommandText = sql.ToString();
            com.CommandType = CommandType.Text;
            com.Connection = con;
            com.CommandTimeout = int.MaxValue;

            con.Open();
            try
            {
                com.Parameters.Add(new SqlParameter("@AccountID", AccountID));
                com.Parameters.Add(new SqlParameter("@MSCApprovalStatusCID", MSCApprovalStatusCID));

                DataTable dt = new DataTable();
                ad.Fill(dt);

                if (dt.Rows.Count > 0)
                {
                    output = new Guid(dt.Rows[0][0].ToString());
                }
                else
                {
                    return null;
                }
            }
            catch (Exception ex)
            {
                throw;
            }
            finally
            {
                con.Close();
            }
            return output;
        }
        private bool IsDifferentApprovalDate(Guid? MSCStatusHistoryID, string ApprovalLetterDate)
        {
            DateTime dApprovalLetterDate = Convert.ToDateTime(ApprovalLetterDate);
            string sApprovalLetterDate = dApprovalLetterDate.ToString("dd/MM/yyyy");
            SqlConnection con = SyncHelper.NewCRMConnection();
            SqlCommand com = new SqlCommand();
            SqlDataAdapter ad = new SqlDataAdapter(com);

            System.Text.StringBuilder sql = new System.Text.StringBuilder();

            sql.AppendLine("SELECT * ");
            sql.AppendLine("FROM MSCStatusHistory ");
            sql.AppendLine("WHERE MSCStatusHistoryID = @MSCStatusHistoryID ");
            sql.AppendLine("AND CONVERT(VARCHAR(12),MSCApprovalDate,103) = @MSCApprovalLetterDate ");

            com.CommandText = sql.ToString();
            com.CommandType = CommandType.Text;
            com.Connection = con;
            com.CommandTimeout = int.MaxValue;

            con.Open();
            try
            {
                com.Parameters.Add(new SqlParameter("@MSCStatusHistoryID", MSCStatusHistoryID));
                com.Parameters.Add(new SqlParameter("@MSCApprovalLetterDate", sApprovalLetterDate));

                DataTable dt = new DataTable();
                ad.Fill(dt);

                if (dt.Rows.Count == 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                throw;
            }
            finally
            {
                con.Close();
            }
        }
        public static Guid? GetAccountIDByFileID(string FileID)
        {
            Guid output = new Guid();
            SqlConnection con = SyncHelper.NewCRMConnection();
            SqlCommand com = new SqlCommand();
            SqlDataAdapter ad = new SqlDataAdapter(com);

            System.Text.StringBuilder sql = new System.Text.StringBuilder();
            sql.AppendLine("SELECT AccountID FROM Account WHERE MSCFileID = @FileID");

            com.CommandText = sql.ToString();
            com.CommandType = CommandType.Text;
            com.Connection = con;
            com.CommandTimeout = int.MaxValue;

            //con.Open()
            try
            {
                com.Parameters.Add(new SqlParameter("@FileID", FileID));

                DataTable dt = new DataTable();
                ad.Fill(dt);

                if (dt.Rows.Count > 0)
                {
                    output = new Guid(dt.Rows[0][0].ToString());
                }
                else
                {
                    return null;
                }
            }
            catch (Exception ex)
            {
                throw;
            }
            finally
            {
                con.Close();
            }
            return output;
        }
        public void CreateApprovalLetterDate(Guid AccountID, Guid? MSCApprovalStatusCID, DateTime MOFApprovalDate)
        {
            SqlConnection con = SyncHelper.NewCRMConnection();
            SqlCommand com = new SqlCommand();

            System.Text.StringBuilder sql = new System.Text.StringBuilder();
            sql.AppendLine("INSERT INTO MSCStatusHistory (");
            sql.AppendLine("MSCStatusHistoryID, ");
            sql.AppendLine("AccountID,");
            sql.AppendLine("MSCApprovalStatusCID,");
            sql.AppendLine("MSCApprovalDate,");
            sql.AppendLine("DataSource,");
            sql.AppendLine("CreatedBy,");
            sql.AppendLine("CreatedByName,");
            sql.AppendLine("CreatedDate,");
            sql.AppendLine("ModifiedBy,");
            sql.AppendLine("ModifiedByName,");
            sql.AppendLine("ModifiedDate");
            sql.AppendLine(")");
            sql.AppendLine("VALUES (");
            sql.AppendLine("@MSCStatusHistoryID, ");
            sql.AppendLine("@AccountID,");
            sql.AppendLine("@MSCApprovalStatusCID,");
            sql.AppendLine("@MSCApprovalDate,");
            sql.AppendLine("@DataSource,");
            sql.AppendLine("@CreatedBy,");
            sql.AppendLine("@CreatedByName,");
            sql.AppendLine("@CreatedDate,");
            sql.AppendLine("@CreatedBy,");
            sql.AppendLine("@CreatedByName,");
            sql.AppendLine("@CreatedDate");
            sql.AppendLine(")");
            com.CommandText = sql.ToString();
            com.CommandType = CommandType.Text;
            com.Connection = con;
            com.CommandTimeout = int.MaxValue;

            Guid MSCStatusHistoryID = Guid.NewGuid();
            com.Parameters.Add(new SqlParameter("@MSCStatusHistoryID", MSCStatusHistoryID));
            com.Parameters.Add(new SqlParameter("@AccountID", AccountID));
            com.Parameters.Add(new SqlParameter("@MSCApprovalStatusCID", MSCApprovalStatusCID));
            com.Parameters.Add(new SqlParameter("@MSCApprovalDate", MOFApprovalDate));
            com.Parameters.Add(new SqlParameter("@DataSource", BOL.Wizard.SyncHelper.DataSource));
            com.Parameters.Add(new SqlParameter("@CreatedBy", SyncHelper.AdminID));
            com.Parameters.Add(new SqlParameter("@CreatedByName", SyncHelper.AdminName));
            com.Parameters.Add(new SqlParameter("@CreatedDate", DateTime.Now));

            try
            {
                con.Open();
                com.ExecuteNonQuery();
                //Company Change Log - MSC History
                BOL.AuditLog.Modules.AccountLog alMgr = new BOL.AuditLog.Modules.AccountLog();
                DataRow newLogData = alMgr.SelectAccountForLog_MSCStatus(MSCStatusHistoryID);
                alMgr.CreateAccountLog_MSCStatusHistory(MSCStatusHistoryID, AccountID, null, newLogData, new Guid(SyncHelper.AdminID), SyncHelper.AdminName);
            }
            catch (Exception ex)
            {
                throw;
            }
            finally
            {
                con.Close();
            }
        }
        public static bool IsNonMSC(Guid AccountID)
        {
            Nullable<Guid> NonMSCAccountTypeCID = SyncHelper.GetCodeMasterID("Non MSC", BOL.AppConst.CodeType.AccountType);
            if (NonMSCAccountTypeCID.HasValue)
            {
                Guid accountTypeCID = GetAccountTypeCID(AccountID);
                return accountTypeCID == NonMSCAccountTypeCID.Value;
            }
            else
            {
                return false;
            }
        }
        private static Guid GetAccountTypeCID(Guid AccountID)
        {
            SqlConnection con = SyncHelper.NewCRMConnection();
            SqlCommand com = new SqlCommand();
            SqlDataAdapter ad = new SqlDataAdapter(com);

            System.Text.StringBuilder sql = new System.Text.StringBuilder();
            sql.AppendLine("SELECT AccountTypeCID FROM Account WHERE AccountID = @AccountID");

            com.CommandText = sql.ToString();
            com.CommandType = CommandType.Text;
            com.Connection = con;
            com.CommandTimeout = int.MaxValue;

            try
            {
                con.Open();
                com.Parameters.Add(new SqlParameter("@AccountID", AccountID));

                DataTable dt = new DataTable();
                ad.Fill(dt);
                return new Guid(dt.Rows[0][0].ToString());
            }
            catch (Exception ex)
            {
                throw;
            }
            finally
            {
                con.Close();
            }
        }

        private DataTable GetWizardApprovedDate()
        {
            SqlConnection con = SyncHelper.NewWizardConnection();
            SqlCommand com = new SqlCommand();
            SqlDataAdapter ad = new SqlDataAdapter(com);
            System.Text.StringBuilder sql = new System.Text.StringBuilder();
            sql.AppendLine("SELECT FileID, CONVERT(nvarchar,AppLtrDate,103) AS 'AppLtrDate' ");
            //* ModifiedDate refer as MSCApprovalDate
            sql.AppendLine("FROM Wizard_EIRAppLtrDatesView");

            com.CommandText = sql.ToString();
            com.CommandType = CommandType.Text;
            com.Connection = con;
            com.CommandTimeout = int.MaxValue;

            con.Open();
            try
            {
                DataTable dt = new DataTable();
                ad.Fill(dt);
                return dt;
            }
            catch (Exception ex)
            {
                throw;
            }
            finally
            {
                con.Close();
            }
        }
    }
}
