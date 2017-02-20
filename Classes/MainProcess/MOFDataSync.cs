using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CRMSync.Classes.MainProcess
{
    class MOFDataSync
    {
        internal static DataTable wizardData = new DataTable();
        
        internal static void Start(DataTable wizardData)
        {
            LogFileHelper.logList = new ArrayList();
            DateTime SyncDate = DateTime.Now;
            if (wizardData.Rows.Count > 0)
            {
                foreach (DataRow row in wizardData.Rows)
                {
                    string FileID = row["FileID"].ToString();
                    string SubmitType = row["SubmitType"].ToString();
                    string CompanyName = row["CompanyName"].ToString();
                    string ROCNumber = row["ROCNumber"].ToString();
                    string OperationalStatus= row["OperationalStatus"].ToString();
                    string CoreActivities = row["CoreActivities"].ToString();
                    switch (SubmitType)
                    {
                        case "S":
                            if (!CheckRecordExist(FileID, SubmitType))
                                InsertIntoMOFMaklumatSyarikat();
                            break;
                        case "A":
                        case "P":
                        case "E":
                        case "N":
                            break;

                    }
                }
            }
        }

        internal static void Start()
        {
            LogFileHelper.logList = new ArrayList();
            DateTime SyncDate = DateTime.Now;
            using (SqlConnection Connection = SQLHelper.GetConnection())
            {
                Console.WriteLine(string.Format("[{0}] : Start Sync MOF Sync Data..", DateTime.Now.ToString()));
                LogFileHelper.logList.Add(string.Format("[{0}] : Start Sync MOF Sync Data...", DateTime.Now.ToString()));
                try
                {
                    //READ FROM SPBigFIle:
                    string SyncedDate = "";
                    wizardData = SelectACApprovedAccountList(out SyncedDate);
                    if (wizardData.Rows.Count > 0)
                    {
                        foreach (DataRow row in wizardData.Rows)
                        {
                            string FileID = row["FileID"].ToString();
                            string SubmitType = row["SubmitType"].ToString();
                            string CompanyName = row["CompanyName"].ToString();
                            string ROCNumber = row["ROCNumber"].ToString();
                            string OperationalStatus = row["OperationalStatus"].ToString();
                            string CoreActivities = row["CoreActivities"].ToString();
                            string URL = row["URL"].ToString();
                            string DateofIncorporation=row["DateOfApproval"].ToString();
                            string YearOfApproval = row["Year"].ToString();
                            switch (SubmitType)
                            {
                                case "S":
                                    if (!CheckRecordExist(FileID, SubmitType))
                                        InsertIntoMOFMaklumatSyarikat();
                                    break;
                                case "A":
                                case "P":
                                case "E":
                                case "N":
                                    break;

                            }
                        }
                    }
                }
                catch
                {

                }
            }
        }
        private static DataTable SelectACApprovedAccountList(out string SyncedDate)
        {
            SqlConnection con = SQLHelper.GetConnection();
            SqlCommand com = new SqlCommand();
            SqlDataAdapter ad = new SqlDataAdapter(com);

            string LastSync = BOL.Common.Modules.Parameter.WIZARD_TMS;
            SyncedDate = LastSync;
            System.Text.StringBuilder sql = new System.Text.StringBuilder();
            sql.AppendLine(ConfigurationSettings.AppSettings["WizardStoredProc"].ToString()).Append(" '").Append(LastSync).Append("'");
            //sql.AppendLine(ConfigurationSettings.AppSettings["WizardStoredProc"].ToString()).Append(" '").Append("2016-12-15 00:00:00 AM").Append("'");
            //SyncedDate = "2016-12-15 00:00:00 AM";
            com.CommandText = sql.ToString();
            com.CommandType = CommandType.Text;
            com.Connection = con;
            com.CommandTimeout = int.MaxValue;

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
        private static bool CheckRecordExist(string fileID, string submitType)
        {
            throw new NotImplementedException();
        }

        private static void InsertIntoMOFMaklumatSyarikat()
        {
            using (SqlConnection Connection = SQLHelper.GetConnectionMOF())
            {
                try
                {
                    SqlCommand com = new SqlCommand();
                    System.Text.StringBuilder sql = new System.Text.StringBuilder();
                    sql.AppendLine("INSERT INTO AccountCluster ");
                    sql.AppendLine("(");
                    sql.AppendLine("[CompanyName],[MSCFileID],[CoreActivities],[ROCNumber],[URL],[ACMeetingDate],[ApprovalLetterDate]");
                    sql.AppendLine(",[YearOfApproval],[DateofIncorporation],[OperationalStatus] ,[FinancialIncentive] ,[Cluster],[BusinessPhone]");
                    sql.AppendLine(",[Fax],[TaxRevenueLoss],[Tier],[SubmitDate],[MSCCertNo],[BA],[SubmitType],[SyncDate]");
                    sql.AppendLine(")");
                    sql.AppendLine("VALUES(");
                    sql.AppendLine("@CompanyName, @MSCFileID, @CoreActivities, @ROCNumber,@URL,@ACMeetingDate,@ApprovalLetterDate");
                    sql.AppendLine("@YearOfApproval, @DateofIncorporation, @OperationalStatus, @FinancialIncentive, @Cluster, @BusinessPhone");
                    sql.AppendLine("@Fax, @TaxRevenueLoss, @Tier, @SubmitDate, @MSCCertNo, @BA,@SubmitType,@SyncDate");
                    sql.AppendLine(")");

                    com.Parameters.Add(new SqlParameter("@CompanyName", ""));
                    com.Parameters.Add(new SqlParameter("@MSCFileID", ""));
                    com.Parameters.Add(new SqlParameter("@CoreActivities", ""));
                    com.Parameters.Add(new SqlParameter("@ROCNumber",""));
                    com.Parameters.Add(new SqlParameter("@URL", ""));
                    com.Parameters.Add(new SqlParameter("@ACMeetingDate", ""));
                    com.Parameters.Add(new SqlParameter("@ApprovalLetterDate", ""));
                    com.Parameters.Add(new SqlParameter("@DateofIncorporation", ""));
                    com.Parameters.Add(new SqlParameter("@OperationalStatus", ""));
                    com.Parameters.Add(new SqlParameter("@FinancialIncentive", ""));
                    com.Parameters.Add(new SqlParameter("@Cluster", ""));
                    com.Parameters.Add(new SqlParameter("@BusinessPhone", ""));
                    com.Parameters.Add(new SqlParameter("@Fax", ""));
                    com.Parameters.Add(new SqlParameter("@TaxRevenueLoss", ""));
                    com.Parameters.Add(new SqlParameter("@Tier", ""));
                    com.Parameters.Add(new SqlParameter("@SubmitDate", ""));
                    com.Parameters.Add(new SqlParameter("@MSCCertNo", ""));
                    com.Parameters.Add(new SqlParameter("@BA", ""));
                    com.Parameters.Add(new SqlParameter("@SubmitType", ""));
                    com.Parameters.Add(new SqlParameter("@SyncDate", ""));

                    com.CommandText = sql.ToString();
                    com.CommandType = CommandType.Text;
                    com.Connection = Connection;
                    com.CommandTimeout = int.MaxValue;
                    com.ExecuteNonQuery();
                }
                catch (Exception)
                {

                    throw;
                }
            }
        }
    }
}
