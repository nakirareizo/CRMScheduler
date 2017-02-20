using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CRMSync.Classes.MainProcess
{
    class SyncShareHolder
    {
        internal static void StartSync()
        {
            string sSource = null;
            string sLog = null;
            string sMachine = null;

            sSource = "Wizard Sync";
            sLog = "Application";
            sMachine = ".";

            using (SqlConnection con = SQLHelper.GetConnection())
            {
                con.Open();
                SqlTransaction trx = null;
                trx = con.BeginTransaction("ShareHolder");

                try
                {
                    Console.WriteLine(string.Format("[{0}] : Start Sync ShareHolder", DateTime.Now.ToString()));
                    LogFileHelper.logList.Add(string.Format("[{0}] : Start Sync ShareHolder", DateTime.Now.ToString()));

                    DataTable wizardData = GetWizardShareHolder(con, trx);

                    LogFileHelper.logList.Add("Successfully get WizardData");
                    //    DateTime DateNow = System.DateTime.Now;
                    int totalRecord = wizardData.Rows.Count;
                    Console.WriteLine(string.Format("ShareHolder Data Counted : [{0}] ", totalRecord.ToString()));
                    LogFileHelper.logList.Add(string.Format("ShareHolder Data Counted : [{0}] ", totalRecord.ToString()));
                    int count = 0;
                    foreach (System.Data.DataRow row in wizardData.Rows)
                    {
                        count += 1;

                        string FileID = row["FileID"].ToString();
                        Console.WriteLine(string.Format("FileID : [{0}]", FileID.ToString()));
                        LogFileHelper.logList.Add(string.Format("FileID : [{0}]", FileID.ToString()));
                        string ShareHolderName = row["ShareHolderName"].ToString().Trim();
                        double? Percentage = SyncHelper.ConvertToDouble(row["Percentage"].ToString());
                        int BumiShare = 0;
                        if (Convert.ToString(row["BumiShare"]).Trim().Equals("True"))
                        {
                            BumiShare = 1;
                        }
                        Guid? CountryRegionID = GetCountryRegionID(con, trx, Convert.ToString(row["RegionName"]).Trim());
                        Console.WriteLine(string.Format("CountryRegionID : [{0}]", CountryRegionID.ToString()));
                        LogFileHelper.logList.Add(string.Format("CountryRegionID : [{0}]", CountryRegionID.ToString()));
                        Nullable<Guid> AccountID = SyncHelper.GetAccountIDByFileID(FileID);

                        if (AccountID.HasValue)
                        {
                            Console.WriteLine(string.Format("AccountID : [{0}]", AccountID.ToString()));
                            LogFileHelper.logList.Add(string.Format("AccountID : [{0}]", AccountID.ToString()));
                            if (SyncHelper.IsNonMSC(AccountID))
                            {
                                continue;
                            }
                            Nullable<Guid> ShareHolderID = GetShareHolderID(con, trx, AccountID, ShareHolderName);

                            if (ShareHolderID.HasValue)
                            {
                                UpdateShareHolder(con, trx, AccountID, ShareHolderID, ShareHolderName, Percentage, BumiShare, CountryRegionID);
                                Console.WriteLine(string.Format("[{0}] {2}/{3} : Update ShareHolder FileID {1}", DateTime.Now.ToString(), FileID, count, totalRecord));
                                LogFileHelper.logList.Add(string.Format("[{0}] {2}/{3} : Update ShareHolder FileID {1}", DateTime.Now.ToString(), FileID, count, totalRecord));
                            }
                            else
                            {
                                CreateShareHolder(con, trx, AccountID, ShareHolderName, Percentage, BumiShare, CountryRegionID, DateTime.Now);
                                Console.WriteLine(string.Format("[{0}] {2}/{3} : Create ShareHolder FileID {1}", DateTime.Now.ToString(), FileID, count, totalRecord));
                                LogFileHelper.logList.Add(string.Format("[{0}] {2}/{3} : Create ShareHolder FileID {1}", DateTime.Now.ToString(), FileID, count, totalRecord));
                            }

                            //Calculation based on Shareholder
                            UpdateAccountJVCategory(con, trx, AccountID, new Guid(SyncHelper.AdminID), SyncHelper.AdminName);
                            Console.WriteLine(string.Format("UpdateAccountJVCategory succeed, AccountID : [{0}]", AccountID.ToString()));
                            LogFileHelper.logList.Add(string.Format("UpdateAccountJVCategory succeed, AccountID : [{0}]", AccountID.ToString()));
                            UpdateAccountBumiClassification(con, trx, AccountID, new Guid(SyncHelper.AdminID), SyncHelper.AdminName);
                            Console.WriteLine(string.Format("UpdateAccountBumiClassification succeed, AccountID : [{0}]", AccountID.ToString()));
                            LogFileHelper.logList.Add(string.Format("UpdateAccountBumiClassification succeed, AccountID : [{0}]", AccountID.ToString()));
                            UpdateAccountClassification(con, trx, AccountID, new Guid(SyncHelper.AdminID), SyncHelper.AdminName);
                            Console.WriteLine(string.Format("UpdateAccountClassification succeed, AccountID : [{0}]", AccountID.ToString()));
                            LogFileHelper.logList.Add(string.Format("UpdateAccountClassification succeed, AccountID : [{0}]", AccountID.ToString()));
                        }
                        else
                        {
                            Console.WriteLine(string.Format("[{0}] {2}/{3} : Record FileID {1} not found", DateTime.Now.ToString(), FileID, count, totalRecord));
                            LogFileHelper.logList.Add(string.Format("[{0}] {2}/{3} : Record FileID {1} not found", DateTime.Now.ToString(), FileID, count, totalRecord));
                        }
                    }

                    trx.Commit();

                    Console.WriteLine(string.Format("[{0}] : End Sync ShareHolder", DateTime.Now.ToString()));
                    LogFileHelper.logList.Add(string.Format("[{0}] : End Sync ShareHolder", DateTime.Now.ToString()));
                }
                catch (Exception ex)
                {
                    trx.Rollback();
                    LogFileHelper.logList.Add("ROLLBACK, ERROR: " + ex.Message);
                }

                con.Close();
            }
        }

        public static int UpdateAccountJVCategory(SqlConnection Connection, SqlTransaction Transaction, Guid? AccountID, Guid ActionBy, string ActionByName)
        {
            int affectedRows = 0;

            Nullable<Guid> currentJVCategoryCID = GetJVCategoryCID(AccountID);
            Nullable<Guid> JVCategoryCID = CalculateJVCategory(AccountID);

            if (!currentJVCategoryCID.Equals(JVCategoryCID))
            {
                using (SqlCommand cmd = new SqlCommand("", Connection))
                {
                    StringBuilder sql = new StringBuilder();
                    sql.AppendLine("UPDATE Account");
                    sql.AppendLine("SET JVCategoryCID = @JVCategoryCID, ModifiedDate = getdate(), ModifiedBy = @ActionBy, ModifiedByName = @ActionByName");
                    sql.AppendLine("WHERE AccountID = @AccountID");
                    cmd.CommandText = sql.ToString();
                    cmd.Parameters.AddWithValue("@AccountID", AccountID);
                    cmd.Parameters.AddWithValue("@JVCategoryCID", ConvertDbNull(JVCategoryCID));
                    cmd.Parameters.AddWithValue("@ActionBy", ActionBy);
                    cmd.Parameters.AddWithValue("@ActionByName", ActionByName);
                    affectedRows += cmd.ExecuteNonQuery();
                }
            }

            return affectedRows;
        }

        public static Nullable<Guid> GetJVCategoryCID(Guid? AccountID)
        {
            using (SqlConnection conn = new SqlConnection(ConfigurationSettings.AppSettings["CRM"].ToString()))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("", conn))
                {
                    using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                    {
                        StringBuilder sql = new StringBuilder();
                        sql.AppendLine("SELECT JVCategoryCID");
                        sql.AppendLine("FROM Account");
                        sql.AppendLine("WHERE AccountID = @AccountID");
                        cmd.CommandText = sql.ToString();
                        cmd.Parameters.AddWithValue("@AccountID", AccountID);
                        DataTable dt = new DataTable();
                        adapter.Fill(dt);
                        if (dt != null && dt.Rows.Count > 0 && !string.IsNullOrEmpty(dt.Rows[0][0].ToString()))
                        {
                            return new Guid(dt.Rows[0][0].ToString());
                        }
                        else
                        {
                            return null;
                        }
                    }
                }
            }
        }

        public static Nullable<Guid> CalculateJVCategory(Guid? AccountID)
        {
            Nullable<Guid> JVCategoryCID = null;
            double foreignSharePercentage = GetForeignCountryShareholderPercentage(AccountID);
            double localSharePercentage = GetLocalCountryShareholderPercentage(AccountID);
            CodeMaster mgr = new CodeMaster();

            if (localSharePercentage == 100 && foreignSharePercentage == 0)
            {
                JVCategoryCID = mgr.GetCodeMasterIDWithNull(BOL.AppConst.CodeType.JVCategory, "100% Malaysia");
            }
            else if (localSharePercentage == 0 && foreignSharePercentage == 100)
            {
                JVCategoryCID = mgr.GetCodeMasterIDWithNull(BOL.AppConst.CodeType.JVCategory, "100% Foreign");
            }
            else if (localSharePercentage == 50 && foreignSharePercentage == 50)
            {
                JVCategoryCID = mgr.GetCodeMasterIDWithNull(BOL.AppConst.CodeType.JVCategory, "50/50");
            }
            else if (localSharePercentage < foreignSharePercentage)
            {
                JVCategoryCID = mgr.GetCodeMasterIDWithNull(BOL.AppConst.CodeType.JVCategory, "Majority Foreign");
            }
            else if (localSharePercentage > foreignSharePercentage)
            {
                JVCategoryCID = mgr.GetCodeMasterIDWithNull(BOL.AppConst.CodeType.JVCategory, "Majority Local");
            }

            return JVCategoryCID;
        }

        private static double GetLocalCountryShareholderPercentage(Guid? AccountID)
        {
            using (SqlConnection conn = new SqlConnection(ConfigurationSettings.AppSettings["CRM"].ToString()))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("", conn))
                {
                    using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                    {
                        StringBuilder sql = new StringBuilder();
                        sql.AppendLine("SELECT ISNULL(SUM(ISNULL(s.Percentage, 0)), 0)");
                        sql.AppendLine("FROM Shareholder s");
                        sql.AppendLine("INNER JOIN Region Country ON Country.RegionID = s.CountryRegionID");
                        sql.AppendLine("AND Country.RegionName = 'Malaysia'");
                        sql.AppendLine("WHERE AccountID = @AccountID");
                        sql.AppendLine("AND [Status] = 1");
                        cmd.CommandText = sql.ToString();
                        cmd.Parameters.AddWithValue("@AccountID", AccountID);
                        DataTable dt = new DataTable();
                        adapter.Fill(dt);
                        if (dt != null && dt.Rows.Count > 0)
                        {
                            return Convert.ToDouble(dt.Rows[0][0]);
                        }
                        else
                        {
                            return 0.0;
                        }
                    }
                }
            }
        }

        public static double GetForeignCountryShareholderPercentage(Guid? AccountID)
        {
            using (SqlConnection conn = SQLHelper.GetConnection())
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("", conn))
                {
                    using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                    {
                        StringBuilder sql = new StringBuilder();
                        sql.AppendLine("SELECT ISNULL(SUM(ISNULL(s.Percentage, 0)), 0)");
                        sql.AppendLine("FROM Shareholder s");
                        sql.AppendLine("LEFT JOIN Region Country ON Country.RegionID = s.CountryRegionID");
                        sql.AppendLine("WHERE AccountID = @AccountID");
                        sql.AppendLine("AND [Status] = 1");
                        sql.AppendLine("AND (Country.RegionName <> 'Malaysia' OR Country.RegionName IS NULL)");
                        cmd.CommandText = sql.ToString();
                        cmd.Parameters.AddWithValue("@AccountID", AccountID);
                        DataTable dt = new DataTable();
                        adapter.Fill(dt);
                        if (dt != null && dt.Rows.Count > 0)
                        {
                            return Convert.ToDouble(dt.Rows[0][0]);
                        }
                        else {
                            return 0.0;
                        }
                    }
                }
            }
        }

        private static string GetCountryName(string v)
        {
            throw new NotImplementedException();
        }

        public static int UpdateAccountBumiClassification(SqlConnection Connection, SqlTransaction Transaction, Guid? AccountID, Guid ActionBy, string ActionByName)
        {
            int affectedRows = 0;

            Nullable<Guid> currentBumiClassificationCID = GetBumiClassificationCID(AccountID);
            Nullable<Guid> BumiClassificationCID = CalculateBumiClassification(AccountID);

            if (!currentBumiClassificationCID.Equals(BumiClassificationCID))
            {
                using (SqlConnection conn = new SqlConnection(ConfigurationSettings.AppSettings["CRM"].ToString()))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand("", conn))
                    {
                        StringBuilder sql = new StringBuilder();
                        sql.AppendLine("UPDATE Account");
                        sql.AppendLine("SET BumiClassificationCID = @BumiClassificationCID, ModifiedDate = getdate(), ModifiedBy = @ActionBy, ModifiedByName = @ActionByName");
                        sql.AppendLine("WHERE AccountID = @AccountID");
                        cmd.CommandText = sql.ToString();
                        cmd.Parameters.AddWithValue("@AccountID", AccountID);
                        cmd.Parameters.AddWithValue("@BumiClassificationCID", ConvertDbNull(BumiClassificationCID));
                        cmd.Parameters.AddWithValue("@ActionBy", ActionBy);
                        cmd.Parameters.AddWithValue("@ActionByName", ActionByName);
                        affectedRows += cmd.ExecuteNonQuery();
                    }
                }
            }

            return affectedRows;
        }

        public static Nullable<Guid> GetBumiClassificationCID(Guid? AccountID)
        {
            using (SqlConnection conn = new SqlConnection(ConfigurationSettings.AppSettings["CRM"].ToString()))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("", conn))
                {
                    using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                    {
                        StringBuilder sql = new StringBuilder();
                        sql.AppendLine("SELECT BumiClassificationCID");
                        sql.AppendLine("FROM Account");
                        sql.AppendLine("WHERE AccountID = @AccountID");
                        cmd.CommandText = sql.ToString();
                        cmd.Parameters.AddWithValue("@AccountID", AccountID);
                        DataTable dt = new DataTable();
                        adapter.Fill(dt);
                        if (dt != null && dt.Rows.Count > 0 && !string.IsNullOrEmpty(dt.Rows[0][0].ToString()))
                        {
                            return new Guid(dt.Rows[0][0].ToString());
                        }
                        else
                        {
                            return null;
                        }
                    }
                }
            }
        }
        public static Guid? CalculateBumiClassification(Guid? AccountID)
        {
            try
            {
                Nullable<Guid> BumiClassificationCID = null;
                decimal bumiSharePercentage = GetBumiShareholderPercentage(AccountID);
                CodeMaster mgr = new CodeMaster();

                if (bumiSharePercentage == 0)
                {
                    BumiClassificationCID = mgr.GetCodeMasterIDWithNull(BOL.AppConst.CodeType.BumiClassification, "Others");
                }
                else if (bumiSharePercentage > 50)
                {
                    BumiClassificationCID = mgr.GetCodeMasterIDWithNull(BOL.AppConst.CodeType.BumiClassification, "Majority Bumi");
                }
                else if (bumiSharePercentage <= 50)
                {
                    BumiClassificationCID = mgr.GetCodeMasterIDWithNull(BOL.AppConst.CodeType.BumiClassification, "Bumi participation");
                }

                return BumiClassificationCID;
            }
            catch (Exception ex)
            {
                //MsgBox(ex.Message, MsgBoxStyle.OkOnly, "CalculateBumiClassification")
                return null;
            }
        }

        private static int GetBumiShareholderPercentage(Guid? AccountID)
        {
            using (SqlConnection conn = new SqlConnection(ConfigurationSettings.AppSettings["CRM"].ToString()))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("", conn))
                {
                    using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                    {
                        StringBuilder sql = new StringBuilder();
                        sql.AppendLine("SELECT ISNULL(SUM(ISNULL(s.Percentage, 0)), 0)");
                        sql.AppendLine("FROM Shareholder s");
                        sql.AppendLine("WHERE AccountID = @AccountID");
                        sql.AppendLine("AND [Status] = 1");
                        sql.AppendLine("AND s.BumiShare = 1");
                        cmd.CommandText = sql.ToString();
                        cmd.Parameters.AddWithValue("@AccountID", AccountID);
                        DataTable dt = new DataTable();
                        adapter.Fill(dt);
                        if (dt != null && dt.Rows.Count > 0)
                        {
                            return Convert.ToInt32(dt.Rows[0][0]);
                        }
                        else
                        {
                            return 0;
                        }
                    }
                }
            }
        }

        private static object ConvertDbNull(object obj)
        {
            return obj == null ? DBNull.Value : obj;
        }
        public static void CreateShareHolder(SqlConnection Connection, SqlTransaction Transaction, Guid? AccountID, string ShareHolderName, double? Percentage, int BumiShare, Guid? CountryRegionID, DateTime DateNow)
        {
            using (SqlCommand com = new SqlCommand("", Connection))
            {
                System.Text.StringBuilder sql = new System.Text.StringBuilder();
                //sql.AppendLine("DELETE FROM ShareHolder")
                //sql.AppendLine("WHERE AccountID = @AccountID ")
                //sql.AppendLine("AND ModifiedDate <> @DateNow ")
                sql.AppendLine("INSERT INTO ShareHolder (");
                sql.AppendLine("ShareHolderID, ");
                sql.AppendLine("AccountID,");
                sql.AppendLine("ShareHolderName,");
                sql.AppendLine("Percentage,");
                sql.AppendLine("BumiShare,");
                sql.AppendLine("CountryRegionID,");
                sql.AppendLine("CreatedBy,");
                sql.AppendLine("CreatedByName,");
                sql.AppendLine("CreatedDate,");
                sql.AppendLine("ModifiedBy,");
                sql.AppendLine("ModifiedByName,");
                sql.AppendLine("ModifiedDate,");
                sql.AppendLine("Status");
                sql.AppendLine(")");
                sql.AppendLine("VALUES (");
                sql.AppendLine("@ShareHolderID, ");
                sql.AppendLine("@AccountID,");
                sql.AppendLine("@ShareHolderName,");
                sql.AppendLine("@Percentage,");
                sql.AppendLine("@BumiShare,");
                sql.AppendLine("@CountryRegionID,");
                sql.AppendLine("@CreatedBy,");
                sql.AppendLine("@CreatedByName,");
                sql.AppendLine("@DateNow,");
                sql.AppendLine("@CreatedBy,");
                sql.AppendLine("@CreatedByName,");
                sql.AppendLine("@DateNow,");
                sql.AppendLine("1");
                sql.AppendLine(")");

                com.CommandText = sql.ToString();
                com.CommandType = CommandType.Text;
                com.Connection = Connection;
                com.Transaction = Transaction;
                com.CommandTimeout = int.MaxValue;

                Guid ShareHolderID = Guid.NewGuid();
                com.Parameters.Add(new SqlParameter("@ShareHolderID", ShareHolderID));
                com.Parameters.Add(new SqlParameter("@AccountID", AccountID));
                com.Parameters.Add(new SqlParameter("@ShareHolderName", ShareHolderName));
                com.Parameters.Add(new SqlParameter("@Percentage", Percentage));
                com.Parameters.Add(new SqlParameter("@BumiShare", BumiShare));
                com.Parameters.Add(new SqlParameter("@DateNow", DateNow));
                com.Parameters.Add(new SqlParameter("@CountryRegionID", CountryRegionID));
                com.Parameters.Add(new SqlParameter("@CreatedBy", SyncHelper.AdminID));
                com.Parameters.Add(new SqlParameter("@CreatedByName", SyncHelper.AdminName));

                try
                {
                    //con.Open()
                    com.ExecuteNonQuery();

                    //Company Change Log - MSC History
                    BOL.AuditLog.Modules.AccountLog alMgr = new BOL.AuditLog.Modules.AccountLog();
                    DataRow newLogData = alMgr.SelectAccountForLog_MSCStatus_Wizard(Connection, Transaction, ShareHolderID);
                    alMgr.CreateAccountLog_Shareholder_Wizard(Connection, Transaction, ShareHolderID, AccountID.Value, null, newLogData, new Guid(SyncHelper.AdminID), SyncHelper.AdminName);

                }
                catch (Exception ex)
                {
                    throw;
                    //Finally
                    //	con.Close()
                }
            }
        }


        private static void UpdateShareHolder(SqlConnection Connection, SqlTransaction Transaction, Guid? AccountID, Guid? ShareHolderID, string ShareHolderName, double? Percentage, int BumiShare, Guid? CountryRegionID)
        {
            using (SqlCommand com = new SqlCommand("", Connection))
            {

                System.Text.StringBuilder sql = new System.Text.StringBuilder();
                sql.AppendLine("UPDATE ShareHolder SET");
                sql.AppendLine("ShareHolderName = @ShareHolderName,");
                sql.AppendLine("Percentage = @Percentage,");
                sql.AppendLine("BumiShare = @BumiShare,");
                sql.AppendLine("Status = 1,");
                sql.AppendLine("CountryRegionID = @CountryRegionID,");
                sql.AppendLine("ModifiedBy = @ModifiedBy,");
                sql.AppendLine("ModifiedByName = @ModifiedByName,");
                sql.AppendLine("ModifiedDate = GETDATE()");
                sql.AppendLine("WHERE ShareHolderID = @ShareHolderID");

                com.CommandText = sql.ToString();
                com.CommandType = CommandType.Text;
                com.Connection = Connection;
                com.Transaction = Transaction;
                com.CommandTimeout = int.MaxValue;

                com.Parameters.Add(new SqlParameter("@ShareHolderID", ShareHolderID));
                com.Parameters.Add(new SqlParameter("@ShareHolderName", ShareHolderName));
                com.Parameters.Add(new SqlParameter("@Percentage", Percentage));
                com.Parameters.Add(new SqlParameter("@BumiShare", BumiShare));
                com.Parameters.Add(new SqlParameter("@CountryRegionID", CountryRegionID));
                com.Parameters.Add(new SqlParameter("@ModifiedBy", SyncHelper.AdminID));
                com.Parameters.Add(new SqlParameter("@ModifiedByName", SyncHelper.AdminName));

                try
                {
                    //con.Open(
                    com.ExecuteNonQuery();

                    //Company Change Log - MSC History
                    BOL.AuditLog.Modules.AccountLog alMgr = new BOL.AuditLog.Modules.AccountLog();
                    DataRow newLogData = alMgr.SelectAccountForLog_MSCStatus_Wizard(Connection, Transaction, ShareHolderID.Value);
                    alMgr.CreateAccountLog_Shareholder_Wizard(Connection, Transaction, ShareHolderID.Value, AccountID.Value, null, newLogData, new Guid(SyncHelper.AdminID), SyncHelper.AdminName);

                }
                catch (Exception ex)
                {
                    throw;

                }
            }
        }

        private static Guid? GetShareHolderID(SqlConnection Connection, SqlTransaction Transaction, Guid? AccountID, string ShareHolderName)
        {
            using (SqlCommand com = new SqlCommand("", Connection))
            {
                using (SqlDataAdapter ad = new SqlDataAdapter(com))
                {

                    System.Text.StringBuilder sql = new System.Text.StringBuilder();
                    sql.AppendLine("SELECT ShareHolderID");
                    sql.AppendLine("FROM ShareHolder");
                    sql.AppendLine("WHERE AccountID = @AccountID");
                    sql.AppendLine("AND ShareHolderName = @ShareHolderName");

                    com.CommandText = sql.ToString();
                    com.CommandType = CommandType.Text;
                    com.Connection = Connection;
                    com.Transaction = Transaction;
                    com.CommandTimeout = int.MaxValue;

                    //con.Open()
                    try
                    {
                        com.Parameters.Add(new SqlParameter("@AccountID", AccountID));
                        com.Parameters.Add(new SqlParameter("@ShareHolderName", ShareHolderName));

                        DataTable dt = new DataTable();
                        ad.Fill(dt);

                        if (dt.Rows.Count > 0)
                        {
                            return new Guid(dt.Rows[0][0].ToString());
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
                        //con.Close()
                    }
                }
            }
        }

        private static Nullable<Guid> GetCountryRegionID(SqlConnection Connection, SqlTransaction Transaction, string RegionName)
        {
            Guid RegionTypeID = GetRegionTypeID(Connection, Transaction, "Country");
            return GetRegionID(Connection, Transaction, RegionTypeID, RegionName);
        }
        private static Guid GetRegionTypeID(SqlConnection Connection, SqlTransaction Transaction, string RegionType)
        {

            using (SqlCommand com = new SqlCommand("", Connection))
            {
                using (SqlDataAdapter ad = new SqlDataAdapter(com))
                {

                    System.Text.StringBuilder sql = new System.Text.StringBuilder();
                    sql.AppendLine("SELECT RegionTypeID FROM RegionType");
                    sql.AppendLine("WHERE RegionType = @RegionType");

                    com.CommandText = sql.ToString();
                    com.CommandType = CommandType.Text;
                    com.Connection = Connection;
                    com.Transaction = Transaction;
                    com.CommandTimeout = int.MaxValue;

                    try
                    {
                        com.Parameters.Add(new SqlParameter("@RegionType", RegionType));

                        DataTable dt = new DataTable();
                        ad.Fill(dt);

                        if (dt.Rows.Count > 0)
                        {
                            return new Guid(dt.Rows[0][0].ToString());
                        }
                        else
                        {
                            return Guid.Empty;
                        }
                    }
                    catch (Exception ex)
                    {
                        throw;
                    }
                    finally
                    {
                    }
                }
            }
        }
        private static Nullable<Guid> GetRegionID(SqlConnection Connection, SqlTransaction Transaction, Guid RegionTypeID, string RegionName)
        {

            using (SqlCommand com = new SqlCommand("", Connection))
            {
                using (SqlDataAdapter ad = new SqlDataAdapter(com))
                {

                    System.Text.StringBuilder sql = new System.Text.StringBuilder();
                    sql.AppendLine("SELECT RegionID FROM Region");
                    sql.AppendLine("WHERE RegionTypeID = @RegionTypeID");
                    sql.AppendLine("AND RegionName = @RegionName");

                    com.CommandText = sql.ToString();
                    com.CommandType = CommandType.Text;
                    com.Connection = Connection;
                    com.Transaction = Transaction;
                    com.CommandTimeout = int.MaxValue;

                    try
                    {
                        com.Parameters.Add(new SqlParameter("@RegionTypeID", RegionTypeID));
                        com.Parameters.Add(new SqlParameter("@RegionName", RegionName));

                        DataTable dt = new DataTable();
                        ad.Fill(dt);

                        if (dt.Rows.Count > 0)
                        {
                            return new Guid(dt.Rows[0][0].ToString());
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
                    }
                }
            }
        }

        private static DataTable GetWizardShareHolder(SqlConnection con, SqlTransaction trx)
        {
            using (SqlCommand com = new SqlCommand("", con))
            {
                using (SqlDataAdapter ad = new SqlDataAdapter(com))
                {
                    StringBuilder sql = new System.Text.StringBuilder();
                    sql.AppendLine("SELECT FileID, OwnerShipSHName AS ShareHolderName, OwnerShipPer AS Percentage, OwnerShipBumi AS BumiShare, OwnerShipCName AS RegionName");
                    sql.AppendLine("FROM IntegrationDB.dbo.EIR_PMSCOwnerShipDtls v");
                    //sql.AppendLine("WHERE UICStatus = 1")

                    com.CommandText = sql.ToString();
                    com.CommandType = CommandType.Text;
                    com.Transaction = trx;
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
                    }
                }
            }
        }

        public static int UpdateAccountClassification(SqlConnection Connection, SqlTransaction Transaction, Guid? AccountID, Guid ActionBy, string ActionByName)
        {
            int affectedRows = 0;

            Nullable<Guid> currentClassificationCID = GetClassificationCID(AccountID);
            Nullable<Guid> ClassificationCID = CalculateClassification(AccountID);

            if (!currentClassificationCID.Equals(ClassificationCID) && !ClassificationCID.Equals(Guid.Empty))
            {
                using (SqlConnection conn = new SqlConnection(ConfigurationSettings.AppSettings["CRM"].ToString()))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand("", conn))
                    {
                        StringBuilder sql = new StringBuilder();
                        sql.AppendLine("UPDATE Account");
                        sql.AppendLine("SET ClassificationCID = @ClassificationCID, ModifiedDate = getdate(), ModifiedBy = @ActionBy, ModifiedByName = @ActionByName");
                        sql.AppendLine("WHERE AccountID = @AccountID");
                        cmd.CommandText = sql.ToString();
                        cmd.Parameters.AddWithValue("@AccountID", AccountID);
                        cmd.Parameters.AddWithValue("@ClassificationCID", ConvertDbNull(ClassificationCID));
                        cmd.Parameters.AddWithValue("@ActionBy", ActionBy);
                        cmd.Parameters.AddWithValue("@ActionByName", ActionByName);
                        affectedRows += cmd.ExecuteNonQuery();
                    }
                }
            }

            return affectedRows;
        }

        public static Nullable<Guid> GetClassificationCID(Guid? AccountID)
        {
            using (SqlConnection conn = new SqlConnection(ConfigurationSettings.AppSettings["CRM"].ToString()))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("", conn))
                {
                    using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                    {
                        StringBuilder sql = new StringBuilder();
                        sql.AppendLine("SELECT ClassificationCID");
                        sql.AppendLine("FROM Account");
                        sql.AppendLine("WHERE AccountID = @AccountID");
                        cmd.CommandText = sql.ToString();
                        cmd.Parameters.AddWithValue("@AccountID", AccountID);
                        DataTable dt = new DataTable();
                        adapter.Fill(dt);
                        if (dt != null && dt.Rows.Count > 0 && !string.IsNullOrEmpty(dt.Rows[0][0].ToString()))
                        {
                            return new Guid(dt.Rows[0][0].ToString());
                        }
                        else
                        {
                            return null;
                        }
                    }
                }
            }
        }

        public static Guid? CalculateClassification(Guid? AccountID)
        {
            try
            {
                Nullable<Guid> ClassificationCID = null;
                double localSharePercentage = GetLocalCountryShareholderPercentage(AccountID);
                double foreignSharePercentage = GetForeignCountryShareholderPercentage(AccountID);
                CodeMaster mgr = new CodeMaster();

                if (localSharePercentage == 50 && foreignSharePercentage == 50)
                {
                    ClassificationCID = mgr.GetCodeMasterIDWithNull(BOL.AppConst.CodeType.Classification, "50/50");
                }
                else if (localSharePercentage > foreignSharePercentage)
                {
                    ClassificationCID = mgr.GetCodeMasterIDWithNull(BOL.AppConst.CodeType.Classification, "Malaysian Owned");
                }
                else if (localSharePercentage < foreignSharePercentage)
                {
                    ClassificationCID = mgr.GetCodeMasterIDWithNull(BOL.AppConst.CodeType.Classification, "Foreign Owned");
                }

                return ClassificationCID;
            }
            catch (Exception ex)
            {
                //MsgBox(ex.Message, MsgBoxStyle.OkOnly, "CalculateClassification")
                return null;
            }
        }
    }
}
