using CRMSync.Classes;
using System;
using System.Collections;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
namespace CRMSync
{
    class ShareHolderCleanUp
    {
        internal static void Start()
        {
            string Filename0 = getFileName();
            DataTable Wizarddata = ExcelToDataTable(Filename0, "WizardData");
            using (SqlConnection Connection = SQLHelper.GetConnection())
            {

                SqlTransaction Transaction = default(SqlTransaction);
                Transaction = Connection.BeginTransaction("WizardSync");
                foreach (DataRow row in Wizarddata.Rows)
                {
                    string FileID = row["FileID"].ToString();
                    //1.Check whether shareholdername in View table existed in ShareHolder ?
                    DataTable dtShareHolder = GetShareHolder(Connection, Transaction, FileID);
                    //2, if existed in ShareHolder table
                    bool Existed = CheckDataExistedInShareHolder(Connection, Transaction, FileID);
                    if (Existed)
                    {
                        string ShareHolderName = dtShareHolder.Rows[0]["OwnershipSHName"].ToString();
                        Double OwnerShipPer = Convert.ToDouble(dtShareHolder.Rows[0]["OwnershipPer"].ToString());
                        //3. Get AccountID
                        Guid? AccountID = GetAccountIDByFileID(Connection, Transaction, FileID);
                        //4. Delete ShareHolder in ShareHolder table
                        DeleteShareholder(Connection, Transaction, AccountID, ShareHolderName);

                    }
                    //5.Check whether ShareHolder Name in View Table existed in ShareHolderDV, if not INSERT if yes just update

                    string ShareholderName = row["OwnershipSHName"].ToString();
                    Guid? AccountDVID = GetAccountDVIDIDByShareHolderName(Connection, Transaction, ShareholderName);
                    Nullable<decimal> Percentage = SyncHelper.ConvertToDecimal(row["OwnershipPer"]);
                    bool BumiShare = SyncHelper.ConvertToBoolean(row["OwnershipBumi"]);
                    Nullable<Guid> CountryRegionID = SyncHelper.GetRegionID(Connection, Transaction, row["OwnershipCName"].ToString());
                    CreateUpdateShareholderDV(Connection, Transaction, AccountDVID, ShareholderName, Percentage, BumiShare, CountryRegionID);
                }
                //Transaction.Commit();

            }
        }

        private static bool CheckDataExistedInShareHolder(SqlConnection Connection, SqlTransaction Transaction, string MSCFileID)
        {
            bool Existed = false;

            Guid? AccountID = GetAccountIDByFileID(Connection, Transaction, MSCFileID);
            SqlCommand com = new SqlCommand();
            SqlDataAdapter ad = new SqlDataAdapter(com);

            System.Text.StringBuilder sql = new System.Text.StringBuilder();
            sql.AppendLine("SELECT * FROM ShareHolder WHERE AccountID = @AccountID");

            com.CommandText = sql.ToString();
            com.CommandType = CommandType.Text;
            com.Connection = Connection;
            com.Transaction = Transaction;
            com.CommandTimeout = int.MaxValue;

            //con.Open()
            try
            {
                com.Parameters.Add(new SqlParameter("@AccountID", AccountID));

                DataTable dt = new DataTable();
                ad.Fill(dt);

                if (dt.Rows.Count > 0)
                {
                    Existed = true;
                }
                else {
                    Existed = false;
                }
            }
            catch (Exception ex)
            {
                throw;
            }
            return Existed;
        }

        private static Guid? GetAccountDVIDIDByShareHolderName(SqlConnection Connection, SqlTransaction Transaction, string ShareholderName)
        {
            SqlCommand com = new SqlCommand();
            SqlDataAdapter ad = new SqlDataAdapter(com);

            System.Text.StringBuilder sql = new System.Text.StringBuilder();
            sql.AppendLine("SELECT AccountDVID FROM ShareHolderDV WHERE UPPER(ShareholderName) = @ShareholderName");

            com.CommandText = sql.ToString();
            com.CommandType = CommandType.Text;
            com.Connection = Connection;
            com.Transaction = Transaction;
            com.CommandTimeout = int.MaxValue;

            //con.Open()
            try
            {
                com.Parameters.Add(new SqlParameter("@ShareholderName", ShareholderName.ToUpper()));

                DataTable dt = new DataTable();
                ad.Fill(dt);

                if (dt.Rows.Count > 0)
                {
                    return new Guid(dt.Rows[0][0].ToString());
                }
                else {
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
        private static void CreateUpdateShareholderDV(SqlConnection Connection, SqlTransaction Transaction, Guid? AccountDVID, string ShareholderName, Nullable<decimal> Percentage, bool BumiShare, Nullable<Guid> CountryRegionID)
        {
            if (!string.IsNullOrEmpty(ShareholderName))
            {
                SqlCommand com = new SqlCommand();
                StringBuilder sql = new StringBuilder();
                Guid? ShareHolderDVID = GetShareHolderDVID(Connection, Transaction, AccountDVID, ShareholderName);

                if (ShareHolderDVID.HasValue)
                {
                    try
                    {
                        sql.AppendLine("UPDATE ShareHolderDV SET ");
                        sql.AppendLine("Percentage = @Percentage,");
                        sql.AppendLine("BumiShare = @BumiShare,");
                        sql.AppendLine("Status = @Status,");
                        sql.AppendLine("CountryRegionID = @CountryRegionID");
                        sql.AppendLine("WHERE ShareHolderDVID = @ShareHolderDVID");

                        com.Parameters.Add(new SqlParameter("@ShareHolderDVID", ShareHolderDVID));
                        com.Parameters.Add(new SqlParameter("@Percentage", SyncHelper.ReturnNull(Percentage)));
                        com.Parameters.Add(new SqlParameter("@BumiShare", BumiShare));
                        com.Parameters.Add(new SqlParameter("@Status", EnumSync.Status.Active));
                        com.Parameters.Add(new SqlParameter("@CountryRegionID", SyncHelper.ReturnNull(CountryRegionID)));

                        com.CommandText = sql.ToString();
                        com.CommandType = CommandType.Text;
                        com.Connection = Connection;
                        com.Transaction = Transaction;
                        com.CommandTimeout = int.MaxValue;


                        //con.Open()
                        com.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {

                    }
                }
                else {
                    try
                    {
                        sql.AppendLine("INSERT INTO ShareHolderDV ");
                        sql.AppendLine("(ShareHolderDVID, AccountDVID, ShareHolderID, ShareholderName, Percentage, BumiShare, Status, CountryRegionID)");
                        sql.AppendLine("VALUES");
                        sql.AppendLine("(@ShareHolderDVID, @AccountDVID, NULL, @ShareholderName, @Percentage, @BumiShare, @Status, @CountryRegionID)");

                        ShareHolderDVID = Guid.NewGuid();
                        com.Parameters.Add(new SqlParameter("@ShareHolderDVID", ShareHolderDVID));
                        com.Parameters.Add(new SqlParameter("@AccountDVID", AccountDVID));
                        com.Parameters.Add(new SqlParameter("@ShareholderName", ShareholderName));
                        com.Parameters.Add(new SqlParameter("@Percentage", SyncHelper.ReturnNull(Percentage)));
                        com.Parameters.Add(new SqlParameter("@BumiShare", BumiShare));
                        com.Parameters.Add(new SqlParameter("@Status", EnumSync.Status.Active));
                        com.Parameters.Add(new SqlParameter("@CountryRegionID", SyncHelper.ReturnNull(CountryRegionID)));

                        com.CommandText = sql.ToString();
                        com.CommandType = CommandType.Text;
                        com.Connection = Connection;
                        com.Transaction = Transaction;
                        com.CommandTimeout = int.MaxValue;


                        //con.Open()
                        com.ExecuteNonQuery();

                    }
                    catch (Exception ex)
                    {
                        throw;
                        //Finally
                        //	con.Close()
                    }
                }
            }
        }
        private static Nullable<Guid> GetShareHolderDVID(SqlConnection Connection, SqlTransaction Transaction, Guid? AccountDVID, string ShareholderName)
        {
            SqlCommand com = new SqlCommand();
            SqlDataAdapter ad = new SqlDataAdapter(com);

            System.Text.StringBuilder sql = new System.Text.StringBuilder();
            sql.AppendLine("SELECT ShareHolderDVID");
            sql.AppendLine("FROM ShareHolderDV");
            sql.AppendLine("WHERE AccountDVID = @AccountDVID");
            sql.AppendLine("AND ShareholderName = @ShareholderName");

            com.CommandText = sql.ToString();
            com.CommandType = CommandType.Text;
            com.Connection = Connection;
            com.Transaction = Transaction;
            com.CommandTimeout = int.MaxValue;

            //con.Open()
            try
            {
                com.Parameters.Add(new SqlParameter("@AccountDVID", AccountDVID));
                com.Parameters.Add(new SqlParameter("@ShareholderName", ShareholderName));
                DataTable dt = new DataTable();
                ad.Fill(dt);

                if (dt.Rows.Count > 0)
                {
                    return new Guid(dt.Rows[0][0].ToString());
                }
                else {
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
        public static Nullable<Guid> GetAccountIDByFileID(SqlConnection Connection, SqlTransaction Transaction, string FileID)
        {
            SqlCommand com = new SqlCommand();
            SqlDataAdapter ad = new SqlDataAdapter(com);

            System.Text.StringBuilder sql = new System.Text.StringBuilder();
            sql.AppendLine("SELECT AccountID FROM Account WHERE MSCFileID = @FileID");

            com.CommandText = sql.ToString();
            com.CommandType = CommandType.Text;
            com.Connection = Connection;
            com.Transaction = Transaction;
            com.CommandTimeout = int.MaxValue;

            //con.Open()
            try
            {
                com.Parameters.Add(new SqlParameter("@FileID", FileID));

                DataTable dt = new DataTable();
                ad.Fill(dt);

                if (dt.Rows.Count > 0)
                {
                    return new Guid(dt.Rows[0][0].ToString());
                }
                else {
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
        private static void DeleteShareholder(SqlConnection Connection, SqlTransaction Transaction, Guid? AccountID, string ShareHolderName)
        {

            SqlCommand com = new SqlCommand();
            System.Text.StringBuilder sql = new System.Text.StringBuilder();

            BOL.AuditLog.Modules.AccountLog alMgr = new BOL.AuditLog.Modules.AccountLog();
            Guid ShareHolderID = Guid.NewGuid();
            DataRow currentLogData = alMgr.SelectAccountForLog_Shareholder_Wizard(Connection, Transaction, ShareHolderID);

            sql.AppendLine("DELETE FROM ShareHolder ");
            sql.AppendLine("WHERE AccountID = @AccountID ");
            sql.AppendLine(" AND  ShareholderName = @ShareholderName ");
            com.Parameters.Add(new SqlParameter("@AccountID", AccountID));
            com.Parameters.Add(new SqlParameter("@ShareholderName", ShareHolderName));
            com.CommandText = sql.ToString();
            com.CommandType = CommandType.Text;
            com.Connection = Connection;
            com.Transaction = Transaction;
            com.CommandTimeout = int.MaxValue;

            try
            {
                //con.Open()
                com.ExecuteNonQuery();

                //Company Change Log - Shareholder
                DataRow newLogData = alMgr.SelectAccountForLog_Shareholder_Wizard(Connection, Transaction, ShareHolderID);
                if (currentLogData != null && newLogData != null)
                {
                    alMgr.CreateAccountLog_Shareholder_Wizard(Connection, Transaction, ShareHolderID, AccountID.Value, currentLogData, newLogData, new Guid(SyncHelper.AdminID), SyncHelper.AdminName);
                }
            }
            catch (Exception ex)
            {
                throw;
                //Finally
                //	con.Close()
            }
        }
        private static DataTable GetShareHolder(SqlConnection Connection, SqlTransaction Transaction, string MSCFileID)
        {
            SqlCommand com = new SqlCommand();
            SqlDataAdapter ad = new SqlDataAdapter(com);

            System.Text.StringBuilder sql = new System.Text.StringBuilder();
            sql.AppendLine("SELECT * ");
            sql.AppendLine("FROM IntegrationDB.dbo.EIR_PMSCOwnerShipDtls ");
            sql.AppendLine("WHERE FileID = @MSCFileID");

            com.CommandText = sql.ToString();
            com.CommandType = CommandType.Text;
            com.Connection = Connection;
            com.Transaction = Transaction;
            com.CommandTimeout = int.MaxValue;

            //con.Open()
            try
            {
                com.Parameters.Add(new SqlParameter("@MSCFileID", MSCFileID));

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
                //con.Close()
            }
        }
        public static string getFileName()
        {
            //LogList = new List<string>();
            string Directory = @"C:\WizardSync\spBigFiles";
            var directory0 = new DirectoryInfo(Directory);
            var myFile0 = (from f in directory0.GetFiles()
                           orderby f.LastWriteTime descending
                           select f).First();
            string Filename0 = Directory + @"\" + myFile0;
            return Filename0;
        }

        public static DataTable ExcelToDataTable(string fullFilename, string tableName)
        {
            string conStr = "";
            string Extension = Path.GetExtension(fullFilename);
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
    }
}
