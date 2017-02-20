using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CRMSync
{
    class SyncHelper
    {

        public const string AdminName = "admin";
        public const string AdminID = "74425431-65A4-498E-A6ED-910A9E20B6FC";
        public const string AQIRAdminID = "086FBDE4-4D4D-4959-BC01-7D2E182D1936";

        public const string DataSource = "WIZ";
        public static SqlConnection NewWizardConnection()
        {
            return new SqlConnection(ConfigurationSettings.AppSettings["Wizard"].ToString());
        }

        public static SqlConnection NewCRMConnection()
        {
            return new SqlConnection(ConfigurationSettings.AppSettings["CRM"].ToString());
        }

        public static object ReturnNull(Nullable<Guid> Value)
        {
            if (Value.HasValue)
            {
                return Value.Value;
            }
            else
            {
                return DBNull.Value;
            }
        }

        public static object ReturnNull(Nullable<DateTime> Value)
        {
            if (Value.HasValue)
            {
                return Value.Value;
            }
            else
            {
                return DBNull.Value;
            }
        }

        public static object ReturnNull(Nullable<decimal> Value)
        {
            if (Value.HasValue)
            {
                return Value.Value;
            }
            else
            {
                return DBNull.Value;
            }
        }

        public static object ReturnNull(Nullable<int> Value)
        {
            if (Value.HasValue)
            {
                return Value.Value;
            }
            else
            {
                return DBNull.Value;
            }
        }

        public static object ReturnNull(string Value)
        {
            if (!string.IsNullOrEmpty(Value))
            {
                return Value;
            }
            else
            {
                return DBNull.Value;
            }
        }

        public static Nullable<DateTime> ConvertStringToDateTime(string value, bool ReturnNowIfException = true)
        {
            if (string.IsNullOrEmpty(value))
                return null;

            try
            {
                System.Globalization.DateTimeFormatInfo format = new System.Globalization.DateTimeFormatInfo();
                format.ShortDatePattern = "dd/MM/yyyy";
                format.DateSeparator = "/";
                return Convert.ToDateTime(value, format);
            }
            catch (Exception ex)
            {
                if (ReturnNowIfException)
                {
                    return DateTime.Now;
                }
                else
                {
                    return null;
                }
            }
        }

        public static Nullable<int> ConvertToInteger(object value)
        {
            if (string.IsNullOrEmpty(value.ToString()))
                return null;

            try
            {
                return Convert.ToInt32(value);
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        internal static bool IsNonMSC(Guid? AccountID)
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

        private static Guid GetAccountTypeCID(Guid? AccountID)
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
                return new Guid( dt.Rows[0][0].ToString());
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

        internal static SqlConnection WizardProductionConnection()
        {
            return new SqlConnection(ConfigurationSettings.AppSettings["WizardProcduction"].ToString());
        }

        internal static decimal ConvertToDecimal(object value)
        {
            if (string.IsNullOrEmpty(value.ToString()))
                return 0;

            try
            {
                return Convert.ToDecimal(value);
            }
            catch (Exception ex)
            {
                return 0;
            }
        }

        internal static Guid? GetAccountIDByFileID(string FileID)
        {
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
                    return new Guid (dt.Rows[0][0].ToString());
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
        }

        public static Nullable<double> ConvertToDouble(string value)
        {
            if (string.IsNullOrEmpty(value.ToString()))
                return null;

            try
            {
                return Convert.ToDouble(value);
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public static bool ConvertToBoolean(object value)
        {
            try
            {
                return Convert.ToBoolean(value);
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        private static Guid CreateCodeMaster(SqlConnection Connection, SqlTransaction Transaction, string Value, string CodeType)
        {
            Guid codeMasterID = Guid.NewGuid();

            //Dim con As SqlClient.SqlConnection = SyncHelper.NewCRMConnection
            SqlCommand com = new SqlCommand();

            System.Text.StringBuilder sql = new System.Text.StringBuilder();
            sql.AppendLine("INSERT INTO CodeMaster (CodeMasterID, SeqNo, CodeType, CodeName,");
            sql.AppendLine("CreatedBy,CreatedByName,ModifiedBy,ModifiedByName,CreatedDate,ModifiedDate) ");
            sql.AppendLine("VALUES (@CodeMasterID, 0, @CodeType, @CodeName, ");
            sql.AppendLine("@CreatedBy, @CreatedByName, @ModifiedBy, @ModifiedByName,@CreatedDate,@ModifiedDate ) ");

            com.CommandText = sql.ToString();
            com.CommandType = CommandType.Text;
            com.Connection = Connection;
            com.Transaction = Transaction;

            codeMasterID = Guid.NewGuid();
            com.Parameters.Add(new SqlParameter("@CodeMasterID", codeMasterID));
            com.Parameters.Add(new SqlParameter("@CodeType", CodeType));
            com.Parameters.Add(new SqlParameter("@CodeName", Value));
            com.Parameters.Add(new SqlParameter("@CreatedBy", SyncHelper.AdminID));
            com.Parameters.Add(new SqlParameter("@CreatedByName", SyncHelper.AdminName));
            com.Parameters.Add(new SqlParameter("@ModifiedBy", SyncHelper.AdminID));
            com.Parameters.Add(new SqlParameter("@ModifiedByName", SyncHelper.AdminName));
            com.Parameters.Add(new SqlParameter("@CreatedDate", DateTime.Now));
            com.Parameters.Add(new SqlParameter("@ModifiedDate", DateTime.Now));

            try
            {
                //con.Open()
                com.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                throw;
            }
            finally
            {
                //con.Close()
            }

            return codeMasterID;
        }

        private static Guid CreateCodeMaster(string Value, string CodeType)
        {
            Guid codeMasterID = Guid.NewGuid();

            SqlConnection con = SyncHelper.NewCRMConnection();
            SqlCommand com = new SqlCommand();

            System.Text.StringBuilder sql = new System.Text.StringBuilder();
            sql.AppendLine("INSERT INTO CodeMaster (CodeMasterID, SeqNo, CodeType, CodeName,");
            sql.AppendLine("CreatedBy,CreatedByName,ModifiedBy,ModifiedByName,CreatedDate,ModifiedDate) ");
            sql.AppendLine("VALUES (@CodeMasterID, 0, @CodeType, @CodeName, ");
            sql.AppendLine("@CreatedBy, @CreatedByName, @ModifiedBy, @ModifiedByName,@CreatedDate,@ModifiedDate ) ");

            com.CommandText = sql.ToString();
            com.CommandType = CommandType.Text;
            com.Connection = con;
            codeMasterID = Guid.NewGuid();
            com.Parameters.Add(new SqlParameter("@CodeMasterID", codeMasterID));
            com.Parameters.Add(new SqlParameter("@CodeType", CodeType));
            com.Parameters.Add(new SqlParameter("@CodeName", Value));
            com.Parameters.Add(new SqlParameter("@CreatedBy", SyncHelper.AdminID));
            com.Parameters.Add(new SqlParameter("@CreatedByName", SyncHelper.AdminName));
            com.Parameters.Add(new SqlParameter("@ModifiedBy", SyncHelper.AdminID));
            com.Parameters.Add(new SqlParameter("@ModifiedByName", SyncHelper.AdminName));
            com.Parameters.Add(new SqlParameter("@CreatedDate", DateTime.Now));
            com.Parameters.Add(new SqlParameter("@ModifiedDate", DateTime.Now));

            try
            {
                con.Open();
                com.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                throw;
            }
            finally
            {
                con.Close();
            }

            return codeMasterID;
        }

        public static Nullable<Guid> GetCodeMasterID(SqlConnection Connection, SqlTransaction Transaction, string Value, string CodeType, bool CreateIfNotExists = false)
        {
            SqlCommand com = new SqlCommand();
            SqlDataAdapter ad = new SqlDataAdapter(com);

            System.Text.StringBuilder sql = new System.Text.StringBuilder();
            sql.AppendLine("SELECT CodeMasterID");
            sql.AppendLine("FROM CodeMaster");
            sql.AppendLine("WHERE CodeType = @CodeType");
            sql.AppendLine("AND CodeName = @Value");

            com.CommandText = sql.ToString();
            com.CommandType = CommandType.Text;
            com.Connection = Connection;
            com.Transaction = Transaction;

            //con.Open()
            try
            {
                com.Parameters.Add(new SqlParameter("@CodeType", CodeType));
                com.Parameters.Add(new SqlParameter("@Value", Value));

                DataTable dt = new DataTable();
                ad.Fill(dt);

                if (dt.Rows.Count > 0)
                {
                    return new Guid(dt.Rows[0][0].ToString());
                }
                else
                {
                    if (CreateIfNotExists)
                    {
                        return CreateCodeMaster(Connection, Transaction, Value, CodeType);
                    }
                    else
                    {
                        return null;
                    }
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

        public static Nullable<Guid> GetCodeMasterID(string Value, string CodeType, bool CreateIfNotExists = false)
        {
            SqlConnection con = new SqlConnection(ConfigurationSettings.AppSettings["CRM"].ToString());
            SqlCommand com = new SqlCommand();
            SqlDataAdapter ad = new SqlDataAdapter(com);

            System.Text.StringBuilder sql = new System.Text.StringBuilder();
            sql.AppendLine("SELECT CodeMasterID");
            sql.AppendLine("FROM CodeMaster");
            sql.AppendLine("WHERE CodeType = @CodeType");
            sql.AppendLine("AND CodeName = @Value");

            com.CommandText = sql.ToString();
            com.CommandType = CommandType.Text;
            com.Connection = con;

            //con.Open()
            try
            {
                com.Parameters.Add(new SqlParameter("@CodeType", CodeType));
                com.Parameters.Add(new SqlParameter("@Value", Value));

                DataTable dt = new DataTable();
                ad.Fill(dt);

                if (dt.Rows.Count > 0)
                {
                    return new Guid(dt.Rows[0][0].ToString());
                }
                else
                {
                    if (CreateIfNotExists)
                    {
                        return CreateCodeMaster(Value, CodeType);
                    }
                    else
                    {
                        return null;
                    }
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

        public static string RemoveDoubleSpace(string Value)
        {
            while (Value.IndexOf("  ") > 0)
            {
                Value = Value.Replace("  ", " ");
            }
            return Value;
        }

        public static List<string> ForceSplit(string Value, int MaxLength)
        {
            List<string> splittedText = new List<string>();
            string newValue = Value;

            while (newValue.Length > MaxLength)
            {
                // look for last space to split, if no last space, force to split
                int lastSpaceIndex = newValue.LastIndexOf(" ", MaxLength);
                if (lastSpaceIndex == -1)
                    lastSpaceIndex = MaxLength;

                string newText = newValue.Substring(0, lastSpaceIndex);
                if (lastSpaceIndex == MaxLength)
                {
                    newValue = newValue.Substring(lastSpaceIndex, newValue.Length - lastSpaceIndex);
                }
                else
                {
                    newValue = newValue.Substring(lastSpaceIndex + 1, newValue.Length - lastSpaceIndex - 1);
                }
                if (newText.Trim().Length > 0)
                    splittedText.Add(newText.Trim());
            }

            // insert last line
            if (newValue.Trim().Length > 0)
                splittedText.Add(newValue.Trim());

            return splittedText;
        }

        public static string GetCity(SqlConnection Connection, SqlTransaction Transaction, string Value)
        {
            SqlCommand com = new SqlCommand();
            SqlDataAdapter ad = new SqlDataAdapter(com);

            System.Text.StringBuilder sql = new System.Text.StringBuilder();
            sql.AppendLine("SELECT r.RegionName");
            sql.AppendLine("FROM Region r");
            sql.AppendLine("JOIN RegionType rt ON r.RegionTypeID = rt.RegionTypeID");
            sql.AppendLine("WHERE r.RegionName = @Value");
            sql.AppendLine("AND rt.RegionType = 'City'");

            com.CommandText = sql.ToString();
            com.CommandType = CommandType.Text;
            com.Connection = Connection;
            com.Transaction = Transaction;

            //con.Open()
            try
            {
                com.Parameters.Add(new SqlParameter("@Value", Value));

                DataTable dt = new DataTable();
                ad.Fill(dt);

                if (dt.Rows.Count > 0)
                {
                    return dt.Rows[0][0].ToString();
                }
                else
                {
                    return string.Empty;
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

        public static string GetCity(string Value)
        {
           SqlConnection con = new SqlConnection(ConfigurationSettings.AppSettings["CRM"].ToString());
            SqlCommand com = new SqlCommand();
            SqlDataAdapter ad = new SqlDataAdapter(com);

            System.Text.StringBuilder sql = new System.Text.StringBuilder();
            sql.AppendLine("SELECT r.RegionName");
            sql.AppendLine("FROM Region r");
            sql.AppendLine("JOIN RegionType rt ON r.RegionTypeID = rt.RegionTypeID");
            sql.AppendLine("WHERE r.RegionName = @Value");
            sql.AppendLine("AND rt.RegionType = 'City'");

            com.CommandText = sql.ToString();
            com.CommandType = CommandType.Text;
            com.Connection = con;

            con.Open();
            try
            {
                com.Parameters.Add(new SqlParameter("@Value", Value));

                DataTable dt = new DataTable();
                ad.Fill(dt);

                if (dt.Rows.Count > 0)
                {
                    return dt.Rows[0][0].ToString();
                }
                else
                {
                    return string.Empty;
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

        public static string GetState(SqlConnection Connection, SqlTransaction Transaction, string Value)
        {
            SqlCommand com = new SqlCommand();
            SqlDataAdapter ad = new SqlDataAdapter(com);

            System.Text.StringBuilder sql = new System.Text.StringBuilder();
            sql.AppendLine("SELECT r.RegionName");
            sql.AppendLine("FROM Region r");
            sql.AppendLine("JOIN RegionType rt ON r.RegionTypeID = rt.RegionTypeID");
            sql.AppendLine("WHERE r.RegionName = @Value");
            sql.AppendLine("AND rt.RegionType = 'State'");

            com.CommandText = sql.ToString();
            com.CommandType = CommandType.Text;
            com.Connection = Connection;
            com.Transaction = Transaction;

            //con.Open()
            try
            {
                com.Parameters.Add(new SqlParameter("@Value", Value));

                DataTable dt = new DataTable();
                ad.Fill(dt);

                if (dt.Rows.Count > 0)
                {
                    return dt.Rows[0][0].ToString();
                }
                else
                {
                    return string.Empty;
                }
            }
            catch (Exception ex)
            {
                throw;
                //Finally
                //	con.Close()
            }
        }

        public static string GetState(string Value)
        {
            SqlConnection con = new SqlConnection(ConfigurationSettings.AppSettings["CRM"].ToString());
            SqlCommand com = new SqlCommand();
            SqlDataAdapter ad = new SqlDataAdapter(com);

            System.Text.StringBuilder sql = new System.Text.StringBuilder();
            sql.AppendLine("SELECT r.RegionName");
            sql.AppendLine("FROM Region r");
            sql.AppendLine("JOIN RegionType rt ON r.RegionTypeID = rt.RegionTypeID");
            sql.AppendLine("WHERE r.RegionName = @Value");
            sql.AppendLine("AND rt.RegionType = 'State'");

            com.CommandText = sql.ToString();
            com.CommandType = CommandType.Text;
            com.Connection = con;

            con.Open();
            try
            {
                com.Parameters.Add(new SqlParameter("@Value", Value));

                DataTable dt = new DataTable();
                ad.Fill(dt);

                if (dt.Rows.Count > 0)
                {
                    return dt.Rows[0][0].ToString();
                }
                else
                {
                    return string.Empty;
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

        public static Nullable<Guid> GetRegionID(SqlConnection Connection, SqlTransaction Transaction, string Value)
        {
            SqlCommand com = new SqlCommand();
            SqlDataAdapter ad = new SqlDataAdapter(com);

            System.Text.StringBuilder sql = new System.Text.StringBuilder();
            sql.AppendLine("SELECT RegionID");
            sql.AppendLine("FROM Region");
            sql.AppendLine("WHERE RegionName = @Value");

            com.CommandText = sql.ToString();
            com.CommandType = CommandType.Text;
            com.Connection = Connection;
            com.Transaction = Transaction;

            //con.Open()
            try
            {
                com.Parameters.Add(new SqlParameter("@Value", Value));

                DataTable dt = new DataTable();
                ad.Fill(dt);

                if (dt.Rows.Count > 0)
                {
                    return new Guid( dt.Rows[0][0].ToString());
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

        public static Nullable<Guid> GetRegionID(string Value)
        {
            SqlConnection con = new SqlConnection(ConfigurationSettings.AppSettings["CRM"].ToString());
            SqlCommand com = new SqlCommand();
            SqlDataAdapter ad = new SqlDataAdapter(com);

            System.Text.StringBuilder sql = new System.Text.StringBuilder();
            sql.AppendLine("SELECT RegionID");
            sql.AppendLine("FROM Region");
            sql.AppendLine("WHERE RegionName = @Value");

            com.CommandText = sql.ToString();
            com.CommandType = CommandType.Text;
            com.Connection = con;

            con.Open();
            try
            {
                com.Parameters.Add(new SqlParameter("@Value", Value));

                DataTable dt = new DataTable();
                ad.Fill(dt);

                if (dt.Rows.Count > 0)
                {
                    return new Guid (dt.Rows[0][0].ToString());
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
        }

        public static Nullable<Guid> GetSubClusterID_Wizard(SqlConnection Connection, SqlTransaction Transaction, string ClusterName)
        {
            SqlCommand com = new SqlCommand();
            SqlDataAdapter ad = new SqlDataAdapter(com);

            System.Text.StringBuilder sql = new System.Text.StringBuilder();
            sql.AppendLine("SELECT ClusterID");
            sql.AppendLine("FROM Cluster");
            sql.AppendLine("WHERE ClusterName = @ClusterName");
            //sql.AppendLine("AND ParentClusterID IS NOT NULL")

            com.CommandText = sql.ToString();
            com.CommandType = CommandType.Text;
            com.Connection = Connection;
            com.Transaction = Transaction;
            com.CommandTimeout = int.MaxValue;

            //con.Open()
            try
            {
                com.Parameters.Add(new SqlParameter("@ClusterName", ClusterName));
                DataTable dt = new DataTable();
                ad.Fill(dt);

                if (dt.Rows.Count > 0)
                {
                    return new Guid( dt.Rows[0][0].ToString());
                }
                else
                {
                    return null;
                }
            }
            catch (Exception ex)
            {
                throw;
                //Finally
                //	con.Close()
            }
        }

        public static Nullable<Guid> GetSubClusterID(string ClusterName)
        {
            SqlConnection con = new SqlConnection(ConfigurationSettings.AppSettings["CRM"].ToString());
            SqlCommand com = new SqlCommand();
            SqlDataAdapter ad = new SqlDataAdapter(com);

            System.Text.StringBuilder sql = new System.Text.StringBuilder();
            sql.AppendLine("SELECT ClusterID");
            sql.AppendLine("FROM Cluster");
            sql.AppendLine("WHERE ClusterName = @ClusterName");
            //sql.AppendLine("AND ParentClusterID IS NOT NULL")

            com.CommandText = sql.ToString();
            com.CommandType = CommandType.Text;
            com.Connection = con;
            com.CommandTimeout = int.MaxValue;

            con.Open();
            try
            {
                com.Parameters.Add(new SqlParameter("@ClusterName", ClusterName));
                DataTable dt = new DataTable();
                ad.Fill(dt);

                if (dt.Rows.Count > 0)
                {
                    return new Guid( dt.Rows[0][0].ToString());
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
        }

        public static Nullable<Guid> GetUserIDByFullName(string FullName)
        {
            SqlConnection con = new SqlConnection(ConfigurationSettings.AppSettings["CRM"].ToString());
            SqlCommand com = new SqlCommand();
            SqlDataAdapter ad = new SqlDataAdapter(com);

            System.Text.StringBuilder sql = new System.Text.StringBuilder();
            sql.AppendLine("SELECT UserID");
            sql.AppendLine("FROM SecurityUserProfile");
            sql.AppendLine("WHERE FullName = @FullName");

            com.CommandText = sql.ToString();
            com.CommandType = CommandType.Text;
            com.Connection = con;
            com.CommandTimeout = int.MaxValue;

            con.Open();
            try
            {
                com.Parameters.Add(new SqlParameter("@FullName", FullName));
                DataTable dt = new DataTable();
                ad.Fill(dt);

                if (dt.Rows.Count > 0)
                {
                    return new Guid( dt.Rows[0][0].ToString());
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
        }

        public static string GetMappingValue_Wizard(SqlConnection Connection, SqlTransaction Transaction, string CodeType, string FromValue)
        {
            SqlCommand com = new SqlCommand();
            SqlDataAdapter ad = new SqlDataAdapter(com);

            System.Text.StringBuilder sql = new System.Text.StringBuilder();
            sql.AppendLine("SELECT ToValue");
            sql.AppendLine("FROM Mapping");
            sql.AppendLine("WHERE CodeType = @CodeType");
            sql.AppendLine("AND FromValue = @FromValue");

            com.CommandText = sql.ToString();
            com.CommandType = CommandType.Text;
            com.Connection = Connection;
            com.Transaction = Transaction;

            //con.Open()
            try
            {
                com.Parameters.Add(new SqlParameter("@CodeType", CodeType));
                com.Parameters.Add(new SqlParameter("@FromValue", FromValue));

                DataTable dt = new DataTable();
                ad.Fill(dt);

                if (dt.Rows.Count > 0)
                {
                    return dt.Rows[0][0].ToString();
                }
                else
                {
                    return string.Empty;
                }
            }
            catch (Exception ex)
            {
                throw;
                //Finally
                //con.Close()
            }
        }

        public static string GetMappingValue(string CodeType, string FromValue)
        {
            SqlConnection con = new SqlConnection(ConfigurationSettings.AppSettings["CRM"].ToString());
            SqlCommand com = new SqlCommand();
            SqlDataAdapter ad = new SqlDataAdapter(com);

            System.Text.StringBuilder sql = new System.Text.StringBuilder();
            sql.AppendLine("SELECT ToValue");
            sql.AppendLine("FROM Mapping");
            sql.AppendLine("WHERE CodeType = @CodeType");
            sql.AppendLine("AND FromValue = @FromValue");

            com.CommandText = sql.ToString();
            com.CommandType = CommandType.Text;
            com.Connection = con;

            con.Open();
            try
            {
                com.Parameters.Add(new SqlParameter("@CodeType", CodeType));
                com.Parameters.Add(new SqlParameter("@FromValue", FromValue));

                DataTable dt = new DataTable();
                ad.Fill(dt);

                if (dt.Rows.Count > 0)
                {
                    return dt.Rows[0][0].ToString();
                }
                else
                {
                    return string.Empty;
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
    }
}
