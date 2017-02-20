using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CRMSync.Classes.MainProcess
{
    class SyncRelocationStatus
    {
        internal Guid? GetRelocationStatusID(SqlConnection Connection, SqlTransaction Transaction, Guid? AccountID, Guid? RelocationStatusID)
        {
            SqlCommand com = new SqlCommand();
            SqlDataAdapter ad = new SqlDataAdapter(com);

            System.Text.StringBuilder sql = new System.Text.StringBuilder();
            sql.AppendLine("SELECT relocationid");
            sql.AppendLine("FROM   relocation");
            sql.AppendLine("WHERE  accountid = @AccountID");
            sql.AppendLine("AND    relocationstatus = @relocationstatus");

            com.CommandText = sql.ToString();
            com.CommandType = CommandType.Text;
            com.Connection = Connection;
            com.Transaction = Transaction;
            com.CommandTimeout = int.MaxValue;

            try
            {
                com.Parameters.Add(new SqlParameter("@AccountID", AccountID));
                com.Parameters.Add(new SqlParameter("@relocationstatus", RelocationStatusID));

                DataTable dt = new DataTable();
                ad.Fill(dt);
                //con.Close()
                if (dt.Rows.Count > 0)
                {
                    return new Guid( dt.Rows[0][0].ToString());
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
        internal void UpdateRelocationDeadline(SqlConnection Connection, SqlTransaction Transaction, Guid? AccountID, Guid? RelocationID, Nullable<DateTime> RelocationDeadline)
        {

            SqlCommand com = new SqlCommand();

            System.Text.StringBuilder sql = new System.Text.StringBuilder();
            sql.AppendLine("UPDATE relocation");
            sql.AppendLine("SET    relocationdeadline = @relocationdeadline, ");
            sql.AppendLine("       modifiedby = @modifiedby, ");
            sql.AppendLine("       modifiedbyname = @modifiedbyname, ");
            sql.AppendLine("       modifieddate = @modifieddate ");
            sql.AppendLine("FROM   relocation");
            sql.AppendLine("WHERE  relocationid = @relocationid");

            com.CommandText = sql.ToString();
            com.CommandType = CommandType.Text;
            com.Connection = Connection;
            com.Transaction = Transaction;
            com.CommandTimeout = int.MaxValue;

            com.Parameters.Add(new SqlParameter("@relocationid", RelocationID));
            com.Parameters.Add(new SqlParameter("@relocationdeadline", ConvertToNull(RelocationDeadline)));
            com.Parameters.Add(new SqlParameter("@modifiedby", SyncHelper.AdminID));
            com.Parameters.Add(new SqlParameter("@modifiedbyname", SyncHelper.AdminName));
            com.Parameters.Add(new SqlParameter("@modifieddate", DateTime.Now));

            try
            {
                com.ExecuteNonQuery();
                BOL.AuditLog.Modules.AccountLog alMgr = new BOL.AuditLog.Modules.AccountLog();
                DataRow newLogData = alMgr.SelectAccountForLog_Relocation_Wizard(Connection, Transaction, RelocationID.Value);
                alMgr.CreateAccountLog_Relocation_Wizard(Connection, Transaction, RelocationID.Value, AccountID.Value, null, newLogData, new Guid(SyncHelper.AdminID), SyncHelper.AdminName);

            }
            catch (Exception ex)
            {
            }
        }
        private static object ConvertToNull(object value)
        {
            if (value == null)
            {
                return DBNull.Value;
            }
            else {
                return value;
            }
        }
        internal void CreateRelocationDeadline(SqlConnection Connection, SqlTransaction Transaction, Guid? AccountID, Guid? RelocationStatusID, Nullable<DateTime> RelocationDeadline)
        {
            SqlCommand com = new SqlCommand();

            StringBuilder sql = new StringBuilder();
            sql.AppendLine("INSERT INTO relocation ");
            sql.AppendLine("            (relocationid, ");
            sql.AppendLine("             accountid, ");
            sql.AppendLine("             relocationstatus, ");
            sql.AppendLine("             relocationdeadline, ");
            sql.AppendLine("             createddate, ");
            sql.AppendLine("             modifieddate, ");
            sql.AppendLine("             createdby, ");
            sql.AppendLine("             modifiedby, ");
            sql.AppendLine("             createdbyname, ");
            sql.AppendLine("             modifiedbyname) ");
            sql.AppendLine("VALUES      (@relocationid, ");
            sql.AppendLine("             @accountid, ");
            sql.AppendLine("             @relocationstatus, ");
            sql.AppendLine("             @relocationdeadline, ");
            sql.AppendLine("             @createddate, ");
            sql.AppendLine("             @modifieddate, ");
            sql.AppendLine("             @createdby, ");
            sql.AppendLine("             @modifiedby, ");
            sql.AppendLine("             @createdbyname, ");
            sql.AppendLine("             @modifiedbyname) ");

            com.CommandText = sql.ToString();
            com.CommandType = CommandType.Text;
            com.Connection = Connection;
            com.Transaction = Transaction;
            com.CommandTimeout = int.MaxValue;

            Guid RelocationID = Guid.NewGuid();
            com.Parameters.Add(new SqlParameter("@relocationid", RelocationID));
            com.Parameters.Add(new SqlParameter("@accountid", AccountID));
            com.Parameters.Add(new SqlParameter("@relocationstatus", RelocationStatusID));
            com.Parameters.Add(new SqlParameter("@relocationdeadline", ConvertToNull(RelocationDeadline)));
            com.Parameters.Add(new SqlParameter("@createdby", SyncHelper.AdminID));
            com.Parameters.Add(new SqlParameter("@createdbyname", SyncHelper.AdminName));
            com.Parameters.Add(new SqlParameter("@createddate", DateTime.Now));
            com.Parameters.Add(new SqlParameter("@modifiedby", SyncHelper.AdminID));
            com.Parameters.Add(new SqlParameter("@modifiedbyname", SyncHelper.AdminName));
            com.Parameters.Add(new SqlParameter("@modifieddate", DateTime.Now));

            try
            {
                //con.Open()
                com.ExecuteNonQuery();

                //Company Change Log - MSC History
                BOL.AuditLog.Modules.AccountLog alMgr = new BOL.AuditLog.Modules.AccountLog();
                DataRow newLogData = alMgr.SelectAccountForLog_Relocation_Wizard(Connection, Transaction, RelocationID);
                alMgr.CreateAccountLog_Relocation_Wizard(Connection, Transaction, RelocationID, AccountID.Value, null, newLogData, new Guid(SyncHelper.AdminID), SyncHelper.AdminName);

            }
            catch (Exception ex)
            {
            }
        }
    }
}
