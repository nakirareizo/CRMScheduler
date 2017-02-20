using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CRMSync.Classes
{
    class SyncMSCStatusHistory
    {
        internal Guid? GetMSCStatusHistoryID(SqlConnection Connection, SqlTransaction Transaction, Guid? AccountID, Guid? MSCApprovalStatusCID)
        {
            SqlCommand com = new SqlCommand();
            SqlDataAdapter ad = new SqlDataAdapter(com);

            System.Text.StringBuilder sql = new System.Text.StringBuilder();
            sql.AppendLine("SELECT MSCStatusHistoryID");
            sql.AppendLine("FROM MSCStatusHistory");
            sql.AppendLine("WHERE AccountID = @AccountID");
            sql.AppendLine("AND MSCApprovalStatusCID = @MSCApprovalStatusCID");

            com.CommandText = sql.ToString();
            com.CommandType = CommandType.Text;
            com.Connection = Connection;
            com.Transaction = Transaction;
            com.CommandTimeout = int.MaxValue;

            //con.Open()
            try
            {
                com.Parameters.Add(new SqlParameter("@AccountID", AccountID));
                com.Parameters.Add(new SqlParameter("@MSCApprovalStatusCID", MSCApprovalStatusCID));

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

        internal void CreateMSCStatusHistory(SqlConnection Connection, SqlTransaction Transaction, Guid? AccountID, Guid? MSCApprovalStatusCID, DateTime? MSCApprovalDate)
        {
            SqlCommand com = new SqlCommand();
           StringBuilder sql = new StringBuilder();
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
            sql.AppendLine("GETDATE(),");
            sql.AppendLine("@CreatedBy,");
            sql.AppendLine("@CreatedByName,");
            sql.AppendLine("GETDATE()");
            sql.AppendLine(")");

            com.CommandText = sql.ToString();
            com.CommandType = CommandType.Text;
            com.Connection = Connection;
            com.Transaction = Transaction;
            com.CommandTimeout = int.MaxValue;

            Guid MSCStatusHistoryID = Guid.NewGuid();
            com.Parameters.Add(new SqlParameter("@MSCStatusHistoryID", MSCStatusHistoryID));
            com.Parameters.Add(new SqlParameter("@AccountID", AccountID));
            com.Parameters.Add(new SqlParameter("@MSCApprovalStatusCID", MSCApprovalStatusCID));
            com.Parameters.Add(new SqlParameter("@MSCApprovalDate", MSCApprovalDate));
            com.Parameters.Add(new SqlParameter("@DataSource", SyncHelper.DataSource));
            com.Parameters.Add(new SqlParameter("@CreatedBy", SyncHelper.AdminID));
            com.Parameters.Add(new SqlParameter("@CreatedByName", SyncHelper.AdminName));

            try
            {
                //con.Open()
                com.ExecuteNonQuery();

                //Company Change Log - MSC History
                BOL.AuditLog.Modules.AccountLog alMgr = new BOL.AuditLog.Modules.AccountLog();
                DataRow newLogData = alMgr.SelectAccountForLog_MSCStatus_Wizard(Connection, Transaction, MSCStatusHistoryID);
                alMgr.CreateAccountLog_MSCStatusHistory_Wizard(Connection, Transaction, MSCStatusHistoryID, AccountID.Value, null, newLogData, new Guid(SyncHelper.AdminID), SyncHelper.AdminName);

            }
            catch (Exception ex)
            {
                throw;
                //Finally
                //	con.Close()
            }
        }

        public void UpdateMSCStatusHistory(SqlConnection Connection, SqlTransaction Transaction, Guid? AccountID, Guid? MSCStatusHistoryID, DateTime? MSCApprovalDate)
        {
            SqlCommand com = new SqlCommand();

            System.Text.StringBuilder sql = new System.Text.StringBuilder();
            sql.AppendLine("UPDATE MSCStatusHistory SET");
            sql.AppendLine("MSCApprovalDate = @MSCApprovalDate,");
            sql.AppendLine("DataSource = @DataSource,");
            sql.AppendLine("ModifiedBy = @ModifiedBy,");
            sql.AppendLine("ModifiedByName = @ModifiedByName,");
            sql.AppendLine("ModifiedDate = GETDATE()");
            sql.AppendLine("WHERE MSCStatusHistoryID = @MSCStatusHistoryID");

            com.CommandText = sql.ToString();
            com.CommandType = CommandType.Text;
            com.Connection = Connection;
            com.Transaction = Transaction;
            com.CommandTimeout = int.MaxValue;

            com.Parameters.Add(new SqlParameter("@MSCStatusHistoryID", MSCStatusHistoryID));
            com.Parameters.Add(new SqlParameter("@MSCApprovalDate", MSCApprovalDate));
            com.Parameters.Add(new SqlParameter("@DataSource",SyncHelper.DataSource));
            com.Parameters.Add(new SqlParameter("@ModifiedBy", SyncHelper.AdminID));
            com.Parameters.Add(new SqlParameter("@ModifiedByName", SyncHelper.AdminName));

            try
            {
                com.ExecuteNonQuery();

                //Company Change Log - MSC History
                BOL.AuditLog.Modules.AccountLog alMgr = new BOL.AuditLog.Modules.AccountLog();
                DataRow newLogData = alMgr.SelectAccountForLog_MSCStatus_Wizard(Connection, Transaction, MSCStatusHistoryID.Value);
                alMgr.CreateAccountLog_MSCStatusHistory_Wizard(Connection, Transaction, MSCStatusHistoryID.Value, AccountID.Value, null, newLogData, new Guid(SyncHelper.AdminID), SyncHelper.AdminName);

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
