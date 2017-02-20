using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CRMSync.Classes
{
    class SyncAccount
    {
        internal static Guid? GetAccountIDByFileID(SqlConnection Connection, SqlTransaction Transaction, string FileID)
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
                con.Close();
            }
        }
    }
}
