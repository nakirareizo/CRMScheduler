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
    class CodeMaster
    {
        public Nullable<Guid> GetCodeMasterIDWithNull(string CodeType, string CodeName)
        {
            SqlConnection con = new SqlConnection(ConfigurationSettings.AppSettings["CRM"].ToString());
            SqlCommand com = new SqlCommand();
            SqlDataAdapter ad = new SqlDataAdapter(com);

            StringBuilder sql = new StringBuilder();
            sql.AppendLine("SELECT CodeMaster.CodeMasterID");
            sql.AppendLine("FROM CodeMaster");
            sql.AppendLine("WHERE CodeType = @CodeType");
            sql.AppendLine("AND CodeName = @CodeName");

            com.CommandText = sql.ToString();
            com.CommandType = CommandType.Text;
            com.Connection = con;

            com.Parameters.Add(new SqlParameter("@CodeType", CodeType));
            com.Parameters.Add(new SqlParameter("@CodeName", CodeName));

            con.Open();
            try
            {
                DataTable dt = new DataTable();
                ad.Fill(dt);
                if (dt != null && dt.Rows.Count > 0)
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

        public Guid GetCodeMasterID_Wizard(SqlConnection Connection, SqlTransaction Transaction, string CodeType, string CodeName)
        {
            SqlCommand com = new SqlCommand();
            SqlDataAdapter ad = new SqlDataAdapter(com);

            StringBuilder sql = new StringBuilder();
            sql.AppendLine("SELECT CodeMaster.CodeMasterID");
            sql.AppendLine("FROM CodeMaster");
            sql.AppendLine("WHERE CodeType = @CodeType");
            sql.AppendLine("AND CodeName = @CodeName");

            com.CommandText = sql.ToString();
            com.CommandType = CommandType.Text;
            com.Connection = Connection;
            com.Transaction = Transaction;

            com.Parameters.Add(new SqlParameter("@CodeType", CodeType));
            com.Parameters.Add(new SqlParameter("@CodeName", CodeName));

            //con.Open()
            try
            {
                DataTable dt = new DataTable();
                ad.Fill(dt);

                if (dt != null && dt.Rows.Count > 0)
                {
                    return new Guid( dt.Rows[0][0].ToString());
                }
                else {
                    return Guid.Empty;
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

        public Nullable<Guid> GetCodeMasterIDWithNull_Wizard(SqlConnection Connection, SqlTransaction Transaction, string CodeType, string CodeName)
        {
            SqlCommand com = new SqlCommand();
            SqlDataAdapter ad = new SqlDataAdapter(com);

            StringBuilder sql = new StringBuilder();
            sql.AppendLine("SELECT CodeMaster.CodeMasterID");
            sql.AppendLine("FROM CodeMaster");
            sql.AppendLine("WHERE CodeType = @CodeType");
            sql.AppendLine("AND CodeName = @CodeName");

            com.CommandText = sql.ToString();
            com.CommandType = CommandType.Text;
            com.Connection = Connection;
            com.Transaction = Transaction;

            com.Parameters.Add(new SqlParameter("@CodeType", CodeType));
            com.Parameters.Add(new SqlParameter("@CodeName", CodeName));

            //con.Open()
            try
            {
                DataTable dt = new DataTable();
                ad.Fill(dt);
                if (dt != null && dt.Rows.Count > 0)
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

    }
}
