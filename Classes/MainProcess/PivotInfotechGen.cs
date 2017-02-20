using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;

namespace CRMSync.Classes.MainProcess
{
    class PivotInfotechGen
    {
        internal static void Start()
        {
            DataTable dtInfotech = getAllData();
            if (dtInfotech.Rows.Count > 0)
            {
                ExcelFileHelper.GenerateExcelFile(dtInfotech, DateTime.Now.ToString("dd-MM-yyyy"));
            }
        }

        private static DataTable getAllData()
        {
            DataTable dtFiltered = new DataTable();
            dtFiltered.Columns.Add("MSCFileID");
            dtFiltered.Columns.Add("Company Name");
            dtFiltered.Columns.Add("Registration Number");
            dtFiltered.Columns.Add("Year of Approval");
            dtFiltered.Columns.Add("Operation Status");
            dtFiltered.Columns.Add("RevenueYr1");
            dtFiltered.Columns.Add("RevenueYr2");
            dtFiltered.Columns.Add("RevenueYr3");
            dtFiltered.Columns.Add("RevenueYr4");
            dtFiltered.Columns.Add("LocalSalesYr1");
            dtFiltered.Columns.Add("LocalSalesYr2");
            dtFiltered.Columns.Add("LocalSalesYr3");
            dtFiltered.Columns.Add("LocalSalesYr4");
            dtFiltered.Columns.Add("ExportSalesYr1");
            dtFiltered.Columns.Add("ExportSalesYr2");
            dtFiltered.Columns.Add("ExportSalesYr3");
            dtFiltered.Columns.Add("ExportSalesYr4");
            dtFiltered.Columns.Add("TotalWoker1");
            dtFiltered.Columns.Add("TotalWoker2");
            dtFiltered.Columns.Add("TotalWoker3");
            dtFiltered.Columns.Add("TotalWoker4");
            SqlConnection con = SQLHelper.GetConnection();
            SqlCommand com = new SqlCommand();
            SqlDataAdapter ad = new SqlDataAdapter(com);
            System.Text.StringBuilder sql = new System.Text.StringBuilder();
            sql.AppendLine("SELECT[File ID] , a.[Company Name] , a.[Registration Number] , a.[Year Of Approval]");
            sql.AppendLine(", a.[Operational Status], f.Year , f.Revenue, f.LocalSales , f.ExportSales , f.LocalWorker");
            sql.AppendLine(", f.ForeignWorker, f.TotalWorker");
            sql.AppendLine("FROM FinancialAndWorkerForecast f");
            sql.AppendLine("Left Join[CRM_PRD].[dbo].[BI_MSC_Company] a on a.AccountID = f.AccountID");
            sql.AppendLine("where f.Year between 2012 and 2015");
            com.CommandText = sql.ToString();
            com.CommandType = CommandType.Text;
            com.Connection = con;
            com.CommandTimeout = int.MaxValue;

            try
            {
                DataTable dt = new DataTable();
                ad.Fill(dt);
                string LastComp = "";
                string LastYearCounted = "";

                string MSCFileID = "", RegNo = "", YearofApproval = "", OperationStatus = ""; 
                string RevenueYr1 = "", RevenueYr2 = "", RevenueYr3 = "", RevenueYr4 = "", LocalSalesYr1 = "", LocalSalesYr2 = "", LocalSalesYr3 = "", LocalSalesYr4 = "", ExportSalesYr1 = "", ExportSalesYr2 = "", ExportSalesYr3 = "", ExportSalesYr4 = "", TotalWoker1 = "", TotalWoker2 = "", TotalWoker3 = "", TotalWoker4 = "";

                foreach (DataRow dr in dt.Rows)
                {
                    string CurrentComp = dr["Company Name"].ToString();
                    string CurrenttYearCounted = dr["Year"].ToString();
                    if (LastComp != CurrentComp)
                    {
                        InsertInDT(ref dtFiltered, MSCFileID, LastComp, RegNo, YearofApproval, OperationStatus, RevenueYr1, RevenueYr2, RevenueYr3, RevenueYr4, LocalSalesYr1, LocalSalesYr2, LocalSalesYr3, LocalSalesYr4, ExportSalesYr1, ExportSalesYr2, ExportSalesYr3, ExportSalesYr4, TotalWoker1, TotalWoker2, TotalWoker3, TotalWoker4);
                        MSCFileID = ""; RegNo = ""; YearofApproval = ""; OperationStatus = "";
                        RevenueYr1 = ""; RevenueYr2 = ""; RevenueYr3 = ""; RevenueYr4 = ""; LocalSalesYr1 = ""; LocalSalesYr2 = ""; LocalSalesYr3 = ""; LocalSalesYr4 = ""; ExportSalesYr1 = ""; ExportSalesYr2 = ""; ExportSalesYr3 = ""; ExportSalesYr4 = ""; TotalWoker1 = ""; TotalWoker2 = ""; TotalWoker3 = ""; TotalWoker4 = "";
                    }
                    string Revenue = dr["Revenue"].ToString();
                    if (RevenueYr1 == "")
                        RevenueYr1 = Revenue;
                    if (RevenueYr2 == "")
                        RevenueYr2 = Revenue;
                    if (RevenueYr3 == "")
                        RevenueYr3 = Revenue;
                    if (MSCFileID == "")
                        MSCFileID = dr["File ID"].ToString();
                    if (LastComp == "")
                        LastComp = dr["Company Name"].ToString();
                    if (RegNo == "")
                        RegNo = dr["Registration Number"].ToString();
                    if (YearofApproval == "")
                        YearofApproval = dr["Year Of Approval"].ToString();
                    if (OperationStatus == "")
                        OperationStatus = dr["Operational Status"].ToString();


                    LastYearCounted = dr["Year"].ToString();
                    LastComp = dr["Company Name"].ToString();
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

            return dtFiltered;
        }

        private static void InsertInDT(ref DataTable dtFiltered, string MSCFileID, string LastComp, string regNo, string yearofApproval, string operationStatus, string revenueYr1, string revenueYr2, string revenueYr3, string revenueYr4, string localSalesYr1, string localSalesYr2, string localSalesYr3, string localSalesYr4, string exportSalesYr1, string exportSalesYr2, string exportSalesYr3, string exportSalesYr4, string TotalWoker1, string TotalWoker2, string TotalWoker3, string TotalWoker4)
        {
            DataRow DR = dtFiltered.NewRow();

            DR["MSCFileID"] = MSCFileID;
            DR["Company Name"] = LastComp;
            DR["Registration Number"] = regNo;
            DR["Year of Approval"] = yearofApproval;
            DR["Operation Status"] = operationStatus;
            DR["RevenueYr1"] = revenueYr1;
            DR["RevenueYr2"] = revenueYr2;
            DR["RevenueYr3"] = revenueYr3;
            DR["RevenueYr4"] = revenueYr4;
            DR["LocalSalesYr1"] = localSalesYr1;
            DR["LocalSalesYr2"] = localSalesYr2;
            DR["LocalSalesYr3"] = localSalesYr3;
            DR["LocalSalesYr4"] = localSalesYr4;
            DR["ExportSalesYr1"] = exportSalesYr1;
            DR["ExportSalesYr2"] = exportSalesYr1;
            DR["ExportSalesYr3"] = exportSalesYr1;
            DR["ExportSalesYr4"] = exportSalesYr1;
            DR["TotalWokerYr1"] = TotalWoker1;
            DR["TotalWokerYr2"] = TotalWoker2;
            DR["TotalWokerYr3"] = TotalWoker3;
            DR["TotalWokerYr4"] = TotalWoker4;
            dtFiltered.Rows.Add(DR);
        }
    }
}
