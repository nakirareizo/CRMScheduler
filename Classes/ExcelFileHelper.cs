using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace CRMSync.Classes
{
    class ExcelFileHelper
    {
        internal static void GenerateExcelFile(DataTable dt, string SyncedDate)
        {
            #region "EXPORT spBigFile INTO EXCEL"
            //Export spBigFile into EXCEL File
            string folderPath = ConfigurationSettings.AppSettings["ExcelLocation"].ToString();
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            string ExcelFile = "";
            string SyncType = "spBigFile";
            if (IsDirectoryEmpty(folderPath))
            {
                ExcelFile = ConfigurationSettings.AppSettings["ExcelLocation"].ToString() + "spBigFile_" + DateTime.Now.ToString("ddMMyyyy") + ".xlsx";
            }
            else
            {
                int LastCounter = getLastFileCounter(folderPath,SyncType);

                if (LastCounter == 0)
                    ExcelFile = ConfigurationSettings.AppSettings["ExcelLocation"].ToString() + "spBigFile_" + DateTime.Now.ToString("ddMMyyyy") + ".xlsx";
                else
                    ExcelFile = ConfigurationSettings.AppSettings["ExcelLocation"].ToString() + "spBigFile_" + DateTime.Now.ToString("ddMMyyyy") + "_" + LastCounter.ToString() + ".xlsx";
            }
            ConvertToExcel(dt, ExcelFile, SyncedDate);
            #endregion
        }

        internal static void GenerateExcelFileApprovalDates(DataTable dt, string SyncedDate)
        {
            #region "EXPORT spBigFile INTO EXCEL"
            //Export spBigFile into EXCEL File
            string folderPath = ConfigurationSettings.AppSettings["ExcelLocation"].ToString();
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            string ExcelFile = "";
            string SyncType = "ApprovalDates";
            if (IsDirectoryEmpty(folderPath))
            {
                ExcelFile = ConfigurationSettings.AppSettings["ExcelLocation"].ToString() + "ApprovalDates_" + DateTime.Now.ToString("ddMMyyyy") + ".xlsx";
            }
            else
            {
                
                int LastCounter = getLastFileCounter(folderPath, SyncType);

                if (LastCounter == 0)
                    ExcelFile = ConfigurationSettings.AppSettings["ExcelLocation"].ToString() + "ApprovalDates_" + DateTime.Now.ToString("ddMMyyyy") + ".xlsx";
                else
                    ExcelFile = ConfigurationSettings.AppSettings["ExcelLocation"].ToString() + "ApprovalDates_" + DateTime.Now.ToString("ddMMyyyy") + "_" + LastCounter.ToString() + ".xlsx";
            }
            ConvertToExcel(dt, ExcelFile, SyncedDate);
            #endregion
        }

        private static bool IsDirectoryEmpty(string path)
        {
            return !Directory.EnumerateFileSystemEntries(path).Any();
        }
        private static void ConvertToExcel(DataTable dt, string ExcelFile, string SyncedDate)
        {
            //"03/20/2016 17:10:57 PM"
            string[] formats = {
                "yyyy-MM-d HH:mm:ss tt",
                "M/d/yyyy HH:mm:ss tt",
                "M/d/yyyy HH:mm tt",
                "yyyy-MM-dd HH:mm:ss tt",
                "yyyy-dd-MM HH:mm:ss tt",
                "MM/dd/yyyy HH:mm:ss tt",
                "dd/MM/yyyy HH:mm:ss tt",
                "MM/dd/yyyy HH:mm:ss",
                "M/d/yyyy h:mm:ss",
                "M/d/yyyy hh:mm tt",
                "M/d/yyyy hh tt",
                "M/d/yyyy h:mm",
                "M/d/yyyy h:mm",
                "MM/dd/yyyy hh:mm",
                "M/dd/yyyy hh:mm",
                "MM/d/yyyy HH:mm:ss.ffffff",
            "dd-MM-yyyy"};
            DateTime myDate = DateTime.ParseExact(SyncedDate, formats, new CultureInfo(Thread.CurrentThread.CurrentCulture.Name), DateTimeStyles.None);
            // DateTime myDate = DateTime.ParseExact(SyncedDate, DateString,provider, CultureInfo.InvariantCulture);

            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dt, "spBigFile_" + myDate.ToString("ddMMyyyy"));
                wb.SaveAs(ExcelFile);
            }
        }

        private static int getLastFileCounter(string folderPath, string SyncType)
        {
            int output = 0;
            var directory = new DirectoryInfo(folderPath);
            var fileName = directory.GetFiles()
            .OrderByDescending(f => f.LastWriteTime)
            .First();

            string[] arrFile = fileName.ToString().Split('_');
            string sDate = arrFile[1].ToString().Substring(0, 8); //21032016
            DateTime dDate = getDateTime(sDate);
            //1.xlsx, 10.xlsx
            string lastCounter = "";
            if (dDate.ToString("ddMMyyyy") == DateTime.Now.ToString("ddMMyyyy"))
            {
                if (arrFile.Count() > 2)
                {
                    if (arrFile[2].Length <= 6)
                        lastCounter = arrFile[2].Substring(0, 1);
                    else
                        lastCounter = arrFile[2].Substring(0, 2);
                }
                if (arrFile.Length == 2)
                    output = 1;
                else if (Convert.ToInt32(lastCounter) == 10)
                {
                    output = 0;
                }
                else
                {
                    output = Convert.ToInt32(lastCounter) + 1;
                }
            }
            else
                output = 0;
            return output;
        }
        private static DateTime getDateTime(string sDate)
        {
            DateTime myDate = new DateTime();
            string[] formats = { "ddMMyyyy" };
            return myDate = DateTime.ParseExact(sDate, formats, new CultureInfo(Thread.CurrentThread.CurrentCulture.Name), DateTimeStyles.None);
        }
    }
}
