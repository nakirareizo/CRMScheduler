using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace CRMSync.Classes
{

    class LogFileHelper
    {
        public static ArrayList logList { get; set; }
        static internal void WriteLog(ArrayList logList, string ModeSync)
        {
            string[] strarr = null;
            string FolderPath = ConfigurationSettings.AppSettings["LogFileLocation"].ToString();
            if ((!System.IO.Directory.Exists(FolderPath)))
            {
                System.IO.Directory.CreateDirectory(FolderPath);
            }
            string FILE_NAME = "";
            if (IsDirectoryEmpty(FolderPath))
            {
                FILE_NAME = ConfigurationSettings.AppSettings["LogFileLocation"].ToString() + ModeSync + DateTime.Now.ToString("dd-MMM-yyyy") + ".txt";
            }
            else
            {
                int LastCounter = getLastFileCounter(FolderPath);
                strarr = (String[])logList.ToArray(typeof(string));

                if (LastCounter == 0)
                    FILE_NAME = ConfigurationSettings.AppSettings["LogFileLocation"].ToString() + ModeSync + DateTime.Now.ToString("dd-MMM-yyyy") + ".txt";
                else
                    FILE_NAME = ConfigurationSettings.AppSettings["LogFileLocation"].ToString() + ModeSync + DateTime.Now.ToString("dd-MMM-yyyy") + "_" + LastCounter.ToString() + ".txt";


            }
            System.IO.StreamWriter objWriter = new System.IO.StreamWriter(FILE_NAME);
            if (strarr != null)
            {
                foreach (string row1 in strarr)
                {
                    objWriter.WriteLine(row1);
                }
                objWriter.Close();
            }
        }
        private static bool IsDirectoryEmpty(string path)
        {
            return !Directory.EnumerateFileSystemEntries(path).Any();
        }

        private static int getLastFileCounter(string folderPath)
        {
            int output = 0;
            var directory = new DirectoryInfo(folderPath);
            var fileName = directory.GetFiles()
            .OrderByDescending(f => f.LastWriteTime)
            .First();

            string[] arrFile = fileName.ToString().Split('_');
            string sDate = arrFile[1].ToString().Substring(0, 11); //21-Mar-2016
            DateTime dDate = getDateTime(sDate);
            //1.txt, 10.txt
            string lastCounter = "";
            if (dDate.ToString("dd-MMM-yyyy") == DateTime.Now.ToString("dd-MMM-yyyy"))
            {
                if (arrFile.Count() > 2)
                {
                    if (arrFile[2].Length <= 5)
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
            string[] formats = { "dd-MMM-yyyy" };
            return myDate = DateTime.ParseExact(sDate, formats, new CultureInfo(Thread.CurrentThread.CurrentCulture.Name), DateTimeStyles.None);
        }
    }
}
