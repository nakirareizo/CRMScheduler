using CRMSync.Classes;
using CRMSync.Classes.MainProcess;
using System;
using System.Collections;
using System.Linq;

namespace CRMSync
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.BackgroundColor = ConsoleColor.DarkRed;
            //Console.WriteLine("Starto");
            //Console.ReadLine();
            //ApprovedMSCComp
            SyncApprovedAccount.StartSync(true);   
            //SyncApprovedAccount.StartSync(false);
            //APPROVEDLETTERDATEs
            //ApprovalLetterDate syncApprovalDates = new ApprovalLetterDate();
            //syncApprovalDates.StartSync();
            //IHL
            //SyncApprovedAccountIHL.StartSync();
            //ShareHolder
            //SyncShareHolder.StartSync();
            //UpdateContactPostMSCChange();
            //ShareHolderCleanUp.Start();
            //SyncApprovedAccountClean.StartSync();
            //UpdateAuditLog.StartSync();
            //PivotInfotechGen.Start();
            //MOF
            //MOFDataSync.Start();

        }


    }
}
