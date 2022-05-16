using MiTools;
using OutlookLib;
using RPABaseAPI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using Excel = Microsoft.Office.Interop.Excel;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Outlook_and_Excel
{
    public class SampleOutlook : MyCartesProcessBase
    {
        private string fExcelFile;
        private OutlookAPI fOutlook;

        public SampleOutlook() : base()
        {
            fExcelFile = null;
            fOutlook = null;
            ShowAbort = true;
        }

        protected override string getNeededRPASuiteVersion()
        {
            return "3.2.1.0";
        }
        protected override string getRPAMainFile()
        {
            return CurrentPath + "\\Cartes\\Outlook and Excel.cartes.rpa";
        }
        protected override void LoadConfiguration(XmlDocument XmlCfg)
        {
            fExcelFile = ToString(XmlCfg.SelectSingleNode("//emails"));
            if (File.Exists(fExcelFile))
                fExcelFile = Path.GetFullPath(fExcelFile);
            else throw new Exception("\"" + fExcelFile + "\" is not a file");
        }
        protected override void DoExecute(ref DateTime start)
        {
            Excel.Application objExl;
            bool ReadOnly = true, lbExit;
            int row, i, j;
            string EmailTo, EmailSubject, EmailMessage;

            objExl = new Excel.Application();
            try
            {
                objExl.DisplayAlerts = false;
                objExl.Visible = true;
                Excel.Workbook wb = objExl.Workbooks.Open(ExcelFile, false, ReadOnly);
                try
                {
                    Excel.Worksheet xlWorkSheet;

                    xlWorkSheet = wb.Worksheets.get_Item(1);
                    Outlook.VisibleMode = true;
                    // I go through Excel to send all the emails
                    lbExit = false;
                    row = 2;
                    do
                    {
                        CheckAbort();
                        EmailTo = ToString(xlWorkSheet.Cells[row, "A"].Text).Trim();
                        if (EmailTo.Length == 0) lbExit = true;
                        else
                        {
                            Balloon(EmailTo);
                            EmailSubject = ToString(xlWorkSheet.Cells[row, "B"].Text).Trim();
                            EmailMessage = ToString(xlWorkSheet.Cells[row, "C"].Text).Trim();
                            Outlook.SendEmail(EmailTo, EmailSubject, EmailMessage);
                        }
                        row++;
                    } while (!lbExit);
                    // I go through all the emails in the inbox
                    List<Outlook.MAPIFolder> folders = new List<Outlook.MAPIFolder>();
                    folders.Add(Outlook.Inbox);
                    j = 0;
                    i = 0;
                    while ((i < folders.Count) && (j < 200))
                    {
                        CheckAbort();
                        Outlook.MAPIFolder folder = folders.ElementAt(i);
                        row = 1; 
                        while (row <= folder.Folders.Count) // I add all the folders
                        {
                            folders.Add(Outlook.Inbox.Folders[row]);
                            row++;
                        }
                        row = 1;
                        while (row <= folder.Items.Count) // I show all the emails
                        {
                            CheckAbort();
                            dynamic item = folder.Items[row];
                            if (item is Outlook.MailItem mail)
                                Balloon(folder.Name + LF + mail.SenderEmailAddress + LF + mail.Subject);
                            else if (item is Outlook.ReportItem report)
                                Balloon(folder.Name + LF + report.Subject);
                            row++;
                            j++;
                        }
                        i++;
                    }
                }
                finally
                {
                    wb.Close(false);
                }
            }
            finally
            {
                try
                {
                    objExl.Visible = false;
                }
                finally
                {
                    try { objExl.Quit(); } catch { }
                }
            }
        }

        public string ExcelFile
        {
            get { return fExcelFile; }
        }
        public OutlookAPI Outlook
        {
            get
            {
                if (fOutlook == null)
                    fOutlook = new OutlookAPI(this);
                return fOutlook;
            }
        }
    }
}
