using Cartes;
using MiTools;
using RPABaseAPI;
using System;
using System.Runtime.InteropServices;
using System.Threading;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookLib
{
    public class OutlookAPI : MyCartesAPIBase
    {
        private Outlook.Application fOutlook;
        private Outlook.NameSpace fSpace;
        private Outlook.MAPIFolder fInbox, fDeleted;
        private Thread fThStack = null;
        private bool fPreviuos, fVisibleMode;
        private RPAWin32Component OutlookMain = null;

        public OutlookAPI(MyCartesProcess owner) : base(owner)
        {
            fPreviuos = false;
            fVisibleMode = true;
            fOutlook = null;
            fSpace = null;
            fInbox = null;
            fDeleted = null;
        }
        ~OutlookAPI()
        {
            try
            {
                CR.WaitOne();
                try
                {
                    fThStack = null;
                    if (fOutlook != null)
                    {
                        if (!fPreviuos) Close();
                        else
                        {
                            Annul(ref fDeleted);
                            Annul(ref fInbox);
                            Annul(ref fSpace);
                            Annul(ref fOutlook);
                            fPreviuos = false;
                        }
                    }
                }
                finally
                {
                    CR.ReleaseMutex();
                }
            }
            catch (Exception e)
            {
                forensic("OutlookAPI.~OutlookAPI", e);
                throw;
            }
        }

        private void Annul(ref Outlook.Application value)
        {
            if (value != null)
            {
                Marshal.FinalReleaseComObject(value);
                value = null;
            }
        }
        private void Annul(ref Outlook.NameSpace value)
        {
            if (value != null)
            {
                Marshal.FinalReleaseComObject(value);
                value = null;
            }
        }
        private void Annul(ref Outlook.MAPIFolder value)
        {
            if (value != null)
            {
                Marshal.FinalReleaseComObject(value);
                value = null;
            }
        }
        private void InitThread()
        {
            CR.WaitOne();
            try
            {
                if ((fThStack == null) || !fThStack.IsAlive)
                {
                    fThStack = new Thread(ProcessBackGroundButtons);
                    fThStack.IsBackground = true;
                    fThStack.Start();
                }
            }
            finally
            {
                CR.ReleaseMutex();
            }
        }
        private Outlook.Application GetOutlook()
        {
            try
            {
                if (fOutlook == null)
                {
                    CR.WaitOne();
                    try
                    {
                        if (fOutlook == null)
                        {
                            reset(OutlookMain);
                            fPreviuos = OutlookMain.ComponentExist();
                            if (fPreviuos && (StringIn(OutlookMain.WindowState, "Minimized", "Maximized") || (OutlookMain.Visible == 0)))
                                    OutlookMain.Show("Restore");
                            fOutlook = new Outlook.Application();
                            InitThread();
                            if (!fPreviuos && VisibleMode)
                                Inbox.Display();
                        }
                    }
                    finally
                    {
                        CR.ReleaseMutex();
                    }
                }
            }
            catch (Exception e)
            {
                forensic("OutlookAPI.GetOutlook", e);
                throw;
            }
            return fOutlook;
        }
        private Outlook.NameSpace GetSpace()
        {
            try
            {
                if (fSpace == null)
                {
                    CR.WaitOne();
                    try
                    {
                        if (fSpace == null)
                            fSpace = Application.GetNamespace("MAPI");
                    }
                    finally
                    {
                        CR.ReleaseMutex();
                    }
                }
            }
            catch (Exception e)
            {
                forensic("OutlookAPI.GetSpace", e);
                throw;
            }
            return fSpace;
        }
        private Outlook.MAPIFolder GetInbox()
        {
            if (fInbox == null)
                fInbox = GetFolder(Outlook.OlDefaultFolders.olFolderInbox);
            return fInbox;
        }
        private Outlook.MAPIFolder GetDeleted()
        {
            if (fDeleted == null)
                fDeleted = GetFolder(Outlook.OlDefaultFolders.olFolderDeletedItems);
            return fDeleted;
        }
        private Outlook.MAPIFolder GetFolder(Outlook.OlDefaultFolders FolderType)
        {
            return Space.GetDefaultFolder(FolderType);
        }
        private Outlook.Folders GetFolders()
        {
            return Space.Folders;
        }

        protected override void MergeLibrariesAndLoadVariables()
        {
            if (!isVariable("$OutlookMain"))
            {
                if (cartes.merge(CurrentPath + "\\Cartes\\Outlook.cartes.rpa") != 1)
                    throw new Exception("I cannot load the Outlook library");
                OutlookMain = GetComponent<RPAWin32Component>("$OutlookMain");
            }
        }
        protected override void UnMergeLibrariesAndUnLoadVariables()
        {
            try
            {
                try
                {
                    fThStack = null;
                }
                finally
                {
                    if (!fPreviuos && (OutlookMain != null))
                        Close();
                    OutlookMain = null;
                }
            }
            catch (Exception e)
            {
                forensic("OutlookAPI.UnMergeLibrariesAndUnLoadVariables", e);
                throw;
            }
        }
        protected override string getNeededRPASuiteVersion() 
        {
            return "3.4.2.0";
        }
        protected virtual void ProcessBackGroundButtons() /* This is the method of a thread responsible for clicking the
            buttons that allow access to Outlook from DCOM. */
        {
            var Cartes = new CartesObj();
            RPAWin32Automation OutlookAllowBtnES = null;
            RPAWin32CheckRadioButton OutlookAllowAccessES = null;
            bool Ready = false;
            DateTime timereset = DateTime.Now.AddSeconds(30);

            while (fThStack != null)
            {
                try
                {
                    if (true && (Ready || Cartes.isVariable("$OutlookMain")))
                    {
                        if (OutlookAllowAccessES == null)
                        {
                            OutlookAllowBtnES = Cartes.GetComponent<RPAWin32Automation>("$OutlookAllowES");
                            OutlookAllowAccessES = Cartes.GetComponent<RPAWin32CheckRadioButton>("$OutlookAllowAccessES");
                        }
                        Ready = true;
                        CR.WaitOne();
                        try
                        {
                            if (OutlookAllowBtnES.ComponentExist() && (OutlookAllowBtnES.enabled() != 0) &&
                                StringIn(OutlookAllowBtnES.name(), "permitir", "allow"))
                            {
                                if (OutlookAllowAccessES.ComponentExist())
                                    OutlookAllowAccessES.Checked = 1;
                                OutlookAllowBtnES.click();
                            }
                            else if (timereset < DateTime.Now)
                            {
                                Cartes.reset(OutlookAllowBtnES);
                                timereset = DateTime.Now.AddSeconds(30);
                            }
                        }
                        finally
                        {
                            CR.ReleaseMutex();
                        }
                    }
                }
                catch (Exception e)
                {
                    forensic("OutlookAPI.ProcessBackGroundButtons", e);
                }
                Thread.Sleep(500);
            }
        }
        protected Outlook.Application Application
        {
            get { return GetOutlook(); }
        }
        protected Outlook.NameSpace Space
        {
            get { return GetSpace(); }
        }

        public delegate bool ProcessMail(Outlook.MailItem mail);
        public override void Close()
        {
            try
            {
                CR.WaitOne();
                try
                {
                    if (fOutlook != null)
                    {
                        fThStack = null;
                        Annul(ref fDeleted);
                        Annul(ref fInbox);
                        Annul(ref fSpace);
                        fOutlook.Quit();
                        Annul(ref fOutlook);
                        fPreviuos = false;
                    }
                }
                finally
                {
                    CR.ReleaseMutex();
                }
            }
            catch (Exception e)
            {
                forensic("OutlookAPI.Close", e);
                throw;
            }
        }
        public override void SendEmail(string to, string subject, string body)
        {
            try
            {
                if (to.Length == 0) throw new Exception("The empty string is not a valid email address.");
                Outlook.MailItem mailItem = NewEmail();
                mailItem.Subject = subject;
                mailItem.Body = body;
                mailItem.To = to;
                mailItem.Send();
            }
            catch (Exception e)
            {
                forensic("OutlookAPI.SendEmail", e);
                throw;
            }
        }
        public virtual Outlook.MAPIFolder GetFolder(string folder) /* Returns the working folder that matches the specified
            path. You can use the split bar to indicate a routing. For example: bots@acme.net/inbox */
        {
            string[] path = folder.ToLower().Replace('/', '\\').Split("\\", StringSplitOptions.None);

            Outlook.MAPIFolder Find(Outlook.Folders carpeta, int indice)
            {
                Outlook.MAPIFolder resultado = null;

                int i = 1;
                while ((i <= carpeta.Count) && (resultado == null))
                {
                    Outlook.MAPIFolder subcarpeta = carpeta[i];
                    if (ToString(subcarpeta.Name).ToLower() == path[indice])
                    {
                        if (indice == path.Length - 1) resultado = subcarpeta;
                        else resultado = Find(subcarpeta.Folders, indice + 1);
                    }
                    i++;
                }
                return resultado;
            }
            return Find(Folders, 0);
        }
        public virtual Outlook.MailItem NewEmail()
        {
            return Application.CreateItem(Outlook.OlItemType.olMailItem);
        }
        public virtual bool ProcessMailFrom(Outlook.MAPIFolder folder, ProcessMail processor, bool subfolders = false) /* The method
            will call the "processor" function with all the mails in the folder while "processor" returns true. If "subfolder"
            is true it will include the subfolders. The method will return the value returned in the last call to "processor". */
        {
            bool resultado;
            int i;

            resultado = true;
            if ((folder != null) && (processor != null)) try
                {
                    i = folder.Items.Count; // It is very important to start with the last email. If the email is deleted, the count would be lost going forward.
                    while ((0 < i) && resultado && !IsAborting)
                    {
                        dynamic item = folder.Items[i];
                        if (item is Outlook.MailItem mail)
                            resultado = processor(mail);
                        else if (item is Outlook.ReportItem report)
                        {
                            // Nothing to do
                        }
                        i--;
                    }
                    if (subfolders)
                    {
                        i = 1;
                        while ((i <= folder.Folders.Count) && resultado && !IsAborting)
                        {
                            resultado = ProcessMailFrom(folder.Folders[i], processor, subfolders);
                            i++;
                        }
                    }
                }
                catch (Exception ex)
                {
                    forensic("OutlookAPI.ProcessMailFrom", ex);
                    throw;
                }
            return resultado;
        }
        public bool ProcessMailFrom(string folder, ProcessMail processor, bool subfolders = false)
        {
            return ProcessMailFrom(GetFolder(folder), processor, subfolders);
        }

        public bool VisibleMode
        {
            get { return fVisibleMode; }
            set { fVisibleMode = value; }
        }
        public Outlook.MAPIFolder Inbox
        {
            get { return GetInbox(); }
        }
        public Outlook.MAPIFolder DeletedItems
        {
            get { return GetDeleted(); }
        }
        public Outlook.MAPIFolder this[Outlook.OlDefaultFolders FolderType]
        {
            get { return GetFolder(FolderType); }
        }
        public Outlook.Folders Folders
        {
            get { return GetFolders(); }
        }
    }
}
