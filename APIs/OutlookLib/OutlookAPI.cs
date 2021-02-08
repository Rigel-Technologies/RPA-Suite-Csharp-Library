using Cartes;
using MiTools;
using RPABaseAPI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookLib
{
    public class OutlookAPI : MyCartesAPIBase
    {
        private Mutex fRC = new Mutex();
        private Outlook.Application fOutlook;
        private Outlook.NameSpace fSpace;
        private Outlook.MAPIFolder fInbox, fDeleted;
        private Thread fThStack = null;
        private bool fPreviuos, fVisibleMode;
        private RPAWin32Component OutlookMain = null;
        private RPAWin32Automation OutlookAllowBtnES = null;
        private RPAWin32CheckRadioButton OutlookAllowAccessES = null;

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
            RC.WaitOne();
            try
            {
                fThStack = null;
                if (!fPreviuos && (fOutlook != null)) Close();
                else
                {
                    fDeleted = null;
                    fInbox = null;
                    fSpace = null;
                    fOutlook = null;
                    fPreviuos = false;
                }
            }
            finally
            {
                RC.ReleaseMutex();
            }
        }

        private Mutex GetRC()
        {
            return fRC;
        }
        private void InitThread()
        {
            RC.WaitOne();
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
                RC.ReleaseMutex();
            }
        }
        private Outlook.Application GetOutlook()
        {
            try
            {
                if (fOutlook == null)
                {
                    RC.WaitOne();
                    try
                    {
                        if (fOutlook == null)
                        {
                            fOutlook = new Outlook.Application();
                            reset(OutlookMain);
                            fPreviuos = OutlookMain.ComponentExist();
                            InitThread();
                            if (!fPreviuos && VisibleMode)
                                Inbox.Display();
                        }
                    }
                    finally
                    {
                        RC.ReleaseMutex();
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
                    RC.WaitOne();
                    try
                    {
                        if (fSpace == null)
                            fSpace = Application.GetNamespace("MAPI");
                    }
                    finally
                    {
                        RC.ReleaseMutex();
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
                OutlookAllowBtnES = GetComponent<RPAWin32Automation>("$OutlookAllowES");
                OutlookAllowAccessES = GetComponent<RPAWin32CheckRadioButton>("$OutlookAllowAccessES");
            }
        }
        protected override void UnMergeLibrariesAndUnLoadVariables()
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
        protected virtual void ProcessBackGroundButtons() /* This is the method of a thread responsible for clicking the
            buttons that allow access to Outlook from DCOM. */
        {
            DateTime timereset = DateTime.Now.AddSeconds(30);
            while (fThStack != null)
            {
                try
                {
                    RC.WaitOne();
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
                            reset(OutlookAllowBtnES);
                            timereset = DateTime.Now.AddSeconds(30);
                        }
                    }
                    finally
                    {
                        RC.ReleaseMutex();
                    }
                }
                catch (Exception e)
                {
                    forensic("OutlookAPI.ProcessBackGroundButtons", e);
                }
                Thread.Sleep(500);
            }
        }
        protected Mutex RC // The Critical section
        {
            get { return GetRC(); }
        }
        protected Outlook.Application Application
        {
            get { return GetOutlook(); }
        }
        protected Outlook.NameSpace Space
        {
            get { return GetSpace(); }
        }

        public override void Close()
        {
            RC.WaitOne();
            try
            {
                Application.Quit();
                fThStack = null;
                fDeleted = null;
                fInbox = null;
                fSpace = null;
                fOutlook = null;
                fPreviuos = false;
            }
            finally
            {
                RC.ReleaseMutex();
            }
        }
        public virtual Outlook.MailItem NewEmail()
        {
            return Application.CreateItem(Outlook.OlItemType.olMailItem);
        }
        public virtual void SendEmail(string EmailTo, string EmailSubject, string EmailMessage)
        {
            try
            {
                if (EmailTo.Length == 0) throw new Exception("The empty string is not a valid email address.");
                Outlook.MailItem mailItem = NewEmail();
                mailItem.Subject = EmailSubject;
                mailItem.Body = EmailMessage;
                mailItem.To = EmailTo;
                mailItem.Send();
            }catch(Exception e)
            {
                forensic("OutlookAPI.SendEmail", e);
                throw;
            }
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
