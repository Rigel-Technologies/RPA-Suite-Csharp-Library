using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Cartes;
using MiTools;
using RPABaseAPI;
using System.Threading;

namespace ChromeLib
{
    public class Chrome : MyCartesAPIBase
    {
        private static bool loaded = false;
        private bool fSpanish, fIncognito;
        private int fHeight, fWidth;
        private string fCertificate;
        private Thread fThStack = null, fTh = null;
        private CredentialStack fPrxyPsw;
        protected RPAWin32Component chm = null, chmURLEdit = null, chmTabs = null, chmClose = null, chmCloseCRC = null;
        protected RPAWin32Component chmProxy = null, chmProxyPsw = null;
        protected RPAWin32Component chmCertificateTitle = null;

        public Chrome(MyCartesProcess owner) : base(owner)
        {
            fPrxyPsw = null;
            fSpanish = true;
            fIncognito = false;
            fWidth = 985;
            fHeight = 732;
            fCertificate = string.Empty;
        }
        ~Chrome()
        {
            SetProxyPassword(null);
        }
        private string CheckURL(string URL)
        {
            string murl = ToString(URL).ToLower();
            if (murl.Contains("http://")) return URL;
            else if (murl.Contains("http:/")) return CheckURL(URL.Substring(6));
            else if (murl.Contains("http:")) return CheckURL(URL.Substring(5));
            else if (murl.Contains("https://")) return URL;
            else if (murl.Contains("https:/")) return CheckURL(URL.Substring(7));
            else if (murl.Contains("https:")) return CheckURL(URL.Substring(6));
            else if (murl.Contains("chrome:")) return URL;
            else return "http://" + URL;
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
        private void InitTask()
        {
            CR.WaitOne();
            try
            {
                if ((fTh == null) || !fTh.IsAlive)
                {
                    fTh = new Thread(ProcessBackGroundTask);
                    fTh.IsBackground = true;
                    fTh.Start();
                }
            }
            finally
            {
                CR.ReleaseMutex();
            }
        }
        private void ProcessBackGroundButtons() /* This is the method of a thread responsible for clicking the
            buttons that allow access to Proxy. */
        {
            const int delay = 30;
            var cartes = new CartesObj();
            DateTime timereset = DateTime.Now.AddSeconds(delay);
            while (fThStack != null)
            {
                try
                {
                    if (timereset < DateTime.Now)
                    {
                        CR.WaitOne();
                        try
                        {
                            cartes.reset(chmProxy);
                        }
                        finally
                        {
                            CR.ReleaseMutex();
                        }
                        timereset = DateTime.Now.AddSeconds(delay);
                    }
                    else
                    {
                        CR.WaitOne();
                        try
                        {
                            CheckProxy(cartes);
                        }
                        finally
                        {
                            CR.ReleaseMutex();
                        }
                    }
                }
                catch (Exception e)
                {
                    forensic("Chrome.ProcessBackGroundButtons", e);
                }
                Thread.Sleep(500);
            }
        }
        private void ProcessBackGroundTask() /* This is the method of a thread responsible for clicking various
            warning buttons. */
        {
            const int delay = 30;
            var cartes = new CartesObj();
            RPAWin32Component chmNoticeTranslateClose = null;
            DateTime timereset = DateTime.Now.AddSeconds(delay);

            chmNoticeTranslateClose = cartes.GetComponent<RPAWin32Component>("$ChromeNoticeTranslateClose");
            while (fTh != null)
            {
                try
                {
                    if (timereset < DateTime.Now)
                    {
                        CR.WaitOne();
                        try
                        {
                            cartes.reset(chmNoticeTranslateClose);
                        }
                        finally
                        {
                            CR.ReleaseMutex();
                        }
                        timereset = DateTime.Now.AddSeconds(delay);
                    }
                    else
                    {
                        CR.WaitOne();
                        try
                        {
                            if (Certificate.Length > 0)
                                SelectCertificate(cartes, Certificate);
                            if (chmNoticeTranslateClose.ComponentExist() && StringIn(chmNoticeTranslateClose.name(), "cerrar", "close"))
                                chmNoticeTranslateClose.click();
                        }
                        finally
                        {
                            CR.ReleaseMutex();
                        }
                    }
                }
                catch (Exception e)
                {
                    forensic("Chrome.ProcessBackGroundTask", e);
                }
                Thread.Sleep(500);
            }
        }
        private bool GetProxyExists(RPAWin32Component Proxy, RPAWin32Component ProxyPsw)
        {
            return Proxy.ComponentExist() && ProxyPsw.ComponentExist() && Proxy.Inside(ProxyPsw);
        }
        private bool GetOpenDigitalCertificateWindow(RPAWin32Component chmCertificateTitle)
        {
            return chmCertificateTitle.ComponentExist() && ToString(chmCertificateTitle.name()).ToLower().Contains("seleccionar un certificado");
        }
        private bool Spanish
        {
            get { return fSpanish; }
        }

        protected override void MergeLibrariesAndLoadVariables()
        {
            if (!isVariable("$Chrome"))
            {
                loaded = cartes.merge(CurrentPath + "\\Cartes\\Chrome Es.rpa") == 1;
            }
            if (chmCertificateTitle == null)
            {
                chm = GetComponent<RPAWin32Component>("$Chrome");
                chmURLEdit = GetComponent<RPAWin32Component>("$ChromeURLEdit");
                chmTabs = GetComponent<RPAWin32Component>("$ChromeTabs");
                chmClose = GetComponent<RPAWin32Component>("$ChromeClose");
                chmCloseCRC = GetComponent<RPAWin32Component>("$ChromeCloseCRC");
                chmProxy = GetComponent<RPAWin32Component>("$ChromePrx");
                chmProxyPsw = GetComponent<RPAWin32Component>("$ChromePrxPsw");
                chmCertificateTitle = GetComponent<RPAWin32Component>("$ChromeCertificateTitle");
                InitTask();
            }
        }
        protected virtual int GetHeight()
        {
            return fHeight;
        }
        protected virtual int GetWidth()
        {
            return fWidth;
        }
        protected virtual string GetCertificate()
        {
            return fCertificate;
        }
        protected virtual void SetCertificate(string value)
        {
            fCertificate = ToString(value);
        }
        protected virtual bool GetIncognito()
        {
            return fIncognito;
        }
        protected virtual void SetIncognito(bool value)
        {
            fIncognito = value;
        }
        protected virtual CredentialStack GetProxyPassword()
        {
            CR.WaitOne();
            try
            {
                return fPrxyPsw;
            }
            finally
            {
                CR.ReleaseMutex();
            }
        }
        protected virtual void SetProxyPassword(CredentialStack value)
        {
            CR.WaitOne();
            try
            {
                fPrxyPsw = value;
                if (fPrxyPsw == null) fThStack = null;
                else InitThread();
            }
            finally
            {
                CR.ReleaseMutex();
            }
        }
        protected virtual bool CheckProxy(CartesObj cartes) /* If the window to set the proxy credentials is visible, this method fills the form
            with the "ProxyPassword" credentials and returns true. Otherwise, it returns false. */
        {
            void NeverSavePassword(RPAWin32Component chmSvPsw)
            {
                try
                {
                    while (chmSvPsw.ComponentExist(3) && StringIn(chmSvPsw.name(), "Nunca"))
                    {
                        CheckAbort();
                        chmSvPsw.click();
                        cartes.reset(chmSvPsw);
                        Thread.Sleep(250);
                    }
                }
                catch (Exception e)
                {
                    forensic("Chrome.NeverSavePassword", e);
                    throw;
                }
            }

            bool result = false;
            try
            {
                if (ProxyPassword != null)
                {
                    CR.WaitOne();
                    try
                    {
                        RPAWin32Component chmProxy = null, chmProxyPsw = null;
                        chmProxy = cartes.GetComponent<RPAWin32Component>("$ChromePrx");
                        chmProxyPsw = cartes.GetComponent<RPAWin32Component>("$ChromePrxPsw");
                        while (GetProxyExists(chmProxy, chmProxyPsw) && StringIn(ToString(chmProxy.name()).ToLower(), "iniciar sesión"))
                        {
                            RPAWin32Component chmProxyUser = null, chmProxyAceptar = null;

                            chmProxyUser = cartes.GetComponent<RPAWin32Component>("$ChromePrxUser");
                            chmProxyAceptar = cartes.GetComponent<RPAWin32Component>("$ChromePrxIniciar");
                            CheckAbort();
                            chmProxyUser.Value = ProxyPassword.User;
                            ProxyPassword.Write(chmProxyPsw);
                            chmProxyAceptar.click();
                            cartes.reset(chmProxy);
                            NeverSavePassword(cartes.GetComponent<RPAWin32Component>("$ChromeSavePSW"));
                            result = true;
                            Thread.Sleep(250);
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
                forensic("ProcessBackGroundButtons.CheckProxy", e);
                throw;
            }
            return result;
        }
        protected virtual bool SelectCertificate(CartesObj cartes, string certificate) /* If the window to choose digital certificate is visible,
            this method chooses the indicated certificate and returns true. Otherwise, it returns false. */
        {
            bool lbOK = false;

            CR.WaitOne();
            try
            {
                if (GetOpenDigitalCertificateWindow(cartes.GetComponent<RPAWin32Component>("$ChromeCertificateTitle")))
                {
                    RPAWin32Component chmCertificateGrid = null, chmCertificateOK = null, chmCertificateCancel = null;
                    chmCertificateGrid = cartes.GetComponent<RPAWin32Component>("$ChromeCertificateGrid");
                    chmCertificateOK = cartes.GetComponent<RPAWin32Component>("$ChromeCertificateOK");
                    chmCertificateCancel = cartes.GetComponent<RPAWin32Component>("$ChromeCertificateCancel");
                    string ls = ToString(certificate).ToLower();
                    int i = 0;
                    while ((i < chmCertificateGrid.descendants) && !lbOK)
                    {
                        var child = chmCertificateGrid.Child<RPAWin32Component>(i);
                        if ((child.IdRole() == 28) && (ToString(child.dochild("\\0", "name")).ToLower() == ls))
                        {
                            RPAParameters parameter = new RPAParameters();
                            parameter.itemAsInteger[0] = 1;
                            child.dochild("\\0", "click", parameter);
                            lbOK = true;
                        }
                        else i++;
                    }
                    if (lbOK) chmCertificateOK.click();
                    else
                    {
                        string msg;
                        chmCertificateCancel.click();
                        if (ls.Length > 0) msg = "The \"" + certificate + "\" certificate is not present in the browser.";
                        else msg = "You need a digital certificate to access this site.";
                        cartes.balloon(msg);
                        cartes.forensic(msg);
                    }
                    cartes.reset(chmCertificateGrid.api());
                    chmCertificateGrid.ComponentNotExist(5);
                }
            }
            finally
            {
                CR.ReleaseMutex();
            }
            return lbOK;
        }

        public override void Close() // It closes Chrome.
        {
            Execute("$cshChrome00 = new TChrome(" + Owner.Abort + ");\r\n" +
                    "$cshChrome00.closeAll();");
        }
        public virtual void SetDimensions(int Width, int Height) // It sets the dimensions to which the main window will be adjusted
        {
            try
            {
                if ((Width < 1) || (Height < 1))
                    throw new Exception("The dimensions of a window cannot be less than 1 pixel.");
                fWidth = Width;
                fHeight = Height;
            }
            catch (Exception e)
            {
                forensic("Chrome::SetDimensions", e);
                throw;
            }
        }
        public virtual void AdjustWindow() /* The method must adjust the main window of the application to the dimensions indicated in the Width
           and Height properties. */
        {
            AdjustWindow(chm, Width, Height);
        }
        public virtual bool GetProxyExists() // This method returns true if the proxy credentials window is visible.
        {
            return GetProxyExists(chmProxy, chmProxyPsw);
        }
        public virtual bool GetOpenDigitalCertificateWindow() // This method returns true if the Digital Certificate Window is visible.
        {
            bool result;
            CR.WaitOne();
            try
            {
                result = GetOpenDigitalCertificateWindow(chmCertificateTitle);
            }
            finally
            {
                CR.ReleaseMutex();
            }
            return result;
        }
        public virtual void OpenURL(string URL, params IRPAComponent[] Components) /* It opens the indicated web page. Components must be a list of components
              of the page that indicates when the page has been loaded: for example, "$googlelogo". */
        {
            bool spread, exit, lbAdjust;
            DateTime timeout;

            string TxtIncognito()
            {
                return Incognito ? " --incognito" : "";
            }

            spread = false;
            lbAdjust = false;
            timeout = DateTime.Now.AddSeconds(120);
            exit = false;
            do
            {
                Balloon("Opening Chrome");
                Thread.Sleep(500);
                reset(chm);
                CheckAbort();
                if (timeout < DateTime.Now) throw new Exception("I can not open Chrome");
                else try
                    {
                        if (chm.ComponentExist())
                        {
                            CR.WaitOne();
                            try
                            {
                                if (GetProxyExists()) CheckProxy(cartes);
                                else if (chmClose.ComponentExist() && chmCloseCRC.ComponentExist() &&
                                    ((ToString(chmClose.name()).Length > 0) || (ToString(chmCloseCRC.name()).Length > 0)) &&
                                    (!StringIn(ToString(chmClose.name()), "cerrar") || !StringIn(ToString(chmCloseCRC.name()), "cerrar"))
                                   )
                                {
                                    fSpanish = false;
                                    spread = true;
                                    throw new MyException(EXIT_ERROR, "ERROR! Your Chrome or your Windows are not in Spanish. This library learned in Spanish buttons.");
                                }
                                else if (lbAdjust || ((Components != null) && ComponentsExist(0, Components)))
                                {
                                    RPAWin32Component chrome = chmURLEdit.Root();
                                    fSpanish = true;
                                    if (!StringIn(ToString(chrome.WindowState), "normal"))
                                        chrome.Show("restore");
                                    else
                                    {
                                        AdjustWindow();
                                        ControlTab(timeout);
                                        exit = lbAdjust || ComponentsExist(0, Components);
                                    }
                                }
                                else if (chmURLEdit.ComponentExist())
                                {
                                    RPAWin32Component chrome = chmURLEdit.Root();

                                    if (!StringIn(ToString(chrome.WindowState), "normal"))
                                        chrome.Show("restore");
                                    else
                                    {
                                        chmURLEdit.focus();
                                        reset(chm);
                                        if ((Components == null) || !ComponentsExist(0, Components))
                                        {
                                            ControlTab(timeout);
                                            chmURLEdit.TypeFromClipboardCheck(CheckURL(URL), 1, 0);
                                            chmURLEdit.TypeKey("Enter");
                                            Thread.Sleep(500);
                                            if (Components == null) lbAdjust = true;
                                            else ComponentsExist(30, Components);
                                        }
                                    }
                                }
                                else
                                {
                                    /* I do not have $ChromeURLEdit, but I have $Chrome */
                                    fSpanish = true;
                                    if (!StringIn(ToString(chm.WindowState), "normal"))
                                        chm.Show("restore");
                                    else Close(); // It is clear that I have a Chrome at least with a rare configuration
                                }
                            }
                            finally
                            {
                                CR.ReleaseMutex();
                            }
                        }
                        else if (ToString(URL).Trim().Length > 0)
                        {
                            cartes.run("chrome.exe" + TxtIncognito() + " \"" + ToString(URL).Trim() + "\"");
                            reset(chm);
                            Thread.Sleep(500);
                            Balloon("Waiting for Chrome");
                            List<IRPAComponent> list = new List<IRPAComponent>();
                            if (Components == null) list.Add(chm);
                            else list.AddRange(Components.ToList());
                            list.Add(chmProxy);
                           exit= ComponentsExist(30, list.ToArray());
                        }
                        else
                        {
                            cartes.run("chrome.exe" + TxtIncognito());
                            reset(chm);
                            Thread.Sleep(500);
                           exit= ComponentsExist(30, chmProxy, chmURLEdit);
                        }
                    }
                    catch (Exception e)
                    {
                        Balloon(e.Message);
                        if (spread) throw;
                        else if (e is MyException me) throw me;
                        else
                        {
                            forensic("Chrome.OpenURL", e);
                            Thread.Sleep(2000);
                        }
                    }
            } while (!exit);
        }
        public void OpenURL(string URL)
        {
            OpenURL(URL, null);
        }
        public void ControlTab(DateTime timeout) // Close all browser tabs until only one is left.
        {
            int borrardesde, final;
            string route = string.Empty;

            Balloon("Closing tabs...");
            CR.WaitOne();
            try
            {
                reset(chmTabs);
                if (chmTabs.ComponentExist())
                {
                    while ((chmTabs.dochild(route, "ComponentExist") == "1") && (1 <= int.Parse(chmTabs.dochild(route, "descendants"))) && !StringIn(chmTabs.dochild(route + "\\0", "idrole"), "37"))
                    {
                        route = route + "\\0";
                        reset(chmTabs);
                    }
                    chmTabs.focus();
                    borrardesde = 1;
                    while (chmTabs.dochild(route + "\\" + borrardesde.ToString(), "idrole") == "37")
                    {
                        if (timeout < Now) throw new Exception("I can not close the tabs of Chrome.");
                        chmTabs.dochild(route + "\\" + borrardesde.ToString(), "click");
                        reset(chmTabs);
                        Thread.Sleep(500);
                        CheckAbort();
                        final = int.Parse(chmTabs.dochild(route + "\\" + borrardesde.ToString(), "descendants")) - 1;
                        while (!(final < 0) &&
                               (chmTabs.dochild(route + "\\" + borrardesde + "\\" + final, "idrole") != "43") &&
                               (chmTabs.dochild(route + "\\" + borrardesde + "\\" + final, "visible") != "1"))
                        {
                            final = final - 1;
                        }
                        Balloon("Closing tabs...");
                        if (!(final < 0)) chmTabs.dochild(route + "\\" + borrardesde + "\\" + final, "click");
                        else throw new Exception("Can't close tab.");
                        reset(chmTabs);
                    }
                }
            }
            finally
            {
                CR.ReleaseMutex();
            }
        }
        public void Login(string email, string password) /* The indicated user is logged into Chrome. */
        {
            Execute("$cshChrome00 = new TChrome(" + Owner.Abort + ");\r\n" +
                    "$cshChrome00.Login(\"" + email.Replace("\"", "\"\"") + "\", \"" + password.Replace("\"", "\"\"") + "\");");
        }
        public void Setup() /* It adjusts the Chrome settings. */
        {
            Execute("$cshChrome00 = new TChrome(" + Owner.Abort + ");\r\n" +
                    "$cshChrome00.Setup;");
        }
        public void BackPage() // Go back page.
        {
            Execute("$cshChrome00 = new TChrome(" + Owner.Abort + ");\r\n" +
                    "$cshChrome00.BackPage;");
        }
        public void RefreshPage() // Refresh the page.
        {
            Execute("$cshChrome00 = new TChrome(" + Owner.Abort + ");\r\n" +
                    "$cshChrome00.RefreshPage;");
        }

        public int Width // Read only. The width to which the browser will wrap.
        {
            get { return GetWidth(); }
        }
        public int Height // Read only. The height to which the browser will wrap.
        {
            get { return GetHeight(); }
        }
        public bool Incognito // Read & write. If the browser opens in incognito mode
        {
            get { return GetIncognito(); }
            set { SetIncognito(value); }
        }
        public CredentialStack ProxyPassword
        {
            get { return GetProxyPassword(); }
            set { SetProxyPassword(value); }
        }
        public string Certificate // It is the name of the digital certificate that you want to use to navigate.
        {
            get { return GetCertificate(); }
            set { SetCertificate(value); }
        }
    }
}
