using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using System.Text;
using System.Threading.Tasks;
using Cartes;
using MiTools;
using RPABaseAPI;
using System.Threading;

namespace ChromeLib
{
    public class Chrome : MyCartesAPIBase
    {
        private static bool loaded = false;
        private bool fSpanish;
        private Mutex fRC = new Mutex();
        private Thread fThStack = null;
        private CredentialStack fPrxyPsw;
        protected RPAWin32Component chm = null, chmURLEdit = null, chmTabs = null, chmClose = null, chmCloseCRC = null;
        protected RPAWin32Component chmSvPsw = null, chmProxy = null, chmProxyUser = null, chmProxyPsw = null, chmProxyAceptar = null;

        public Chrome(MyCartesProcess owner) : base(owner)
        {
            fPrxyPsw = null;
            fSpanish = true;
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
        public void ControlTab(DateTime timeout)
        {
            int borrardesde, final;
            string route = string.Empty;

            Balloon("Closing tabs...");
            reset(chmTabs);
            while ((chmTabs.dochild(route, "descendants") == "1") && !StringIn(chmTabs.dochild(route + "\\0", "idrole"), "37"))
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

        protected override void MergeLibrariesAndLoadVariables()
        {
            if (!loaded || !isVariable("$Chrome"))
            {
                loaded = cartes.merge(CurrentPath + "\\Cartes\\Chrome Es.rpa") == 1;
            }
            if (chm == null)
            {
                chm = GetComponent<RPAWin32Component>("$Chrome");
                chmURLEdit = GetComponent<RPAWin32Component>("$ChromeURLEdit");
                chmTabs = GetComponent<RPAWin32Component>("$ChromeTabs");
                chmClose = GetComponent<RPAWin32Component>("$ChromeClose");
                chmCloseCRC = GetComponent<RPAWin32Component>("$ChromeCloseCRC");
                chmSvPsw = GetComponent<RPAWin32Component>("$ChromeSavePSW");
                chmProxy = GetComponent<RPAWin32Component>("$ChromePrx");
                chmProxyUser = GetComponent<RPAWin32Component>("$ChromePrxUser");
                chmProxyPsw = GetComponent<RPAWin32Component>("$ChromePrxPsw");
                chmProxyAceptar = GetComponent<RPAWin32Component>("$ChromePrxIniciar");
            }
        }
        protected virtual CredentialStack GetProxyPassword()
        {
            RC.WaitOne();
            try
            {
                return fPrxyPsw;
            }
            finally
            {
                RC.ReleaseMutex();
            }
        }
        protected virtual void SetProxyPassword(CredentialStack value)
        {
            RC.WaitOne();
            try
            {
                fPrxyPsw = value;
                if (fPrxyPsw == null) fThStack = null;
                else InitThread();
            }
            finally
            {
                RC.ReleaseMutex();
            }
        }
        protected virtual void ProcessBackGroundButtons() /* This is the method of a thread responsible for clicking the
            buttons that allow access to Outlook from DCOM. */
        {
            const int delay = 30;
            DateTime timereset = DateTime.Now.AddSeconds(delay);
            while (fThStack != null)
            {
                try
                {
                    if (timereset < DateTime.Now)
                    {
                        reset(chmProxy);
                        timereset = DateTime.Now.AddSeconds(delay);
                    }
                    else
                    {
                        RC.WaitOne();
                        try
                        {
                            CheckProxy();
                        }
                        finally
                        {
                            RC.ReleaseMutex();
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
        protected virtual void NeverSavePassword()
        {
            try
            {
                while (chmSvPsw.ComponentExist(3) && StringIn(chmSvPsw.name(), "Nunca"))
                {
                    CheckAbort();
                    chmSvPsw.click();
                    reset(chmSvPsw);
                    Thread.Sleep(250);
                }
            }catch(Exception e)
            {
                forensic("Chrome.NeverSavePassword", e);
                throw;
            }
        }
        protected virtual void CheckProxy()
        {
            try
            {
                if (ProxyPassword != null)
                {
                    while (chmProxy.ComponentExist() && StringIn(chmProxy.name(), "Iniciar sesión"))
                    {
                        CheckAbort();
                        chmProxyUser.Value = ProxyPassword.User;
                        ProxyPassword.Write((RPAComponent)chmProxyPsw);
                        chmProxyAceptar.click();
                        reset(chmProxy);
                        NeverSavePassword();
                        Thread.Sleep(250);
                    }
                }
            }catch(Exception e)
            {
                forensic("ProcessBackGroundButtons.CheckProxy", e);
                throw;
            }
        }
        protected Mutex RC // The Critical section
        {
            get { return GetRC(); }
        }

        public override void Close() // It closes Chrome.
        {
            MergeLibrariesAndLoadVariables();
            Execute("$cshChrome00 = new TChrome(" + Abort + ");\r\n" +
                    "$cshChrome00.closeAll();");
        }
        public void OpenURL(string URL, params IRPAComponent[] Components) /* It opens the indicated web page. Components must be a list of components of the page
              that indicates when the page has been loaded: for example, "$googlelogo". */
        {
            bool spread, exit;
            DateTime timeout;

            spread = false;
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
                            RC.WaitOne();
                            try
                            {
                                if (chmProxy.ComponentExist()) CheckProxy();
                                else if (chmClose.ComponentExist() && chmCloseCRC.ComponentExist() &&
                                    ((ToString(chmClose.name()).Length > 0) || (ToString(chmCloseCRC.name()).Length > 0)) &&
                                    (!StringIn(ToString(chmClose.name()), "cerrar") || !StringIn(ToString(chmCloseCRC.name()), "cerrar"))
                                   )
                                {
                                    fSpanish = false;
                                    spread = true;
                                    throw new Exception("ERROR! Your Chrome or your Windows are not in Spanish. This library learned in Spanish buttons.");
                                }
                                else if (ComponentsExist(0, Components))
                                {
                                    RPAWin32Component chrome = chmURLEdit.Root();
                                    fSpanish = true;
                                    if (!StringIn(ToString(chrome.WindowState), "normal"))
                                        chrome.Show("restore");
                                    chrome.Move(0, 0);
                                    chrome.ReSize(985, 732);
                                    ControlTab(timeout);
                                    exit = ComponentsExist(0, Components);
                                }
                                else if (chmURLEdit.ComponentExist())
                                {
                                    ControlTab(timeout);
                                    chmURLEdit.TypeFromClipboardCheck(CheckURL(URL), 1, 0);
                                    chmURLEdit.TypeKey("Enter");
                                    Thread.Sleep(500);
                                    ComponentsExist(30, Components);
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
                                RC.ReleaseMutex();
                            }
                        }
                        else if (ToString(URL).Trim().Length > 0)
                        {
                            cartes.run("chrome.exe \"" + ToString(URL).Trim() + "\"");
                            reset(chm);
                            Thread.Sleep(500);
                            Balloon("Waiting for Chrome");
                            List<IRPAComponent> list = Components.ToList();
                            list.Add(chmProxy);
                            ComponentsExist(30, list.ToArray());
                        }
                        else
                        {
                            cartes.run("chrome.exe");
                            reset(chm);
                            Thread.Sleep(500);
                            ComponentsExist(30, chmProxy, chmURLEdit);
                        }
                    }
                    catch (Exception e)
                    {
                        Balloon(e.Message);
                        if (spread) throw;
                        else
                        {
                            forensic("Chrome.OpenURL", e);
                            Thread.Sleep(2000);
                        }
                    }
            } while (!exit);
        }
        public void Login(string email, string password) /* The indicated user is logged into Chrome. */
        {
            MergeLibrariesAndLoadVariables();
            Execute("$cshChrome00 = new TChrome(" + Abort + ");\r\n" +
                    "$cshChrome00.Login(\"" + email.Replace("\"", "\"\"") + "\", \"" + password.Replace("\"", "\"\"") + "\");");
        }
        public void Setup() /* It adjusts the Chrome settings. */
        {
            MergeLibrariesAndLoadVariables();
            Execute("$cshChrome00 = new TChrome(" + Abort + ");\r\n" +
                    "$cshChrome00.Setup;");
        }

        public CredentialStack ProxyPassword
        {
            get { return GetProxyPassword(); }
            set { SetProxyPassword(value); }
        }
    }
}
