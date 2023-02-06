using Cartes;
using ChromeLib;
using RPABaseAPI;
using System;
using System.Xml;

namespace Miscellaneous
{
    class SampleMiscellaneous : MyCartesProcessBase
    {
        private Chrome fChrome;

        public SampleMiscellaneous() : base()
        {
            fChrome = null;
        }

        protected override string getNeededRPASuiteVersion()
        {
            return "3.2.1.0";
        }
        protected override string getRPAMainFile()
        {
            return CurrentPath + "\\Cartes\\Miscellaneous.cartes.rpa";
        }
        protected override void LoadConfiguration(XmlDocument XmlCfg)
        {
        }
        protected override void DoExecute(ref DateTime start)
        {
            Chrome GBrowser = Chrome;
            RPAWin32Component crmlogo = GetComponent<RPAWin32Component>("$ChrmRPALogo");

            LoopForever = true;
            GBrowser.Incognito = false;
            GBrowser.OpenURL("www.roboticprocessautomation.net", crmlogo);
            Balloon("Browser is already open");
            WaitFor(5);
        }
        protected Chrome GetChrome()
        {
            if (fChrome == null)
            {
                CR.WaitOne();
                try
                {
                    if (fChrome == null)
                        fChrome = new Chrome(this);
                }
                finally
                {
                    CR.ReleaseMutex();
                }
            }
            return fChrome;
        }

        public Chrome Chrome
        {
            get { return GetChrome(); }
        }
    }
}
