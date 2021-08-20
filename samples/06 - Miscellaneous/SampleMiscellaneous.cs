using Cartes;
using ChromeLib;
using IExplorerLib;
using RPABaseAPI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace Miscellaneous
{
    class SampleMiscellaneous : MyCartesProcessBase
    {
        private Chrome fChrome;
        private IExplorer fIExplorer;

        public SampleMiscellaneous() : base()
        {
            fChrome = null;
            fIExplorer = null;
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
            IExplorer MBrowser = IExplorer;
            RPAWin32Component crmlogo = GetComponent<RPAWin32Component>("$ChrmRigelLogo");

            GBrowser.Incognito = false;
            GBrowser.OpenURL("www.rigeltechnologies.net", crmlogo);
            MBrowser.OpenURL("www.rigeltechnologies.net", "$IExplrRigelLogo");
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
        protected IExplorer GetIExplorer()
        {
            if (fIExplorer == null)
            {
                CR.WaitOne();
                try
                {
                    if (fIExplorer == null)
                        fIExplorer = new IExplorer(this);
                }
                finally
                {
                    CR.ReleaseMutex();
                }
            }
            return fIExplorer;
        }

        public Chrome Chrome
        {
            get { return GetChrome(); }
        }
        public IExplorer IExplorer
        {
            get { return GetIExplorer(); }
        }
    }
}
