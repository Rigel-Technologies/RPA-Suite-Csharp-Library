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

namespace ChromeLib
{
    public class Chrome : MyCartesAPIBase
    {
        private static bool loaded = false;

        public Chrome(MyCartesProcess owner) : base(owner)
        {
            MergeLibrariesAndLoadVariables();
        }
        protected override void MergeLibrariesAndLoadVariables()
        {
            if (!loaded || !isVariable("$Chrome"))
            {
                loaded = cartes.merge(CurrentPath + "\\Cartes\\Chrome Es.rpa") == 1;
            }
        }

        public override void Close() // It closes Chrome.
        {
            MergeLibrariesAndLoadVariables();
            Execute("$cshChrome00 = new TChrome(" + Abort + ");\r\n" +
                    "$cshChrome00.closeAll();");
        }
        public void OpenURL(string URL, string csvComponente) /* It opens the indicated web page. "csvComponente" must be a component of the page
              that indicates when the page has been loaded: for example, "$googlelogo". */
        {
            MergeLibrariesAndLoadVariables();
            Execute("$cshChrome00 = new TChrome(" + Abort + ");\r\n" +
                    "$cshChrome00.OpenURL(\"" + URL.Replace("\"", "\"\"") + "\", " + csvComponente.Replace("\"", "\"\"") + ");");
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
    }
}
