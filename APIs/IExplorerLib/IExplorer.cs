using MiTools;
using RPABaseAPI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IExplorerLib
{
    public class IExplorer : MyCartesAPIBase
    {
        private static bool loaded = false;

        public IExplorer(MyCartesProcess owner) : base(owner)
        {
            MergeLibrariesAndLoadVariables();
        }
        protected override void MergeLibrariesAndLoadVariables()
        {
            if (!loaded || !isVariable("$InternetExplorer"))
            {
                loaded = cartes.merge(CurrentPath + "\\Cartes\\Internet Explorer ES.rpa") == 1;
            }
        }

        public override void Close() // It closes Internet Explorer.
        {
            Execute("$cshIExplorer00 = new TInternetExplorer(" + Owner.Abort + ");\r\n" +
                    "$cshIExplorer00.closeAll();");
        }
        public void CloseOptionPanel()
        {
            Execute("$cshIExplorer00 = new TInternetExplorer(" + Owner.Abort + ");\r\n" +
                    "$cshIExplorer00.CloseOptionPanel();");
        }
        public void OpenURL(string URL, string csvComponente) /* It opens the indicated web page. "csvComponente" must be a component of the page
              that indicates when the page has been loaded: for example, "$googlelogo". */
        {
            Execute("$cshIExplorer00 = new TInternetExplorer(" + Owner.Abort + ");\r\n" +
                    "$cshIExplorer00.OpenURL(\"" + URL.Replace("\"", "\"\"") + "\", " + csvComponente.Replace("\"", "\"\"") + ");");
        }
    }
}
