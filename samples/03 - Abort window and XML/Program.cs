using Cartes;
using CE_Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using System.Xml;

namespace AbortAndXML
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            CartesObj cartes = new CartesObj();
            RPADataString Abort;
            string workingFile;

            workingFile = Environment.CurrentDirectory;
            cartes.open(workingFile + "\\rpa\\AbortAndXML.rpa");
            try
            {
                XmlDocument doc = new XmlDocument();
                XmlNode usersNode = doc.CreateElement("users");
                XMLFile datos2 = new XMLFile(); // Cartes class from CE_Data

                Abort = (RPADataString) cartes.component("$Abort");
                Abort.ShowAbortDialog("Presss the button to end", "Bye", "Abort");
                if (Abort.Value == "0")
                {
                    XmlNode userNode = null;
                    XmlNode phoneNode = null;

                    // I create the XML with the native class of C #
                    doc.AppendChild(usersNode);
                    userNode = doc.CreateElement("name");
                    userNode.InnerText = "Federico Codd";
                    usersNode.AppendChild(userNode);
                    phoneNode = doc.CreateElement("telephone");
                    phoneNode.InnerText = "985124753";
                    usersNode.AppendChild(phoneNode);
                    phoneNode = doc.CreateElement("telephone");
                    phoneNode.InnerText = "654357951";
                    usersNode.AppendChild(phoneNode);
                    // I create the XML with the Cartes class
                    datos2.AsString["name"] = userNode.InnerText;
                    datos2.getKey("telephone").listAsString[0] = "985124753";
                    datos2.getKey("telephone").listAsString[1] = "985124753";
                    do
                    {
                        Thread.Sleep(2000);
                    }
                    while (Abort.Value == "0");
                    doc.Save(workingFile + "\\datos1.xml");
                    datos2.SaveToFile(workingFile + "\\datos2.xml");
                }
            }
            finally
            {
                cartes.close();
            }
        }
    }
}
