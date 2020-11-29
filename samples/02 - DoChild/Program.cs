using Cartes;
using MiTools;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace DoChild
{
    static class Program
    {
        static private CartesObj cartes = new CartesObj();

        [STAThread]
        static void Main()
        {
            DateTime start;
            RPAParameters parameters = new RPAParameters();
            IRPAComponent notepad;
            RPAWin32Component editor;
            String workingFile, editorRoute;

            try
            {
                start = DateTime.Now;
                workingFile = Environment.CurrentDirectory;
                cartes.open(workingFile + "\\library\\notepad.rpa");
                editor = cartes.GetComponent<RPAWin32Component>("$NotePadEditor");
                if (editor.ComponentNotExist())
                {
                    cartes.run("notepad.exe");
                    editor.waitforcomponent(30);
                }
                editorRoute = editor.route();
                notepad = editor.Root();
                parameters.item[0] = WalkTree(0, notepad, "");
                notepad.dochild(editorRoute, "TypeFromClipboard", parameters);
                editor.focus();
                cartes.balloon("This example has opened the notepad and has presented you the component tree with its structure.");
                cartes.RegisterIteration(start, "ok", "<task>Put your trace here in xml</task>", 1);
                cartes.forensic("This is a trace for the swarm log, and the Windows event viewer.");
                MessageBox.Show("End");
            }
            finally
            {
                cartes.close();
            }
        }

        public static string WalkTree(int level, IRPAComponent component, string path)
        {
            RPAParameters param = new RPAParameters();
            string margin = "";
            string result = "";
            int i;

            cartes.balloon("Component : " + path + "\r\nClass : " + component.dochild(path, "class", param));
            i = 0;
            while (i < level)
            {
                margin = margin + "  ";
                i = i + 1;
            }
            result = margin + level + " - " + component.dochild(path, "name", param) + " " + "[" + component.dochild(path, "class", param) + "] = " + component.dochild(path, "wrapper", param);
            i = 0;
            while (i < int.Parse(component.dochild(path, "descendants", param)))
            {
                result = result + "\r\n" + WalkTree(level + 1, component, path + "\\" + i.ToString());
                i = i + 1;
            }
            return result;
        }
    }
}
