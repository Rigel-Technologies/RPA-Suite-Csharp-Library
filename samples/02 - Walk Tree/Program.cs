using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Cartes;
using MiTools;

namespace Walk_tree
{
    static class MostBasic
    {
        [STAThread]
        static void Main()
        {
            DateTime start;
            CartesObj robot = new CartesObj();
            Cartes.RPAWin32Component notepad;
            String workingFile;
 
            try
            {
                start = DateTime.Now;
                workingFile = Environment.CurrentDirectory;
                robot.open(workingFile + "\\library\\notepad.rpa");
                notepad = robot.GetComponent<RPAWin32Component>("$NotePadEditor");
                if (notepad.componentexist(0) == 0)
                {
                    robot.run("notepad.exe");
                    notepad.waitforcomponent(30);
                }
                notepad.Value = WalkTree(0, notepad.Root()); ;
                notepad.focus();
                robot.balloon("This example has opened the notepad and has presented you the component tree with its structure.");
                robot.RegisterIteration(start, "ok", "<task>Put your trace here in xml</task>", 1);
                robot.forensic("This is a trace for the swarm log, and the Windows event viewer.");
                MessageBox.Show("End");
            }
            finally
            {
                robot.close();
            }
        }


        public static string WalkTree (int level, RPAWin32Component component)
        {
            RPAParameters param  = new RPAParameters();
            string margin = "";
            string result = "";
 
            int i;

            i = 0;
            while (i < level)
            {
                margin = margin + "  ";
                i = i + 1;
            }
            result = margin + level + " - " + component.name() + " " + "[" + component.Execute("class", param) + "] = " + component.Wrapper();
            i = 0;
            while (i < component.descendants)
            {
                result = result + "\r\n" + WalkTree(level + 1, (RPAWin32Component) component.child(i));
                i = i + 1;
             }
            return result;
        }
    }
}
