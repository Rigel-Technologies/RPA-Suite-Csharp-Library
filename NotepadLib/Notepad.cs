using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Cartes;
using MiTools;
using RPABaseAPI;

namespace NotepadLib
{
    public class Notepad : MyCartesAPIBase
    {
        private static bool loaded = false;
        private RPAWin32Component notepad = null, notepadclose = null, notepadnosave = null;

        public Notepad(MyCartesProcess owner) : base(owner)
        {
        }
        protected override void MergeLibrariesAndLoadVariables()
        {
            if (!loaded || (Execute("isVariable(\"$NotePadEditor\");") != "1"))
            {
                notepad = null;
                loaded = cartes.merge(CurrentPath + "\\Cartes\\NotepadLib.cartes.rpa") == 1;
            }
            if (notepad == null)
            {
                notepad = GetComponent<RPAWin32Component>("$NotePadEditor");
                notepadclose = GetComponent<RPAWin32Component>("$NotePadClose");
                notepadnosave = GetComponent<RPAWin32Component>("$NotePadNoSave");
            }
        }

        protected override string getNeededRPASuiteVersion()
        {
            return "3.0.1.0";
        }
        protected void open()
        {
            bool exit;
            DateTime timeout;
            RPAParameters parameters;

            parameters = new RPAParameters();
            timeout = DateTime.Now.AddSeconds(60);
            exit = false;
            do
            {
                cartes.reset(notepad.api()); /* "reset" warns Cartes of changes in
                  screen applications. Cartes reduces consumption of C.P.U. with
                  this notice. The A.I. of Cartes presupposes these opportune
                  notices in your source code. */
                if (timeout < DateTime.Now) throw new Exception("Exhausted timeout opening Notepad.");
                else if (notepad.componentexist(0) == 1)
                {
                    parameters.clear();
                    parameters.item[0] = "900";
                    parameters.item[1] = "580";
                    notepad.doroot("resize", parameters);
                    exit = true;
                }
                else
                {
                    cartes.run("notepad.exe");
                    notepad.waitforcomponent(30);
                }
            } while (!exit);
        }

        public override void Close()
        {
            bool exit;
            DateTime timeout;

            timeout = DateTime.Now.AddSeconds(120);
            exit = false;
            do
            {
                reset(notepad); /* "reset" warns Cartes of changes in
                  screen applications. Cartes reduces consumption of C.P.U. with
                  this notice. The A.I. of Cartes presupposes these opportune
                  notices in your source code. */
                Thread.Sleep(250);
                if (timeout < DateTime.Now) throw new Exception("Time out closing Notepad.");
                else if (notepadnosave.ComponentExist()) notepadnosave.click();
                else if (notepadclose.ComponentExist()) notepadclose.click();
                else exit = !notepad.ComponentExist();
            }
            while (!exit);
        }
        public bool exists() // The method returns true if notepad exists
        {
            reset(notepad);
            return notepad.componentexist(0) == 1;
        }
        public void write (string value)
        {
            bool exit;
            DateTime timeout;

            timeout = DateTime.Now.AddSeconds(60);
            exit = false;
            do
            {
                cartes.reset(notepad.api()); /* "reset" warns Cartes of changes in
                  screen applications. Cartes reduces consumption of C.P.U. with
                  this notice. The A.I. of Cartes presupposes these opportune
                  notices in your source code. */
                if (timeout < DateTime.Now) throw new Exception("Exhausted timeout opening Notepad.");
                else if (notepad.componentexist(0) == 1)
                {
                    AssignValueInsistently(timeout, notepad, value, false);
                    notepad.focus();
                    exit = true;
                }
                else open();
            }while (!exit);
        }
    }
}
