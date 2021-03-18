using Cartes;
using MiTools;
using NotepadLib;
using RPABaseAPI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace Demo_Plugin_Visual_Studio
{
    class SampleNotepad : MyCartesProcessBase
    {
        private string fMessage;
        private Notepad fNotepad;

        public SampleNotepad() : base()
        {
            fMessage = "";
            fNotepad = null;
            ShowAbort = true;
        }

        protected override string getNeededRPASuiteVersion()
        {
            return "3.0.1.0";
        }
        protected override string getRPAMainFile()
        {
            return CurrentPath + "\\Cartes\\Project Structure.cartes.rpa";
        }
        protected override void LoadConfiguration(XmlDocument XmlCfg)
        {
            fMessage = ToString(XmlCfg.SelectSingleNode("//demo//message"));
        }
        protected override void DoExecute(ref DateTime start)
        {
            bool exit;
            DateTime timeout;

            try
            {
                notepad.Write(fMessage.Replace("%user%", CartesScriptExecute("SessionUser;")));
                if (IsSwarmExecution) cartes.balloon("This example shows how to use C# libraries in a process launched by the swarm.");
                else cartes.balloon("This example shows how to use C# libraries in a process that you could launch with the swarm.");
                forensic("This is a trace for the swarm log and the Windows event viewer.");
                timeout = DateTime.Now.AddSeconds(60);
                exit = false;
                do
                {
                    Thread.Sleep(500);
                    CheckAbort();
                    if (timeout < DateTime.Now) exit = true;
                    else if (!notepad.Exists()) exit = true;
                } while (!exit);
                RegisterIteration(start, EXIT_OK, "<task>Put your trace here in xml</task>");
            }catch(Exception e)
            {
                if (e is MyException me) RegisterIteration(start, me.code, e.Message, true);
                else RegisterIteration(start, EXIT_ERROR, e.Message, true);
            }
            start = DateTime.Now;
        }

        public Notepad notepad
        {
            get
            {
                if (fNotepad == null)
                    fNotepad = new Notepad(this);
                return fNotepad;
            }
        }
    }
}
