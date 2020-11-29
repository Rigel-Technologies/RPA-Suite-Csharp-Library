using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Cartes;
using MiTools;

namespace MostBasic
{
	static class MostBasic
	{
		[STAThread]
		static void Main()
		{
			CartesObj robot = new CartesObj();
			String primes;
			DateTime start;
			Cartes.RPAWin32Component notepad;
			String workingFile;
			String primesFile;
			
			try{
			  start = DateTime.Now;
              workingFile = Environment.CurrentDirectory;
              primesFile = workingFile + "\\primes.txt";
			  robot.open(workingFile + "\\library\\notepad.rpa");
			  robot.Execute("visualmode(true);");
              primes = robot.Execute("LoadFromTxtFile(\"" + primesFile + "\");");
			  notepad = robot.GetComponent<RPAWin32Component>("$NotePadEditor");
			  if (notepad.componentexist(0) == 0){
				robot.run("notepad.exe");
				notepad.waitforcomponent(30);
                robot.reset(notepad.api()); /* "reset" warns Cartes of changes in
                  screen applications. Cartes reduces consumption of C.P.U. with
                  this notice. The A.I. of Cartes presupposes these opportune
                  notices in the source code. */
                }
              notepad.Value = notepad.ActiveXClass() + "\r\n" + primes;
              notepad.focus();
              robot.balloon("This example has opened the Notepad to write the first prime numbers.");
			  robot.RegisterIteration(start, "ok", "<task>Put your trace here in xml</task>", 1);
              robot.forensic("This is a trace for the swarm log, and the Windows event viewer.");
              MessageBox.Show ("End");
            }
            finally{
		      robot.close();
		     }
	    }
    }
}
