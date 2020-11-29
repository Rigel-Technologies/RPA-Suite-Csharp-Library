using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;
using Cartes;
using NotepadLib;

namespace Demo_Plugin_Visual_Studio
{
    static class Program
    {
        /// <summary>
        /// Punto de entrada principal para la aplicación.
        /// </summary>
        [STAThread]
        static void Main()
        {
            SampleNotepad sample = null;

            sample = new SampleNotepad();
            sample.Execute();
        }
    }
}
