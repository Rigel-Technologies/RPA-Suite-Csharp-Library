using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Miscellaneous
{
    static class Program
    {
        /// <summary>
        /// </summary>
        [STAThread]
        static void Main()
        {
            SampleMiscellaneous sample = null;

            sample = new SampleMiscellaneous();
            sample.Execute();
        }
    }
}
