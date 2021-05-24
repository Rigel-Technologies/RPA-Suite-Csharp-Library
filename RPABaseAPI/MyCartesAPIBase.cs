using Cartes;
using MiTools;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace RPABaseAPI
{
    public abstract class MyCartesAPIBase : MyCartesAPI
    {
        public MyCartesAPIBase(MyCartesProcess owner) : base(owner)
        {
        }
        protected override void UnMergeLibrariesAndUnLoadVariables() // Normally, you don't need to do anything to download the library.
        {
        }
    }

    public class ClassVisualStudio : MyCartesAPIBase
    {
        private static bool loaded = false;
        private RPAWin32Component vsWindow = null;
        private RPAWin32Automation vsClose = null;

        public ClassVisualStudio(MyCartesProcess owner) : base(owner)
        {
        }
        
        protected override void MergeLibrariesAndLoadVariables()
        {
            if (!loaded || !isVariable("$VisualStudio"))
            {
                loaded = cartes.merge(CurrentPath + "\\Cartes\\VisualStudio.rpa") == 1;
            }
            if (vsWindow == null)
            {
                vsWindow = GetComponent<RPAWin32Component>("$VisualStudio");
                vsClose = GetComponent<RPAWin32Automation>("$VisualStudioClose");
            }
        }
        protected override void UnMergeLibrariesAndUnLoadVariables()
        {
            loaded = false;
            vsWindow = null;
        }

       

        public override void Close()
        {
            bool exit;
            DateTime timeout;

            timeout = DateTime.Now.AddSeconds(120);
            exit = false;
            while (!exit)
            {
                cartes.reset(vsWindow.api());
                Thread.Sleep(200);
                if (timeout < DateTime.Now) throw new Exception("Time out closing Visual Studio.");
                else try
                    {
                        if (vsClose.componentexist(0) == 1) vsClose.click(0);
                        else if (vsWindow.componentexist(0) != 1) exit = true;
                    }
                    catch (Exception e)
                    {
                        forensic("ClassVisualStudio::Close", e);
                    }
            }
        }
        public virtual bool Minimize() // Minimize the Visual Studio
        {
            bool result = false;
            
            if (vsWindow.componentexist(0) == 1)
            {
                RPAWin32Component lpWindow = vsWindow.Root();
                if ((lpWindow.Visible == 1) && !StringIn(lpWindow.WindowState, "Minimized"))
                    lpWindow.Show("Minimize");
            }
            return result;
        }
        public virtual bool Restore() // Restore the Visual Studio
        {
            bool result = false;

            if (vsWindow.componentexist(0) == 1)
            {
                RPAWin32Component lpWindow = vsWindow.Root();
                if (StringIn(lpWindow.WindowState, "Minimized", "Maximized"))
                    lpWindow.Show("Restore");
            }
            return result;
        }
    }

    public abstract class MyCartesProcessBase : MyCartesProcess
    {
        private ClassVisualStudio fVS = null;
        
        public MyCartesProcessBase() : base("$Abort")
        {
        }

        protected override void MergeLibrariesAndLoadVariables()
        {
            try
            {
                cartes.merge(CurrentPath + "\\Cartes\\RPABaseAPI.cartes.rpa");
                base.MergeLibrariesAndLoadVariables();
            }
            catch (Exception e)
            {
                forensic("MyCartesProcessBase.MergeLibrariesAndLoadVariables", e);
                throw new MyException(EXIT_ERROR, e.Message);
            }
        }
        protected override bool DoInit()
        {
            VisualStudio.Minimize();
            return base.DoInit();
        }
        protected override void DoEnd() 
        {
            VisualStudio.Restore();
            base.DoEnd();
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint cButtons, uint dwExtraInfo);
        //Mouse actions
        private const int MOUSEEVENTF_LEFTDOWN = 0x02;
        private const int MOUSEEVENTF_LEFTUP = 0x04;
        private const int MOUSEEVENTF_RIGHTDOWN = 0x08;
        private const int MOUSEEVENTF_RIGHTUP = 0x10;

        public void DoMouseClick() //Do a click without a component focused in the cursor's current position
        {
            //Call the imported function with the cursor's current position
            uint X = (uint)Cursor.Position.X;
            uint Y = (uint)Cursor.Position.Y;
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, X, Y, 0, 0);
        }

        protected static DateTime GetFormatDateTime(string mask, string value)
        {
            CE_Data.DateTime dt = new CE_Data.DateTime();
            dt.Text[mask] = value;
            return new DateTime(int.Parse(dt.Text["yyyy"]), int.Parse(dt.Text["mm"]), int.Parse(dt.Text["dd"]),
                                int.Parse(dt.Text["hh"]), int.Parse(dt.Text["nn"]), int.Parse(dt.Text["ss"]));
        }
        protected virtual ClassVisualStudio GetVisualStudio()
        {
            if (fVS == null) fVS = new ClassVisualStudio(this);
            return fVS;
        }
        public override void Close()
        {
            int i;

            i = 0;
            while (i < GetApis())
            {
                try
                {
                    if (GetApi(i) is ClassVisualStudio vs)
                    {
                        // Nothing to do
                    }
                    else GetApi(i).Close();
                }
                catch (Exception e)
                {
                    forensic("MyCartesProcessBase::Close()\r\n" + "GetApi(" + i.ToString() + ").Close()", e);
                }
                i++;
            }
        }

        public ClassVisualStudio VisualStudio
        {
            get { return GetVisualStudio(); }
        } // Read
    }
}
