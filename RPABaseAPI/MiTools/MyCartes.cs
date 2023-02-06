using Cartes;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net.Mail;
using System.Text;
using System.Threading;
using System.Xml;
using Microsoft.Win32;
using System.Windows.Forms;
using System.Linq;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;

//////////////////////
// 2023/01/15
//////////////////////

namespace MiTools
{
    public abstract class MyCartes : MyObject // This abstract class is the basis for working and grouping methods for Cartes.
    {
        public enum SwarmCommand { none, execute, finish };
        protected const string EXIT_ABORT = "Abort";
        protected const string EXIT_ERROR = "KO";
        private const string EXIT_MERGE_KO = "LOADING_" + EXIT_ERROR;
        private const string EXIT_UNMERGE_KO = "UNLOADING_" + EXIT_ERROR;
        protected const string EXIT_OK = "OK";

        private static Version fVersion = null;
        private Version flVersion = null;
        private NumberFormatInfo fDoubleFormatProvider = null;
        private KeyValuePair<DateTime, string> fLastBalloon = new KeyValuePair<DateTime, string>(DateTime.Now, string.Empty);

        public MyCartes()
        {
            flVersion = null;
        }

        internal void Load()
        {
            try
            {
                MergeLibrariesAndLoadVariables();
            }
            catch (Exception e)
            {
                forensic(ClassName + ".MergeLibrariesAndLoadVariables", e);
                if (e is MyException pp) throw;
                else throw new MyException(EXIT_MERGE_KO, "I cannot load the " + ClassName + " library" + LF + e.Message);
            }
        }
        internal void UnLoad()
        {
            try
            {
                UnMergeLibrariesAndUnLoadVariables();
            }
            catch (Exception e)
            {
                forensic("MyCartes.UnMergeLibrariesAndUnLoadVariables", e);
                if (e is MyException pp) throw;
                else throw new MyException(EXIT_UNMERGE_KO, "I cannot unload the " + ClassName + " library" + LF + e.Message);
            }
        }

        private Version getNeededRPASuiteVersionP() // It returns the version of RPA Suite needed by this library
        {
            try
            {
                if (flVersion == null)
                {
                    if (!Version.TryParse(ToString(getNeededRPASuiteVersion()), out flVersion))
                        throw new Exception(getNeededRPASuiteVersion() + " is not a valid version number.");
                }
            }catch(Exception e)
            {
                forensic("MyCartes.getNeededRPASuiteVersionP", e);
                throw;
            }
            return flVersion;
        }
        private bool GetIsRPASuiteInstalled()
        {
            bool result;
            try
            {
                Version v = CurrentRPASuiteVersion;
                result = v.ToString().Length > 0;
            }
            catch
            {
                result = false;
            }
            return result;
        }
        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        private static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint cButtons, uint dwExtraInfo);

        protected abstract Mutex GetRC();
        protected abstract CartesObj getCartes();
        protected abstract string getCartesPath();  // It returns the file of Cartes
        protected abstract void MergeLibrariesAndLoadVariables();  // Rewrite this method to load the libraries and Cartesa variables that your class handles.
        protected abstract void UnMergeLibrariesAndUnLoadVariables(); // Rewrite this method to unload the libraries and Cartesa variables that your class handles.
        protected virtual Version getCurrentRPASuiteVersion()  // It returns the version of RPA Suite
        {
            if (fVersion == null)
            {
                object lvalue;

                lvalue = Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Cartes", "Product Version", null);
                if (lvalue == null)
                    lvalue = Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Cartes", "Product Version", null);
                if (!Version.TryParse(ToString(lvalue), out fVersion))
                    throw new Exception("RPA Suite is not installed. Please install RPA Suite.");
            }
            return fVersion;
        }
        protected virtual string getNeededRPASuiteVersion() // It returns a string with the version of RPA Suite needed by this library
        {
            return "3.4.3.0";
        }
        protected virtual void CheckRPASuiteVersion() // It checks if the current version and needed are OK
        {
            if (CurrentRPASuiteVersion.CompareTo(NeededRPASuiteVersion) < 0)
                throw new Exception("You need RPA Suite v" + NeededRPASuiteVersion + " or higher.");
        }
        protected virtual NumberFormatInfo getDoubleFormatProvider()
        {
            if (fDoubleFormatProvider == null)
            {
                fDoubleFormatProvider = new NumberFormatInfo();
                fDoubleFormatProvider.NumberDecimalSeparator = ",";
                fDoubleFormatProvider.NumberGroupSeparator = ".";
                fDoubleFormatProvider.NumberGroupSizes = new int[] { 3 };
            }
            return fDoubleFormatProvider;
        }
        protected abstract string getProjectId(); // Returns the ID of the loaded project in Cartes. If Cartes does not have a loaded project, it returns the empty string.
        protected string GetProjectName() // returns the Name of the loaded project in Cartes. If Cartes does not have a loaded project, it returns the empty string.
        {
            return ToString(Execute("Name;"));
        }
        protected abstract SwarmCommand GetSCommand();
        protected abstract bool GetIsSwarmExecution();
        protected virtual void reset(IRPAComponent component) // Reset the API of te component
        {
            CR.WaitOne(); // Be careful, if you use concurrent threads, protect them with this critical region. Especially ...
            try // ... the "reset" command.
            {
                cartes.reset(component);
            }
            finally
            {
                CR.ReleaseMutex();
            }
        }
        protected bool isVariable(string VariableName)  // If a variable-component exists in the rpa project, returns true
        {
            try
            {
                return Execute("isVariable(\"" + VariableName + "\");") == "1";
            }
            catch
            {
                return false;
            }
        }
        protected virtual RPAComponent GetComponent<RPAComponent>(string variablename) // returns the Cartes variable indicated
            where RPAComponent : class, IRPAComponent
        {
            return cartes.GetComponent<RPAComponent>(variablename);
        }
        protected virtual string Execute(string command) // It Executes a Cartes Script in cartes.execute and check if errors
        {
            string result;

            CR.WaitOne(); // Be careful, if you use concurrent threads, protect them with this critical region. Especially ...
            try // ... the "reset" command.
            {
                result = ToString(cartes.Execute(command));
                if (ToString(cartes.LastError()).Length > 0)
                    throw new Exception(cartes.LastError());
            }
            finally
            {
                CR.ReleaseMutex();
            }
            return result;
        }
        protected virtual string CartesScriptExecute(string command)
        {
            return Execute(command);
        }
        protected virtual void Balloon(string message)
        {
            bool ok;

            CR.WaitOne();
            try
            {
                if (fLastBalloon.Value == message)
                    ok = fLastBalloon.Key < DateTime.Now.AddSeconds(-10);
                else ok = message != string.Empty;
                if (ok)
                {
                    cartes.balloon(message);
                    fLastBalloon = new KeyValuePair<DateTime, string>(DateTime.Now, message);
                }
            }
            finally
            {
                CR.ReleaseMutex();
            }
        }
        protected virtual new void forensic(string message) // It writes "message" in the swarm log and in the windows event viewer.
        {
            CR.WaitOne();
            try
            {
                cartes.forensic(message);
            }
            finally
            {
                CR.ReleaseMutex();
            }
        }
        protected virtual new void forensic(string message, Exception e)
        {
            forensic(message + "\r\n" + e.Message);
        }
        protected abstract void CheckAbort(); // If abort is requested, the method should throw an exception.
        protected abstract bool GetIsAborting(); // The method returns true if abort has been requested
        protected void DoMouseClick() //Do a click in the cursor's current position
        {
            //Mouse actions
            const int MOUSEEVENTF_LEFTDOWN = 0x02;
            const int MOUSEEVENTF_LEFTUP = 0x04;
            //const int MOUSEEVENTF_RIGHTDOWN = 0x08;
            //const int MOUSEEVENTF_RIGHTUP = 0x10;

            //Call the imported function with the cursor's current position
            uint X = (uint)Cursor.Position.X;
            uint Y = (uint)Cursor.Position.Y;
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, X, Y, 0, 0);
        }
        protected virtual bool ComponentsExist(int seconds, params IRPAComponent[] components) /* The method waits for the indicated
              seconds until one of the components exists. If any of the components exist, returns true. */
        {
            DateTime timeout;
            bool exit, result = false;
            List<string> lAPI = new List<string>();

            try
            {
                if (components == null) result = true;
                else
                {
                    exit = false;
                    timeout = DateTime.Now.AddSeconds(seconds);
                    do
                    {
                        CheckAbort();
                        foreach (IRPAComponent component in components)
                        {
                            if (component.componentexist(0) == 1)
                            {
                                result = true;
                                exit = true;
                                break;
                            }
                            else if (!lAPI.Contains(component.api()))
                                lAPI.Add(component.api());
                        }
                        if (!exit)
                        {
                            if (timeout < DateTime.Now) exit = true;
                            Thread.Sleep(400);
                            foreach (string api in lAPI)
                            {
                                CR.WaitOne(); // Be careful, if you use concurrent threads, protect them with this critical region. Especially ...
                                try // ... the "reset" command.
                                {
                                    cartes.reset(api);
                                }
                                finally
                                {
                                    CR.ReleaseMutex();
                                }
                            }
                        }
                    } while (!exit);
                }
            }
            catch (Exception e)
            {
                forensic("MyCartes.ComponentsExist", e);
                throw;
            }
            return result;
        }
        protected virtual string WaitForCartesMethodValue(IRPAComponent component, string method, int seconds) /* It waits until the indicated method has
            value and returns it. Once the waiting time has been exceeded, it throws an exception. */
        {
            RPAParameters parameters = new RPAParameters();
            DateTime timeout;
            string result;

            try
            {
                timeout = DateTime.Now.AddSeconds(seconds);
                do
                {
                    CheckAbort();
                    cartes.reset(component.api());
                    Thread.Sleep(400);
                    if (timeout < DateTime.Now) throw new Exception("Timeout");
                    else result = ToString(component.Execute(method, parameters));
                } while (result.Length == 0);
                return result;
            }
            catch (Exception e)
            {
                forensic("MyCartes::WaitForCartesMethodValue", e);
                throw;
            }
        }
        protected virtual void AssignValueInsistently(DateTime timeout, IRPAWin32Component component, string value, bool typed = false) /* Assign the indicated value to the
            component. If it does not succeed, it will insist until the system time exceeds "timeot". */
        {
            RPAWin32Component root = null;
            RPAParameters parameters = new RPAParameters();
            string path = string.Empty;

            void RecuperarRuta()
            {
                do
                {
                    root = component.Root();
                    if (root != null)
                    {
                        path = component.route();
                        break;
                    }
                    else if (timeout < DateTime.Now) throw new Exception("I can't assign the value \"" + value + "\" to the component.");
                    else
                    {
                        reset(component);
                        CheckAbort();
                        Thread.Sleep(1000);
                    }
                } while (true);
            }

            CR.WaitOne();
            try
            {
                RecuperarRuta();
                do
                {
                    parameters.clear();
                    if (typed)
                        try
                        {
                            parameters.item[0] = value;
                            parameters.itemAsInteger[1] = 0;
                            parameters.itemAsInteger[2] = 0;
                            root.dochild(path, "TypeFromClipboard", parameters);
                        }
                        catch
                        {
                            parameters.item[0] = value;
                            parameters.itemAsInteger[1] = 0;
                            parameters.itemAsInteger[2] = 0;
                            root.dochild(path, "TypeWord", parameters);
                        }
                    else
                    {
                        parameters.item[0] = value;
                        root.dochild(path, "Value", parameters);
                    }
                    CheckAbort();
                    Thread.Sleep(1000);
                    reset(component);
                    if (ToString(root.dochild(path, "Value")).ToLower() == value.ToLower()) break;
                    else
                    {
                        if (timeout < DateTime.Now) throw new Exception("I can't assign the value \"" + value + "\" to the component.");
                        else RecuperarRuta(); // Prrfff!!! Unbelievable
                        Thread.Sleep(1000);
                    }
                } while (true);
            }
            finally
            {
                CR.ReleaseMutex();
            }
        }
        protected virtual double SimilarStrings(string a, string b)
        {
            NumberFormatInfo fDoubleFormat = new NumberFormatInfo();
            string sresult;
            double dresult;

            try
            {
                fDoubleFormat.NumberDecimalSeparator = ".";
                fDoubleFormat.NumberGroupSeparator = ",";
                fDoubleFormat.NumberGroupSizes = new int[] { 3 };
                sresult = cartes.Execute("similarstrings(\"spa\", \"" + a.Replace("\"", "\"\"") + "\", \"" + b.Replace("\"", "\"\"") + "\");");
                dresult = Convert.ToDouble(sresult, fDoubleFormat);
            }
            catch(Exception e)
            {
                forensic("MyCartes.SimilarStrings", e);
                throw;
            }
            return dresult;
        }
        protected virtual void AdjustWindow(IRPAWin32Component component, int x, int y, int width, int height) // Adjusts the component main window to the indicated size and coordinates.
        {
            CR.WaitOne();
            try
            {
                if (component.ComponentExist())
                {
                    RPAWin32Component lpWindow = component.Root();

                    if (StringIn(lpWindow.WindowState, "Minimized", "Maximized") || (lpWindow.Visible == 0))
                        lpWindow.Show("Restore");
                    if ((lpWindow.width != width) || (lpWindow.height != height))
                        lpWindow.ReSize(width, height);
                    if ((lpWindow.x != x) || (lpWindow.y != y))
                        lpWindow.Move(x, y);
                }
            }
            finally
            {
                CR.ReleaseMutex();
            }
        }
        protected void AdjustWindow(IRPAWin32Component component, int width, int height) // Adjusts the component main window to the indicated size.
        {
            AdjustWindow(component, 0, 0, width, height);
        }
        protected void CenterWindow(IRPAWin32Component component, int width, int height) // Center the component's main window
        {
            CR.WaitOne();
            try
            {
                if (component.ComponentExist())
                {
                    RPAWin32Component lpWindow = component.Root();
                    if (StringIn(lpWindow.WindowState, "Minimized", "Maximized") || (lpWindow.Visible == 0))
                        lpWindow.Show("Restore");
                    Screen pantalla = Screen.FromHandle((IntPtr)lpWindow.handle());
                    Rectangle Position = pantalla.WorkingArea;
                    int x = Position.X + (Position.Width - width) / 2,
                        y = Position.Y + (Position.Height - height) / 2;
                    AdjustWindow(lpWindow, x, y, width, height);
                }
            }
            finally
            {
                CR.ReleaseMutex();
            }
        }
        protected virtual void CloseWindow(DateTime timeout, IRPAWin32Component CloseButton, string name) // Press the button until it disappears. "name" is the name of the window.
        {
            reset(CloseButton);
            if (CloseButton.ComponentExist())
            {
                DateTime init = DateTime.Now;
                int handle = CloseButton.handle();
                while (CloseButton.ComponentExist() && (handle == CloseButton.handle()))
                {
                    Balloon("Closing \"" + name + "\"...");
                    if (timeout < DateTime.Now)
                    {
                        int seconds = Convert.ToInt32(Math.Round((DateTime.Now - init).TotalSeconds));
                        throw new Exception("Timeout, " + seconds.ToString() + " seconds. I can't close the \"" + name + "\" window.");
                    }
                    if (CloseButton.focused() != 1)
                        CloseButton.focus();
                    CloseButton.click(0);
                    CloseButton.ComponentNotExist(5);
                    Thread.Sleep(250);
                    reset(CloseButton);
                }
            }
        }
        protected void CloseWindow(IRPAWin32Component CloseButton, string name)
        {
            CloseWindow(DateTime.Now.AddSeconds(60), CloseButton, name);
        }
        protected void CloseWindow(IRPAWin32Component CloseButton)
        {
            CloseWindow(CloseButton, ToString(CloseButton.Root().Value).Trim());
        }
        protected void CloseWindows(IRPAWin32Component CloseButton)
        {
            while (CloseButton.ComponentExist())
            {
                CloseWindow(CloseButton);
            }
        }
        protected virtual void scrollUp(int mouseX, int mouseY, IRPAWin32Accessibility component)
        {
            RPAParameters parametros = new RPAParameters();
            DateTime timeout;
            int n, y;

            try
            {
                parametros.item[0] = mouseX.ToString();
                parametros.item[1] = mouseY.ToString();
                component.focus();
                component.doroot("SetMouse", parametros);
                n = 0;
                y = component.y;
                timeout = DateTime.Now.AddSeconds(30);
                while (component.OffScreen == 1)
                {
                    if (timeout < DateTime.Now) throw new Exception("Timeout waiting by scrolling.");
                    CheckAbort();
                    cartes.balloon("Scroll...");
                    component.doroot("SetMouse", parametros);
                    component.Up();
                    Thread.Sleep(500);
                    if (n > 3) throw new Exception("I can not scroll up.");
                    if (y == component.y) n++;
                    else n = 0;
                }
            }
            catch (Exception e)
            {
                forensic("MyCartes.scrollUp", e);
                throw;
            }
        }
        protected virtual void scrollDown(int mouseX, int mouseY, IRPAWin32Accessibility component)
        {
            RPAParameters parameters = new RPAParameters();
            DateTime timeout;
            int n, y;

            try
            {
                parameters.item[0] = mouseX.ToString();
                parameters.item[1] = mouseY.ToString();
                component.focus();
                component.doroot("SetMouse", parameters);
                n = 0;
                y = component.y;
                timeout = DateTime.Now.AddSeconds(30);
                while (component.OffScreen == 1)
                {
                    if (timeout < DateTime.Now) throw new Exception("Timeout waiting by scrolling.");
                    CheckAbort();
                    cartes.balloon("Scroll...");
                    component.doroot("SetMouse", parameters);
                    component.down();
                    Thread.Sleep(500);
                    if (n > 3) throw new Exception("I can not scroll down.");
                    if (y == component.y) n++;
                    else n = 0;
                }
            }
            catch (Exception e)
            {
                cartes.forensic("MyCartes.scrollDown(RPAWin32Accessibility)\r\n" + e.Message);
                throw;
            }
        }
        protected virtual void scrollDown(int mouseX, int y, int height, IRPAWin32Component component)
        {
            RPAParameters parametros = new RPAParameters();
            DateTime timeout;

            try
            {
                parametros.item[0] = mouseX.ToString();
                parametros.item[1] = (y + 2).ToString();
                component.focus();
                component.doroot("SetMouse", parametros);
                timeout = DateTime.Now.AddSeconds(20);
                while (component.y < y)
                {
                    if (timeout < DateTime.Now) throw new Exception("Timeout waiting by scrolling.");
                    CheckAbort();
                    Balloon("Scroll up...");
                    component.Up();
                    Thread.Sleep(500);
                }
                while (y + height < component.y + component.height)
                {
                    if (timeout < DateTime.Now) throw new Exception("Timeout waiting by scrolling.");
                    CheckAbort();
                    Balloon("Scroll down...");
                    component.down();
                    Thread.Sleep(500);
                }
            }
            catch (Exception e)
            {
                forensic("MyCartes.scrollDown(RPAWin32Component)", e);
                throw;
            }
        }

        public override double ToDouble(string value)
        {
            return Convert.ToDouble(value, getDoubleFormatProvider());
        }
        public abstract void WaitFor(int seconds); // It waits the indicated seconds, except if the process is aborted or the swarm asks it to finish.
        public delegate bool MeetsCriteria(); // If the criteria is satisfied, this prototype must return true, otherwise return false.
        public abstract void SendWarning(string subject, string body); /* Sends the notice to the email account assigned to the running
          swarm process. If you don't run a swarm process, the function ignores the warning. */
        public abstract void SendEmail(string to, string subject, string body); // It sends an email. To do this, it uses the configured email account on the RPA Server.

        public Mutex CR // An instance of Mutex to control the critical regions in competition.
        {
            get { return GetRC(); }
        }
        public bool IsRPASuiteInstalled
        {
            get { return GetIsRPASuiteInstalled(); }
        }  // Read Only. It returns if RPA Suite is installed
        public Version CurrentRPASuiteVersion
        {
            get { return getCurrentRPASuiteVersion(); }
        }  // Read Only. It returns the version of RPA Suite
        public Version NeededRPASuiteVersion
        {
            get { return getNeededRPASuiteVersionP(); }
        }  // Read Only. It returns the version of RPA Suite needed by this library
        public string CartesPath
        {
            get
            {
                return getCartesPath();
            }
        } // Read. It returns the file of Cartes
        public CartesObj cartes
        {
            get { return getCartes(); }
        }  // Read Only
        public string ProjectId
        {
            get { return getProjectId(); }
        }  // Read Only. Returns the ID of the loaded project in Cartes. If Cartes does not have a loaded project, it returns the empty string.
        public string Name
        {
            get { return GetProjectName(); }
        }  // Read Only. Returns the Name of the loaded project in Cartes. If Cartes does not have a loaded project, it returns the empty string.
        public bool IsSwarmExecution
        {
            get { return GetIsSwarmExecution(); }
        } // Read Only. It returns true if the execution of the current process has been commanded by the swarm
        public SwarmCommand Command
        {
            get { return GetSCommand(); }
        } // Read Only. The order received by the swarm
        public bool IsAborting // Read. The method returns true if abort has been requested
        {
            get { return GetIsAborting(); }
        }
        public NumberFormatInfo DoubleFormatProvider
        {
            get { return getDoubleFormatProvider(); }
        }  // Read Only
    }

    public abstract class MyCartesProcess : MyCartes // This abstract class allows you to create processes using APIS from MyCartesAPI.
    {
        public enum ProcessState { rest, starting, iterating, ending }; // The process must exist in some of these execution states.
        protected const string EXIT_SETTINGS_KO = "Settings_" + EXIT_ERROR;
        private static string fCartesPath = null;
        private static CartesObj fCartes = null;
        private ProcessState fState;
        private SwarmCommand fSCommand;
        private bool fAllowSendReport;
        private DateTime fStart;
        private Dictionary<string, int> fTypifycationsCounter = null;
        private List<string> fTypifycations = null;
        private string fAbort; // Cartes Script Variable for know when the process musts abort
        private List<MyCartesAPI> apis;
        private string fFileSettings;
        private bool fShowAbort, fVisibleMode, fExecuting, fLoopForever;
        private RPADataString frpaAbort = null;
        protected SmtpClient fSMTP;

        private class MyCartesForensic : MyForensic
        {
            CartesObj fCartes;

            public MyCartesForensic() : base()
            {
                fCartes = new CartesObj();
            }
            protected override bool forensic(EventLogEntryType type, string message) // It writes "message" in the windows event viewer.
            {
                bool result = true;

                switch (type)
                {
                    case EventLogEntryType.Error: Cartes.forensic(message); break;
                    case EventLogEntryType.Warning: Cartes.forensic(message); break;
                    case EventLogEntryType.Information: Cartes.forensic(message); break;
                    default: result = base.forensic(type, message); break;
                }
                return result;
            }

            public CartesObj Cartes
            {
                get { return fCartes; }
            }
        }

        public MyCartesProcess(string csvAbort) : base()  // "csvAbort" is a Cartes variable that when valid one will indicate to the instance that it must abort.
        {
            fState = ProcessState.rest;
            fSCommand = SwarmCommand.none;
            fAllowSendReport = true;
            fStart = DateTime.Now;
            fTypifycationsCounter = new Dictionary<string, int>();
            fTypifycations = new List<string>();
            fAbort = ToString(csvAbort).Trim();
            apis = new List<MyCartesAPI>();
            fFileSettings = null;
            fShowAbort = true;
            fVisibleMode = true;
            fExecuting = false;
            fLoopForever = false;
            fSMTP = null;
        }

        internal void iCheckAbort()
        {
            CheckAbort();
        }
        internal void AddAPI(MyCartesAPI api)
        {
            try
            {
                if (apis.IndexOf(api) < 0)
                {
                    apis.Add(api);
                    if (ProjectId.Length > 0)
                        api.Load();
                }
            }catch(Exception e)
            {
                forensic("MyCartesProcess.AddAPI", e);
                throw;
            }
        }
        internal void DeleteAPI(MyCartesAPI api)
        {
            apis.Remove(api);
        }
 
        private void LoadConfiguration() // It loads the process configuration.   
        {
            XmlDocument lpXmlCfg = null;

            try
            {
                lpXmlCfg = new XmlDocument();
                lpXmlCfg.Load(SettingsFile);
                LoadConfiguration(lpXmlCfg);
            }catch(Exception e)
            {
                if (e is MyException m) throw m;
                else throw new MyException(EXIT_SETTINGS_KO, "Loading settings:" + LF + e.Message);
            }
        }

        protected override Mutex GetRC()
        {
            return CartesObjExtensions.CR;
        }
        protected override CartesObj getCartes()
        {
            CartesObj resultado;

            if (fCartes == null)
            {
                CR.WaitOne();
                try
                {
                    if (fCartes == null)
                    {
                        if ((CartesPath.Length > 0) && File.Exists(CartesPath))
                        {
                            bool ok;
                            string CartesName = Path.GetFileNameWithoutExtension(CartesPath);
                            System.Diagnostics.Process current = System.Diagnostics.Process.GetCurrentProcess();
                            System.Diagnostics.Process[] ap = System.Diagnostics.Process.GetProcessesByName(CartesName);
                            ok = false;
                            foreach (System.Diagnostics.Process item in ap)
                            {
                                if (item.SessionId == current.SessionId)
                                {
                                    ok = true;
                                    break;
                                }
                            }
                            if (!ok)
                            {
                                System.Diagnostics.Process.Start(CartesPath);
                                System.Threading.Thread.Sleep(3000);
                            }
                            MyCartesForensic lpCoroner = new MyCartesForensic();
                            setCoroner(lpCoroner);
                            fCartes = lpCoroner.Cartes;
                            CartesObjExtensions.GlobalCartes = fCartes;
                        }
                        else throw new Exception("Cartes is not installed. Please install Robot Cartes from the RPA Suite installation.");
                    }
                }
                finally
                {
                    CR.ReleaseMutex();
                }
            }
            CR.WaitOne();
            try
            {
                resultado = fCartes;
            }
            finally
            {
                CR.ReleaseMutex();
            }
            return resultado;
        }
        protected override string getCartesPath()  
        {
            object InstallPath;

            if (fCartesPath == null)
            {
                InstallPath = Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Cartes", "Cartes Client", null);
                if (InstallPath == null)
                    InstallPath = Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Cartes", "Cartes Client", null);
                fCartesPath = ToString(InstallPath);
            }
            return fCartesPath;
        }
        protected override void MergeLibrariesAndLoadVariables()
        {
            try
            {
                foreach (MyCartes item in apis)
                    item.Load();
                if (Abort.Length > 0)
                {
                    if (isVariable(Abort)) frpaAbort = (RPADataString)cartes.component(Abort);
                    else throw new MyException(EXIT_ERROR, Abort + " does not exist!");
                }
                else frpaAbort = null;
            }
            catch (Exception e)
            {
                forensic("MyCartesProcess.MergeLibrariesAndLoadVariables", e);
                throw new MyException(EXIT_ERROR, e.Message);
            }
        }
        protected override void UnMergeLibrariesAndUnLoadVariables()
        {
            try
            {
                foreach (MyCartes item in apis)
                    item.UnLoad();
            }
            catch (Exception e)
            {
                forensic("MyCartesProcess.UnMergeLibrariesAndUnLoadVariables", e);
                throw new MyException(EXIT_ERROR, e.Message);
            }
        }
        protected override string getProjectId() 
        {
            return ToString(Execute("ProjectId;"));
        }
        protected override bool GetIsSwarmExecution()
        {
            bool result = false;

            try
            {
                result = fExecuting && (Execute("IsSwarmExecution;") == "1");
            }catch(Exception e)
            {
                forensic("MyCartesForensic.GetIsSwarmExecution", e);
            }
            return result;
        }
        protected override SwarmCommand GetSCommand()
        {
            SwarmCommand result = fSCommand;

            switch (cartes.SwarmContinuity)
            {
                case 2:
                    result = SwarmCommand.finish;
                    break;
                case 1:
                    result = SwarmCommand.execute;
                    break;
                default:
                    result = SwarmCommand.none;
                    break;
            }
            return result;
        }
        protected override void CheckAbort() // It checks if the variable to abort is 1 to throw an exception
        {
            if (GetIsAborting())
                throw new MyException(EXIT_ABORT, "Abort by user.");
        }
        protected override bool GetIsAborting() 
        {
            return (Abort.Length > 0) && (Execute(Abort + ".value;") == "1");
        }
        protected virtual int GetApis()
        {
            return apis.Count;
        }
        protected virtual MyCartesAPI GetApi(int index)
        {
            try
            {
                return apis.ElementAt(index);
            }
            catch (Exception e)
            {
                forensic("MyCartesProcess::GetApi(" + index.ToString() + ")", e);
                throw;
            }
        }
        protected SmtpClient GetSMTP() // It returns a SMTP Client to send emails.
        {
            if (fSMTP == null)
            {
                fSMTP = new SmtpClient();
                fSMTP.Credentials = new System.Net.NetworkCredential("myaccount@gmail.com", "mypassword");
                fSMTP.Port = 587;
                fSMTP.Host = "smtp.gmail.com";
                fSMTP.EnableSsl = true;
            }
            return fSMTP;
        }
        protected abstract string getRPAMainFile(); // Here you must return the main ".rpa" file
        protected abstract void LoadConfiguration(XmlDocument XmlCfg); // Here you must load the configuration of the process
        protected virtual void ShowAbortDialog(IRPADataString abort)
        {
            abort.ShowAbortDialog("Press button to abort", "Closing...", "Abort");
        }
        protected virtual int RegisterIteration(DateTime start, string typify, string data, bool screenShot = false)
        {
            int result, counter;
            try
            {
                CR.WaitOne();
                try
                {
                    result = cartes.RegisterIteration(start, typify, data, screenShot ? 1 : 0);
                    if (result == 2) fSCommand = SwarmCommand.finish;
                    else fSCommand = SwarmCommand.execute;
                }
                finally
                {
                    CR.ReleaseMutex();
                }
            }
            finally
            {
                if (fTypifycationsCounter.TryGetValue(typify, out counter))
                    fTypifycationsCounter[typify] = counter + 1;
                else fTypifycationsCounter[typify] = 1;
                fTypifycations.Insert(fTypifycations.Count, typify);
                if (AllowSendReports)
                    SendReport(fStart, State, fTypifycationsCounter, fTypifycations);
            }
            return result;
        }
        protected virtual void SwarmDelay()
        {
            CR.WaitOne();
            try
            {
                if (IsSwarmExecution)
                {
                    cartes.swarmdelaydefault();
                    WaitFor(50);
                    fSCommand = SwarmCommand.finish;
                }
            }
            finally
            {
                CR.ReleaseMutex();
            }
        }
        protected virtual void SendReport(DateTime start, ProcessState state, Dictionary<string, int> TypifycationsCounter, List<string> Typifycations)
        /* This method is used to send execution status reports to the email enabled in the RPA Center as recipient of the warnings.
           Redefine it to send your own reports. The parameters are "start" (process execution start time), "status" (process status),
           "TypificationsCounter" (number of operations per typifycation) and "Typifycations" (listed in order of execution of the
           operations performed) . */
        {

            string Resumen()
            {
                string result = string.Empty;

                foreach (KeyValuePair<string, int> cell in TypifycationsCounter)
                    result = result + ToString(cell.Key) + " : " + ToString(cell.Value) + LF;
                if (result.Length > 0) return "Number of operations for each typifycation." + LF + result;
                else return result;
            }

            string text = string.Empty;

            switch (state)
            {
                case ProcessState.ending:
                    string timing = "It started at " + start.ToShortTimeString() + " and finished at " + DateTime.Now.ToShortTimeString() + ".";
                    if (Typifycations.Count > 0)
                    {
                        text = "\"" + Name + "\" is over." + LF + 
                               timing + LF + LF +
                               Resumen();
                        SendWarning("\"" + Name + "\" is over.", text);
                    }
                    else
                    {
                        text = "\"" + Name + "\" is over without processing anything." + LF +
                               timing + LF;
                        SendWarning("\"" + Name + "\" is over without processing anything.", text);
                    }
                    break;
            }
        }
        protected virtual bool DoInit() // If DoExecute must be invoked, this method returns true.
        {
            if (!IsDebug)
                Close();
            return true;
        }
        protected abstract void DoExecute(ref DateTime start); // Here you must execute the process. The process have already loaded the configuration
        protected virtual void DoEnd() // This method is invoked after running DoExecute.
        {
            if (!IsDebug)
                Close();
        }

        public override void WaitFor(int seconds)
        {
            DateTime timeout = Now.AddSeconds(seconds), lifesignal = Now.AddSeconds(10);
            
            while ((Now < timeout) && !IsAborting && (Command != SwarmCommand.finish))
            {
                if (lifesignal < Now)
                {
                    Execute("SendLifeSignal;");
                    lifesignal = Now.AddSeconds(60);
                }
                Thread.Sleep(125);
            }
        }
        public override void SendWarning(string subject, string body)
        {
            cartes.SendWarning(ToString(subject), ToString(body));
        }
        public override void SendEmail(string to, string subject, string body)
        {
            string command = "SendEmail(\"" + ToString(to).Replace("\"", "\"\"") + "\", \"" + ToString(subject).Replace("\"", "\"\"") + "\", \"" + ToString(body).Replace("\"", "\"\"") + "\");";
            try
            {
                Execute(command);
            }catch(Exception e)
            {
                forensic("MyCartesProcess::SendEmail", e);
                throw;
            }
        }
        public bool Execute()  // Execute the process. if succesfull return True, else return false
        {
            DateTime start;
            bool result = false;
            string lsMainFile;
            bool enter = false;
            Mutex mutex = null;

            void TratarCatch (Exception e)
            {
                MyException mye;

                Balloon(e.Message);
                forensic(e.Message);
                if (e is MyException m) mye = m;
                else mye = new MyException(EXIT_ERROR, e.Message);
                try
                {
                    if (mye.code == EXIT_SETTINGS_KO)
                        RegisterIteration(start, EXIT_SETTINGS_KO, "<task>\r\n" +
                                                        "  <error>I can not load the configuration file \"" + SettingsFile + "\"</error>\r\n" +
                                                        "  <message>" + e.Message + "</message>\r\n" +
                                                        "</task>");
                    else
                        RegisterIteration(start, mye.code, "<task>\r\n" +
                                                        "  <data>" + e.Message + "</data>\r\n" +
                                                        "</task>", true);
                }
                catch (Exception e2)
                {
                    Coroner.write("MyCartesProcess::Execute", e);
                    Coroner.write("MyCartesProcess::Execute", e2);
                    forensic("MyCartesProcess::Execute" + LF + "Unexpected communication failure with RPA Server.", e2);
                    throw;
                }
            }

            try
            {
                mutex = new Mutex(true, "Processing area of \"RPA Suite\".", out enter);
                try
                {
                    if (enter)
                    {
                        fState = ProcessState.starting;
                        fSCommand = SwarmCommand.execute;
                        try
                        {
                            fStart = DateTime.Now;
                            start = fStart;
                            fTypifycationsCounter.Clear();
                            fTypifycations.Clear();
                            CheckRPASuiteVersion();
                            if (IsDebug) Balloon("Attention: This run is in debug mode.");
                            else Balloon("I'm opening the project...");
                            lsMainFile = RPAMainFile;
                            if (File.Exists(CurrentPath + "\\" + lsMainFile)) cartes.open(CurrentPath + "\\" + lsMainFile);
                            else if (File.Exists(CurrentPath + "\\Cartes\\" + lsMainFile)) cartes.open(CurrentPath + "\\Cartes\\" + lsMainFile);
                            else cartes.open(RPAMainFile);
                            try
                            {
                                fExecuting = true;
                                try
                                {
                                    MergeLibrariesAndLoadVariables();
                                    Balloon("Reading settings...");
                                    LoadConfiguration();
                                    Balloon(Name);
                                    try
                                    {
                                        if (VisibleMode)
                                            Execute("visualmode(1);");
                                        try
                                        {
                                            if (DoInit())
                                            {
                                                bool exit;
                                                if (ShowAbort && (frpaAbort != null))
                                                    ShowAbortDialog(frpaAbort);
                                                if (AllowSendReports)
                                                    SendReport(fStart, State, fTypifycationsCounter, fTypifycations);
                                                exit = false;
                                                fState = ProcessState.iterating;
                                                do
                                                {
                                                    try
                                                    {
                                                        DoExecute(ref start);
                                                    }
                                                    catch(Exception e)
                                                    {
                                                        if (LoopForever)
                                                        {
                                                            TratarCatch(e);
                                                            try
                                                            {
                                                                Close();
                                                            }
                                                            catch { }
                                                        }
                                                        else throw;
                                                    }
                                                    if (!LoopForever) exit = true;
                                                    else if (Command == SwarmCommand.finish) exit = true;
                                                    else if (IsAborting) exit = true;
                                                    else start = DateTime.Now;
                                                } while (!exit);
                                            }
                                        }
                                        finally
                                        {
                                            fState = ProcessState.ending;
                                            DoEnd();
                                        }
                                    }
                                    finally
                                    {
                                        if (VisibleMode)
                                            Execute("visualmode(0);");
                                        if (AllowSendReports)
                                            SendReport(fStart, State, fTypifycationsCounter, fTypifycations);
                                    }
                                    Balloon("\"" + Name + "\" is over.");
                                }
                                catch (Exception e)
                                {
                                    fState = ProcessState.ending;
                                    TratarCatch(e);
                                }
                            }
                            finally
                            {
                                try
                                {
                                    UnMergeLibrariesAndUnLoadVariables();
                                }
                                finally
                                {
                                    fExecuting = false;
                                    cartes.close();
                                }
                            }
                        }
                        finally
                        {
                            fSCommand = SwarmCommand.none;
                            fState = ProcessState.rest;
                        }
                        result = true;
                    }
                    else
                    {
                        string sname = Name,
                               message = "You will not run another process until it finishes.";
                        if (0 < sname.Length) message = "Process \"" + sname + "\" is already running." + LF + message;
                        else message = "There is already a process running." + LF + message;
                        Balloon(message);
                        forensic(message);
                    }
                }
                finally
                {
                    mutex.Close();
                }
            }
            catch (Exception e)
            {
                Balloon(e.Message);
                forensic(e.Message);
#if DEBUG
                MessageBox.Show(e.Message);
#endif
            }
            return result;
        }
        public virtual void Close()
        {
            int i;

            i = 0;
            while (i < GetApis())
            {
                try
                {
                    GetApi(i).Close();
                }
                catch (Exception e)
                {
                    forensic("MyCartesProcess::Close()\r\n" + "GetApi(" + i.ToString() + ").Close()", e);
                }
                i++;
            }
        }

        public bool LoopForever
        {
            get { return fLoopForever; }
            set { fLoopForever = value; }
        } // Read & Write. The execution of the process will be repeated over and over again until it is aborted or the swarm asks to leave.
        public bool AllowSendReports
        {
            get { return fAllowSendReport; }
            set { fAllowSendReport = value; }
        } // Read & Write. The process will send execution status reports to the email enabled in the RPA Center.
        public bool VisibleMode
        {
            get { return fVisibleMode; }
            set { fVisibleMode = value; }
        } // Read & Write. It controls the visible mode of Carte. 
        public string SettingsFile
        {
            get
            {
                if (fFileSettings == null)
                    fFileSettings = CurrentPath + "\\settings.xml";
                return fFileSettings;
            }
            set { fFileSettings = value; }
        } // Read & Write
        public bool ShowAbort
        {
            get { return fShowAbort; }
            set { fShowAbort = value; }
        } // Read & Write. It controls the appearance of the window to abort. 
        public string Abort
        {
            get { return fAbort; }
        }  // Read Only. The property returns the name of the variable used for the abort control.
        public ProcessState State
        {
            get { return fState; }
        } // Read Only. The property returns the state of the process
        public string RPAMainFile
        {
            get { return getRPAMainFile(); }
        } // Read
        public SmtpClient SMTP
        {
            get { return GetSMTP(); }
        } // Read
    }

    public abstract class MyCartesAPI : MyCartes // This abstract class allows you to inherit to create application APIs (Chrome, SAP ...) using Cartes for MyCartesProcess.
    {
        private MyCartesProcess fowner;
        private bool fChecked;

        public MyCartesAPI(MyCartesProcess owner) : base()
        {
            fowner = owner;
            fChecked = false;
            fowner.AddAPI(this);
        }
        ~MyCartesAPI()
        {
            fowner.DeleteAPI(this);
        }
        protected override Mutex GetRC()
        {
            return Owner.CR;
        }
        protected override CartesObj getCartes()
        {
            if (!fChecked)
            {
                CheckRPASuiteVersion();
                fChecked = true;
            }
            return Owner.cartes;
        }
        protected override string getCartesPath()
        {
            return Owner.CartesPath;
        }
        protected override NumberFormatInfo getDoubleFormatProvider()
        {
            return Owner.DoubleFormatProvider;
        }
        protected override string getProjectId()
        {
            return Owner.ProjectId;
        }
        protected override bool GetIsSwarmExecution()
        {
            return Owner.IsSwarmExecution;
        }
        protected override SwarmCommand GetSCommand()
        {
            return Owner.Command;
        }
        protected override void CheckAbort()
        {
            Owner.iCheckAbort();
        }
        protected override bool GetIsAborting()
        {
            return Owner.IsAborting;
        }
        protected virtual MyCartesProcess GetOwner()
        {
            return fowner;
        }

        public override double ToDouble(string value)
        {
            return Owner.ToDouble(value);
        }
        public override void WaitFor(int seconds)
        {
            Owner.WaitFor(seconds);
        }
        public override void SendWarning(string subject, string body)
        {
            Owner.SendWarning(subject, body);
        }
        public override void SendEmail(string to, string subject, string body)
        {
            Owner.SendEmail(to, subject, body);
        }
        public abstract void Close(); // This method should close all the application windows

        public MyCartesProcess Owner
        {
            get { return GetOwner(); }
        }
    }

    public static class CartesObjExtensions
    {
        internal static CartesObj GlobalCartes = null;
        private static Mutex gRC = new Mutex();
        private static T Casting<T>(IRPAComponent component) where T : class, IRPAComponent
        {
            try
            {
                if (component == null) return null;
                else if (component is T result) return result;
                else throw new Exception("Casting() is a " + component.ActiveXClass());
            }
            catch (Exception e)
            {
                MyObject.Coroner.write("T CartesObjExtensions::Casting<T>", e);
                throw;
            }
        }

        public static Mutex CR // A golbal instance of Mutex to control the critical regions in competition.
        {
            get { return gRC; }
        }
        public static void Write(this ICredentialStack credential, IRPAWin32Component component)
        {
            IRPAComponent lpC = component;
            credential.Write((RPAComponent)lpC);
        }
        public static T GetComponent<T>(this CartesObj cartes, string variablename) where T : class, IRPAComponent
        {
            IRPAComponent component = cartes.component(variablename);

            if (component == null) return null;
            else if (component is T result) return result;
            else throw new Exception(variablename + " is a " + component.ActiveXClass());
        }
        public static void reset(this CartesObj cartes, IRPAComponent component)
        {
            if (component != null)
            {
                CR.WaitOne();
                try
                {
                    string api = component.api();
                    cartes.reset(api);
                }
                finally
                {
                    CR.ReleaseMutex();
                }
            }
        }
        public static bool isVariable(this CartesObj cartes, string VariableName) // If a variable-component exists in the rpa project, returns true
        {
            bool resultado;
            CR.WaitOne();
            try
            {
                try
                {
                    resultado = MyObject.ToString(cartes.Execute("isVariable(\"" + VariableName + "\");")) == "1";
                    if ((cartes.LastError() != null) && (cartes.LastError().Length > 0))
                        throw new Exception(cartes.LastError());
                }
                catch
                {
                    resultado = false;
                }
            }
            finally
            {
                CR.ReleaseMutex();
            }
            return resultado;
        }
        public static bool ComponentExist(this IRPAComponent component, int TimeOut = 0)
        {
            bool result = false;
            try
            {
                result = component.componentexist(TimeOut) != 0;
            }catch(Exception e)
            {
                MyObject.Coroner.write("CartesObjExtensions.ComponentExist", e);
            }
            return result;
        }
        public static bool ComponentNotExist(this IRPAComponent component, int TimeOut = 0)
        {
            bool result = false;

            try
            {
                if (TimeOut <= 0) result = !component.ComponentExist(TimeOut);
                else if (!component.ComponentExist(1)) result = true;
                else
                {
                    DateTime timeout = DateTime.Now.AddSeconds(TimeOut);
                    while (!result)
                    {
                        GlobalCartes.reset(component.api());
                        if (!component.ComponentExist(1)) result = true;
                        else if (timeout < DateTime.Now) break;
                        else Thread.Sleep(250);
                    }
                }
            }
            catch(Exception e)
            {
                MyObject.Coroner.write("CartesObjExtensions.ComponentNotExist", e);
                throw;
            }

            return result;
        }
        public static List<int> RouteInt(this IRPAComponent component) // Returns the component path from the root.
        {
            string[] splited = component.route().Split('\\');
            List<int> iRoute = new List<int>();

            for (int i = 1; i < splited.Length; i++)
                iRoute.Add(int.Parse(splited[i]));
            return iRoute;
        }
        public static RPAWin32Component Root(this IRPAWin32Component component)
        {
            return Casting<RPAWin32Component>(component.getComponentRoot());
        }
        public static IRPAJava32Component Root(this IRPAJava32Component component)
        {
            return Casting<IRPAJava32Component>(component.getComponentRoot());
        }
        public static IRPAMSHTMLComponent Root(this IRPAMSHTMLComponent component)
        {
            return Casting<IRPAMSHTMLComponent>(component.getComponentRoot());
        }
        public static IRPASapComponent Root(this IRPASapComponent component)
        {
            return Casting<IRPASapComponent>(component.getComponentRoot());
        }
        public static string doroot(this IRPAComponent component, string method)
        {
            return component.doroot(method, null);
        }
        public static T Child<T>(this IRPAComponent component, string route) where T : class, IRPAComponent
        {
            string[] splited = route.Split('\\');
            List<int> iRoute = new List<int>();

            for(int i = 1; i < splited.Length; i++)
                iRoute.Add(int.Parse(splited[i]));
            return Child<T>(component, iRoute.ToArray());
        }
        public static T Child<T>(this IRPAComponent component, params int[] route) where T : class, IRPAComponent
        {
            IRPAComponent current = component;

            for (int i = 0; i < route.Length; i++)
                current = current.child(route[i]);
            return Casting<T>(current);
        }
        public static T Child<T>(this IRPAWin32Component component, int index) where T : class, IRPAWin32Component
        {
            return Casting<T>(component.child(index));
        }
        public static T Child<T>(this IRPAJava32Component component, int index) where T : class, IRPAJava32Component
        {
            return Casting<T>(component.child(index));
        }
        public static T Child<T>(this IRPAMSHTMLComponent component, int index) where T : class, IRPAMSHTMLComponent
        {
            return Casting<T>(component.child(index));
        }
        public static T Child<T>(this IRPASapComponent component, int index) where T : class, IRPASapComponent
        {
            return Casting<T>(component.child(index));
        }
        public static string dochild(this IRPAComponent component, string route, string method)
        {
            return component.dochild(route, method, null);
        }
        public static string dochild(this IRPAComponent component, string route, string method, params string[] parameters)
        {
            RPAParameters List = new RPAParameters();

            foreach (string item in parameters)
                List.item[List.items] = item;
            return component.dochild(route, method, List);
        }
        public static string dochild(this IRPAComponent component, string route, string method, int parameter)
        {
            return component.dochild(route, method, parameter.ToString());
        }
        public static void click(this IRPAWin32Component component)
        {
            component.click(0);
        }
        public static void TypeKey(this IRPAWin32Component component, string key)
        {
            component.TypeKey(key, "", "");
        }
        public static Rectangle FindPicture(this IRPAWin32Component component, params string[] ImageFiles)
        {
            RPAParameters parameters = new RPAParameters(), output;
            Rectangle result = Rectangle.Empty;

            for (int i = 0; i < ImageFiles.Count(); i++)
                parameters.item[i] = ImageFiles[i];
            output = component.FindPicture(parameters);
            if (output.item[0] == "1")
                result = new Rectangle(int.Parse(output.item[1]), int.Parse(output.item[2]), int.Parse(output.item[3]), int.Parse(output.item[4]));
            return result;
        }
        public static Rectangle FindPicture(this IRPAWin32Component component, List<string> ImageFiles)
        {
            return component.FindPicture(ImageFiles.ToArray());
        }
        public static void ClickOnImage(this IRPAWin32Component component, bool MoveMouse, params string[] ImageFiles)
        {
            RPAParameters parameters = new RPAParameters();

            for (int i = 0; i < ImageFiles.Count(); i++)
                parameters.item[i] = ImageFiles[i];
            component.clickonimage(parameters, MoveMouse ? 1 : 0);
        }
        public static void ClickOnImage(this IRPAWin32Component component, bool MoveMouse, List<string> ImageFiles)
        {
            component.ClickOnImage(MoveMouse, ImageFiles.ToArray());
        }
        public static void DoubleClickOnImage(this IRPAWin32Component component, bool MoveMouse, params string[] ImageFiles)
        {
            RPAParameters parameters = new RPAParameters();

            for (int i = 0; i < ImageFiles.Count(); i++)
                parameters.item[i] = ImageFiles[i];
            component.doubleclickonimage(parameters, MoveMouse ? 1 : 0);
        }
        public static void DoubleClickOnImage(this IRPAWin32Component component, bool MoveMouse, List<string> ImageFiles)
        {
            component.DoubleClickOnImage(MoveMouse, ImageFiles.ToArray());
        }
        public static bool GetReadOnly(this IRPAWin32Component component)
        {
            int ro = 64;

            return ((component.IdSituation() & ro) == ro);
        }
        public static bool Inside(this IRPAWin32Component root, IRPAWin32Component child) // Returns true if the screen coordinates of "child" are within "root".
        {
            if (root.ComponentExist() && child.ComponentExist())
                return ((root.x <= child.x) && (child.x + child.width <= root.x + root.width)) &&
                    ((root.y <= child.y) && (child.y + child.height <= root.y + root.height));
            else return false;
        }
        public static bool Inside(this IRPAJava32Component root, IRPAJava32Component child)
        {
            if (root.ComponentExist() && child.ComponentExist())
                return ((root.x <= child.x) && (child.x + child.width <= root.x + root.width)) &&
                    ((root.y <= child.y) && (child.y + child.height <= root.y + root.height));
            else return false;
        }
        public static bool Inside(this IRPASapControl root, IRPASapControl child)
        {
            if (root.ComponentExist() && child.ComponentExist())
                return ((root.x() <= child.x()) && (child.x() + child.width() <= root.x() + root.width())) &&
                    ((root.y() <= child.y()) && (child.y() + child.height() <= root.y() + root.height()));
            else return false;
        }
        public static bool Inside(this IRPAMSHTMLComponent root, IRPAMSHTMLComponent child)
        {
            if (root.ComponentExist() && child.ComponentExist())
                return ((root.x() <= child.x()) && (child.x() + child.width() <= root.x() + root.width())) &&
                    ((root.y() <= child.y()) && (child.y() + child.height() <= root.y() + root.height()));
            else return false;
        }
    }
}
