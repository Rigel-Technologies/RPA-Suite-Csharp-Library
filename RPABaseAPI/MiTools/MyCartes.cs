﻿using Cartes;
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
using System.Windows;
using System.Diagnostics;
using System.Drawing;

//////////////////////
// 2021/03/18
//////////////////////

namespace MiTools
{
    public abstract class MyCartes : MyObject // This abstract class is the basis for working and grouping methods for Cartes.
    {
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
                forensic(ClassName + ".UnMergeLibrariesAndUnLoadVariables", e);
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
            return "3.3.6.0";
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
        protected abstract bool GetIsSwarmExecution();
        protected virtual void reset(IRPAComponent component) // Reset the API of te component
        {
            cartes.reset(component.api());
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

            result = ToString(cartes.Execute(command));
            if ((cartes.LastError() != null) && (cartes.LastError().Length > 0))
                throw new Exception(cartes.LastError());
            else if (result == null) return "";
            else return result;
        }
        protected virtual string CartesScriptExecute(string command)
        {
            return Execute(command);
        }
        protected virtual void Balloon(string message)
        {
            bool ok;

            if (fLastBalloon.Value == message)
                ok = fLastBalloon.Key < DateTime.Now.AddSeconds(-10);
            else ok = true;
            if (ok)
            {
                cartes.balloon(message);
                fLastBalloon = new KeyValuePair<DateTime, string>(DateTime.Now, message);
            }
        }
        protected virtual new void forensic(string message) // It writes "message" in the swarm log and in the windows event viewer.
        {
            cartes.forensic(message);
        }
        protected virtual new void forensic(string message, Exception e)
        {
            forensic(message + "\r\n" + e.Message);
        }
        protected abstract void CheckAbort(); // If abort is requested, the method should throw an exception.
        protected abstract bool GetIsAborting(); // The method returns true if abort has been requested
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
                            foreach(string api in lAPI)
                                cartes.reset(api);
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
            do
            {
                if (typed)
                    try
                    {
                        component.TypeFromClipboardCheck(value, 0, 0);
                    }
                    catch 
                    {
                        component.TypeWordCheck(value, 0, 0);
                    }
                else component.Value = value;
                CheckAbort();
                Thread.Sleep(1000);
                reset(component);
                if (ToString(component.Value).ToLower() == value.ToLower()) break;
                else
                {
                    if (timeout < DateTime.Now) throw new Exception("I can't assign the value \"" + value + "\" to the component.");
                    Thread.Sleep(1000);
                }
            } while (true);
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
        protected virtual void AdjustWindow(IRPAWin32Component component, int width, int height) // Adjusts the main component window to the indicated size.
        {
            RPAWin32Component lpWindow = component.Root();

            if (lpWindow.ComponentExist())
            {
                if (StringIn(lpWindow.WindowState, "Minimized", "Maximized") || (lpWindow.Visible == 0))
                    lpWindow.Show("Restore");
                if ((lpWindow.width != width) || (lpWindow.height != height))
                    lpWindow.ReSize(width, height);
                if ((lpWindow.x != 0) || (lpWindow.y != 0))
                    lpWindow.Move(0, 0);
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
        protected const string EXIT_SETTINGS_KO = "Settings_" + EXIT_ERROR;
        private static string fCartesPath = null;
        private static Mutex fRC = new Mutex();
        private static CartesObj fCartes = null;
        private string fAbort; // Cartes Script Variable for know when the process musts abort
        private List<MyCartesAPI> apis;
        private string fFileSettings;
        private bool fShowAbort, fVisibleMode, fExecuting;
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
            fAbort = ToString(csvAbort).Trim();
            apis = new List<MyCartesAPI>();
            fFileSettings = null;
            fShowAbort = true;
            fVisibleMode = true;
            fExecuting = false;
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
            return fRC;
        }
        protected override CartesObj getCartes()
        {
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
            return fCartes;
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
        protected override void CheckAbort() // It checks if the variable to abort is 1 to throw an exception
        {
            if (GetIsAborting())
                throw new MyException(EXIT_ABORT, "Abort by user.");
        }
        protected override bool GetIsAborting() 
        {
            return (Abort.Length > 0) && (cartes.Execute(Abort + ".value;") == "1");
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
            return cartes.RegisterIteration(start, typify, data, screenShot ? 1 : 0);
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
        protected string Abort
        {
            get { return fAbort; }
        }  // Read Only

        public bool Execute()  // Execute the process. if succesfull return True, else return false
        {
            DateTime start;
            bool result = false;
            string lsMainFile;
            bool enter = false;
            Mutex mutex = null;

            try
            {
                mutex = new Mutex(true, "Processing area of \"RPA Suite\".", out enter);
                try
                {
                    if (enter)
                    {
                        start = DateTime.Now;
                        CheckRPASuiteVersion();
                        Balloon("I'm opening the project...");
                        lsMainFile = RPAMainFile;
                        if (File.Exists(CurrentPath + "\\" + lsMainFile)) cartes.open(CurrentPath + "\\" + lsMainFile);
                        else if (File.Exists(CurrentPath + "\\Cartes\\" + lsMainFile)) cartes.open(CurrentPath + "\\Cartes\\" + lsMainFile);
                        else cartes.open(RPAMainFile);
                        try
                        {
                            fExecuting = true;
                            Balloon(Name);
                            try
                            {
                                Balloon("Reading settings...");
                                MergeLibrariesAndLoadVariables();
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
                                            if (ShowAbort && (frpaAbort != null))
                                                ShowAbortDialog(frpaAbort);
                                            DoExecute(ref start);
                                        }
                                    }
                                    finally
                                    {
                                        DoEnd();
                                    }
                                }
                                finally
                                {
                                    if (VisibleMode)
                                        Execute("visualmode(0);");
                                }
                                Balloon("\"" + Name + "\" is over.");
                            }
                            catch (Exception e)
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
                        result = true;
                    }
                    else
                    {
                        string message = "Process \"" + Name + "\" is already running." + LF +
                                         "You will not run another process until it finishes.";
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
        public bool VisibleMode
        {
            get { return fVisibleMode; }
            set { fVisibleMode = value; }
        } // Read & Write. It controls the visible mode of Carte. 
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
        public abstract void Close(); // This method should close all the application windows

        public MyCartesProcess Owner
        {
            get { return GetOwner(); }
        }
    }

    public static class CartesObjExtensions
    {
        internal static CartesObj GlobalCartes = null;

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
        public static void Write(this ICredentialStack credential, IRPAWin32Component component)
        {
            IRPAComponent lpC = component;
            credential.Write((RPAComponent)lpC);
        }
        public static T GetComponent<T>(this CartesObj cartesObj, string variablename) where T : class, IRPAComponent
        {
            IRPAComponent component = cartesObj.component(variablename);

            if (component == null) return null;
            else if (component is T result) return result;
            else throw new Exception(variablename + " is a " + component.ActiveXClass());
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
                else
                {
                    DateTime timeout = DateTime.Now.AddSeconds(TimeOut);
                    while (!result)
                    {
                        if (!component.ComponentExist(TimeOut)) result = true;
                        else if (timeout < DateTime.Now) break;
                        else
                        {
                            GlobalCartes.reset(component.api());
                            Thread.Sleep(250);
                        }
                    }
                }
            }catch(Exception e)
            {
                MyObject.Coroner.write("CartesObjExtensions.ComponentNotExist", e);
                throw;
            }

            return result;
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
    }
}
