using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml;
using System.Diagnostics;

namespace MiTools
{
    [ComVisible(false)]
    public abstract class MyObject : Object// This abstract class allows you to group some very useful methods
    {
        private static System.Threading.Mutex fRCel = new System.Threading.Mutex();
        private static MyForensic fCoroner = new MyForensic();
        private static bool fVerbose = false;
        private static string[] sgbt = { "true", "1" }, sgbf = { "false", "0" };
        private static string fCurrentPath = null, fCurrentApp = null;
        private static Version fCurrent = null;
        private static int fInstances = 0;

        // Implementation of MyObject class
        public MyObject()
        {
            fRCel.WaitOne();
            try
            {
                if (fCurrentPath == null)
                {
                    fVerbose = IsDebug;
                    fCurrentPath = Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName);
                    fCurrentApp = Path.GetFileName(Process.GetCurrentProcess().MainModule.FileName);
                }
            }
            finally
            {
                fInstances++;
                fRCel.ReleaseMutex();
            }
        }
        ~MyObject()
        {
            fRCel.WaitOne();
            try
            {
                fInstances--;
            }
            finally
            {
                fRCel.ReleaseMutex();
            }
        }

        private static bool GetIsDebug()
        {
#if (DEBUG)
            return true;
#else
            return false;
#endif
        }
        private Version getLibraryVersionP() // It returns the version of this library
        {
            if (fCurrent == null)
            {
                if (!Version.TryParse(ToString(getLibraryVersion()), out fCurrent))
                    throw new Exception(getLibraryVersion() + " is not a valid version number.");
            }
            return fCurrent;
        }

        protected static string getLibraryVersion()
        {
            return "1.6.5.51";
        }
        protected virtual string getCurrentFile()
        {
            return fCurrentApp;
        }
        protected static void setCoroner(MyForensic instance)
        {
            fCoroner = instance;
        }
        protected virtual bool forensic(EventLogEntryType type, string message) // It writes "message" in the windows event viewer.
        {
            return fCoroner.forensic(type, message);
        }
        protected void forensic(string message)
        {
            forensic(EventLogEntryType.Information, message);
        }
        protected virtual void forensic(string message, Exception e)
        {
            if (Verbose)
            {
                string lsF = string.Empty;

                if (e is MyException me)
                    lsF = "Code      : " + me.code + "\r\n";
                lsF = message + "\r\n" +
                      "P.I.D.    : " + Process.GetCurrentProcess().Id.ToString() + "\r\n" +
                      "Exception : " + e.GetType().FullName + "\r\n" +
                      lsF +
                      "Message   : " + e.Message;
                forensic(EventLogEntryType.Error, lsF);
            }
        }
        protected virtual DateTime getNow() // Returns the system date and time
        {
            return DateTime.Now;
        }

        public static T Casting<T>(object value) where T : class
        {
            try
            {
                if (value == null) return null;
                else if (value is T result) return result;
                else if (value is MyObject myvalue) throw new Exception("object is a " + myvalue.ClassName);
                else throw new Exception("object is a " + value.GetType().FullName);
            }
            catch (Exception e)
            {
                Coroner.write("T MyObject::Casting<T>(object)", e);
                throw;
            }
        }
        public static string AlignRight(string text, char character, int length)
        {
            if (text == null) return AlignRight(string.Empty, character, length);
            else while (text.Length < length)
                    text = character + text;
            return text;
        }
        public static string ConcatenateSpecial(string a, string separator, string b)
        {
            if (a == null) return ConcatenateSpecial(string.Empty, separator, b);
            else if (separator == null) return ConcatenateSpecial(a, string.Empty, b);
            else if (b == null) return ConcatenateSpecial(a, separator, string.Empty);
            else if (a.Length == 0) return b;
            else if (separator.Length == 0) return a + b;
            else if (b.Length == 0) return a;
            else return a + separator + b;
        }
        public static string ConcatenateSpecial(string[] a, string separator)
        {
            string result = string.Empty;

            foreach (string item in a)
                result = ConcatenateSpecial(result, separator, item);
            return result;
        }
        public static string NoDiacritics(string a) // It returns a string without diacritic symbols
        {
            StringBuilder resultado = new StringBuilder(ToString(a).Trim());
            int i;

            i = resultado.Length - 1;
            while (i >= 0)
            {
                switch (resultado[i])
                {
                    case 'á': resultado[i] = 'a'; break;
                    case 'é': resultado[i] = 'e'; break;
                    case 'í': resultado[i] = 'i'; break;
                    case 'ó': resultado[i] = 'o'; break;
                    case 'ú': resultado[i] = 'u'; break;
                    case 'ñ': resultado[i] = 'n'; break;
                    case 'à': resultado[i] = 'a'; break;
                    case 'è': resultado[i] = 'e'; break;
                    case 'ì': resultado[i] = 'i'; break;
                    case 'ò': resultado[i] = 'o'; break;
                    case 'ù': resultado[i] = 'u'; break;
                    case 'ü': resultado[i] = 'u'; break;
                }
                i--;
            }
            return resultado.ToString().Replace("  ", " ").Trim();
        }
        public static bool StringIn(string text, params string[] texts) // If "text" is in "texts" (case insensitive), it returns true
        {
            bool result = false;
            string sAux;

            sAux = ToString(text).ToLower();
            foreach (string item in texts)
            {
                if (text == null)
                {
                    if (item == null)
                    {
                        result = true;
                        break;
                    }
                }
                else if ((item != null) && (sAux == item.ToLower()))
                {
                    result = true;
                    break;
                }
            }
            return result;
        }
        public static bool StringIn(string text, List<string> texts)
        {
            return StringIn(text, texts.ToArray());
        }
        public static bool IntegerIn(int value, params int[] values) // If "value" is in "values", it returns true
        {
            bool result = false;

            foreach (int item in values)
            {
                if (value == item)
                {
                    result = true;
                    break;
                }
            }
            return result;
        }
        public static string ToString(object value) // If "value" is null, it returns an empty string.
        {
            if (value == null) return string.Empty;
            else if (value is string s) return s;
            else if (value is XmlNode node) return node == null ? string.Empty : ToString(node.InnerText);
            else if (value is IEnumerable<KeyValuePair<string, int>> dtsi)
            {
                string result = "";

                foreach (KeyValuePair<string, int> pair in dtsi)
                {
                    result += pair.Key + "=" + pair.Value.ToString() + LF;
                }
                return result;
            }
            else if (value is string[] ast) return ConcatenateSpecial(ast, ", ");
            else if (value is ICollection<string> ics)
            {
                string result = "";

                foreach (string item in ics)
                {
                    result += item + LF;
                }
                return result;
            }
            else return value.ToString();
        }
        public static bool ToBool(string text) // It converts to bool
        {
            if (text == null) return false;
            else if (StringIn(text, sgbt)) return true;
            else if (StringIn(text, sgbf)) return false;
            else throw new Exception("\"" + text + "\" is not a valid boolean value.");
        }
        public static bool ToBool(int value) // It converts to bool
        {
            if (value == 0) return false;
            else return true;
        }
        public virtual int ToIntDef(string value, int defaultvalue)
        {
            int result;

            if (int.TryParse(value, out result)) return result;
            else return defaultvalue;
        }
        public virtual double ToDouble(string value)
        {
            return Convert.ToDouble(value);
        }
        public double ToDoubleDef(string value, double defaultvalue)
        {
            double resultado = 0;

            try
            {
                resultado = ToDouble(value);
            }
            catch
            {
                resultado = defaultvalue;
            }
            return resultado;
        }

        public string ClassName
        {
            get { return GetType().FullName; }
        } // Returns the instance class name
        public Version LibraryVersion
        {
            get { return getLibraryVersionP(); }
        }  // Read Only. It returns the version of this library
        public static int Instances // Instances counter
        {
            get { return fInstances; }
        }
        public static bool IsDebug // Read Only. If the build is in debug mode, it returns true.
        {
            get { return GetIsDebug(); }
        }
        public static bool Verbose // Enables or disables the ability to write exceptions thrown during execution to the Windows Event Viewer.
        {
            get { return fVerbose; }
            set { fVerbose = value; }
        }
        public static MyForensic Coroner // This object writes to the log
        {
            get { return fCoroner; }
        }
        public string CurrentPath // The directory of the executable
        {
            get { return fCurrentPath; }
        }
        public string CurrentFile // The name of the executable file
        {
            get { return getCurrentFile(); }
        }
        public static string LF // Returns a string with a line break and carriage return
        {
            get { return "\r\n"; }
        }
        public DateTime Now // Returns the system date and time
        {
            get { return getNow(); }
        }
    }

    public class MyException : Exception // This class allows you to throw exceptions with an error code
    {
        private string fCode;
        public MyException(string code, string message) : base(message)
        {
            fCode = MyObject.ToString(code).Trim();
            fCode = fCode.Substring(0, Math.Min(fCode.Length, 16));
            Data["code"] = fCode;
        }

        public string code
        {
            get { return fCode;  }
        }
    }

    public class MyForensic : MyObject  // This class writes messages in the windows event viewer.
    {
        private static System.Threading.Mutex fRCel = new System.Threading.Mutex();
        private static IntPtr fEventLog = IntPtr.Zero;
        private static int fForensics = 0;
        private static string gSourceFile = null;

        // Functions for writing in the Event Viewer
        [DllImport("advapi32.dll", EntryPoint = "RegisterEventSourceW", CharSet = CharSet.Unicode,
                   CallingConvention = CallingConvention.StdCall)]
        private static extern IntPtr RegisterEventSource(string lpUNCServerName, string lpSourceName);
        [DllImport("advapi32.dll", EntryPoint = "DeregisterEventSource", CharSet = CharSet.Unicode,
                   CallingConvention = CallingConvention.StdCall)]
        private static extern bool DeregisterEventSource(IntPtr handle);
        [DllImport("advapi32.dll", EntryPoint = "ReportEventW", CharSet = CharSet.Unicode,
                   CallingConvention = CallingConvention.StdCall)]
        private static extern bool ReportEvent(IntPtr hEventLog, ushort wType, ushort wCategory, int dwEventID,
                                               IntPtr lpUserSid, ushort wNumStrings, uint dwDataSize,
                                               string[] lpStrings, byte[] lpRawData);

        public MyForensic() : base()
        {
            RC.WaitOne();
            try
            {
                fForensics++;
                if (gSourceFile == null)
                    try
                    {
                        gSourceFile = Path.GetFileName(Assembly.GetEntryAssembly().Location);
                    }
                    catch
                    {
                        gSourceFile = Path.GetFileName(Assembly.GetExecutingAssembly().Location);
                    }
            }
            finally
            {
                RC.ReleaseMutex();
            }
        }
        ~MyForensic()
        {
            fRCel.WaitOne();
            try
            {
                fForensics--;
                if ((fForensics == 0) && (IntPtr.Zero != fEventLog))
                {
                    if (DeregisterEventSource(fEventLog))
                        fEventLog = IntPtr.Zero;
                }
            }
            finally
            {
                fRCel.ReleaseMutex();
            }
        }

        protected override bool forensic(EventLogEntryType type, string message) // It writes "message" in the windows event viewer.
        {
            bool result = false;

            RC.WaitOne();
            try
            {
                if (IntPtr.Zero == fEventLog)
                    fEventLog = RegisterEventSource(null, gSourceFile);
                if (IntPtr.Zero != fEventLog)
                {
                    string[] messages = new string[1];
                    byte[] rawData = null;
                    ushort liType;

                    messages[0] = message;
                    switch (type)
                    {
                        case EventLogEntryType.Error: liType = 0x0001; break;
                        case EventLogEntryType.Warning: liType = 0x0002; break;
                        case EventLogEntryType.Information: liType = 0x0004; break;
                        case EventLogEntryType.SuccessAudit: liType = 0x0008; break;
                        case EventLogEntryType.FailureAudit: liType = 0x0016; break;
                        default: liType = 0x0001; break;
                    }
                    result = ReportEvent(fEventLog, liType, 0, 0, IntPtr.Zero, 1, 0, messages, rawData);
                }
            }
            finally
            {
                RC.ReleaseMutex();
            }
            return result;
        }
        protected virtual string GetSource()
        {
            return gSourceFile;
        }
        protected virtual void SetSource(string value)
        {
            if (gSourceFile != value)
            {
                RC.WaitOne();
                try
                {
                    if (ToString(value).Length == 0) throw new Exception("The source name for the event viewer cannot be an empty string.");
                    gSourceFile = ToString(value);
                }
                finally
                {
                    RC.ReleaseMutex();
                }
            }
        }

        protected System.Threading.Mutex RC
        {
            get { return fRCel; }
        }

        public void write(EventLogEntryType type, string message)
        {
            forensic(type, message);
        }
        public void write(string message)
        {
            forensic(message);
        }
        public void write(string message, Exception e)
        {
            forensic(message, e);
        }

        public string Source // It's the source used to log the messages in the event viewer.
        {
            get { return GetSource(); }
        }
    }

    public static class MiToolsExtensions
    {
        public static string Concatenate(this string[] instance, string separator)
        {
            return MyObject.ConcatenateSpecial(instance, separator);
        }
        public static void AddSynonyms<TKey, TValue>(this Dictionary<TKey, TValue> instance, params TKey[] synonyms)
        {
            TValue value;

            foreach (TKey item in synonyms)
            {
                if (instance.TryGetValue(item, out value))
                {
                    foreach (TKey pair in synonyms)
                    {
                        instance[pair] = value;
                    }
                    break;
                }
            }
        }
        public static string ToXML<TValue>(this Dictionary<string, TValue> instance, string bleeding)
        {
            string result = string.Empty, s;

            foreach (KeyValuePair<string, TValue> cell in instance)
            {
                s = MyObject.ToString(cell.Value).Trim();
                if (s.Length > 0)
                    result = result + bleeding + "<" + cell.Key + ">" + s + "</" + cell.Key + ">" + "\r\n";
                else result = result + bleeding + "</" + cell.Key + ">" + "\r\n";
            }
            return result;
        }
        public static bool IsIn(this int a, params int[] values)
        {
            return MyObject.IntegerIn(a, values);
        }
        public static bool IsIn(this string a, params string[] texts)
        {
            return MyObject.StringIn(a, texts);
        }
        public static bool IsIn(this string a, List<string> texts)
        {
            return MyObject.StringIn(a, texts);
        }
        public static bool IsIn(this char a, params char[] chars)
        {
            bool result = false;

            foreach (char item in chars)
            {
                if (a == item)
                {
                    result = true;
                    break;
                }
            }
            return result;
        }
    }
}
