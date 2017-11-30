using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace MyRibbonAddIn
{
    public class Logger
    {
        private static string wanted_path = string.Empty;
        /// <summary>
        /// For Write Error in LogWriter for the Application
        /// </summary>
        /// <param name="logMessage"></param>
        public static void LogWriter(string logMessage)
        {
            bool IsWrite = false;
            try
            {
                if (IsWrite == true)
                {
                    string wanted_path = Path.GetDirectoryName(Path.GetDirectoryName(System.IO.Directory.GetCurrentDirectory()));
                    using (StreamWriter txtWriter = File.AppendText(wanted_path + "\\LoggerError.txt"))
                    {
                        txtWriter.Write("\r\nLog Entry : ");
                        txtWriter.WriteLine("{0} {1}", DateTime.Now.ToLongTimeString(),
                            DateTime.Now.ToLongDateString());
                        txtWriter.WriteLine("  :");
                        txtWriter.WriteLine("  :{0}", logMessage);
                        txtWriter.WriteLine("-------------------------------");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message.ToString());
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sbTrace"></param>
        public static void SaveLoggerTrace(StringBuilder sbTrace)
        {
            bool IsWrite = false;
            try
            {
                if (IsWrite == true)
                {
                    string wanted_path = Path.GetDirectoryName(Path.GetDirectoryName(System.IO.Directory.GetCurrentDirectory()));
                    using (StreamWriter txtWriter = File.AppendText(wanted_path + "\\LoggerTraceError.txt"))
                    {
                        txtWriter.Write("\r\nLog Entry : ");
                        txtWriter.WriteLine("{0} {1}", DateTime.Now.ToLongTimeString(),
                        DateTime.Now.ToLongDateString());
                        txtWriter.WriteLine("  :");
                        txtWriter.WriteLine("  :{0}", sbTrace);
                        txtWriter.WriteLine("-------------------------------");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message.ToString());
            }
        }
    }
}
