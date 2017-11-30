using MyRibbonAddIn.ALS_FWW_Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyRibbonAddIn
{
    public class Template
    {
        frmMainForm form = null;
        public Template()
        {
            form = new frmMainForm();
        }

        private static Template _instance;
        public static Template GetInstance()
        {
            if (_instance == null) _instance = new Template();
            return _instance;
        }
       public void DisplayBJLetter()
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {
                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);
                if (form != null)
                {
                    if (!form.Visible)
                    {
                        form.TopMost = true;
                        form.Show();
                    }
                    else
                    {
                        form.TopMost = true;
                        form.BringToFront();
                    }
                }
            }
            catch (Exception ex)
            {
                sbTrace.Clear();
                sbTrace.AppendLine("Exception" + ex);
                Logger.SaveLoggerTrace(sbTrace);
                Logger.LogWriter(ex.StackTrace);
            }
            finally
            {
                sbTrace.Clear();
                sbTrace.AppendLine("End");
                Logger.SaveLoggerTrace(sbTrace);
            }
        }
    }
}
