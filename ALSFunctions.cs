using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyRibbonAddIn.ALS_FWW_Word
{
    /// <summary>
    /// 
    /// </summary>
    public class ALSFunctions
    {
        public Microsoft.Office.Interop.Word.Application appWord;
        const string DBPath = "D:\\Templates\\";
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        internal ADODB.Connection oConn()
        {
            StringBuilder sbTrace = new StringBuilder();
            ADODB.Connection functionReturnValue = default(ADODB.Connection);
            try
            {
                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);
                functionReturnValue = new ADODB.Connection();
                var _with1 = functionReturnValue;
                //.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\szielinski\Downloads\FWW templates\author.accdb;Persist Security Info=False;"
                _with1.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=DSN=MS Access Database;DBQ=" + DBPath + "author.accdb;DefaultDir=" + DBPath + ";DriverId=25;FIL=MS Access;MaxBufferSize=2048;PageTimeout=5;UID=admin;Initial Catalog=" + DBPath + "author.accdb";
                _with1.Open();
                return functionReturnValue;
            }
            catch (Exception ex)
            {

                sbTrace.Clear();
                sbTrace.AppendLine("Exception" + ex);
                Logger.SaveLoggerTrace(sbTrace);
                Logger.LogWriter(ex.StackTrace);
                return functionReturnValue;
            }
            finally
            {

                sbTrace.Clear();
                sbTrace.AppendLine("End");
                Logger.SaveLoggerTrace(sbTrace);
            }
            //Check for live connection
            //If oConn.State <> 0 Then
            //Response = MsgBox("Unable to open ADODB connection", MsgBoxStyle.Exclamation)
            //Exit Function
            //End If
        }
        /// <summary>
        /// 
        /// </summary>
        public void UpdateAll()
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {
                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);
                SetAppWord();
                //Microsoft.Office.Interop.Word.Section sec = default(Microsoft.Office.Interop.Word.Section);

                var _with1 = appWord;
                _with1.ActiveDocument.Fields.Update();
                foreach (Microsoft.Office.Interop.Word.Section sec in _with1.ActiveDocument.Sections)
                {
                    var _with2 = sec;
                    _with2.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Fields.Update();
                    _with2.Headers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Fields.Update();
                    _with2.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Fields.Update();
                    _with2.Footers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Fields.Update();
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
        /// <summary>
        /// 
        /// </summary>
        internal void SetAppWord()
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {
                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);
                appWord = Globals.ThisAddIn.Application;
            }
            catch (Exception ex)
            {
                sbTrace.Clear();
                sbTrace.AppendLine("Exception" + ex);
                Logger.SaveLoggerTrace(sbTrace);
                Logger.LogWriter(ex.StackTrace);
                Debug.Print("No Document Found");
                //No doc
                Interaction.MsgBox("Open Word Application not found; document assembly cannot continue.", MsgBoxStyle.Critical, "Word not Found" + ex.Message.ToString());
                //end
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
