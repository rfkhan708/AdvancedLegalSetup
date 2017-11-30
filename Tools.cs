using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualBasic;
using Microsoft.Office.Interop.Word;
using System.Diagnostics;
using System.Reflection;

namespace MyRibbonAddIn
{
    public class Tools
    {
        private static Tools instance = null;
        private Tools()
        {
        }
        public static Tools Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new Tools();
                }
                return instance;
            }
        }
        /// <summary>
        /// For Calling SmartApostropheToStraight method in SmartApostropheToStraight content under Legal Ribbon
        /// </summary>
        public void SmartApostropheToStraight()
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);
                bool repQuotes = false;
                Microsoft.Office.Interop.Word.Application appWord = default(Microsoft.Office.Interop.Word.Application);
                appWord = Globals.ThisAddIn.Application;
                repQuotes = appWord.Options.AutoFormatAsYouTypeReplaceQuotes;

                if (repQuotes == true)
                {
                    var _with1 = appWord.Options;
                    _with1.AutoFormatAsYouTypeReplaceQuotes = false;
                }
                appWord.Selection.Find.ClearFormatting();
                appWord.Selection.Find.Replacement.ClearFormatting();
                var _with2 = appWord.Selection.Find;
                _with2.Text = "’";
                _with2.Replacement.Text = "'";
                _with2.Forward = true;
                _with2.Wrap = WdFindWrap.wdFindAsk;
                _with2.Format = false;
                _with2.MatchCase = false;
                _with2.MatchWholeWord = false;
                _with2.MatchWildcards = false;
                _with2.MatchSoundsLike = false;
                _with2.MatchAllWordForms = false;
                appWord.Selection.Find.Execute(Replace: WdReplace.wdReplaceAll);
                if (repQuotes == true)
                {
                    var _with3 = appWord.Options;
                    _with3.AutoFormatAsYouTypeReplaceQuotes = true;

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
        /// For Calling t=StraightApostropheToSmart method in StraightApostropheToSmart content under legal Ribbon
        /// </summary>
        public void StraightApostropheToSmart()
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);
                bool repQuotes = false;
                Microsoft.Office.Interop.Word.Application appWord = default(Microsoft.Office.Interop.Word.Application);
                appWord = Globals.ThisAddIn.Application;

                repQuotes = appWord.Options.AutoFormatAsYouTypeReplaceQuotes;

                if (repQuotes == true)
                {
                    var _with1 = appWord.Options;
                    _with1.AutoFormatAsYouTypeReplaceQuotes = false;
                }

                appWord.Selection.Find.ClearFormatting();
                appWord.Selection.Find.Replacement.ClearFormatting();
                var _with2 = appWord.Selection.Find;
                _with2.Text = "'";
                _with2.Replacement.Text = "’";
                _with2.Forward = true;
                _with2.Wrap = WdFindWrap.wdFindAsk;
                _with2.Format = false;
                _with2.MatchCase = false;
                _with2.MatchWholeWord = false;
                _with2.MatchWildcards = false;
                _with2.MatchSoundsLike = false;
                _with2.MatchAllWordForms = false;
                appWord.Selection.Find.Execute(Replace: WdReplace.wdReplaceAll);
                if (repQuotes == true)
                {
                    var _with3 = appWord.Options;
                    _with3.AutoFormatAsYouTypeReplaceQuotes = true;
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
        /// For Calling HiddnPara method in HiddnPara content under legal Ribbon
        /// </summary>
        public void HiddnPara()
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);

                string hidPara = null;
                string headNo = null;
                Microsoft.Office.Interop.Word.Application appWord = default(Microsoft.Office.Interop.Word.Application);
                appWord = Globals.ThisAddIn.Application;
                //Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;

                headNo = appWord.Selection.get_Style().ToString();
                headNo = headNo.PadRight(1);
                hidPara = "HiddenPara" + headNo;

                appWord.Selection.TypeParagraph();
                appWord.Selection.MoveUp(Unit: WdUnits.wdLine, Count: 1);
                appWord.Selection.EndKey(Unit: WdUnits.wdLine);
                appWord.Selection.MoveRight(Unit: WdUnits.wdCharacter, Count: 1, Extend: WdMovementType.wdExtend);
                var _with1 = appWord.Selection.Font;
                _with1.Hidden = 1;
                _with1.ColorIndex = WdColorIndex.wdBlue;
                appWord.Selection.MoveRight(Unit: WdUnits.wdCharacter, Count: 1);
                object StyleActiveDoc = appWord.ActiveDocument.Styles["Normal"];
                appWord.Selection.set_Style(ref StyleActiveDoc);
                // ERROR: Not supported in C#: OnErrorStatement

                appWord.ActiveDocument.Styles.Add(Name: hidPara, Type: WdStyleType.wdStyleTypeParagraph);
                appWord.ActiveDocument.Styles[hidPara].AutomaticallyUpdate = false;
                appWord.Selection.set_Style(appWord.ActiveDocument.Styles[hidPara]);


                appWord.Selection.set_Style(appWord.ActiveDocument.Styles[hidPara]);
                var _with2 = appWord.ActiveDocument.Styles[hidPara];
                _with2.AutomaticallyUpdate = false;
                object objbasestyle = "Normal";
                _with2.set_BaseStyle(ref objbasestyle);
                object objParaGraph = "Body Text";
                _with2.set_NextParagraphStyle(ref objParaGraph);

                var _with3 = appWord.ActiveDocument.Styles[hidPara].ParagraphFormat;
                _with3.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                _with3.SpaceBefore = 0;
                _with3.SpaceAfter = 12;
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
        /// For calling PasteSpecialUnformatted in PasteSpecialUnformatted content under legal Ribbon
        /// </summary>
        public void PasteSpecialUnformatted()
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);

                Microsoft.Office.Interop.Word.Application appWord = default(Microsoft.Office.Interop.Word.Application);
                appWord = Globals.ThisAddIn.Application;
                appWord.Selection.PasteSpecial(Link: false, DataType: WdPasteDataType.wdPasteText, Placement: WdOLEPlacement.wdInLine, DisplayAsIcon: false);
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
        /// For calling SmartToStraightQuotes method in SmartToStraightQuotes content under legal ribbon
        /// </summary>
        public void SmartToStraightQuotes()
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);
                bool repQuotes = false;
                Microsoft.Office.Interop.Word.Application appWord = default(Microsoft.Office.Interop.Word.Application);
                appWord = Globals.ThisAddIn.Application;
                // Replaces smart quotes with straight quotes

                repQuotes = appWord.Options.AutoFormatAsYouTypeReplaceQuotes;

                if (repQuotes == true)
                {
                    var _with1 = appWord.Options;
                    _with1.AutoFormatAsYouTypeReplaceQuotes = false;
                }

                appWord.Selection.Find.ClearFormatting();
                appWord.Selection.Find.Replacement.ClearFormatting();
                var _with2 = appWord.Selection.Find;
                _with2.Forward = true;
                _with2.Wrap = WdFindWrap.wdFindContinue;
                _with2.Text = "\"";
                _with2.Replacement.Text = "\"";
                appWord.Selection.Find.Execute(Replace: WdReplace.wdReplaceAll);

                if (repQuotes == true)
                {
                    var _with3 = appWord.Options;
                    _with3.AutoFormatAsYouTypeReplaceQuotes = true;
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
        /// For Calling StraightToSmartQuotes Method in StraightToSmart QuotesContent in legal Ribbon
        /// </summary>
        public void StraightToSmartQuotes()
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);
                bool repQuotes = false;
                Microsoft.Office.Interop.Word.Application appWord = default(Microsoft.Office.Interop.Word.Application);
                appWord = Globals.ThisAddIn.Application;
                // Replaces straight quotes with smart quotes

                repQuotes = appWord.Options.AutoFormatAsYouTypeReplaceQuotes;

                if (repQuotes == false)
                {
                    var _with1 = appWord.Options;
                    _with1.AutoFormatAsYouTypeReplaceQuotes = true;
                }

                appWord.Selection.Find.ClearFormatting();
                appWord.Selection.Find.Replacement.ClearFormatting();
                var _with2 = appWord.Selection.Find;
                _with2.Forward = true;
                _with2.Wrap = WdFindWrap.wdFindContinue;
                _with2.Text = "\"";
                _with2.Replacement.Text = "\"";
                appWord.Selection.Find.Execute(Replace: WdReplace.wdReplaceAll);

                if (repQuotes == false)
                {
                    var _with3 = appWord.Options;
                    _with3.AutoFormatAsYouTypeReplaceQuotes = false;
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
        /// We ar calling Help Method for Help Content under Legal Rebbon
        /// </summary>
        public void Help()
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);

                Process.Start("http://www.expertsourcing.com/alrt");
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
        /// Call New Document in Legal Ribbon
        /// </summary>
        public void NewDocument()
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);

                object oMissing = System.Reflection.Missing.Value;
                object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */

                //Start Word and create a new document.
                Microsoft.Office.Interop.Word._Application oWord;
                Microsoft.Office.Interop.Word._Document oDoc;
                oWord = new Microsoft.Office.Interop.Word.Application();
                oWord.Visible = true;
                oDoc = oWord.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing);

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