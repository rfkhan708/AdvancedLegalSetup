using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Word;
using System.Diagnostics;

namespace MyRibbonAddIn
{
    public partial class ThisAddIn
    {
        //bool initialized = false;
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            // return Globals.Factory.GetRibbonFactory().CreateRibbonManager(new Microsoft.Office.Tools.Ribbon.IRibbonExtension[] { new LocalRibbon() });
            return new LocalRibbon();
        }
        // Word.Application wb;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                // wb = this.Application;
               // Microsoft.Office.Interop.Word.ApplicationEvents2_Event wdEvents2 = (Microsoft.Office.Interop.Word.ApplicationEvents2_Event)this.Application;
                //wdEvents2.DocumentOpen += new Word.ApplicationEvents2_DocumentOpenEventHandler(ThisDocument_Open);
                // wb.DocumentOpen += new Word.ApplicationEvents4_DocumentOpenEventHandler(Application_DocumentOpen);
                //var wordApplication = new this.Application() { Visible = true };
                // Listen for documents open
                //wb.DocumentOpen += WordApplicationDocumentOpen;
                // Listen for documents close
                // wb.DocumentBeforeClose += WordApplicationDocumentBeforeClose;
                // Debug.Assert(initialized);
                //this.Application.DocumentOpen += new Word.ApplicationEvents4_DocumentOpenEventHandler(ThisDocument_Open);
               //wdEvents2.NewDocument += new Word.ApplicationEvents2_NewDocumentEventHandler(wdEvents2_NewDocument);
            }
            catch (Exception ex)
            {
                Logger.LogWriter(ex.StackTrace);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="Doc"></param>
        void wdEvents2_NewDocument(Word.Document Doc)
        {
            try
            {
                //string templateUsed = Doc.get_AttachedTemplate().FullName;
                string templateName = Doc.BuiltInDocumentProperties[Microsoft.Office.Interop.Word.WdBuiltInProperty.wdPropertyTemplate].Value.ToString();
                //InstallLocation = Path.GetDirectoryName(new Uri(System.Reflection.Assembly.GetExecutingAssembly().CodeBase).LocalPath);
                if (GlobalEnumClass.PortlandLetterheadforblankpaper.ToUpper() == templateName.ToUpper() || GlobalEnumClass.CentralOregonLetterheadforblankpaper.ToUpper() == templateName.ToUpper())
                {
                    Template.GetInstance().DisplayBJLetter();
                }
            }
            catch (Exception ex)
            {
                Logger.LogWriter(ex.StackTrace);
            }
        }
        private void InitializeCustom()
        {
            //Microsoft.Office.Interop.Word.ApplicationEvents2_Event wdEvents2 = (Microsoft.Office.Interop.Word.ApplicationEvents2_Event)this.Application;
            //wdEvents2.DocumentOpen += new Word.ApplicationEvents2_DocumentOpenEventHandler(ThisDocument_Open);
            //Globals.ThisAddIn.Application.DocumentOpen += new Word.ApplicationEvents4_DocumentOpenEventHandler(Application_DocumentOpen);
            // Globals.ThisAddIn.Application.WindowActivate += new Word.ApplicationEvents4_WindowActivateEventHandler(Application_WindowActivate);
            //initialized = true;
            Globals.ThisAddIn.Application.DocumentOpen += new Word.ApplicationEvents4_DocumentOpenEventHandler(Application_DocumentOpen);
            Globals.ThisAddIn.Application.WindowActivate += new Word.ApplicationEvents4_WindowActivateEventHandler(Application_WindowActivate);
            Globals.ThisAddIn.Application.DocumentBeforeClose += WordApplicationDocumentBeforeClose;
        }
        void ThisDocument_Open(Microsoft.Office.Interop.Word.Document Doc)
        {
            //initialized = true;
            Template.GetInstance().DisplayBJLetter();
        }

        void Application_DocumentOpen(Microsoft.Office.Interop.Word.Document Doc)
        {
            try
            {

                Word.Document doc = this.Application.ActiveDocument;
                if (String.IsNullOrWhiteSpace(doc.Path))
                {

                }
                else
                {
                    //initialized = true;
                    Template.GetInstance().DisplayBJLetter();
                }
            }
            catch (Exception ex)
            {
                Logger.LogWriter(ex.StackTrace);
            }

        }
        void Application_WindowActivate(Microsoft.Office.Interop.Word.Document Doc, Microsoft.Office.Interop.Word.Window Wn)
        {
            try
            {
               // Template.GetInstance().DisplayBJLetter();
                Word.Document docCurr = this.Application.ActiveDocument;
                if (!String.IsNullOrWhiteSpace(docCurr.Path))
                {
                    //Template.GetInstance().DisplayBJLetter();
                    if (!TestDocu.ContainsKey(Doc))
                    {
                        TestDocu.Add(Doc, true);
                        Template.GetInstance().DisplayBJLetter();
                    }
                    // Otherwise, the doc is already in the set of open documents, hence we know the document is already open
                    else
                    {
                        if (TestDocu[Doc] == false)
                            //Console.WriteLine(doc.Name + " is already open!");
                            Template.GetInstance().DisplayBJLetter();
                    }
                }
                //if (initialized == false)
                //{
                //    Word.Document doc = this.Application.ActiveDocument;
                //    if (String.IsNullOrWhiteSpace(doc.Path))
                //    {

                //    }
                //    else
                //    {
                //        initialized = true;
                //        Template.GetInstance().DisplayBJLetter();
                //    }
                //}
            }
            catch (Exception ex)
            {
                Logger.LogWriter(ex.StackTrace);
            }

        }
        private readonly HashSet<Microsoft.Office.Interop.Word.Document> OpenDocuments = new HashSet<Microsoft.Office.Interop.Word.Document>();
        private Dictionary<Microsoft.Office.Interop.Word.Document, bool> TestDocu = new Dictionary<Word.Document, bool>();
        void WordApplicationDocumentBeforeClose(Microsoft.Office.Interop.Word.Document doc, ref bool cancel)
        {
            if(TestDocu.ContainsKey(doc))
            TestDocu.Remove(doc);
            // OpenDocuments.Remove(doc);
            // Console.WriteLine(doc.Name + " closed!");
        }

        void WordApplicationDocumentOpen(Microsoft.Office.Interop.Word.Document doc)
        {
            // If this returns true, the doc is not in the set of open documents, hence the doc is not already open
            if (OpenDocuments.Add(doc))
            {
                OpenDocuments.Add(doc);
                Template.GetInstance().DisplayBJLetter();
            }
            // Otherwise, the doc is already in the set of open documents, hence we know the document is already open
            else
            {
                //Console.WriteLine(doc.Name + " is already open!");
                Template.GetInstance().DisplayBJLetter();
            }
        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
