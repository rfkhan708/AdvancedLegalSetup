using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Office.Interop;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using System.Drawing;


// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new LocalRibbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace MyRibbonAddIn
{
    /// <summary>
    /// 
    /// </summary>
    [ComVisible(true)]
    public class LocalRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        public ADLSNumberingForm form;
        AboutAddIn objAboutAddIn;
        /// <summary>
        /// 
        /// </summary>
        public LocalRibbon()
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);
                form = ADLSNumberingForm.GetInstance();
                //  System.Threading.Tasks.Task.Factory.StartNew(() => fun());

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
       
        #region IRibbonExtensibility Members
        /// <summary>
        /// 
        /// </summary>
        /// <param name="ribbonID"></param>
        /// <returns></returns>
        public string GetCustomUI(string ribbonID)
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);
                return GetResourceText("MyRibbonAddIn.LocalRibbon.xml");
            }
            catch (Exception ex)
            {
                sbTrace.Clear();
                sbTrace.AppendLine("Exception" + ex);
                Logger.SaveLoggerTrace(sbTrace);
                Logger.LogWriter(ex.StackTrace);
                return GetResourceText("MyRibbonAddIn.LocalRibbon.xml");
            }
            finally
            {
                sbTrace.Clear();
                sbTrace.AppendLine("End");
                Logger.SaveLoggerTrace(sbTrace);
            }
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226
        /// <summary>
        /// 
        /// </summary>
        /// <param name="ribbonUI"></param>
        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);
                this.ribbon = ribbonUI;
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

        #endregion

        #region Helpers
        /// <summary>
        /// 
        /// </summary>
        /// <param name="resourceName"></param>
        /// <returns></returns>
        private static string GetResourceText(string resourceName)
        {
            StringBuilder sbTrace = new StringBuilder();
            StreamReader resourceReader = null;
            try
            {
                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);
                Assembly asm = Assembly.GetExecutingAssembly();
                string[] resourceNames = asm.GetManifestResourceNames();
                for (int i = 0; i < resourceNames.Length; ++i)
                {
                    if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                    {
                        using ( resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                        {
                            if (resourceReader != null)
                            {
                                return resourceReader.ReadToEnd();
                            }
                        }
                    }
                }
                return resourceReader.ReadToEnd();
            }
            catch (Exception ex)
            {
                sbTrace.Clear();
                sbTrace.AppendLine("Exception" + ex);
                Logger.SaveLoggerTrace(sbTrace);
                Logger.LogWriter(ex.StackTrace);
                return null;
            }
            finally
            {
                sbTrace.Clear();
                sbTrace.AppendLine("End");
                Logger.SaveLoggerTrace(sbTrace);
            }
        }

        #endregion
        /// <summary>
        /// 
        /// </summary>
        /// <param name="control"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public IEnumerable<Control> GetAll(Control control, Type type)
        {
            IEnumerable<Control> controls = null;
            StringBuilder sbTrace = new StringBuilder();
            try
            {
                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);
                controls = control.Controls.Cast<Control>();
                return controls.SelectMany(ctrls => GetAll(ctrls, type)).Concat(controls).Where(c => c.GetType() == type);
            }
            catch (Exception ex)
            {
                sbTrace.Clear();
                sbTrace.AppendLine("Exception" + ex);
                Logger.SaveLoggerTrace(sbTrace);
                Logger.LogWriter(ex.StackTrace);
                return controls;
            }
            finally
            {
                sbTrace.Clear();
                sbTrace.AppendLine("End");
                Logger.SaveLoggerTrace(sbTrace);
            }
        }
        /// <summary>
        /// Call OnNumberingButton Method
        /// </summary>
        /// <param name="control"></param>
        public void OnNumberingButton(Office.IRibbonControl control)
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);
                form = ADLSNumberingForm.GetInstance();
                if (!form.Visible)
                {
                    form.Show();
                }
                else
                {
                    form.BringToFront();
                }
                var cntls = GetAll(form, typeof(RadioButton));
                foreach (Control cntrl in cntls)
                {
                    RadioButton _rb = (RadioButton)cntrl;
                    if (_rb.Checked)
                    {
                        _rb.Checked = false;
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
        /// <summary>
        /// For Display format numbering Form 
        /// </summary>
        public void ParNumValues()
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);
                Microsoft.Office.Interop.Word.Application appWord = default(Microsoft.Office.Interop.Word.Application);
                appWord = Globals.ThisAddIn.Application;
                if (appWord.ActiveWindow.View.SplitSpecial == WdSpecialPane.wdPaneNone)
                {
                    appWord.ActiveWindow.ActivePane.View.Type = WdViewType.wdNormalView;
                }
                else
                {
                    appWord.ActiveWindow.View.Type = WdViewType.wdNormalView;
                }
                if (form.opt1.Checked)
                {
                    //ParagraphStyle1
                    NumberingClass1.Instance.FormatNumbering();
                }
                else if (form.opt2.Checked)
                {
                    //ParagraphStyle2
                    NumberingClass2.Instance.FormatNumbering();
                }
                else if (form.opt3.Checked)
                {
                    //ParagraphStyle3
                    NumberingClass3.Instance.FormatNumbering();
                }
                else if (form.opt4.Checked)
                {
                    //ParagraphStyle4
                    NumberingClass4.Instance.FormatNumbering();
                }
                else if (form.opt5.Checked)
                {
                    //ParagraphStyle5
                    NumberingClass5.Instance.FormatNumbering();
                }
                else if (form.opt6.Checked)
                {
                    //ParagraphStyle6
                    NumberingClass6.Instance.FormatNumbering();
                }
                else if (form.opt7.Checked)
                {
                    //ParagraphStyle7
                    NumberingClass7.Instance.FormatNumbering();
                }
                else if (form.opt8.Checked)
                {
                    //ParagraphStyle8
                    NumberingClass8.Instance.FormatNumbering();
                }
                else if (form.opt9.Checked)
                {
                    //ParagraphStyle9
                    NumberingClass9.Instance.FormatNumbering();
                }
                else if (form.opt10.Checked)
                {
                    //ParagraphStyle10
                    NumberingClass10.Instance.FormatNumbering();
                }
                else if (form.opt11.Checked)
                {
                    //ParagraphStyle11
                    NumberingClass11.Instance.FormatNumbering();
                }
                else if (form.opt12.Checked)
                {
                    //ParagraphStyle12
                    NumberingClass12.Instance.FormatNumbering();
                }
                else if (form.opt13.Checked)
                {
                    //ParagraphStyle13
                    NumberingClass13.Instance.FormatNumbering();
                }
                else if (form.opt14.Checked)
                {
                    //ParagraphStyle14
                    NumberingClass14.Instance.FormatNumbering();
                }
                else if (form.opt15.Checked)
                {
                    //ParagraphStyle15
                    NumberingClass15.Instance.FormatNumbering();
                }

                else if (form.opt16.Checked)
                {
                    //ParagraphStyle16
                    NumberingClass16.Instance.FormatNumbering();
                }
                else if (form.opt17.Checked)
                {
                    //ParagraphStyle17
                    NumberingClass17.Instance.FormatNumbering();
                }
                else if (form.opt18.Checked)
                {
                    //ParagraphStyle18
                    NumberingClass18.Instance.FormatNumbering();
                }
                else if (form.opt19.Checked)
                {
                    //ParagraphStyle19
                    NumberingClass19.Instance.FormatNumbering();
                }
                else if (form.opt20.Checked)
                {
                    //ParagraphStyle20
                    NumberingClass20.Instance.FormatNumbering();
                }
                else if (form.opt21.Checked)
                {
                    //ParagraphStyle21
                    NumberingClass21.Instance.FormatNumbering();
                }
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading1].LinkToListTemplate(ListTemplate: appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6], ListLevelNumber: 1);

                if (appWord.ActiveWindow.View.SplitSpecial == WdSpecialPane.wdPaneNone)
                {
                    appWord.ActiveWindow.ActivePane.View.Type = WdViewType.wdPrintView;
                }
                else
                {
                    appWord.ActiveWindow.View.Type = WdViewType.wdPrintView;
                }

                appWord.Selection.HomeKey(Unit: WdUnits.wdStory);
                appWord.Application.GoBack();
                //form.Hide();
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
        /// For apply formating in OnPasteUnFormated Content under Legal Ribbon
        /// </summary>
        /// <param name="control"></param>
        public void OnPasteUnFormated(Office.IRibbonControl control)
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);
                Tools.Instance.PasteSpecialUnformatted();
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
        /// Call OnHiddenParagraphButton Method for legal rebbon
        /// </summary>
        /// <param name="control"></param>
        public void OnHiddenParagraphButton(Office.IRibbonControl control)
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);
                Tools.Instance.HiddnPara();
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
        /// Call OnStrTosmrtQuotes method for legal ribbon
        /// </summary>
        /// <param name="control"></param>
        public void OnStrTosmrtQuotes(Office.IRibbonControl control)
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);

                Tools.Instance.StraightToSmartQuotes();
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
        /// Call OnSmrtToStrQuotes Method in Legoal Ribbon
        /// </summary>
        /// <param name="control"></param>
        public void OnSmrtToStrQuotes(Office.IRibbonControl control)
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);

                Tools.Instance.SmartToStraightQuotes();
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
        /// Call to OnSmrtToStrApostrophe for legal ribbon
        /// </summary>
        /// <param name="control"></param>
        public void OnSmrtToStrApostrophe(Office.IRibbonControl control)
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);
                Tools.Instance.SmartApostropheToStraight();
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
        /// Call OnStrTosmrtApostrophe in legal ribbon
        /// </summary>
        /// <param name="control"></param>
        public void OnStrTosmrtApostrophe(Office.IRibbonControl control)
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);
                Tools.Instance.StraightApostropheToSmart();
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
        /// Call OnNewDocumnet in legal ribbon
        /// </summary>
        /// <param name="control"></param>
        public void OnNewDocument(Office.IRibbonControl control)
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);
                Tools.Instance.NewDocument();
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
        /// Call OnAboutAddIn method in legal ribbon
        /// </summary>
        /// <param name="control"></param>
        public void OnAboutAddIn(Office.IRibbonControl control)
        {

            StringBuilder sbTrace = new StringBuilder();
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);
                //Tools.Instance.Help();
                objAboutAddIn = AboutAddIn.GetInstance();
                if (!objAboutAddIn.Visible)
                {
                    objAboutAddIn.Show();
                }
                else
                {
                    objAboutAddIn.BringToFront();
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

        public stdole.IPictureDisp GetCustomImage(Office.IRibbonControl control)
        {
            StringBuilder sbTrace = new StringBuilder();
            stdole.IPictureDisp pictureDisp = null;
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);

                
            switch (control.Id)
            {
                case "textButton":
                    //Image image = Image.FromFile(Environment.CurrentDirectory+ "\\Image\\legal-numbering.bmp");
                    pictureDisp = ImageConverter.Convert(Properties.Resources.legalnumbering);
                    break;
                case "SmartToStraightButton":
                    //Image image2 = Image.FromFile(Environment.CurrentDirectory + "\\Image\\smart-to-straight-quotes.bmp");
                    pictureDisp = ImageConverter.Convert(Properties.Resources.smarttostraightquotes);
                    break;
                case "StraightToSmartButton":
                    //Image image3 = Image.FromFile(Environment.CurrentDirectory + "\\Image\\straight-to-smart-quotes.bmp");
                    pictureDisp = ImageConverter.Convert(Properties.Resources.straighttosmartquotes);
                    break;
                case "SmartToStraightButton2":
                    // Image image4 = Image.FromFile(Environment.CurrentDirectory + "\\Image\\smart-to-straight-apostrophes.bmp");
                    pictureDisp = ImageConverter.Convert(Properties.Resources.smarttostraightapostrophes);
                    break;
                case "SmartToStraightButton23":
                    // Image image5 = Image.FromFile(Environment.CurrentDirectory + "\\Image\\straight-to-smart-apostrophes.bmp");
                    pictureDisp = ImageConverter.Convert(Properties.Resources.straighttosmartapostrophes);
                    break;
                case "btnBlankDoc":
                    // Image image6 = Image.FromFile(Environment.CurrentDirectory + "\\Image\\new-blank-document.bmp");
                    pictureDisp = ImageConverter.Convert(Properties.Resources.newblankdocument);
                    break;

            }
                return pictureDisp;
            }
            catch (Exception ex)
            {
                sbTrace.Clear();
                sbTrace.AppendLine("Exception" + ex);
                Logger.SaveLoggerTrace(sbTrace);
                Logger.LogWriter(ex.StackTrace);
                return pictureDisp;
            }
            finally
            {
                sbTrace.Clear();
                sbTrace.AppendLine("End");
                Logger.SaveLoggerTrace(sbTrace);
            }
            
        }
        public void OnTemplateEdit(Office.IRibbonControl control)
        {

            StringBuilder sbTrace = new StringBuilder();
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);
                Template.GetInstance().DisplayBJLetter();
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
    internal class ImageConverter : System.Windows.Forms.AxHost
    {
        private ImageConverter() : base(null)
        {
        }

        public static stdole.IPictureDisp Convert(System.Drawing.Image image)
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);
                return (stdole.IPictureDisp)AxHost.GetIPictureDispFromPicture(image);
            }
            catch (Exception ex)
            {
                sbTrace.Clear();
                sbTrace.AppendLine("Exception" + ex);
                Logger.SaveLoggerTrace(sbTrace);
                Logger.LogWriter(ex.StackTrace);
                return (stdole.IPictureDisp)AxHost.GetIPictureDispFromPicture(image);
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
