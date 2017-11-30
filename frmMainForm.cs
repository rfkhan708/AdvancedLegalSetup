using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic;
using MyRibbonAddIn.ALS_FWW_Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MyRibbonAddIn.ALS_FWW_Word
{
    /// <summary>
    /// 
    /// </summary>
    public partial class frmMainForm : Form
    {
        public System.Data.DataTable gAuthDT = new System.Data.DataTable();
        public System.Data.DataTable gTypDt = new System.Data.DataTable();
        frmAuthorFind Authorfrm = null;
        frmTypistFind TypistObjfrm = null;
        public string[] strTitleList = null;//String.Split("Attorney~Administrator~Marketing Manager~Paralegal~Administrative Assistant", '~',Convert.ToChar(Constants.vbTextCompare));
        public System.Drawing.Color BGColorD;
        public System.Drawing.Color BGColorE;
        public MyRibbonAddIn.ALS_FWW_Word.ALSFunctions als = new MyRibbonAddIn.ALS_FWW_Word.ALSFunctions();
        public Microsoft.Office.Interop.Word.Application appWord;
        //public event cmbAuthor_SelectedIndexChanged1EventHandler cmbAuthor_SelectedIndexChanged1;
        public delegate void cmbAuthor_SelectedIndexChanged1EventHandler(object sender, System.EventArgs e);
        //public event cmbTypist_SelectedIndexChanged1EventHandler cmbTypist_SelectedIndexChanged1;
        public delegate void cmbTypist_SelectedIndexChanged1EventHandler(object sender, System.EventArgs e);
        //internal System.Windows.Forms.Label lblRe;
        //internal System.Windows.Forms.Label lblCC;
        // private System.Windows.Forms.CheckBox withEventsField_chkRe;
        public frmMainForm()
        {
            InitializeComponent();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cmbAuthor_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cmbTypist_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void butAuthDelete_Click(object sender, EventArgs e)
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);
                if (this.gAuthDT.Rows.Count == 0)
                    return;

                using (StreamWriter sw = new StreamWriter("D:\\Templates\\kfautlst.ini"))
                {
                    MsgBoxResult AutDel = default(MsgBoxResult);

                    AutDel = Interaction.MsgBox("Are you sure you want to delete " + this.cmbAuthor.Text + " from your Author list?", MsgBoxStyle.YesNo, "Delete Author?");

                    if (AutDel == MsgBoxResult.Yes)
                    {
                        DataRowView selecteditem = (DataRowView)this.cmbAuthor.SelectedItem;
                        int selectedindex = this.cmbAuthor.SelectedIndex;

                        this.gAuthDT.Rows.RemoveAt(selectedindex);
                        this.gAuthDT.AcceptChanges();
                        this.cmbAuthor.Refresh();

                        if (this.gAuthDT.Rows.Count == 0)
                        {
                            this.gAuthDT.Rows.Add("Click Find to add names", "0");
                            this.gAuthDT.AcceptChanges();
                        }
                        // DataRow row = default(DataRow);
                        foreach (DataRow row in gAuthDT.Rows)
                        {
                            sw.Write(Strings.Chr(34) + row["Author"].ToString() + Strings.Chr(34));
                            sw.Write(",");
                            sw.WriteLine(row["ID"]);
                            //sw.WriteLine(ListView1.SelectedItems(0).SubItems(1).Text)
                        }

                        sw.Flush();
                        //sw.Close();
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
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void butAuthFind_Click(object sender, EventArgs e)
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);
                if (Authorfrm == null)
                {
                    Authorfrm = new frmAuthorFind();
                }
                if (!Authorfrm.Visible)
                {
                    Authorfrm.TopMost = true;
                    Authorfrm.Show();
                }
                else
                {
                    Authorfrm.TopMost = true;
                    Authorfrm.BringToFront();
                }
                Authorfrm.hHandle = this.Handle;
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
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void butTypDelete_Click(object sender, EventArgs e)
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);
                if (gTypDt.Rows.Count == 0)
                    return;

                using (StreamWriter sw = new StreamWriter("D:\\Templates\\kftyplst.ini"))
                {
                    MsgBoxResult TypDel = default(MsgBoxResult);

                    TypDel = Interaction.MsgBox("Are you sure you want to delete " + this.cmbTypist.Text + " from your Typist list?", MsgBoxStyle.YesNo, "Delete Author?");

                    if (TypDel == MsgBoxResult.Yes)
                    {
                        DataRowView selecteditem = (DataRowView)this.cmbTypist.SelectedItem;
                        int selectedindex = this.cmbTypist.SelectedIndex;

                        gTypDt.Rows.RemoveAt(selectedindex);
                        gTypDt.AcceptChanges();
                        this.cmbTypist.Refresh();

                        if (gTypDt.Rows.Count == 0)
                        {
                            gTypDt.Rows.Add("Click Find to add names", "0");
                            gTypDt.AcceptChanges();
                        }

                        //DataRow row = default(DataRow);

                        foreach (DataRow row in gTypDt.Rows)
                        {
                            sw.Write(Strings.Chr(34) + row["Typist"].ToString() + Strings.Chr(34));
                            sw.Write(",");
                            sw.WriteLine(row["ID"]);
                        }

                        sw.Flush();
                        //sw.Close();
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
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void butTypFind_Click(object sender, EventArgs e)
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);
                if (TypistObjfrm == null)
                {
                    TypistObjfrm = new frmTypistFind();
                }
                if (!TypistObjfrm.Visible)
                {
                    TypistObjfrm.TopMost = true;
                    TypistObjfrm.Show();
                }
                else
                {
                    TypistObjfrm.TopMost = true;
                    TypistObjfrm.BringToFront();
                }
                
                TypistObjfrm.hHandle = this.Handle;
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
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmMainForm_Load(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "©2017 by Advanced Legal Systems, Inc. v" + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();
            // ssALS.Items[0].Text = "©2008 by Advanced Legal Systems, Inc. v" + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();
            BGColorD = System.Drawing.Color.DarkGray;
            BGColorE = System.Drawing.Color.White;
            StringBuilder sbTrace = new StringBuilder();
            try
            {
                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);
                appWord = default(Microsoft.Office.Interop.Word.Application);
                appWord = MyRibbonAddIn.Globals.ThisAddIn.Application;
                //_Document doc = appWord.ActiveDocument;
                SetPersAuthorList();
                SetPersTypistList();
                NestedLoaded();
            }
            catch (Exception ex)
            {
                sbTrace.Clear();
                sbTrace.AppendLine("Exception" + ex);
                Logger.SaveLoggerTrace(sbTrace);
                Logger.LogWriter(ex.StackTrace);
                Debug.Print("No Document Found");
                //No doc
                MessageBox.Show("Open Word Application not found; document assembly cannnot continue.Word not Found" + ex.Message.ToString());
                Close();
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
        public void NestedLoaded()
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);

                this.butCancel.Top = this.butOk.Top;
                this.butCancel.Left = this.butOk.Left + this.butOk.Width + 6;
                var _with1 = this;
                _with1.chkRe.Checked = true;
                _with1.chkDirectFax.Checked = false;
                _with1.chkDelivery.Checked = false;
                _with1.txtDelivery.Text = "";
                _with1.txtDelivery.BackColor = BGColorD;
                //&HE0E0E0
                _with1.lstDelivery.Enabled = false;
                _with1.lstDelivery.BackColor = BGColorD;
                //&HE0E0E0
                _with1.radAddr.Checked = true;
                if (txtRe.Enabled)
                {
                    txtRe.BackColor = BGColorE;
                }
                else
                {
                    txtRe.BackColor = BGColorD;
                }
                if (txtCC.Enabled)
                {
                    txtCC.BackColor = BGColorE;
                }
                else
                {
                    txtCC.BackColor = BGColorD;
                }
                if (txtBCC.Enabled)
                {
                    txtBCC.BackColor = BGColorE;
                }
                else
                {
                    txtBCC.BackColor = BGColorD;
                }
                var _with2 = lstDelivery;
                _with2.Items.Clear();
                _with2.Items.AddRange(new object[] {
            "Personal and Confidential",
            "Privileged and Confidential",
            "Via Certified Mail - Return Receipt Requested",
            "Via Certified Mail - Return Receipt Requested and First Class Mail",
            "Via Email",
            "Via Email and Regular Mail",
            "Via Facsimile and Regular Mail",
            "Via Facsimile",
            "Via Hand Delivery and Regular Mail",
            "Via Hand Delivery",
            "Via Overnight Delivery",
            "Via Telecopier"
        });
                string chkFile = null;

                // chkFile = FileSystem.Dir("D:\\Templates\\deflist.ini"); later we need to disscuss with client about deflist.ini file
                chkFile = FileSystem.Dir("D:\\Templates\\deflist.ini");
                if (chkFile == "deflist.ini")
                {
                    using (StreamReader sr = new StreamReader("D:\\Templates\\deflist.ini"))
                    {
                        var _with3 = sr;
                        string txtTest = null;
                        //System.Data.DataRow row = default(System.Data.DataRow);

                        txtTest = _with3.ReadLine();
                        foreach (System.Data.DataRow row in gAuthDT.Rows)
                        {
                            if (row["Author"].ToString() == txtTest)
                                cmbAuthor.Text = txtTest;
                        }

                        txtTest = _with3.ReadLine();
                        foreach (System.Data.DataRow row in gTypDt.Rows)
                        {
                            if (row["Typist"].ToString() == txtTest)
                                cmbTypist.Text = txtTest;
                        }

                        this.chkDelivery.Checked = Convert.ToBoolean(_with3.ReadLine());
                        this.txtDelivery.Text = _with3.ReadLine();
                        this.lstDelivery.SelectedIndex = Convert.ToInt32(_with3.ReadLine());
                        this.radAddrManually.Checked = Convert.ToBoolean(_with3.ReadLine());
                        this.radAddrMultiple.Checked = Convert.ToBoolean(_with3.ReadLine());
                        this.chkRe.Checked = Convert.ToBoolean(_with3.ReadLine());
                        this.chkCC.Checked = Convert.ToBoolean(_with3.ReadLine());
                        this.chkBCC.Checked = Convert.ToBoolean(_with3.ReadLine());
                        this.chkEnc.Checked = Convert.ToBoolean(_with3.ReadLine());
                        this.chkDirectFax.Checked = Convert.ToBoolean(_with3.ReadLine());
                        this.chkAdmittedTo.Checked = Convert.ToBoolean(_with3.ReadLine());
                        // _with3.Close();
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
        /// 
        /// </summary>
        internal void SetPersAuthorList()
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);
                cmbAuthor.Items.Clear();

                gAuthDT = GetAuthData();

                cmbAuthor.DataSource = gAuthDT;
                cmbAuthor.DisplayMember = "Author";
                cmbAuthor.ValueMember = "ID";
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
        internal void SetPersTypistList()
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);
                cmbTypist.Items.Clear();

                gTypDt = GetTypistData();

                cmbTypist.DataSource = gTypDt;
                cmbTypist.DisplayMember = "Typist";
                cmbTypist.ValueMember = "ID";
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
        /// <returns></returns>
        internal System.Data.DataTable GetAuthData()
        {
            StringBuilder sbTrace = new StringBuilder();
            System.Data.DataTable _AuthData = new System.Data.DataTable();
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);

                //AuthData = Nothing
                _AuthData.Columns.Add("Author");
                _AuthData.Columns.Add("ID");

                if (File.Exists("D:\\Templates\\kfautlst.ini"))
                {
                    using (Microsoft.VisualBasic.FileIO.TextFieldParser csvread = new Microsoft.VisualBasic.FileIO.TextFieldParser("D:\\Templates\\kfautlst.ini"))
                    {
                        csvread.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited;
                        csvread.SetDelimiters(",");

                        string[] strRow = null;
                        string[,] strAuthAr = new string[3, 2];

                        while (!csvread.EndOfData)
                        {
                            try
                            {
                                strRow = csvread.ReadFields();
                                _AuthData.Rows.Add(strRow);
                            }
                            catch (Microsoft.VisualBasic.FileIO.MalformedLineException ex)
                            {
                                MessageBox.Show("Error in c:\\kfauthlst.ini file at line " + ex.Message + ".");
                            }
                        }
                        //csvread.Close();
                    }
                }
                return _AuthData;
            }
            catch (Exception ex)
            {
                sbTrace.Clear();
                sbTrace.AppendLine("Exception" + ex);
                Logger.SaveLoggerTrace(sbTrace);
                Logger.LogWriter(ex.StackTrace);
                return _AuthData;
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
        /// <returns></returns>
        internal System.Data.DataTable GetTypistData()
        {
            System.Data.DataTable _TypData = new System.Data.DataTable();
            StringBuilder sbTrace = new StringBuilder();
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);
                _TypData.Columns.Add("Typist");
                _TypData.Columns.Add("ID");

                if (File.Exists("D:\\Templates\\kftyplst.ini"))
                {
                    using (Microsoft.VisualBasic.FileIO.TextFieldParser csvread = new Microsoft.VisualBasic.FileIO.TextFieldParser("D:\\Templates\\kftyplst.ini"))
                    {
                        csvread.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited;
                        csvread.SetDelimiters(",");

                        string[] strRow = null;
                        string[,] strTypAr = new string[3, 2];

                        while (!csvread.EndOfData)
                        {
                            try
                            {
                                strRow = csvread.ReadFields();
                                _TypData.Rows.Add(strRow);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Error in D:\\Templates\\kftyplst.ini file at line " + ex.Message + ".");
                            }
                        }
                        // csvread.Close();
                    }
                }
                return _TypData;
            }
            catch (Exception ex)
            {
                sbTrace.Clear();
                sbTrace.AppendLine("Exception" + ex);
                Logger.SaveLoggerTrace(sbTrace);
                Logger.LogWriter(ex.StackTrace);
                return _TypData;
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
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void butDefault_Click(object sender, EventArgs e)
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);

                using (StreamWriter sw = new StreamWriter("D:\\Templates\\deflist.ini"))
                {
                    var _with11 = sw;
                    _with11.WriteLine(cmbAuthor.Text);
                    _with11.WriteLine(cmbTypist.Text);
                    _with11.WriteLine(chkDelivery.Checked);
                    _with11.WriteLine(txtDelivery.Text);
                    _with11.WriteLine(lstDelivery.SelectedIndex);
                    _with11.WriteLine(radAddrManually.Checked);
                    _with11.WriteLine(radAddrMultiple.Checked);
                    _with11.WriteLine(chkRe.Checked);
                    _with11.WriteLine(chkCC.Checked);
                    _with11.WriteLine(chkBCC.Checked);
                    _with11.WriteLine(chkEnc.Checked);
                    _with11.WriteLine(chkDirectFax.Checked);
                    _with11.WriteLine(chkAdmittedTo.Checked);
                    // _with11.Close();
                }
                Interaction.MsgBox("Default settings saved.", MsgBoxStyle.Information, "Default settings saved.");
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
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void lstDelivery_SelectedIndexChanged(object sender, EventArgs e)
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);

                var _with5 = this;
                if (_with5.chkDelivery.Checked)
                {
                    _with5.txtDelivery.Text = _with5.lstDelivery.Text;
                }
                else
                {
                    _with5.txtDelivery.Text = "";
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
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void chkRe_CheckedChanged(object sender, EventArgs e)
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);
                txtRe.Enabled = chkRe.Checked;
                if (txtRe.Enabled)
                {
                    txtRe.BackColor = BGColorE;
                }
                else
                {
                    txtRe.BackColor = BGColorD;
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
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void chkCC_CheckedChanged(object sender, EventArgs e)
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);
                txtCC.Enabled = chkCC.Checked;
                if (txtCC.Enabled)
                {
                    txtCC.BackColor = BGColorE;
                }
                else
                {
                    txtCC.BackColor = BGColorD;
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
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void chkBCC_CheckedChanged(object sender, EventArgs e)
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);
                txtBCC.Enabled = chkBCC.Checked;
                if (txtBCC.Enabled)
                {
                    txtBCC.BackColor = BGColorE;
                }
                else
                {
                    txtBCC.BackColor = BGColorD;
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
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void chkEnc_CheckedChanged(object sender, EventArgs e)
        {

        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void chkDirectFax_CheckedChanged(object sender, EventArgs e)
        {

        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void chkAdmittedTo_CheckedChanged(object sender, EventArgs e)
        {

        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void butOk_Click(object sender, EventArgs e)
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {
                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);
                string checkaut = null;
                string checktyp = null;
                int autrecord = 0;
                int typrecord = 0;
                //Control obj = default(Control);

                checkaut = this.cmbAuthor.Text;
                checktyp = this.cmbTypist.Text;

                if (checkaut == "Click Find to add names")
                {
                    Interaction.MsgBox("You must select a valid Author! Click Find to select from the Master List", MsgBoxStyle.Critical, "No Author selected");
                    return;
                }
                if (checktyp == "Click Find to add names")
                {
                    Interaction.MsgBox("You must select a valid Typist! Click Find to select from the Master List", MsgBoxStyle.Critical, "No Typist selected");
                    return;
                }

                //MsgBox(cmbAuthor.Text & " " & cmbAuthor.SelectedValue.ToString)
                autrecord = Convert.ToInt32(this.cmbAuthor.SelectedValue);
                typrecord = Convert.ToInt32(this.cmbTypist.SelectedValue);

                var _with6 = appWord;
                _with6.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekCurrentPageHeader;
                _with6.ActiveWindow.ActivePane.View.NextHeaderFooter();
                _with6.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdParagraph, 2);
                _with6.Selection.InsertDateTime("MMMM d, yyyy", true);
                _with6.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekMainDocument;
                foreach (Control obj in this.Controls)
                {
                    if (obj.Enabled)
                    {
                        if (obj is TextBox)
                        {
                            Debug.Print(obj.Name);
                            if (!string.IsNullOrEmpty(obj.Text))
                            {
                                _with6.ActiveDocument.Bookmarks[obj.Name].Range.Text = obj.Text;
                            }
                        }
                        else if (obj is System.Windows.Forms.CheckBox)
                        {
                            //Dim chk As CheckBox
                            //chk = obj
                            //If Not chk.Checked Then
                            //    .Selection.GoTo(wdGoToBookmark, , , chk.Name.Substring(3))
                            //    MsgBox("TEST: " & chk.Name.Substring(3))
                            //    .Selection.Delete()
                            //End If
                        }
                    }
                }
                object Which = "", Count = "", What = -1, Name = "";
                if (chkRe.Checked)
                {
                    //do nothing
                }
                else
                {
                    Name = "Re";
                    //_with6.Selection.GoTo(ref What, ref Which, ref Count, ref Name);
                    _with6.Selection.Delete();
                }

                if (chkCC.Checked)
                {
                    //do nothing
                }
                else
                {
                    Name = "CC";
                    _with6.Selection.GoTo(ref What, ref Which, ref Count, ref Name);
                    _with6.Selection.Delete();
                }

                if (chkBCC.Checked)
                {
                    //do nothing
                }
                else
                {
                    Name = "BCC";
                    //_with6.Selection.GoTo(ref What, ref Which, ref Count, ref Name);
                    _with6.Selection.Delete();
                }

                if (chkEnc.Checked)
                {
                    //Do Nothing
                }
                else
                {
                    Name = "Enc";
                    //_with6.Selection.GoTo(ref What, ref Which, ref Count, ref Name);
                    _with6.Selection.Delete();
                }

                if (chkDelivery.Checked)
                {
                    //Do Nothing
                }
                else
                {
                    Name = "Delivery";
                    //_with6.Selection.GoTo(ref What, ref Which, ref Count, ref Name);
                    _with6.Selection.Delete();
                }

                //Get Author/Typist Info
                ADODB.Connection cn = default(ADODB.Connection);
                ADODB.Recordset rs = new ADODB.Recordset();
                ALSFunctions als = new ALSFunctions();
                string addType = "";
                string strTitle = null;

                cn = als.oConn();
                var _with7 = rs;
                _with7.Open("SELECT * FROM tblAuthor WHERE recordCounter = " + autrecord, cn);
                strTitle = _with7.Fields["Title"].Value.ToString();
                strTitle = Strings.Trim(strTitle);

                if (Information.IsDBNull(_with7.Fields["Closing"].Value))
                {
                    appWord.ActiveDocument.Bookmarks["Closing"].Range.Text = "";
                }
                else
                {
                    appWord.ActiveDocument.Bookmarks["Closing"].Range.Text = _with7.Fields["Closing"].Value.ToString();
                }

                if (Information.IsDBNull(_with7.Fields["ClosingName"].Value))
                {
                    appWord.ActiveDocument.Bookmarks["ClosingName2"].Range.Text = "";
                    if (appWord.ActiveDocument.Bookmarks.Exists("Name"))
                    {
                        appWord.ActiveDocument.Bookmarks["Name"].Range.Text = "";
                    }
                }
                else
                {
                    appWord.ActiveDocument.Bookmarks["ClosingName2"].Range.Text = _with7.Fields["ClosingName"].Value.ToString();
                    if (inHeader(strTitle))
                    {
                        if (appWord.ActiveDocument.Bookmarks.Exists("Name"))
                        {
                            appWord.ActiveDocument.Bookmarks["Name"].Range.Text = _with7.Fields["ClosingName"].Value.ToString();
                        }
                    }
                }

                if (Information.IsDBNull(_with7.Fields["Initials"].Value))
                {
                    appWord.ActiveDocument.Bookmarks["AutInitials"].Range.Text = "";
                }
                else
                {
                    appWord.ActiveDocument.Bookmarks["AutInitials"].Range.Text = _with7.Fields["Initials"].Value.ToString();
                }

                if (Information.IsDBNull(_with7.Fields["Title"].Value))
                {
                    if (appWord.ActiveDocument.Bookmarks.Exists("Title"))
                    {
                        appWord.ActiveDocument.Bookmarks["Title"].Range.Text = "";
                    }
                    appWord.Selection.GoTo(ref What, ref Which, ref Count, "Title2");
                    appWord.Selection.Delete();
                }
                else
                {
                    if (inHeader(strTitle))
                    {
                        if (appWord.ActiveDocument.Bookmarks.Exists("Title"))
                        {
                            appWord.ActiveDocument.Bookmarks["Title"].Range.Text = strTitle + Strings.Chr(13) + _with7.Fields["admitted"].Value;
                        }
                        if (Strings.InStr(strTitle, "Attorney") == 0)
                        {
                            appWord.ActiveDocument.Bookmarks["Title1"].Range.Text = strTitle;
                        }
                        else
                        {
                            appWord.Selection.GoTo(WdGoToItem.wdGoToBookmark, ref Which, ref Count, "Title2");
                            appWord.Selection.Delete();
                        }
                        if (appWord.ActiveDocument.Bookmarks.Exists("EMail"))
                        {
                            appWord.Selection.GoTo(WdGoToItem.wdGoToBookmark, ref Which, ref Count, "EMail");
                            appWord.Selection.Delete();
                        }
                    }
                    else
                    {
                        if (appWord.ActiveDocument.Bookmarks.Exists("Title1"))
                        {
                            appWord.ActiveDocument.Bookmarks["Title1"].Range.Text = strTitle;
                        }

                        if (appWord.ActiveDocument.Bookmarks.Exists("Email1"))
                        {
                            appWord.ActiveDocument.Bookmarks["Email1"].Range.Text = _with7.Fields["HomeNo"].Value.ToString();
                        }
                    }
                }

                if (Information.IsDBNull(_with7.Fields["HomeNo"].Value) | !inHeader(strTitle))
                {
                    if (appWord.ActiveDocument.Bookmarks.Exists("HomeNo"))
                    {
                        appWord.ActiveDocument.Bookmarks["HomeNo"].Range.Text = "";
                    }
                }
                else
                {
                    if (appWord.ActiveDocument.Bookmarks.Exists("HomeNo"))
                    {
                        appWord.ActiveDocument.Bookmarks["HomeNo"].Range.Text = _with7.Fields["HomeNo"].Value.ToString();
                    }
                }
                rs.Close();
                cn.Close();

                cn = als.oConn();
                var _with8 = rs;
                _with8.Open("SELECT * FROM tblAuthor WHERE recordCounter = " + typrecord, cn);
                _with8.MoveFirst();

                if (Information.IsDBNull(rs.Fields["Initials"].Value))
                {
                    appWord.ActiveDocument.Bookmarks["TypInitials"].Range.Text = "";
                }
                else
                {
                    appWord.ActiveDocument.Bookmarks["TypInitials"].Range.Text = _with8.Fields["Initials"].Value.ToString();
                }
                rs.Close();
                cn.Close();

                if (radAddrMultiple.Checked)
                {
                    addType = "M";
                }

                if (radAddrManually.Checked)
                {
                    addType = "N";
                }

                if (radAddr.Checked)
                {
                    addType = "";
                }

                _with6.ActiveDocument.Bookmarks["address"].Range.Text = GetAddress(addType);
                // _with6.Selection.GoTo(WdGoToItem.wdGoToBookmark, ref Which, ref Count, "Address");

                var _with9 = _with6.ActiveDocument.Bookmarks;
                _with9.DefaultSorting = WdBookmarkSortBy.wdSortByName;
                _with9.ShowHidden = false;
                _with6.Selection.EndKey(WdUnits.wdLine, WdMovementType.wdExtend);
                _with6.Selection.MoveLeft(WdUnits.wdCharacter, 1, WdMovementType.wdExtend);
                var _with10 = _with6.ActiveDocument.Bookmarks;
                _with10.Add("FirstAddress", appWord.Selection.Range);
                _with10.DefaultSorting = WdBookmarkSortBy.wdSortByName;
                _with10.ShowHidden = false;
                _with6.Selection.HomeKey(WdUnits.wdLine);

                _with6.Selection.GoTo(ref What, ref Which, ref Count, "Text");
                _with6.Selection.NextField().Select();

                als.UpdateAll();
                this.Close();
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
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void butCancel_Click(object sender, EventArgs e)
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);

                //appWord.ActiveDocument.Close();
                //Tools.Instance.NewDocument();
                this.Hide();
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
        /// <param name="Type"></param>
        /// <returns></returns>
        private string GetAddress(string Type)
        {
            string functionReturnValue = null;
            string letaddr = null;
            string finaddr = "";
            bool Again = true;
            MsgBoxResult Answer = default(MsgBoxResult);
            StringBuilder sbTrace = new StringBuilder();
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);
                if (Type == "N")
                {
                    functionReturnValue = "";
                    return functionReturnValue;
                }

                while (Again == true)
                {
                    letaddr = appWord.GetAddress(Name: "", DisplaySelectDialog: 1);
                    //letaddr = appWord.GetAddress(, , True, True, , , , )

                    if (Type == "M")
                    {
                        finaddr = finaddr + Strings.Chr(13) + Strings.Chr(13) + letaddr;
                        Answer = Interaction.MsgBox("Select another address?", MsgBoxStyle.YesNo, "Multiple Addresses");
                        switch (Answer)
                        {
                            case MsgBoxResult.Yes:
                                Again = true;
                                break;
                            case MsgBoxResult.No:
                                Again = false;
                                break;
                        }
                    }
                    else if (Type == "N")
                    {
                        finaddr = "";
                        Again = false;
                    }
                    else
                    {
                        finaddr = letaddr;
                        Again = false;
                    }
                }

                functionReturnValue = finaddr;
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
            //.Trim(Chr(13))
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="Title"></param>
        /// <returns></returns>
        private bool inHeader(string Title)
        {
            bool functionReturnValue = false;
            string strTitleLst = null;
            int a = 0;

            Title = Strings.Trim(Title);
            functionReturnValue = false;
            StringBuilder sbTrace = new StringBuilder();
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);
                for (a = 0; a <= Information.UBound(strTitleList); a++)
                {
                    strTitleLst = Strings.Trim(strTitleList[a]);
                    if (Strings.InStr(1, Title, strTitleLst, CompareMethod.Text) > 0)
                    {
                        functionReturnValue = true;
                        return functionReturnValue;
                    }
                }
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
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="Type"></param>
        private void RemoveDefault(string Type)
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);

                string chkFile = null;
                string[] item = new string[14];
                int a = 0;
                chkFile = FileSystem.Dir("D:\\Templates\\deflist.ini");

                if (chkFile == "deflist.ini")
                {
                    using (StreamReader sr = new StreamReader("D:\\Templates\\deflist.ini"))
                    {
                        var _with12 = sr;
                        for (a = 1; a <= 13; a++)
                        {
                            item[a] = _with12.ReadLine();
                        }
                        //_with12.Close();
                    }
                }

                switch (Type)
                {
                    case "T":
                        item[2] = " ";
                        break;
                    case "A":
                        item[1] = " ";
                        break;
                }
                using (StreamWriter sw = new StreamWriter("D:\\Templates\\deflist.ini"))
                {
                    var _with13 = sw;
                    for (a = 1; a <= 13; a++)
                    {
                        _with13.WriteLine(item[a]);
                    }
                    //_with13.Close();
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
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void chkDelivery_CheckedChanged(object sender, EventArgs e)
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);
                var _with4 = this;
                if (!chkDelivery.Checked)
                {
                    _with4.txtDelivery.Text = "";
                    _with4.txtDelivery.Enabled = false;
                    _with4.txtDelivery.BackColor = BGColorD;
                    _with4.lstDelivery.Enabled = false;
                    _with4.lstDelivery.BackColor = BGColorD;
                }
                else
                {
                    _with4.txtDelivery.Text = _with4.lstDelivery.Text;
                    _with4.txtDelivery.Enabled = true;
                    _with4.txtDelivery.BackColor = BGColorE;
                    _with4.lstDelivery.Enabled = true;
                    _with4.lstDelivery.BackColor = BGColorE;
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
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmMainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);
                //appWord.ActiveDocument.Close();
                //Tools.Instance.NewDocument();
                this.Hide();
                this.Parent = null;
                e.Cancel = true;
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
