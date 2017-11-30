using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
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
    public partial class frmAuthorFind : Form
    {
        /// <summary>
        /// 
        /// </summary>
        public frmAuthorFind()
        {
            InitializeComponent();
        }
        /// <summary>
        /// 
        /// </summary>
        public IntPtr hHandle;
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmAuthorFind_Load(object sender, System.EventArgs e)
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {
                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);
                ADODB.Recordset rs = new ADODB.Recordset();
                ALSFunctions als = new ALSFunctions();
                ADODB.Connection cn = default(ADODB.Connection);
                ColumnHeader lvwColumn = default(ColumnHeader);
                ListViewItem itmListItem = default(ListViewItem);

                cn = als.oConn();
                rs.Open("qryAuthor", cn);

                ListView1.Clear();

                lvwColumn = new ColumnHeader();
                lvwColumn.Text = "Name";
                ListView1.Columns.Add(lvwColumn);

                while (!rs.EOF)
                {
                    itmListItem = new ListViewItem();
                    if (Information.IsDBNull(rs.Fields[1].Value))
                    {
                        itmListItem.Text = "";
                    }
                    else
                    {
                    }
                    itmListItem.Text = rs.Fields[1].Value.ToString();

                    if (Information.IsDBNull(rs.Fields[0].Value))
                    {
                        itmListItem.Tag = "";
                        //itmListItem.SubItems.Add("")
                    }
                    else
                    {
                        itmListItem.Tag = rs.Fields[0].Value;
                        //itmListItem.SubItems.Add(rs.Fields(0).Value)
                    }

                    ListView1.Items.Add(itmListItem);
                    //ListView1.Columns.Item[0].AutoResize(ColumnHeaderAutoResizeStyle.ColumnContent);
                    rs.MoveNext();
                }
                rs.Close();
                cn.Close();
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
        private void butAdd_Click(System.Object sender, System.EventArgs e)
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {
                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);
                using (StreamWriter sw = File.AppendText("D:\\Templates\\kfautlst.ini"))
                {
                    sw.Write(Strings.Chr(34) + ListView1.SelectedItems[0].Text + Strings.Chr(34));
                    sw.Write(",");
                    sw.WriteLine(ListView1.SelectedItems[0].Tag.ToString());
                    //sw.WriteLine(ListView1.SelectedItems(0).SubItems(1).Text)
                    //sw.Close();
                }


                frmMainForm frm = default(frmMainForm);
                frm = (frmMainForm)frmMainForm.FromHandle(hHandle);

                string[] strRow = null;
                strRow = Strings.Split(ListView1.SelectedItems[0].Text + "|" + ListView1.SelectedItems[0].Tag.ToString(), "|");
                // dest.Columns.Add("NoOfNighs", typeof(String));

                frm.gAuthDT.Rows.Add(strRow);
                //frm.gAuthDT.LoadDataRow(strRow, true);
                frm.gAuthDT.AcceptChanges();

                //frm.BackColor = Drawing.Color.Aqua
                Interaction.MsgBox(ListView1.SelectedItems[0].Text + " added to you Author list.");
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
        private void butClose_Click(System.Object sender, System.EventArgs e)
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {
                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);
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

        private void frmAuthorFind_FormClosing(object sender, FormClosingEventArgs e)
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);
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
