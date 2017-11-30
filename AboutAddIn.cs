using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MyRibbonAddIn
{
    /// <summary>
    /// 
    /// </summary>
    public partial class AboutAddIn : Form
    {
        /// <summary>
        /// 
        /// </summary>
        public AboutAddIn()
        {
            InitializeComponent();
        }
        private static AboutAddIn _instance;
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public static AboutAddIn GetInstance()
        {
            if (_instance == null) _instance = new AboutAddIn();
            return _instance;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);
                //this.Close();
                AboutAddIn form = AboutAddIn.GetInstance();
                form.Hide();
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
        private void label5_Click(object sender, EventArgs e)
        {

        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void label3_Click(object sender, EventArgs e)
        {

        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AboutAddIn_FormClosing(object sender, FormClosingEventArgs e)
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {

                sbTrace.AppendLine("Start");
                
                this.Hide();
                this.Parent = null;
                e.Cancel = true;
            }
            catch (Exception ex)
            {
                
                Logger.LogWriter(ex.StackTrace);
            }
            finally
            {
                sbTrace.AppendLine("AboutAddIn_FormClosing Complete");
                Logger.SaveLoggerTrace(sbTrace);
                sbTrace.Clear();
            }
        }
    }
}
