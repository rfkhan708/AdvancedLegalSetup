namespace MyRibbonAddIn.ALS_FWW_Word
{
    partial class frmMainForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.grpTypist = new System.Windows.Forms.GroupBox();
            this.butTypFind = new System.Windows.Forms.Button();
            this.butTypDelete = new System.Windows.Forms.Button();
            this.cmbTypist = new System.Windows.Forms.ComboBox();
            this.grpAuthor = new System.Windows.Forms.GroupBox();
            this.butAuthFind = new System.Windows.Forms.Button();
            this.butAuthDelete = new System.Windows.Forms.Button();
            this.cmbAuthor = new System.Windows.Forms.ComboBox();
            this.ssALS = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.butCancel = new System.Windows.Forms.Button();
            this.butOk = new System.Windows.Forms.Button();
            this.radAddrManually = new System.Windows.Forms.RadioButton();
            this.butDefault = new System.Windows.Forms.Button();
            this.radAddrMultiple = new System.Windows.Forms.RadioButton();
            this.radAddr = new System.Windows.Forms.RadioButton();
            this.chkAdmittedTo = new System.Windows.Forms.CheckBox();
            this.chkDirectFax = new System.Windows.Forms.CheckBox();
            this.chkDelivery = new System.Windows.Forms.CheckBox();
            this.chkBCC = new System.Windows.Forms.CheckBox();
            this.chkCC = new System.Windows.Forms.CheckBox();
            this.txtCC = new System.Windows.Forms.TextBox();
            this.GrpAddress = new System.Windows.Forms.GroupBox();
            this.chkRe = new System.Windows.Forms.CheckBox();
            this.txtBCC = new System.Windows.Forms.TextBox();
            this.txtDelivery = new System.Windows.Forms.TextBox();
            this.lblBCC = new System.Windows.Forms.Label();
            this.grpIncludes = new System.Windows.Forms.GroupBox();
            this.chkEnc = new System.Windows.Forms.CheckBox();
            this.lblCC = new System.Windows.Forms.Label();
            this.txtRe = new System.Windows.Forms.TextBox();
            this.lblRe = new System.Windows.Forms.Label();
            this.lstDelivery = new System.Windows.Forms.ListBox();
            this.grpTypist.SuspendLayout();
            this.grpAuthor.SuspendLayout();
            this.ssALS.SuspendLayout();
            this.GrpAddress.SuspendLayout();
            this.grpIncludes.SuspendLayout();
            this.SuspendLayout();
            // 
            // grpTypist
            // 
            this.grpTypist.Controls.Add(this.butTypFind);
            this.grpTypist.Controls.Add(this.butTypDelete);
            this.grpTypist.Controls.Add(this.cmbTypist);
            this.grpTypist.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grpTypist.Location = new System.Drawing.Point(232, 3);
            this.grpTypist.Name = "grpTypist";
            this.grpTypist.Size = new System.Drawing.Size(214, 77);
            this.grpTypist.TabIndex = 33;
            this.grpTypist.TabStop = false;
            this.grpTypist.Text = "Typist Name";
            // 
            // butTypFind
            // 
            this.butTypFind.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.butTypFind.Location = new System.Drawing.Point(133, 46);
            this.butTypFind.Name = "butTypFind";
            this.butTypFind.Size = new System.Drawing.Size(75, 23);
            this.butTypFind.TabIndex = 2;
            this.butTypFind.Text = "&Find";
            this.butTypFind.UseVisualStyleBackColor = true;
            this.butTypFind.Click += new System.EventHandler(this.butTypFind_Click);
            // 
            // butTypDelete
            // 
            this.butTypDelete.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.butTypDelete.Location = new System.Drawing.Point(52, 46);
            this.butTypDelete.Name = "butTypDelete";
            this.butTypDelete.Size = new System.Drawing.Size(75, 23);
            this.butTypDelete.TabIndex = 1;
            this.butTypDelete.Text = "Dele&te";
            this.butTypDelete.UseVisualStyleBackColor = true;
            this.butTypDelete.Click += new System.EventHandler(this.butTypDelete_Click);
            // 
            // cmbTypist
            // 
            this.cmbTypist.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbTypist.FormattingEnabled = true;
            this.cmbTypist.Location = new System.Drawing.Point(6, 19);
            this.cmbTypist.Name = "cmbTypist";
            this.cmbTypist.Size = new System.Drawing.Size(202, 21);
            this.cmbTypist.TabIndex = 0;
            this.cmbTypist.SelectedIndexChanged += new System.EventHandler(this.cmbTypist_SelectedIndexChanged);
            // 
            // grpAuthor
            // 
            this.grpAuthor.Controls.Add(this.butAuthFind);
            this.grpAuthor.Controls.Add(this.butAuthDelete);
            this.grpAuthor.Controls.Add(this.cmbAuthor);
            this.grpAuthor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grpAuthor.Location = new System.Drawing.Point(12, 3);
            this.grpAuthor.Name = "grpAuthor";
            this.grpAuthor.Size = new System.Drawing.Size(214, 77);
            this.grpAuthor.TabIndex = 32;
            this.grpAuthor.TabStop = false;
            this.grpAuthor.Text = "Author Name";
            // 
            // butAuthFind
            // 
            this.butAuthFind.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.butAuthFind.Location = new System.Drawing.Point(131, 46);
            this.butAuthFind.Name = "butAuthFind";
            this.butAuthFind.Size = new System.Drawing.Size(75, 23);
            this.butAuthFind.TabIndex = 2;
            this.butAuthFind.Text = "&Find";
            this.butAuthFind.UseVisualStyleBackColor = true;
            this.butAuthFind.Click += new System.EventHandler(this.butAuthFind_Click);
            // 
            // butAuthDelete
            // 
            this.butAuthDelete.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.butAuthDelete.Location = new System.Drawing.Point(50, 46);
            this.butAuthDelete.Name = "butAuthDelete";
            this.butAuthDelete.Size = new System.Drawing.Size(75, 23);
            this.butAuthDelete.TabIndex = 1;
            this.butAuthDelete.Text = "Dele&te";
            this.butAuthDelete.UseVisualStyleBackColor = true;
            this.butAuthDelete.Click += new System.EventHandler(this.butAuthDelete_Click);
            // 
            // cmbAuthor
            // 
            this.cmbAuthor.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbAuthor.FormattingEnabled = true;
            this.cmbAuthor.Location = new System.Drawing.Point(6, 19);
            this.cmbAuthor.Name = "cmbAuthor";
            this.cmbAuthor.Size = new System.Drawing.Size(199, 21);
            this.cmbAuthor.TabIndex = 0;
            this.cmbAuthor.SelectedIndexChanged += new System.EventHandler(this.cmbAuthor_SelectedIndexChanged);
            // 
            // ssALS
            // 
            this.ssALS.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1});
            this.ssALS.Location = new System.Drawing.Point(0, 526);
            this.ssALS.Name = "ssALS";
            this.ssALS.Size = new System.Drawing.Size(458, 22);
            this.ssALS.TabIndex = 35;
            this.ssALS.Text = "toolStripStatusLabel1";
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(225, 17);
            this.toolStripStatusLabel1.Text = "©2017 by Advanced Legal Systems, Inc. v";
            // 
            // butCancel
            // 
            this.butCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.butCancel.Location = new System.Drawing.Point(371, 500);
            this.butCancel.Name = "butCancel";
            this.butCancel.Size = new System.Drawing.Size(75, 23);
            this.butCancel.TabIndex = 34;
            this.butCancel.Text = "Cancel";
            this.butCancel.UseVisualStyleBackColor = true;
            this.butCancel.Click += new System.EventHandler(this.butCancel_Click);
            // 
            // butOk
            // 
            this.butOk.Location = new System.Drawing.Point(284, 500);
            this.butOk.Name = "butOk";
            this.butOk.Size = new System.Drawing.Size(75, 23);
            this.butOk.TabIndex = 40;
            this.butOk.Text = "&OK";
            this.butOk.UseVisualStyleBackColor = true;
            this.butOk.Click += new System.EventHandler(this.butOk_Click);
            // 
            // radAddrManually
            // 
            this.radAddrManually.AutoSize = true;
            this.radAddrManually.Location = new System.Drawing.Point(6, 65);
            this.radAddrManually.Name = "radAddrManually";
            this.radAddrManually.Size = new System.Drawing.Size(124, 17);
            this.radAddrManually.TabIndex = 2;
            this.radAddrManually.TabStop = true;
            this.radAddrManually.Text = "Add&ress Manually";
            this.radAddrManually.UseVisualStyleBackColor = true;
            // 
            // butDefault
            // 
            this.butDefault.Location = new System.Drawing.Point(364, 94);
            this.butDefault.Name = "butDefault";
            this.butDefault.Size = new System.Drawing.Size(75, 23);
            this.butDefault.TabIndex = 41;
            this.butDefault.Text = "&Set Default";
            this.butDefault.UseVisualStyleBackColor = true;
            this.butDefault.Click += new System.EventHandler(this.butDefault_Click);
            // 
            // radAddrMultiple
            // 
            this.radAddrMultiple.AutoSize = true;
            this.radAddrMultiple.Location = new System.Drawing.Point(6, 42);
            this.radAddrMultiple.Name = "radAddrMultiple";
            this.radAddrMultiple.Size = new System.Drawing.Size(131, 17);
            this.radAddrMultiple.TabIndex = 1;
            this.radAddrMultiple.TabStop = true;
            this.radAddrMultiple.Text = "&Multiple Addresses";
            this.radAddrMultiple.UseVisualStyleBackColor = true;
            // 
            // radAddr
            // 
            this.radAddr.AutoSize = true;
            this.radAddr.Location = new System.Drawing.Point(6, 19);
            this.radAddr.Name = "radAddr";
            this.radAddr.Size = new System.Drawing.Size(109, 17);
            this.radAddr.TabIndex = 0;
            this.radAddr.TabStop = true;
            this.radAddr.Text = "Sin&gle Address";
            this.radAddr.UseVisualStyleBackColor = true;
            // 
            // chkAdmittedTo
            // 
            this.chkAdmittedTo.AutoSize = true;
            this.chkAdmittedTo.Location = new System.Drawing.Point(6, 134);
            this.chkAdmittedTo.Name = "chkAdmittedTo";
            this.chkAdmittedTo.Size = new System.Drawing.Size(166, 17);
            this.chkAdmittedTo.TabIndex = 5;
            this.chkAdmittedTo.Text = "A&uthor\'s Admitted to Info";
            this.chkAdmittedTo.UseVisualStyleBackColor = true;
            this.chkAdmittedTo.CheckedChanged += new System.EventHandler(this.chkAdmittedTo_CheckedChanged);
            // 
            // chkDirectFax
            // 
            this.chkDirectFax.AutoSize = true;
            this.chkDirectFax.Location = new System.Drawing.Point(6, 111);
            this.chkDirectFax.Name = "chkDirectFax";
            this.chkDirectFax.Size = new System.Drawing.Size(181, 17);
            this.chkDirectFax.TabIndex = 4;
            this.chkDirectFax.Text = "&Author\'s Direct Fax Number";
            this.chkDirectFax.UseVisualStyleBackColor = true;
            this.chkDirectFax.CheckedChanged += new System.EventHandler(this.chkDirectFax_CheckedChanged);
            // 
            // chkDelivery
            // 
            this.chkDelivery.AutoSize = true;
            this.chkDelivery.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkDelivery.Location = new System.Drawing.Point(17, 83);
            this.chkDelivery.Name = "chkDelivery";
            this.chkDelivery.Size = new System.Drawing.Size(142, 17);
            this.chkDelivery.TabIndex = 42;
            this.chkDelivery.Text = "&Delivery Instructions";
            this.chkDelivery.UseVisualStyleBackColor = true;
            this.chkDelivery.CheckedChanged += new System.EventHandler(this.chkDelivery_CheckedChanged);
            // 
            // chkBCC
            // 
            this.chkBCC.AutoSize = true;
            this.chkBCC.Location = new System.Drawing.Point(6, 65);
            this.chkBCC.Name = "chkBCC";
            this.chkBCC.Size = new System.Drawing.Size(140, 17);
            this.chkBCC.TabIndex = 2;
            this.chkBCC.Text = "&Blind Carbon Copies";
            this.chkBCC.UseVisualStyleBackColor = true;
            this.chkBCC.CheckedChanged += new System.EventHandler(this.chkBCC_CheckedChanged);
            // 
            // chkCC
            // 
            this.chkCC.AutoSize = true;
            this.chkCC.Location = new System.Drawing.Point(6, 42);
            this.chkCC.Name = "chkCC";
            this.chkCC.Size = new System.Drawing.Size(108, 17);
            this.chkCC.TabIndex = 1;
            this.chkCC.Text = "&Carbon Copies";
            this.chkCC.UseVisualStyleBackColor = true;
            this.chkCC.CheckedChanged += new System.EventHandler(this.chkCC_CheckedChanged);
            // 
            // txtCC
            // 
            this.txtCC.Enabled = false;
            this.txtCC.Location = new System.Drawing.Point(253, 447);
            this.txtCC.Multiline = true;
            this.txtCC.Name = "txtCC";
            this.txtCC.Size = new System.Drawing.Size(205, 47);
            this.txtCC.TabIndex = 50;
            // 
            // GrpAddress
            // 
            this.GrpAddress.Controls.Add(this.radAddrManually);
            this.GrpAddress.Controls.Add(this.radAddrMultiple);
            this.GrpAddress.Controls.Add(this.radAddr);
            this.GrpAddress.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.GrpAddress.Location = new System.Drawing.Point(241, 288);
            this.GrpAddress.Name = "GrpAddress";
            this.GrpAddress.Size = new System.Drawing.Size(200, 86);
            this.GrpAddress.TabIndex = 46;
            this.GrpAddress.TabStop = false;
            this.GrpAddress.Text = "Addressing:";
            // 
            // chkRe
            // 
            this.chkRe.AutoSize = true;
            this.chkRe.Location = new System.Drawing.Point(6, 19);
            this.chkRe.Name = "chkRe";
            this.chkRe.Size = new System.Drawing.Size(70, 17);
            this.chkRe.TabIndex = 0;
            this.chkRe.Text = "&Re: line";
            this.chkRe.UseVisualStyleBackColor = true;
            this.chkRe.CheckedChanged += new System.EventHandler(this.chkRe_CheckedChanged);
            // 
            // txtBCC
            // 
            this.txtBCC.AcceptsReturn = true;
            this.txtBCC.Enabled = false;
            this.txtBCC.Location = new System.Drawing.Point(21, 447);
            this.txtBCC.Multiline = true;
            this.txtBCC.Name = "txtBCC";
            this.txtBCC.Size = new System.Drawing.Size(205, 47);
            this.txtBCC.TabIndex = 52;
            // 
            // txtDelivery
            // 
            this.txtDelivery.Location = new System.Drawing.Point(17, 102);
            this.txtDelivery.Name = "txtDelivery";
            this.txtDelivery.Size = new System.Drawing.Size(208, 20);
            this.txtDelivery.TabIndex = 43;
            // 
            // lblBCC
            // 
            this.lblBCC.AutoSize = true;
            this.lblBCC.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblBCC.Location = new System.Drawing.Point(18, 431);
            this.lblBCC.Name = "lblBCC";
            this.lblBCC.Size = new System.Drawing.Size(31, 13);
            this.lblBCC.TabIndex = 51;
            this.lblBCC.Text = "BCC";
            // 
            // grpIncludes
            // 
            this.grpIncludes.Controls.Add(this.chkAdmittedTo);
            this.grpIncludes.Controls.Add(this.chkDirectFax);
            this.grpIncludes.Controls.Add(this.chkEnc);
            this.grpIncludes.Controls.Add(this.chkBCC);
            this.grpIncludes.Controls.Add(this.chkCC);
            this.grpIncludes.Controls.Add(this.chkRe);
            this.grpIncludes.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grpIncludes.Location = new System.Drawing.Point(241, 123);
            this.grpIncludes.Name = "grpIncludes";
            this.grpIncludes.Size = new System.Drawing.Size(200, 157);
            this.grpIncludes.TabIndex = 45;
            this.grpIncludes.TabStop = false;
            this.grpIncludes.Text = "Will you be including?";
            // 
            // chkEnc
            // 
            this.chkEnc.AutoSize = true;
            this.chkEnc.Location = new System.Drawing.Point(6, 88);
            this.chkEnc.Name = "chkEnc";
            this.chkEnc.Size = new System.Drawing.Size(88, 17);
            this.chkEnc.TabIndex = 3;
            this.chkEnc.Text = "E&nclosures";
            this.chkEnc.UseVisualStyleBackColor = true;
            this.chkEnc.CheckedChanged += new System.EventHandler(this.chkEnc_CheckedChanged);
            // 
            // lblCC
            // 
            this.lblCC.AutoSize = true;
            this.lblCC.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblCC.Location = new System.Drawing.Point(250, 431);
            this.lblCC.Name = "lblCC";
            this.lblCC.Size = new System.Drawing.Size(27, 13);
            this.lblCC.TabIndex = 49;
            this.lblCC.Text = "CC:";
            // 
            // txtRe
            // 
            this.txtRe.AcceptsReturn = true;
            this.txtRe.Enabled = false;
            this.txtRe.Location = new System.Drawing.Point(20, 385);
            this.txtRe.Multiline = true;
            this.txtRe.Name = "txtRe";
            this.txtRe.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtRe.Size = new System.Drawing.Size(419, 43);
            this.txtRe.TabIndex = 48;
            // 
            // lblRe
            // 
            this.lblRe.AutoSize = true;
            this.lblRe.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblRe.Location = new System.Drawing.Point(18, 369);
            this.lblRe.Name = "lblRe";
            this.lblRe.Size = new System.Drawing.Size(27, 13);
            this.lblRe.TabIndex = 47;
            this.lblRe.Text = "Re:";
            // 
            // lstDelivery
            // 
            this.lstDelivery.FormattingEnabled = true;
            this.lstDelivery.Location = new System.Drawing.Point(37, 132);
            this.lstDelivery.Name = "lstDelivery";
            this.lstDelivery.Size = new System.Drawing.Size(188, 225);
            this.lstDelivery.TabIndex = 44;
            this.lstDelivery.SelectedIndexChanged += new System.EventHandler(this.lstDelivery_SelectedIndexChanged);
            // 
            // frmMainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(458, 548);
            this.Controls.Add(this.butOk);
            this.Controls.Add(this.butDefault);
            this.Controls.Add(this.chkDelivery);
            this.Controls.Add(this.txtCC);
            this.Controls.Add(this.GrpAddress);
            this.Controls.Add(this.txtBCC);
            this.Controls.Add(this.txtDelivery);
            this.Controls.Add(this.lblBCC);
            this.Controls.Add(this.grpIncludes);
            this.Controls.Add(this.lblCC);
            this.Controls.Add(this.txtRe);
            this.Controls.Add(this.lblRe);
            this.Controls.Add(this.lstDelivery);
            this.Controls.Add(this.grpTypist);
            this.Controls.Add(this.grpAuthor);
            this.Controls.Add(this.ssALS);
            this.Controls.Add(this.butCancel);
            this.Name = "frmMainForm";
            this.Text = "Letterhead Wizard";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmMainForm_FormClosing);
            this.Load += new System.EventHandler(this.frmMainForm_Load);
            this.grpTypist.ResumeLayout(false);
            this.grpAuthor.ResumeLayout(false);
            this.ssALS.ResumeLayout(false);
            this.ssALS.PerformLayout();
            this.GrpAddress.ResumeLayout(false);
            this.GrpAddress.PerformLayout();
            this.grpIncludes.ResumeLayout(false);
            this.grpIncludes.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        internal System.Windows.Forms.GroupBox grpTypist;
        internal System.Windows.Forms.Button butTypFind;
        internal System.Windows.Forms.Button butTypDelete;
        internal System.Windows.Forms.ComboBox cmbTypist;
        internal System.Windows.Forms.GroupBox grpAuthor;
        internal System.Windows.Forms.Button butAuthFind;
        internal System.Windows.Forms.Button butAuthDelete;
        internal System.Windows.Forms.ComboBox cmbAuthor;
        internal System.Windows.Forms.StatusStrip ssALS;
        internal System.Windows.Forms.Button butCancel;
        internal System.Windows.Forms.Button butOk;
        internal System.Windows.Forms.RadioButton radAddrManually;
        internal System.Windows.Forms.Button butDefault;
        internal System.Windows.Forms.RadioButton radAddrMultiple;
        internal System.Windows.Forms.RadioButton radAddr;
        internal System.Windows.Forms.CheckBox chkAdmittedTo;
        internal System.Windows.Forms.CheckBox chkDirectFax;
        internal System.Windows.Forms.CheckBox chkDelivery;
        internal System.Windows.Forms.CheckBox chkBCC;
        internal System.Windows.Forms.CheckBox chkCC;
        internal System.Windows.Forms.TextBox txtCC;
        internal System.Windows.Forms.GroupBox GrpAddress;
        internal System.Windows.Forms.CheckBox chkRe;
        internal System.Windows.Forms.TextBox txtBCC;
        internal System.Windows.Forms.TextBox txtDelivery;
        internal System.Windows.Forms.Label lblBCC;
        internal System.Windows.Forms.GroupBox grpIncludes;
        internal System.Windows.Forms.CheckBox chkEnc;
        internal System.Windows.Forms.Label lblCC;
        internal System.Windows.Forms.TextBox txtRe;
        internal System.Windows.Forms.Label lblRe;
        internal System.Windows.Forms.ListBox lstDelivery;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
    }
}