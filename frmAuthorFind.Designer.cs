namespace MyRibbonAddIn.ALS_FWW_Word
{
    partial class frmAuthorFind
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
            this.grpAuthor = new System.Windows.Forms.GroupBox();
            this.ListView1 = new System.Windows.Forms.ListView();
            this.butAdd = new System.Windows.Forms.Button();
            this.butClose = new System.Windows.Forms.Button();
            this.grpAuthor.SuspendLayout();
            this.SuspendLayout();
            // 
            // grpAuthor
            // 
            this.grpAuthor.Controls.Add(this.ListView1);
            this.grpAuthor.Location = new System.Drawing.Point(10, 10);
            this.grpAuthor.Name = "grpAuthor";
            this.grpAuthor.Size = new System.Drawing.Size(289, 275);
            this.grpAuthor.TabIndex = 2;
            this.grpAuthor.TabStop = false;
            this.grpAuthor.Text = "Author";
            // 
            // ListView1
            // 
            this.ListView1.Location = new System.Drawing.Point(3, 16);
            this.ListView1.MultiSelect = false;
            this.ListView1.Name = "ListView1";
            this.ListView1.Size = new System.Drawing.Size(280, 253);
            this.ListView1.TabIndex = 0;
            this.ListView1.UseCompatibleStateImageBehavior = false;
            this.ListView1.View = System.Windows.Forms.View.Details;
            // 
            // butAdd
            // 
            this.butAdd.Location = new System.Drawing.Point(305, 23);
            this.butAdd.Name = "butAdd";
            this.butAdd.Size = new System.Drawing.Size(75, 23);
            this.butAdd.TabIndex = 3;
            this.butAdd.Text = "A&dd";
            this.butAdd.UseVisualStyleBackColor = true;
            this.butAdd.Click += new System.EventHandler(this.butAdd_Click);
            // 
            // butClose
            // 
            this.butClose.Location = new System.Drawing.Point(305, 52);
            this.butClose.Name = "butClose";
            this.butClose.Size = new System.Drawing.Size(75, 23);
            this.butClose.TabIndex = 4;
            this.butClose.Text = "&Close";
            this.butClose.UseVisualStyleBackColor = true;
            this.butClose.Click += new System.EventHandler(this.butClose_Click);
            // 
            // frmAuthorFind
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(391, 295);
            this.Controls.Add(this.grpAuthor);
            this.Controls.Add(this.butAdd);
            this.Controls.Add(this.butClose);
            this.Name = "frmAuthorFind";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "frmAuthorFind";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmAuthorFind_FormClosing);
            this.Load += new System.EventHandler(this.frmAuthorFind_Load);
            this.grpAuthor.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        internal System.Windows.Forms.GroupBox grpAuthor;
        internal System.Windows.Forms.ListView ListView1;
        internal System.Windows.Forms.Button butAdd;
        internal System.Windows.Forms.Button butClose;
    }
}