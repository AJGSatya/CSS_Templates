namespace OAMPS.Office.Word.Views.Word
{
    partial class TemplateUpdateChecker
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            this.dgUpdatedItems = new System.Windows.Forms.DataGridView();
            this.Title = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ReleaseNotes = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.UsedDate = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.DateModified = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.label1 = new System.Windows.Forms.Label();
            this.chkHide = new System.Windows.Forms.CheckBox();
            this.btnClose = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgUpdatedItems)).BeginInit();
            this.SuspendLayout();
            // 
            // dgUpdatedItems
            // 
            this.dgUpdatedItems.AllowUserToAddRows = false;
            this.dgUpdatedItems.AllowUserToDeleteRows = false;
            this.dgUpdatedItems.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgUpdatedItems.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.dgUpdatedItems.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgUpdatedItems.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Title,
            this.ReleaseNotes,
            this.UsedDate,
            this.DateModified});
            this.dgUpdatedItems.Location = new System.Drawing.Point(0, 71);
            this.dgUpdatedItems.Name = "dgUpdatedItems";
            this.dgUpdatedItems.Size = new System.Drawing.Size(903, 410);
            this.dgUpdatedItems.TabIndex = 2;
            // 
            // Title
            // 
            this.Title.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCellsExceptHeader;
            this.Title.HeaderText = "Template Section";
            this.Title.Name = "Title";
            this.Title.ReadOnly = true;
            this.Title.Width = 5;
            // 
            // ReleaseNotes
            // 
            this.ReleaseNotes.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.ReleaseNotes.DefaultCellStyle = dataGridViewCellStyle3;
            this.ReleaseNotes.HeaderText = "Release Notes";
            this.ReleaseNotes.Name = "ReleaseNotes";
            // 
            // UsedDate
            // 
            this.UsedDate.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCellsExceptHeader;
            this.UsedDate.HeaderText = "Date section was used in this document";
            this.UsedDate.Name = "UsedDate";
            this.UsedDate.ReadOnly = true;
            this.UsedDate.Width = 5;
            // 
            // DateModified
            // 
            this.DateModified.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCellsExceptHeader;
            this.DateModified.HeaderText = "Date section was modified by the CSS Administrator";
            this.DateModified.Name = "DateModified";
            this.DateModified.ReadOnly = true;
            this.DateModified.Width = 5;
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(9, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(784, 48);
            this.label1.TabIndex = 3;
            this.label1.Text = "The following sections of the underlying template have been updated since you cre" +
    "ated this document.  Read the Release Notes below to appreciate any risk to your" +
    " client.";
            // 
            // chkHide
            // 
            this.chkHide.AutoSize = true;
            this.chkHide.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkHide.Location = new System.Drawing.Point(324, 493);
            this.chkHide.Name = "chkHide";
            this.chkHide.Size = new System.Drawing.Size(469, 21);
            this.chkHide.TabIndex = 4;
            this.chkHide.Text = "I understand these changes.  Please don\'t show these changes again.";
            this.chkHide.UseVisualStyleBackColor = true;
            this.chkHide.CheckedChanged += new System.EventHandler(this.chkHide_CheckedChanged);
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(816, 489);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 29);
            this.btnClose.TabIndex = 5;
            this.btnClose.Text = "&Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // TemplateUpdateChecker
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(903, 528);
            this.ControlBox = false;
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.chkHide);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dgUpdatedItems);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "TemplateUpdateChecker";
            this.Text = " IMPORTANT";
            this.Load += new System.EventHandler(this.TemplateUpdateChecker_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgUpdatedItems)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dgUpdatedItems;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox chkHide;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.DataGridViewTextBoxColumn Title;
        private System.Windows.Forms.DataGridViewTextBoxColumn ReleaseNotes;
        private System.Windows.Forms.DataGridViewTextBoxColumn UsedDate;
        private System.Windows.Forms.DataGridViewTextBoxColumn DateModified;
    }
}