namespace OAMPS.Office.Word.Views.Wizards
{
    partial class ThemesOnlyWizard
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ThemesOnlyWizard));
            this.tbcWizardScreens = new System.Windows.Forms.TabControl();
            this.tvcPolicy = new Aga.Controls.Tree.TreeColumn();
            this.tvcOrderPolicy = new Aga.Controls.Tree.TreeColumn();
            this.tvcCurrent = new Aga.Controls.Tree.TreeColumn();
            this.tvcReccomended = new Aga.Controls.Tree.TreeColumn();
            this.tvcReccomendedId = new Aga.Controls.Tree.TreeColumn();
            this.tvcCurrentId = new Aga.Controls.Tree.TreeColumn();
            this._checked = new Aga.Controls.Tree.NodeControls.NodeCheckBox();
            this._name = new Aga.Controls.Tree.NodeControls.NodeTextBox();
            this._current = new Aga.Controls.Tree.NodeControls.NodeTextBox();
            this._reccommended = new Aga.Controls.Tree.NodeControls.NodeTextBox();
            this._reccommendedId = new Aga.Controls.Tree.NodeControls.NodeTextBox();
            this._currentId = new Aga.Controls.Tree.NodeControls.NodeTextBox();
            this._orderPolicy = new Aga.Controls.Tree.NodeControls.NodeTextBox();
            this.errorProvider = new System.Windows.Forms.ErrorProvider(this.components);
            this.lblCoverPageTitle = new System.Windows.Forms.Label();
            this.lblLogoTitle = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.btnBack = new System.Windows.Forms.Button();
            this.btnNext = new System.Windows.Forms.Button();
            this.lblSpeciality = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // tbcWizardScreens
            // 
            this.tbcWizardScreens.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tbcWizardScreens.Location = new System.Drawing.Point(5, 70);
            this.tbcWizardScreens.Name = "tbcWizardScreens";
            this.tbcWizardScreens.SelectedIndex = 0;
            this.tbcWizardScreens.Size = new System.Drawing.Size(625, 511);
            this.tbcWizardScreens.TabIndex = 34;
            // 
            // tvcPolicy
            // 
            this.tvcPolicy.Header = "Policy Class";
            this.tvcPolicy.MinColumnWidth = 250;
            this.tvcPolicy.SortOrder = System.Windows.Forms.SortOrder.None;
            this.tvcPolicy.TooltipText = null;
            this.tvcPolicy.Width = 250;
            // 
            // tvcOrderPolicy
            // 
            this.tvcOrderPolicy.Header = "Order";
            this.tvcOrderPolicy.SortOrder = System.Windows.Forms.SortOrder.None;
            this.tvcOrderPolicy.TooltipText = null;
            // 
            // tvcCurrent
            // 
            this.tvcCurrent.Header = "Current Insurer";
            this.tvcCurrent.MinColumnWidth = 150;
            this.tvcCurrent.SortOrder = System.Windows.Forms.SortOrder.None;
            this.tvcCurrent.TooltipText = "";
            this.tvcCurrent.Width = 150;
            // 
            // tvcReccomended
            // 
            this.tvcReccomended.Header = "Recommended Insurer";
            this.tvcReccomended.MinColumnWidth = 150;
            this.tvcReccomended.SortOrder = System.Windows.Forms.SortOrder.None;
            this.tvcReccomended.TooltipText = null;
            this.tvcReccomended.Width = 150;
            // 
            // tvcReccomendedId
            // 
            this.tvcReccomendedId.Header = "RecommendedId";
            this.tvcReccomendedId.MaxColumnWidth = 1;
            this.tvcReccomendedId.SortOrder = System.Windows.Forms.SortOrder.None;
            this.tvcReccomendedId.TooltipText = null;
            this.tvcReccomendedId.Width = 1;
            // 
            // tvcCurrentId
            // 
            this.tvcCurrentId.Header = "CurrentId";
            this.tvcCurrentId.MaxColumnWidth = 1;
            this.tvcCurrentId.SortOrder = System.Windows.Forms.SortOrder.None;
            this.tvcCurrentId.TooltipText = null;
            this.tvcCurrentId.Width = 1;
            // 
            // _checked
            // 
            this._checked.DataPropertyName = "Checked";
            this._checked.EditEnabled = true;
            this._checked.LeftMargin = 0;
            this._checked.ParentColumn = this.tvcPolicy;
            // 
            // _name
            // 
            this._name.DataPropertyName = "Text";
            this._name.EditEnabled = true;
            this._name.EditOnClick = true;
            this._name.IncrementalSearchEnabled = true;
            this._name.LeftMargin = 3;
            this._name.ParentColumn = this.tvcPolicy;
            // 
            // _current
            // 
            this._current.DataPropertyName = "Current";
            this._current.IncrementalSearchEnabled = true;
            this._current.LeftMargin = 3;
            this._current.ParentColumn = this.tvcCurrent;
            // 
            // _reccommended
            // 
            this._reccommended.DataPropertyName = "Reccommended";
            this._reccommended.IncrementalSearchEnabled = true;
            this._reccommended.LeftMargin = 3;
            this._reccommended.ParentColumn = this.tvcReccomended;
            // 
            // _reccommendedId
            // 
            this._reccommendedId.DataPropertyName = "ReccommendedId";
            this._reccommendedId.IncrementalSearchEnabled = true;
            this._reccommendedId.LeftMargin = 3;
            this._reccommendedId.ParentColumn = this.tvcReccomendedId;
            // 
            // _currentId
            // 
            this._currentId.DataPropertyName = "CurrentId";
            this._currentId.IncrementalSearchEnabled = true;
            this._currentId.LeftMargin = 3;
            this._currentId.ParentColumn = this.tvcCurrentId;
            // 
            // _orderPolicy
            // 
            this._orderPolicy.DataPropertyName = "OrderPolicy";
            this._orderPolicy.EditEnabled = true;
            this._orderPolicy.EditOnClick = true;
            this._orderPolicy.IncrementalSearchEnabled = true;
            this._orderPolicy.LeftMargin = 3;
            this._orderPolicy.ParentColumn = this.tvcOrderPolicy;
            // 
            // errorProvider
            // 
            this.errorProvider.BlinkStyle = System.Windows.Forms.ErrorBlinkStyle.NeverBlink;
            this.errorProvider.ContainerControl = this;
            // 
            // lblCoverPageTitle
            // 
            this.lblCoverPageTitle.AutoSize = true;
            this.lblCoverPageTitle.Location = new System.Drawing.Point(491, 663);
            this.lblCoverPageTitle.Name = "lblCoverPageTitle";
            this.lblCoverPageTitle.Size = new System.Drawing.Size(59, 13);
            this.lblCoverPageTitle.TabIndex = 39;
            this.lblCoverPageTitle.Text = "Woodshed";
            this.lblCoverPageTitle.Visible = false;
            // 
            // lblLogoTitle
            // 
            this.lblLogoTitle.AutoSize = true;
            this.lblLogoTitle.Location = new System.Drawing.Point(195, 663);
            this.lblLogoTitle.Name = "lblLogoTitle";
            this.lblLogoTitle.Size = new System.Drawing.Size(176, 13);
            this.lblLogoTitle.TabIndex = 38;
            this.lblLogoTitle.Text = "Arthur J. Gallagher & Co (Aus) Limited";
            this.lblLogoTitle.Visible = false;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.pictureBox1.Image = global::OAMPS.Office.Word.Properties.Resources.Gallagher;
            this.pictureBox1.InitialImage = ((System.Drawing.Image)(resources.GetObject("pictureBox1.InitialImage")));
            this.pictureBox1.Location = new System.Drawing.Point(158, -3);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(240, 70);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 37;
            this.pictureBox1.TabStop = false;
            // 
            // btnBack
            // 
            this.btnBack.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnBack.Enabled = false;
            this.btnBack.Location = new System.Drawing.Point(12, 583);
            this.btnBack.Name = "btnBack";
            this.btnBack.Size = new System.Drawing.Size(75, 26);
            this.btnBack.TabIndex = 35;
            this.btnBack.Text = "&Back";
            this.btnBack.UseVisualStyleBackColor = true;
            this.btnBack.Click += new System.EventHandler(this.btnBack_Click);
            // 
            // btnNext
            // 
            this.btnNext.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnNext.Location = new System.Drawing.Point(545, 586);
            this.btnNext.Name = "btnNext";
            this.btnNext.Size = new System.Drawing.Size(75, 23);
            this.btnNext.TabIndex = 36;
            this.btnNext.Text = "&Finish";
            this.btnNext.UseVisualStyleBackColor = true;
            this.btnNext.Click += new System.EventHandler(this.btnNext_Click);
            // 
            // lblSpeciality
            // 
            this.lblSpeciality.AutoSize = true;
            this.lblSpeciality.Location = new System.Drawing.Point(396, 663);
            this.lblSpeciality.Name = "lblSpeciality";
            this.lblSpeciality.Size = new System.Drawing.Size(0, 13);
            this.lblSpeciality.TabIndex = 40;
            this.lblSpeciality.Visible = false;
            // 
            // ThemesOnlyWizard
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(631, 610);
            this.Controls.Add(this.tbcWizardScreens);
            this.Controls.Add(this.lblCoverPageTitle);
            this.Controls.Add(this.lblLogoTitle);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.btnBack);
            this.Controls.Add(this.btnNext);
            this.Controls.Add(this.lblSpeciality);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "ThemesOnlyWizard";
            this.Text = "Update Images";
            this.Load += new System.EventHandler(this.ThemesOnlyWizard_Load);
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TabControl tbcWizardScreens;
        private Aga.Controls.Tree.TreeColumn tvcPolicy;
        private Aga.Controls.Tree.TreeColumn tvcOrderPolicy;
        private Aga.Controls.Tree.TreeColumn tvcCurrent;
        private Aga.Controls.Tree.TreeColumn tvcReccomended;
        private Aga.Controls.Tree.TreeColumn tvcReccomendedId;
        private Aga.Controls.Tree.TreeColumn tvcCurrentId;
        private Aga.Controls.Tree.NodeControls.NodeCheckBox _checked;
        private Aga.Controls.Tree.NodeControls.NodeTextBox _name;
        private Aga.Controls.Tree.NodeControls.NodeTextBox _current;
        private Aga.Controls.Tree.NodeControls.NodeTextBox _reccommended;
        private Aga.Controls.Tree.NodeControls.NodeTextBox _reccommendedId;
        private Aga.Controls.Tree.NodeControls.NodeTextBox _currentId;
        private Aga.Controls.Tree.NodeControls.NodeTextBox _orderPolicy;
        private System.Windows.Forms.ErrorProvider errorProvider;
        private System.Windows.Forms.Label lblCoverPageTitle;
        private System.Windows.Forms.Label lblLogoTitle;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Button btnBack;
        public System.Windows.Forms.Button btnNext;
        private System.Windows.Forms.Label lblSpeciality;

    }
}