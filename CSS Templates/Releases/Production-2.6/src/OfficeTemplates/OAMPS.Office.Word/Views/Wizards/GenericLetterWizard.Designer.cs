namespace OAMPS.Office.Word.Views.Wizards
{
    partial class GenericLetterWizard : BaseWizardForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(GenericLetterWizard));
            this.tvcPolicy = new Aga.Controls.Tree.TreeColumn();
            this.tvcCurrent = new Aga.Controls.Tree.TreeColumn();
            this.tvcRecommended = new Aga.Controls.Tree.TreeColumn();
            this._checked = new Aga.Controls.Tree.NodeControls.NodeCheckBox();
            this._name = new Aga.Controls.Tree.NodeControls.NodeTextBox();
            this._current = new Aga.Controls.Tree.NodeControls.NodeTextBox();
            this._recomended = new Aga.Controls.Tree.NodeControls.NodeTextBox();
            this.lblLogoTitle = new System.Windows.Forms.Label();
            this.lblCoverPageTitle = new System.Windows.Forms.Label();
            this.btnBack = new System.Windows.Forms.Button();
            this.pbOampsLogoFull = new System.Windows.Forms.PictureBox();
            this.btnNext = new System.Windows.Forms.Button();
            this.tbcWizardScreens = new System.Windows.Forms.TabControl();
            this.tabClient = new System.Windows.Forms.TabPage();
            this.gbClient = new System.Windows.Forms.GroupBox();
            this.txtClientAddress3 = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.requiredFieldLabel1 = new OAMPS.Office.Word.Helpers.Controls.RequiredFieldLabel();
            this.txtAddressee = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtClientAddress2 = new System.Windows.Forms.TextBox();
            this.lblCltAddress2 = new System.Windows.Forms.Label();
            this.txtSalutation = new System.Windows.Forms.TextBox();
            this.lblSalutation = new System.Windows.Forms.Label();
            this.txtClientAddress = new System.Windows.Forms.TextBox();
            this.lblCltAddress = new System.Windows.Forms.Label();
            this.txtClientName = new System.Windows.Forms.TextBox();
            this.lblClientName = new System.Windows.Forms.Label();
            this.tabExecutive = new System.Windows.Forms.TabPage();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.lblDepartment = new System.Windows.Forms.Label();
            this.txtExecutiveDepartment = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtBranchPhone = new System.Windows.Forms.TextBox();
            this.lblPostalAddress2 = new System.Windows.Forms.Label();
            this.txtPostal2 = new System.Windows.Forms.TextBox();
            this.lblAddress2 = new System.Windows.Forms.Label();
            this.txtBranchAddress2 = new System.Windows.Forms.TextBox();
            this.lblFax = new System.Windows.Forms.Label();
            this.btnAccountExecutiveLookup = new System.Windows.Forms.Button();
            this.txtFax = new System.Windows.Forms.TextBox();
            this.lblPostalAddress1 = new System.Windows.Forms.Label();
            this.txtPostal1 = new System.Windows.Forms.TextBox();
            this.lblAddress1 = new System.Windows.Forms.Label();
            this.txtBranchAddress1 = new System.Windows.Forms.TextBox();
            this.lblMobile = new System.Windows.Forms.Label();
            this.txtExecutiveMobile = new System.Windows.Forms.TextBox();
            this.lblOwnerPhone = new System.Windows.Forms.Label();
            this.txtExecutivePhone = new System.Windows.Forms.TextBox();
            this.lblOwnerName = new System.Windows.Forms.Label();
            this.lblOwnerEmail = new System.Windows.Forms.Label();
            this.lblOwnerTitle = new System.Windows.Forms.Label();
            this.txtExecutiveEmail = new System.Windows.Forms.TextBox();
            this.txtExecutiveTitle = new System.Windows.Forms.TextBox();
            this.txtExecutiveName = new System.Windows.Forms.TextBox();
            this.tabLetterInfo = new System.Windows.Forms.TabPage();
            this.chkPrePrint = new System.Windows.Forms.CheckBox();
            this.gpAdDocs = new System.Windows.Forms.GroupBox();
            this.label5 = new System.Windows.Forms.Label();
            this.txtPDSVersion = new System.Windows.Forms.TextBox();
            this.chkSatutory = new System.Windows.Forms.CheckBox();
            this.chkRisks = new System.Windows.Forms.CheckBox();
            this.chkPrivacy = new System.Windows.Forms.CheckBox();
            this.chkFSG = new System.Windows.Forms.CheckBox();
            this.chkWarning = new System.Windows.Forms.CheckBox();
            this.lblAdDocs = new System.Windows.Forms.Label();
            this.gbPolicy = new System.Windows.Forms.GroupBox();
            this.txtSubject = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.dateReport = new System.Windows.Forms.DateTimePicker();
            this.lblDateReport = new System.Windows.Forms.Label();
            this.txtReference = new System.Windows.Forms.TextBox();
            this.lblReference = new System.Windows.Forms.Label();
            this.errorProvider1 = new System.Windows.Forms.ErrorProvider(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.pbOampsLogoFull)).BeginInit();
            this.tbcWizardScreens.SuspendLayout();
            this.tabClient.SuspendLayout();
            this.gbClient.SuspendLayout();
            this.tabExecutive.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.tabLetterInfo.SuspendLayout();
            this.gpAdDocs.SuspendLayout();
            this.gbPolicy.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider1)).BeginInit();
            this.SuspendLayout();
            // 
            // tvcPolicy
            // 
            this.tvcPolicy.Header = "Policy Class";
            this.tvcPolicy.MinColumnWidth = 240;
            this.tvcPolicy.SortOrder = System.Windows.Forms.SortOrder.Descending;
            this.tvcPolicy.TooltipText = null;
            this.tvcPolicy.Width = 240;
            // 
            // tvcCurrent
            // 
            this.tvcCurrent.Header = "Current Insurer";
            this.tvcCurrent.MinColumnWidth = 140;
            this.tvcCurrent.SortOrder = System.Windows.Forms.SortOrder.None;
            this.tvcCurrent.TooltipText = null;
            this.tvcCurrent.Width = 140;
            // 
            // tvcRecommended
            // 
            this.tvcRecommended.Header = "Recommended Insurer";
            this.tvcRecommended.MinColumnWidth = 140;
            this.tvcRecommended.SortOrder = System.Windows.Forms.SortOrder.None;
            this.tvcRecommended.TooltipText = null;
            this.tvcRecommended.Width = 140;
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
            // _recomended
            // 
            this._recomended.DataPropertyName = "Recommended";
            this._recomended.IncrementalSearchEnabled = true;
            this._recomended.LeftMargin = 3;
            this._recomended.ParentColumn = this.tvcRecommended;
            // 
            // lblLogoTitle
            // 
            this.lblLogoTitle.AutoSize = true;
            this.lblLogoTitle.Location = new System.Drawing.Point(298, 646);
            this.lblLogoTitle.Name = "lblLogoTitle";
            this.lblLogoTitle.Size = new System.Drawing.Size(152, 13);
            this.lblLogoTitle.TabIndex = 64;
            this.lblLogoTitle.Text = "OAMPS Insurance Brokers Ltd";
            this.lblLogoTitle.Visible = false;
            // 
            // lblCoverPageTitle
            // 
            this.lblCoverPageTitle.AutoSize = true;
            this.lblCoverPageTitle.Location = new System.Drawing.Point(148, 646);
            this.lblCoverPageTitle.Name = "lblCoverPageTitle";
            this.lblCoverPageTitle.Size = new System.Drawing.Size(68, 13);
            this.lblCoverPageTitle.TabIndex = 63;
            this.lblCoverPageTitle.Text = "Boulder Opal";
            this.lblCoverPageTitle.Visible = false;
            // 
            // btnBack
            // 
            this.btnBack.Enabled = false;
            this.btnBack.Location = new System.Drawing.Point(8, 641);
            this.btnBack.Name = "btnBack";
            this.btnBack.Size = new System.Drawing.Size(75, 23);
            this.btnBack.TabIndex = 28;
            this.btnBack.Text = "&Back";
            this.btnBack.UseVisualStyleBackColor = true;
            this.btnBack.Click += new System.EventHandler(this.btnBack_Click);
            // 
            // pbOampsLogoFull
            // 
            this.pbOampsLogoFull.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.pbOampsLogoFull.Image = ((System.Drawing.Image)(resources.GetObject("pbOampsLogoFull.Image")));
            this.pbOampsLogoFull.InitialImage = ((System.Drawing.Image)(resources.GetObject("pbOampsLogoFull.InitialImage")));
            this.pbOampsLogoFull.Location = new System.Drawing.Point(162, 12);
            this.pbOampsLogoFull.Name = "pbOampsLogoFull";
            this.pbOampsLogoFull.Size = new System.Drawing.Size(190, 90);
            this.pbOampsLogoFull.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pbOampsLogoFull.TabIndex = 59;
            this.pbOampsLogoFull.TabStop = false;
            // 
            // btnNext
            // 
            this.btnNext.Location = new System.Drawing.Point(499, 643);
            this.btnNext.Name = "btnNext";
            this.btnNext.Size = new System.Drawing.Size(79, 23);
            this.btnNext.TabIndex = 29;
            this.btnNext.Text = "Next";
            this.btnNext.UseVisualStyleBackColor = true;
            this.btnNext.Click += new System.EventHandler(this.Next_Click);
            // 
            // tbcWizardScreens
            // 
            this.tbcWizardScreens.Controls.Add(this.tabClient);
            this.tbcWizardScreens.Controls.Add(this.tabExecutive);
            this.tbcWizardScreens.Controls.Add(this.tabLetterInfo);
            this.tbcWizardScreens.Location = new System.Drawing.Point(4, 110);
            this.tbcWizardScreens.Multiline = true;
            this.tbcWizardScreens.Name = "tbcWizardScreens";
            this.tbcWizardScreens.SelectedIndex = 0;
            this.tbcWizardScreens.Size = new System.Drawing.Size(569, 525);
            this.tbcWizardScreens.TabIndex = 58;
            this.tbcWizardScreens.Tag = "Letter Information";
            // 
            // tabClient
            // 
            this.tabClient.Controls.Add(this.gbClient);
            this.tabClient.Location = new System.Drawing.Point(4, 22);
            this.tabClient.Name = "tabClient";
            this.tabClient.Padding = new System.Windows.Forms.Padding(3);
            this.tabClient.Size = new System.Drawing.Size(561, 499);
            this.tabClient.TabIndex = 0;
            this.tabClient.Text = "Client Information";
            this.tabClient.UseVisualStyleBackColor = true;
            // 
            // gbClient
            // 
            this.gbClient.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gbClient.Controls.Add(this.txtClientAddress3);
            this.gbClient.Controls.Add(this.label6);
            this.gbClient.Controls.Add(this.requiredFieldLabel1);
            this.gbClient.Controls.Add(this.txtAddressee);
            this.gbClient.Controls.Add(this.label3);
            this.gbClient.Controls.Add(this.txtClientAddress2);
            this.gbClient.Controls.Add(this.lblCltAddress2);
            this.gbClient.Controls.Add(this.txtSalutation);
            this.gbClient.Controls.Add(this.lblSalutation);
            this.gbClient.Controls.Add(this.txtClientAddress);
            this.gbClient.Controls.Add(this.lblCltAddress);
            this.gbClient.Controls.Add(this.txtClientName);
            this.gbClient.Controls.Add(this.lblClientName);
            this.gbClient.Location = new System.Drawing.Point(6, 6);
            this.gbClient.Name = "gbClient";
            this.gbClient.Size = new System.Drawing.Size(534, 240);
            this.gbClient.TabIndex = 55;
            this.gbClient.TabStop = false;
            this.gbClient.Text = "Client Information";
            // 
            // txtClientAddress3
            // 
            this.txtClientAddress3.Location = new System.Drawing.Point(137, 163);
            this.txtClientAddress3.Name = "txtClientAddress3";
            this.txtClientAddress3.Size = new System.Drawing.Size(299, 20);
            this.txtClientAddress3.TabIndex = 4;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(10, 167);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(106, 13);
            this.label6.TabIndex = 68;
            this.label6.Text = "Client Address Line3:";
            // 
            // requiredFieldLabel1
            // 
            this.requiredFieldLabel1.AutoSize = true;
            this.requiredFieldLabel1.Field = this.gbClient;
            this.requiredFieldLabel1.Location = new System.Drawing.Point(82, 63);
            this.requiredFieldLabel1.Name = "requiredFieldLabel1";
            this.requiredFieldLabel1.Size = new System.Drawing.Size(11, 13);
            this.requiredFieldLabel1.TabIndex = 66;
            this.requiredFieldLabel1.Text = "*";
            // 
            // txtAddressee
            // 
            this.txtAddressee.Location = new System.Drawing.Point(137, 19);
            this.txtAddressee.Name = "txtAddressee";
            this.txtAddressee.Size = new System.Drawing.Size(299, 20);
            this.txtAddressee.TabIndex = 1;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(10, 23);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(60, 13);
            this.label3.TabIndex = 65;
            this.label3.Text = "Addressee:";
            // 
            // txtClientAddress2
            // 
            this.txtClientAddress2.Location = new System.Drawing.Point(137, 129);
            this.txtClientAddress2.Name = "txtClientAddress2";
            this.txtClientAddress2.Size = new System.Drawing.Size(299, 20);
            this.txtClientAddress2.TabIndex = 3;
            // 
            // lblCltAddress2
            // 
            this.lblCltAddress2.AutoSize = true;
            this.lblCltAddress2.Location = new System.Drawing.Point(10, 133);
            this.lblCltAddress2.Name = "lblCltAddress2";
            this.lblCltAddress2.Size = new System.Drawing.Size(106, 13);
            this.lblCltAddress2.TabIndex = 60;
            this.lblCltAddress2.Text = "Client Address Line2:";
            // 
            // txtSalutation
            // 
            this.txtSalutation.Location = new System.Drawing.Point(137, 197);
            this.txtSalutation.Name = "txtSalutation";
            this.txtSalutation.Size = new System.Drawing.Size(299, 20);
            this.txtSalutation.TabIndex = 5;
            // 
            // lblSalutation
            // 
            this.lblSalutation.AutoSize = true;
            this.lblSalutation.Location = new System.Drawing.Point(10, 201);
            this.lblSalutation.Name = "lblSalutation";
            this.lblSalutation.Size = new System.Drawing.Size(98, 13);
            this.lblSalutation.TabIndex = 58;
            this.lblSalutation.Text = "Salutation: (Dear...)";
            // 
            // txtClientAddress
            // 
            this.txtClientAddress.Location = new System.Drawing.Point(137, 91);
            this.txtClientAddress.Name = "txtClientAddress";
            this.txtClientAddress.Size = new System.Drawing.Size(299, 20);
            this.txtClientAddress.TabIndex = 2;
            // 
            // lblCltAddress
            // 
            this.lblCltAddress.AutoSize = true;
            this.lblCltAddress.Location = new System.Drawing.Point(10, 95);
            this.lblCltAddress.Name = "lblCltAddress";
            this.lblCltAddress.Size = new System.Drawing.Size(106, 13);
            this.lblCltAddress.TabIndex = 56;
            this.lblCltAddress.Text = "Client Address Line1:";
            // 
            // txtClientName
            // 
            this.txtClientName.Location = new System.Drawing.Point(137, 55);
            this.txtClientName.Name = "txtClientName";
            this.txtClientName.Size = new System.Drawing.Size(299, 20);
            this.txtClientName.TabIndex = 1;
            // 
            // lblClientName
            // 
            this.lblClientName.AutoSize = true;
            this.lblClientName.Location = new System.Drawing.Point(10, 59);
            this.lblClientName.Name = "lblClientName";
            this.lblClientName.Size = new System.Drawing.Size(67, 13);
            this.lblClientName.TabIndex = 49;
            this.lblClientName.Text = "Client Name:";
            // 
            // tabExecutive
            // 
            this.tabExecutive.Controls.Add(this.groupBox2);
            this.tabExecutive.Location = new System.Drawing.Point(4, 22);
            this.tabExecutive.Name = "tabExecutive";
            this.tabExecutive.Size = new System.Drawing.Size(561, 499);
            this.tabExecutive.TabIndex = 2;
            this.tabExecutive.Text = "Account Exec";
            this.tabExecutive.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox2.Controls.Add(this.lblDepartment);
            this.groupBox2.Controls.Add(this.txtExecutiveDepartment);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.txtBranchPhone);
            this.groupBox2.Controls.Add(this.lblPostalAddress2);
            this.groupBox2.Controls.Add(this.txtPostal2);
            this.groupBox2.Controls.Add(this.lblAddress2);
            this.groupBox2.Controls.Add(this.txtBranchAddress2);
            this.groupBox2.Controls.Add(this.lblFax);
            this.groupBox2.Controls.Add(this.btnAccountExecutiveLookup);
            this.groupBox2.Controls.Add(this.txtFax);
            this.groupBox2.Controls.Add(this.lblPostalAddress1);
            this.groupBox2.Controls.Add(this.txtPostal1);
            this.groupBox2.Controls.Add(this.lblAddress1);
            this.groupBox2.Controls.Add(this.txtBranchAddress1);
            this.groupBox2.Controls.Add(this.lblMobile);
            this.groupBox2.Controls.Add(this.txtExecutiveMobile);
            this.groupBox2.Controls.Add(this.lblOwnerPhone);
            this.groupBox2.Controls.Add(this.txtExecutivePhone);
            this.groupBox2.Controls.Add(this.lblOwnerName);
            this.groupBox2.Controls.Add(this.lblOwnerEmail);
            this.groupBox2.Controls.Add(this.lblOwnerTitle);
            this.groupBox2.Controls.Add(this.txtExecutiveEmail);
            this.groupBox2.Controls.Add(this.txtExecutiveTitle);
            this.groupBox2.Controls.Add(this.txtExecutiveName);
            this.groupBox2.Location = new System.Drawing.Point(4, 3);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(546, 430);
            this.groupBox2.TabIndex = 59;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Account Exec:";
            // 
            // lblDepartment
            // 
            this.lblDepartment.AutoSize = true;
            this.lblDepartment.Location = new System.Drawing.Point(6, 373);
            this.lblDepartment.Name = "lblDepartment";
            this.lblDepartment.Size = new System.Drawing.Size(65, 13);
            this.lblDepartment.TabIndex = 55;
            this.lblDepartment.Text = "Department:";
            // 
            // txtExecutiveDepartment
            // 
            this.txtExecutiveDepartment.Location = new System.Drawing.Point(136, 366);
            this.txtExecutiveDepartment.Name = "txtExecutiveDepartment";
            this.txtExecutiveDepartment.Size = new System.Drawing.Size(300, 20);
            this.txtExecutiveDepartment.TabIndex = 54;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(7, 213);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(78, 13);
            this.label1.TabIndex = 53;
            this.label1.Text = "Branch Phone:";
            // 
            // txtBranchPhone
            // 
            this.txtBranchPhone.Location = new System.Drawing.Point(136, 206);
            this.txtBranchPhone.Name = "txtBranchPhone";
            this.txtBranchPhone.Size = new System.Drawing.Size(300, 20);
            this.txtBranchPhone.TabIndex = 8;
            // 
            // lblPostalAddress2
            // 
            this.lblPostalAddress2.AutoSize = true;
            this.lblPostalAddress2.Location = new System.Drawing.Point(6, 341);
            this.lblPostalAddress2.Name = "lblPostalAddress2";
            this.lblPostalAddress2.Size = new System.Drawing.Size(123, 13);
            this.lblPostalAddress2.TabIndex = 51;
            this.lblPostalAddress2.Text = "Branch Postal Address2:";
            // 
            // txtPostal2
            // 
            this.txtPostal2.Location = new System.Drawing.Point(136, 334);
            this.txtPostal2.Name = "txtPostal2";
            this.txtPostal2.Size = new System.Drawing.Size(300, 20);
            this.txtPostal2.TabIndex = 12;
            // 
            // lblAddress2
            // 
            this.lblAddress2.AutoSize = true;
            this.lblAddress2.Location = new System.Drawing.Point(5, 277);
            this.lblAddress2.Name = "lblAddress2";
            this.lblAddress2.Size = new System.Drawing.Size(91, 13);
            this.lblAddress2.TabIndex = 49;
            this.lblAddress2.Text = "Branch Address2:";
            // 
            // txtBranchAddress2
            // 
            this.txtBranchAddress2.Location = new System.Drawing.Point(136, 270);
            this.txtBranchAddress2.Name = "txtBranchAddress2";
            this.txtBranchAddress2.Size = new System.Drawing.Size(300, 20);
            this.txtBranchAddress2.TabIndex = 10;
            // 
            // lblFax
            // 
            this.lblFax.AutoSize = true;
            this.lblFax.Location = new System.Drawing.Point(7, 181);
            this.lblFax.Name = "lblFax";
            this.lblFax.Size = new System.Drawing.Size(64, 13);
            this.lblFax.TabIndex = 47;
            this.lblFax.Text = "Branch Fax:";
            // 
            // btnAccountExecutiveLookup
            // 
            this.btnAccountExecutiveLookup.Location = new System.Drawing.Point(441, 12);
            this.btnAccountExecutiveLookup.Name = "btnAccountExecutiveLookup";
            this.btnAccountExecutiveLookup.Size = new System.Drawing.Size(75, 23);
            this.btnAccountExecutiveLookup.TabIndex = 2;
            this.btnAccountExecutiveLookup.Text = "Lookup";
            this.btnAccountExecutiveLookup.UseVisualStyleBackColor = true;
            this.btnAccountExecutiveLookup.Click += new System.EventHandler(this.btnAccountExecutiveLookup_Click);
            // 
            // txtFax
            // 
            this.txtFax.Location = new System.Drawing.Point(136, 174);
            this.txtFax.Name = "txtFax";
            this.txtFax.Size = new System.Drawing.Size(300, 20);
            this.txtFax.TabIndex = 7;
            // 
            // lblPostalAddress1
            // 
            this.lblPostalAddress1.AutoSize = true;
            this.lblPostalAddress1.Location = new System.Drawing.Point(6, 309);
            this.lblPostalAddress1.Name = "lblPostalAddress1";
            this.lblPostalAddress1.Size = new System.Drawing.Size(123, 13);
            this.lblPostalAddress1.TabIndex = 43;
            this.lblPostalAddress1.Text = "Branch Postal Address1:";
            // 
            // txtPostal1
            // 
            this.txtPostal1.Location = new System.Drawing.Point(136, 302);
            this.txtPostal1.Name = "txtPostal1";
            this.txtPostal1.Size = new System.Drawing.Size(300, 20);
            this.txtPostal1.TabIndex = 11;
            // 
            // lblAddress1
            // 
            this.lblAddress1.AutoSize = true;
            this.lblAddress1.Location = new System.Drawing.Point(5, 245);
            this.lblAddress1.Name = "lblAddress1";
            this.lblAddress1.Size = new System.Drawing.Size(91, 13);
            this.lblAddress1.TabIndex = 41;
            this.lblAddress1.Text = "Branch Address1:";
            // 
            // txtBranchAddress1
            // 
            this.txtBranchAddress1.Location = new System.Drawing.Point(136, 238);
            this.txtBranchAddress1.Name = "txtBranchAddress1";
            this.txtBranchAddress1.Size = new System.Drawing.Size(300, 20);
            this.txtBranchAddress1.TabIndex = 9;
            // 
            // lblMobile
            // 
            this.lblMobile.AutoSize = true;
            this.lblMobile.Location = new System.Drawing.Point(6, 149);
            this.lblMobile.Name = "lblMobile";
            this.lblMobile.Size = new System.Drawing.Size(41, 13);
            this.lblMobile.TabIndex = 39;
            this.lblMobile.Text = "Mobile:";
            // 
            // txtExecutiveMobile
            // 
            this.txtExecutiveMobile.Location = new System.Drawing.Point(136, 142);
            this.txtExecutiveMobile.Name = "txtExecutiveMobile";
            this.txtExecutiveMobile.Size = new System.Drawing.Size(300, 20);
            this.txtExecutiveMobile.TabIndex = 6;
            // 
            // lblOwnerPhone
            // 
            this.lblOwnerPhone.AutoSize = true;
            this.lblOwnerPhone.Location = new System.Drawing.Point(7, 117);
            this.lblOwnerPhone.Name = "lblOwnerPhone";
            this.lblOwnerPhone.Size = new System.Drawing.Size(41, 13);
            this.lblOwnerPhone.TabIndex = 37;
            this.lblOwnerPhone.Text = "Phone:";
            // 
            // txtExecutivePhone
            // 
            this.txtExecutivePhone.Location = new System.Drawing.Point(136, 110);
            this.txtExecutivePhone.Name = "txtExecutivePhone";
            this.txtExecutivePhone.Size = new System.Drawing.Size(300, 20);
            this.txtExecutivePhone.TabIndex = 5;
            // 
            // lblOwnerName
            // 
            this.lblOwnerName.AutoSize = true;
            this.lblOwnerName.Location = new System.Drawing.Point(5, 21);
            this.lblOwnerName.Name = "lblOwnerName";
            this.lblOwnerName.Size = new System.Drawing.Size(38, 13);
            this.lblOwnerName.TabIndex = 35;
            this.lblOwnerName.Text = "Name:";
            // 
            // lblOwnerEmail
            // 
            this.lblOwnerEmail.AutoSize = true;
            this.lblOwnerEmail.Location = new System.Drawing.Point(6, 85);
            this.lblOwnerEmail.Name = "lblOwnerEmail";
            this.lblOwnerEmail.Size = new System.Drawing.Size(35, 13);
            this.lblOwnerEmail.TabIndex = 34;
            this.lblOwnerEmail.Text = "Email:";
            // 
            // lblOwnerTitle
            // 
            this.lblOwnerTitle.AutoSize = true;
            this.lblOwnerTitle.Location = new System.Drawing.Point(6, 53);
            this.lblOwnerTitle.Name = "lblOwnerTitle";
            this.lblOwnerTitle.Size = new System.Drawing.Size(30, 13);
            this.lblOwnerTitle.TabIndex = 33;
            this.lblOwnerTitle.Text = "Title:";
            // 
            // txtExecutiveEmail
            // 
            this.txtExecutiveEmail.Location = new System.Drawing.Point(136, 78);
            this.txtExecutiveEmail.Name = "txtExecutiveEmail";
            this.txtExecutiveEmail.Size = new System.Drawing.Size(300, 20);
            this.txtExecutiveEmail.TabIndex = 4;
            // 
            // txtExecutiveTitle
            // 
            this.txtExecutiveTitle.Location = new System.Drawing.Point(136, 46);
            this.txtExecutiveTitle.Name = "txtExecutiveTitle";
            this.txtExecutiveTitle.Size = new System.Drawing.Size(300, 20);
            this.txtExecutiveTitle.TabIndex = 3;
            // 
            // txtExecutiveName
            // 
            this.txtExecutiveName.Location = new System.Drawing.Point(136, 14);
            this.txtExecutiveName.Name = "txtExecutiveName";
            this.txtExecutiveName.Size = new System.Drawing.Size(300, 20);
            this.txtExecutiveName.TabIndex = 1;
            // 
            // tabLetterInfo
            // 
            this.tabLetterInfo.Controls.Add(this.chkPrePrint);
            this.tabLetterInfo.Controls.Add(this.gpAdDocs);
            this.tabLetterInfo.Controls.Add(this.gbPolicy);
            this.tabLetterInfo.Location = new System.Drawing.Point(4, 22);
            this.tabLetterInfo.Name = "tabLetterInfo";
            this.tabLetterInfo.Padding = new System.Windows.Forms.Padding(3);
            this.tabLetterInfo.Size = new System.Drawing.Size(561, 499);
            this.tabLetterInfo.TabIndex = 1;
            this.tabLetterInfo.Text = "Letter Information";
            this.tabLetterInfo.UseVisualStyleBackColor = true;
            // 
            // chkPrePrint
            // 
            this.chkPrePrint.AutoSize = true;
            this.chkPrePrint.Location = new System.Drawing.Point(24, 331);
            this.chkPrePrint.Name = "chkPrePrint";
            this.chkPrePrint.Size = new System.Drawing.Size(197, 17);
            this.chkPrePrint.TabIndex = 71;
            this.chkPrePrint.Text = "Are you using pre-printed stationery?";
            this.chkPrePrint.UseVisualStyleBackColor = true;
            this.chkPrePrint.CheckedChanged += new System.EventHandler(this.chkPrePrint_CheckedChanged);
            // 
            // gpAdDocs
            // 
            this.gpAdDocs.Controls.Add(this.label5);
            this.gpAdDocs.Controls.Add(this.txtPDSVersion);
            this.gpAdDocs.Controls.Add(this.chkSatutory);
            this.gpAdDocs.Controls.Add(this.chkRisks);
            this.gpAdDocs.Controls.Add(this.chkPrivacy);
            this.gpAdDocs.Controls.Add(this.chkFSG);
            this.gpAdDocs.Controls.Add(this.chkWarning);
            this.gpAdDocs.Controls.Add(this.lblAdDocs);
            this.gpAdDocs.Location = new System.Drawing.Point(6, 127);
            this.gpAdDocs.Name = "gpAdDocs";
            this.gpAdDocs.Size = new System.Drawing.Size(542, 194);
            this.gpAdDocs.TabIndex = 63;
            this.gpAdDocs.TabStop = false;
            this.gpAdDocs.Text = "Enclosed Documents";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(8, 26);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(409, 13);
            this.label5.TabIndex = 1000;
            this.label5.Text = "PDS description and version no (i.e.\'NTI Fleet Insurance Policy (NTI166(15/3/2013" +
    ")):";
            // 
            // txtPDSVersion
            // 
            this.txtPDSVersion.Location = new System.Drawing.Point(11, 48);
            this.txtPDSVersion.Name = "txtPDSVersion";
            this.txtPDSVersion.Size = new System.Drawing.Size(167, 20);
            this.txtPDSVersion.TabIndex = 63;
            // 
            // chkSatutory
            // 
            this.chkSatutory.AutoSize = true;
            this.chkSatutory.Location = new System.Drawing.Point(18, 167);
            this.chkSatutory.Name = "chkSatutory";
            this.chkSatutory.Size = new System.Drawing.Size(109, 17);
            this.chkSatutory.TabIndex = 67;
            this.chkSatutory.Text = "Important Notices";
            this.chkSatutory.UseVisualStyleBackColor = true;
            this.chkSatutory.CheckedChanged += new System.EventHandler(this.chkSatutory_CheckedChanged);
            // 
            // chkRisks
            // 
            this.chkRisks.AutoSize = true;
            this.chkRisks.Location = new System.Drawing.Point(18, 144);
            this.chkRisks.Name = "chkRisks";
            this.chkRisks.Size = new System.Drawing.Size(149, 17);
            this.chkRisks.TabIndex = 66;
            this.chkRisks.Text = "Uninsured Risks Checklist";
            this.chkRisks.UseVisualStyleBackColor = true;
            this.chkRisks.CheckedChanged += new System.EventHandler(this.chkRisks_CheckedChanged);
            // 
            // chkPrivacy
            // 
            this.chkPrivacy.AutoSize = true;
            this.chkPrivacy.Location = new System.Drawing.Point(18, 121);
            this.chkPrivacy.Name = "chkPrivacy";
            this.chkPrivacy.Size = new System.Drawing.Size(112, 17);
            this.chkPrivacy.TabIndex = 65;
            this.chkPrivacy.Text = "Privacy Statement";
            this.chkPrivacy.UseVisualStyleBackColor = true;
            this.chkPrivacy.CheckedChanged += new System.EventHandler(this.chkPrivacy_CheckedChanged);
            // 
            // chkFSG
            // 
            this.chkFSG.AutoSize = true;
            this.chkFSG.Location = new System.Drawing.Point(18, 98);
            this.chkFSG.Name = "chkFSG";
            this.chkFSG.Size = new System.Drawing.Size(143, 17);
            this.chkFSG.TabIndex = 64;
            this.chkFSG.Text = "Financial Services Guide";
            this.chkFSG.UseVisualStyleBackColor = true;
            this.chkFSG.CheckedChanged += new System.EventHandler(this.chkFSG_CheckedChanged);
            // 
            // chkWarning
            // 
            this.chkWarning.AutoSize = true;
            this.chkWarning.Location = new System.Drawing.Point(207, 87);
            this.chkWarning.Name = "chkWarning";
            this.chkWarning.Size = new System.Drawing.Size(142, 17);
            this.chkWarning.TabIndex = 68;
            this.chkWarning.Text = "General Advice Warning";
            this.chkWarning.UseVisualStyleBackColor = true;
            this.chkWarning.Visible = false;
            this.chkWarning.CheckedChanged += new System.EventHandler(this.chkWarning_CheckedChanged);
            // 
            // lblAdDocs
            // 
            this.lblAdDocs.AutoSize = true;
            this.lblAdDocs.Location = new System.Drawing.Point(8, 76);
            this.lblAdDocs.Name = "lblAdDocs";
            this.lblAdDocs.Size = new System.Drawing.Size(164, 13);
            this.lblAdDocs.TabIndex = 1000;
            this.lblAdDocs.Text = "Additional Docs available to print:";
            // 
            // gbPolicy
            // 
            this.gbPolicy.Controls.Add(this.txtSubject);
            this.gbPolicy.Controls.Add(this.label2);
            this.gbPolicy.Controls.Add(this.dateReport);
            this.gbPolicy.Controls.Add(this.lblDateReport);
            this.gbPolicy.Controls.Add(this.txtReference);
            this.gbPolicy.Controls.Add(this.lblReference);
            this.gbPolicy.Location = new System.Drawing.Point(6, 6);
            this.gbPolicy.Name = "gbPolicy";
            this.gbPolicy.Size = new System.Drawing.Size(542, 115);
            this.gbPolicy.TabIndex = 58;
            this.gbPolicy.TabStop = false;
            this.gbPolicy.Text = "Policy Information";
            // 
            // txtSubject
            // 
            this.txtSubject.Location = new System.Drawing.Point(136, 19);
            this.txtSubject.Name = "txtSubject";
            this.txtSubject.Size = new System.Drawing.Size(299, 20);
            this.txtSubject.TabIndex = 54;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 22);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(46, 13);
            this.label2.TabIndex = 63;
            this.label2.Text = "Subject:";
            // 
            // dateReport
            // 
            this.dateReport.CustomFormat = "d-MMM-yyyy";
            this.dateReport.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dateReport.Location = new System.Drawing.Point(136, 71);
            this.dateReport.Name = "dateReport";
            this.dateReport.Size = new System.Drawing.Size(200, 20);
            this.dateReport.TabIndex = 62;
            // 
            // lblDateReport
            // 
            this.lblDateReport.AutoSize = true;
            this.lblDateReport.Location = new System.Drawing.Point(8, 75);
            this.lblDateReport.Name = "lblDateReport";
            this.lblDateReport.Size = new System.Drawing.Size(61, 13);
            this.lblDateReport.TabIndex = 0;
            this.lblDateReport.Text = "Letter date:";
            // 
            // txtReference
            // 
            this.txtReference.Location = new System.Drawing.Point(136, 45);
            this.txtReference.Name = "txtReference";
            this.txtReference.Size = new System.Drawing.Size(299, 20);
            this.txtReference.TabIndex = 55;
            // 
            // lblReference
            // 
            this.lblReference.AutoSize = true;
            this.lblReference.Location = new System.Drawing.Point(6, 48);
            this.lblReference.Name = "lblReference";
            this.lblReference.Size = new System.Drawing.Size(80, 13);
            this.lblReference.TabIndex = 0;
            this.lblReference.Text = "Our Reference:";
            // 
            // errorProvider1
            // 
            this.errorProvider1.BlinkStyle = System.Windows.Forms.ErrorBlinkStyle.NeverBlink;
            this.errorProvider1.ContainerControl = this;
            // 
            // GenericLetterWizard
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.ClientSize = new System.Drawing.Size(590, 678);
            this.Controls.Add(this.lblLogoTitle);
            this.Controls.Add(this.lblCoverPageTitle);
            this.Controls.Add(this.btnBack);
            this.Controls.Add(this.pbOampsLogoFull);
            this.Controls.Add(this.btnNext);
            this.Controls.Add(this.tbcWizardScreens);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "GenericLetterWizard";
            this.Text = "Generic Client Letter";
            this.Load += new System.EventHandler(this.RenewalLetter_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pbOampsLogoFull)).EndInit();
            this.tbcWizardScreens.ResumeLayout(false);
            this.tabClient.ResumeLayout(false);
            this.gbClient.ResumeLayout(false);
            this.gbClient.PerformLayout();
            this.tabExecutive.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.tabLetterInfo.ResumeLayout(false);
            this.tabLetterInfo.PerformLayout();
            this.gpAdDocs.ResumeLayout(false);
            this.gpAdDocs.PerformLayout();
            this.gbPolicy.ResumeLayout(false);
            this.gbPolicy.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.PictureBox pbOampsLogoFull;
        private System.Windows.Forms.TabPage tabLetterInfo;
        private System.Windows.Forms.GroupBox gpAdDocs;
        private System.Windows.Forms.Label lblAdDocs;
        private System.Windows.Forms.GroupBox gbPolicy;
        private System.Windows.Forms.TextBox txtReference;
        private System.Windows.Forms.Label lblReference;
        private System.Windows.Forms.TabPage tabClient;
        private System.Windows.Forms.GroupBox gbClient;
        private System.Windows.Forms.TextBox txtSalutation;
        private System.Windows.Forms.Label lblSalutation;
        private System.Windows.Forms.TextBox txtClientAddress;
        private System.Windows.Forms.Label lblCltAddress;
        private System.Windows.Forms.TextBox txtClientName;
        private System.Windows.Forms.Label lblClientName;
        private System.Windows.Forms.TabControl tbcWizardScreens;
        private System.Windows.Forms.Button btnBack;
        private System.Windows.Forms.Button btnNext;
        private System.Windows.Forms.TextBox txtClientAddress2;
        private System.Windows.Forms.Label lblCltAddress2;
        private System.Windows.Forms.CheckBox chkSatutory;
        private System.Windows.Forms.CheckBox chkRisks;
        private System.Windows.Forms.CheckBox chkPrivacy;
        private System.Windows.Forms.CheckBox chkFSG;
        private System.Windows.Forms.CheckBox chkWarning;
        private System.Windows.Forms.TabPage tabExecutive;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label lblPostalAddress1;
        private System.Windows.Forms.TextBox txtPostal1;
        private System.Windows.Forms.Label lblAddress1;
        private System.Windows.Forms.TextBox txtBranchAddress1;
        private System.Windows.Forms.Label lblMobile;
        private System.Windows.Forms.TextBox txtExecutiveMobile;
        private System.Windows.Forms.Button btnAccountExecutiveLookup;
        private System.Windows.Forms.Label lblOwnerPhone;
        private System.Windows.Forms.TextBox txtExecutivePhone;
        private System.Windows.Forms.Label lblOwnerName;
        private System.Windows.Forms.Label lblOwnerEmail;
        private System.Windows.Forms.Label lblOwnerTitle;
        private System.Windows.Forms.TextBox txtExecutiveEmail;
        private System.Windows.Forms.TextBox txtExecutiveTitle;
        private System.Windows.Forms.TextBox txtExecutiveName;
        private System.Windows.Forms.Label lblFax;
        private System.Windows.Forms.TextBox txtFax;
        private System.Windows.Forms.Label lblAddress2;
        private System.Windows.Forms.TextBox txtBranchAddress2;
        private System.Windows.Forms.Label lblPostalAddress2;
        private System.Windows.Forms.TextBox txtPostal2;
        private Aga.Controls.Tree.TreeColumn tvcPolicy;
        private Aga.Controls.Tree.TreeColumn tvcCurrent;
        private Aga.Controls.Tree.TreeColumn tvcRecommended;
        private Aga.Controls.Tree.NodeControls.NodeCheckBox _checked;
        private Aga.Controls.Tree.NodeControls.NodeTextBox _name;
        private Aga.Controls.Tree.NodeControls.NodeTextBox _current;
        private Aga.Controls.Tree.NodeControls.NodeTextBox _recomended;
        private System.Windows.Forms.Label lblCoverPageTitle;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtBranchPhone;
        private System.Windows.Forms.Label lblDateReport;
        private System.Windows.Forms.ErrorProvider errorProvider1;
        private System.Windows.Forms.DateTimePicker dateReport;
        private System.Windows.Forms.RadioButton rdoUFINo;
        private System.Windows.Forms.RadioButton rdoUFIYes;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.Label lblDepartment;
        private System.Windows.Forms.TextBox txtExecutiveDepartment;
        private System.Windows.Forms.TextBox txtPDSVersion;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtAddressee;
        private System.Windows.Forms.Label label3;
        private Helpers.Controls.RequiredFieldLabel requiredFieldLabel1;
        private System.Windows.Forms.TextBox txtClientAddress3;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.CheckBox chkPrePrint;
        private System.Windows.Forms.TextBox txtSubject;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label lblLogoTitle;

    }
}
