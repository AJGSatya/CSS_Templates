namespace OAMPS.Office.Word.Views.Wizards
{
    partial class QuoteSlipWizard
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(QuoteSlipWizard));
            this.tvcPolicy = new Aga.Controls.Tree.TreeColumn();
            this._checked = new Aga.Controls.Tree.NodeControls.NodeCheckBox();
            this._name = new Aga.Controls.Tree.NodeControls.NodeTextBox();
            this.lblCoverPageTitle = new System.Windows.Forms.Label();
            this.lblLogoTitle = new System.Windows.Forms.Label();
            this.btnBack = new System.Windows.Forms.Button();
            this.btnNext = new System.Windows.Forms.Button();
            this.pbOampsLogoFull = new System.Windows.Forms.PictureBox();
            this.tbcWizardScreens = new System.Windows.Forms.TabControl();
            this.tabClientInfo = new System.Windows.Forms.TabPage();
            this.groupBox7 = new System.Windows.Forms.GroupBox();
            this.clbQuestions = new System.Windows.Forms.CheckedListBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label18 = new System.Windows.Forms.Label();
            this.dtpDateRequired = new System.Windows.Forms.DateTimePicker();
            this.requiredFieldLabel2 = new OAMPS.Office.Word.Helpers.Controls.RequiredFieldLabel();
            this.txtClientCommonName = new System.Windows.Forms.TextBox();
            this.lblClientCommonName = new System.Windows.Forms.Label();
            this.dtpPeriodOfInsuranceFrom = new System.Windows.Forms.DateTimePicker();
            this.lblPeriodOfInsuranceTo = new System.Windows.Forms.Label();
            this.dtpPeriodOfInsuranceTo = new System.Windows.Forms.DateTimePicker();
            this.lblPeriodOfInsuranceFrom = new System.Windows.Forms.Label();
            this.txtClientName = new System.Windows.Forms.TextBox();
            this.lblClientName = new System.Windows.Forms.Label();
            this.tabAccountExecutive = new System.Windows.Forms.TabPage();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.lblDepartment = new System.Windows.Forms.Label();
            this.txtExecutiveDepartment = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
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
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.label13 = new System.Windows.Forms.Label();
            this.txtAssitantExecDepartment = new System.Windows.Forms.TextBox();
            this.btnAssistantAccountExecutiveLookup = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.txtAssistantExecutivePhone = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.txtAssistantExecutiveEmail = new System.Windows.Forms.TextBox();
            this.txtAssistantExecutiveTitle = new System.Windows.Forms.TextBox();
            this.txtAssistantExecutiveName = new System.Windows.Forms.TextBox();
            this.tabPolicyClass = new System.Windows.Forms.TabPage();
            this.txtFindPolicyClass = new System.Windows.Forms.TextBox();
            this.label19 = new System.Windows.Forms.Label();
            this.tvaPolicies = new Aga.Controls.Tree.TreeViewAdv();
            this.tabQuestions = new System.Windows.Forms.TabPage();
            this.label16 = new System.Windows.Forms.Label();
            this.clbQustions = new System.Windows.Forms.CheckedListBox();
            this.errorProvider = new System.Windows.Forms.ErrorProvider(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.pbOampsLogoFull)).BeginInit();
            this.tbcWizardScreens.SuspendLayout();
            this.tabClientInfo.SuspendLayout();
            this.groupBox7.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.tabAccountExecutive.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.tabPolicyClass.SuspendLayout();
            this.tabQuestions.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider)).BeginInit();
            this.SuspendLayout();
            // 
            // tvcPolicy
            // 
            this.tvcPolicy.Header = "Policy Class";
            this.tvcPolicy.MinColumnWidth = 250;
            this.tvcPolicy.SortOrder = System.Windows.Forms.SortOrder.None;
            this.tvcPolicy.TooltipText = null;
            this.tvcPolicy.Width = 250;
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
            // lblCoverPageTitle
            // 
            this.lblCoverPageTitle.AutoSize = true;
            this.lblCoverPageTitle.Location = new System.Drawing.Point(368, 663);
            this.lblCoverPageTitle.Name = "lblCoverPageTitle";
            this.lblCoverPageTitle.Size = new System.Drawing.Size(68, 13);
            this.lblCoverPageTitle.TabIndex = 42;
            this.lblCoverPageTitle.Text = "Boulder Opal";
            this.lblCoverPageTitle.Visible = false;
            // 
            // lblLogoTitle
            // 
            this.lblLogoTitle.AutoSize = true;
            this.lblLogoTitle.Location = new System.Drawing.Point(139, 663);
            this.lblLogoTitle.Name = "lblLogoTitle";
            this.lblLogoTitle.Size = new System.Drawing.Size(152, 13);
            this.lblLogoTitle.TabIndex = 41;
            this.lblLogoTitle.Text = "OAMPS Insurance Brokers Ltd";
            this.lblLogoTitle.Visible = false;
            // 
            // btnBack
            // 
            this.btnBack.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnBack.Enabled = false;
            this.btnBack.Location = new System.Drawing.Point(11, 678);
            this.btnBack.Name = "btnBack";
            this.btnBack.Size = new System.Drawing.Size(76, 26);
            this.btnBack.TabIndex = 39;
            this.btnBack.Text = "&Back";
            this.btnBack.UseVisualStyleBackColor = true;
            this.btnBack.Click += new System.EventHandler(this.btnBack_Click);
            // 
            // btnNext
            // 
            this.btnNext.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnNext.Location = new System.Drawing.Point(526, 679);
            this.btnNext.Name = "btnNext";
            this.btnNext.Size = new System.Drawing.Size(75, 23);
            this.btnNext.TabIndex = 40;
            this.btnNext.Text = "&Next";
            this.btnNext.UseVisualStyleBackColor = true;
            this.btnNext.Click += new System.EventHandler(this.btnNext_Click);
            // 
            // pbOampsLogoFull
            // 
            this.pbOampsLogoFull.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.pbOampsLogoFull.Image = ((System.Drawing.Image)(resources.GetObject("pbOampsLogoFull.Image")));
            this.pbOampsLogoFull.InitialImage = ((System.Drawing.Image)(resources.GetObject("pbOampsLogoFull.InitialImage")));
            this.pbOampsLogoFull.Location = new System.Drawing.Point(165, 2);
            this.pbOampsLogoFull.Name = "pbOampsLogoFull";
            this.pbOampsLogoFull.Size = new System.Drawing.Size(140, 58);
            this.pbOampsLogoFull.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pbOampsLogoFull.TabIndex = 36;
            this.pbOampsLogoFull.TabStop = false;
            // 
            // tbcWizardScreens
            // 
            this.tbcWizardScreens.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tbcWizardScreens.Controls.Add(this.tabClientInfo);
            this.tbcWizardScreens.Controls.Add(this.tabAccountExecutive);
            this.tbcWizardScreens.Controls.Add(this.tabPolicyClass);
            this.tbcWizardScreens.Controls.Add(this.tabQuestions);
            this.tbcWizardScreens.Location = new System.Drawing.Point(5, 66);
            this.tbcWizardScreens.Name = "tbcWizardScreens";
            this.tbcWizardScreens.SelectedIndex = 0;
            this.tbcWizardScreens.Size = new System.Drawing.Size(608, 610);
            this.tbcWizardScreens.TabIndex = 33;
            // 
            // tabClientInfo
            // 
            this.tabClientInfo.Controls.Add(this.groupBox7);
            this.tabClientInfo.Controls.Add(this.groupBox1);
            this.tabClientInfo.Location = new System.Drawing.Point(4, 22);
            this.tabClientInfo.Name = "tabClientInfo";
            this.tabClientInfo.Padding = new System.Windows.Forms.Padding(3);
            this.tabClientInfo.Size = new System.Drawing.Size(600, 584);
            this.tabClientInfo.TabIndex = 5;
            this.tabClientInfo.Text = "Client Information";
            this.tabClientInfo.UseVisualStyleBackColor = true;
            // 
            // groupBox7
            // 
            this.groupBox7.Controls.Add(this.clbQuestions);
            this.groupBox7.Location = new System.Drawing.Point(10, 167);
            this.groupBox7.Name = "groupBox7";
            this.groupBox7.Size = new System.Drawing.Size(582, 389);
            this.groupBox7.TabIndex = 55;
            this.groupBox7.TabStop = false;
            this.groupBox7.Text = "Policy";
            // 
            // clbQuestions
            // 
            this.clbQuestions.CheckOnClick = true;
            this.clbQuestions.FormattingEnabled = true;
            this.clbQuestions.Location = new System.Drawing.Point(5, 16);
            this.clbQuestions.Name = "clbQuestions";
            this.clbQuestions.Size = new System.Drawing.Size(576, 394);
            this.clbQuestions.TabIndex = 1;
            this.clbQuestions.SelectedIndexChanged += new System.EventHandler(this.clbQuestions_SelectedIndexChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.label18);
            this.groupBox1.Controls.Add(this.dtpDateRequired);
            this.groupBox1.Controls.Add(this.requiredFieldLabel2);
            this.groupBox1.Controls.Add(this.txtClientCommonName);
            this.groupBox1.Controls.Add(this.lblClientCommonName);
            this.groupBox1.Controls.Add(this.dtpPeriodOfInsuranceFrom);
            this.groupBox1.Controls.Add(this.lblPeriodOfInsuranceTo);
            this.groupBox1.Controls.Add(this.dtpPeriodOfInsuranceTo);
            this.groupBox1.Controls.Add(this.lblPeriodOfInsuranceFrom);
            this.groupBox1.Controls.Add(this.txtClientName);
            this.groupBox1.Controls.Add(this.lblClientName);
            this.groupBox1.Location = new System.Drawing.Point(9, 8);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(583, 153);
            this.groupBox1.TabIndex = 53;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Insurance Information";
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Location = new System.Drawing.Point(11, 131);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(100, 13);
            this.label18.TabIndex = 56;
            this.label18.Text = "Quote Required By:";
            // 
            // dtpDateRequired
            // 
            this.dtpDateRequired.Location = new System.Drawing.Point(159, 127);
            this.dtpDateRequired.Name = "dtpDateRequired";
            this.dtpDateRequired.Size = new System.Drawing.Size(299, 20);
            this.dtpDateRequired.TabIndex = 55;
            // 
            // requiredFieldLabel2
            // 
            this.requiredFieldLabel2.AutoSize = true;
            this.requiredFieldLabel2.Field = this.groupBox1;
            this.requiredFieldLabel2.Location = new System.Drawing.Point(121, 24);
            this.requiredFieldLabel2.Name = "requiredFieldLabel2";
            this.requiredFieldLabel2.Size = new System.Drawing.Size(11, 13);
            this.requiredFieldLabel2.TabIndex = 33;
            this.requiredFieldLabel2.Text = "*";
            // 
            // txtClientCommonName
            // 
            this.txtClientCommonName.Location = new System.Drawing.Point(159, 47);
            this.txtClientCommonName.Name = "txtClientCommonName";
            this.txtClientCommonName.Size = new System.Drawing.Size(299, 20);
            this.txtClientCommonName.TabIndex = 2;
            // 
            // lblClientCommonName
            // 
            this.lblClientCommonName.Location = new System.Drawing.Point(10, 42);
            this.lblClientCommonName.Name = "lblClientCommonName";
            this.lblClientCommonName.Size = new System.Drawing.Size(144, 26);
            this.lblClientCommonName.TabIndex = 54;
            this.lblClientCommonName.Text = "Client Abbreviated name (common name):";
            // 
            // dtpPeriodOfInsuranceFrom
            // 
            this.dtpPeriodOfInsuranceFrom.Location = new System.Drawing.Point(159, 74);
            this.dtpPeriodOfInsuranceFrom.Name = "dtpPeriodOfInsuranceFrom";
            this.dtpPeriodOfInsuranceFrom.Size = new System.Drawing.Size(299, 20);
            this.dtpPeriodOfInsuranceFrom.TabIndex = 3;
            this.dtpPeriodOfInsuranceFrom.ValueChanged += new System.EventHandler(this.dtpPeriodOfInsuranceFrom_ValueChanged);
            // 
            // lblPeriodOfInsuranceTo
            // 
            this.lblPeriodOfInsuranceTo.AutoSize = true;
            this.lblPeriodOfInsuranceTo.Location = new System.Drawing.Point(11, 105);
            this.lblPeriodOfInsuranceTo.Name = "lblPeriodOfInsuranceTo";
            this.lblPeriodOfInsuranceTo.Size = new System.Drawing.Size(120, 13);
            this.lblPeriodOfInsuranceTo.TabIndex = 51;
            this.lblPeriodOfInsuranceTo.Text = "Period Of Insurance To:";
            // 
            // dtpPeriodOfInsuranceTo
            // 
            this.dtpPeriodOfInsuranceTo.Location = new System.Drawing.Point(159, 101);
            this.dtpPeriodOfInsuranceTo.Name = "dtpPeriodOfInsuranceTo";
            this.dtpPeriodOfInsuranceTo.Size = new System.Drawing.Size(299, 20);
            this.dtpPeriodOfInsuranceTo.TabIndex = 4;
            // 
            // lblPeriodOfInsuranceFrom
            // 
            this.lblPeriodOfInsuranceFrom.AutoSize = true;
            this.lblPeriodOfInsuranceFrom.Location = new System.Drawing.Point(11, 76);
            this.lblPeriodOfInsuranceFrom.Name = "lblPeriodOfInsuranceFrom";
            this.lblPeriodOfInsuranceFrom.Size = new System.Drawing.Size(130, 13);
            this.lblPeriodOfInsuranceFrom.TabIndex = 50;
            this.lblPeriodOfInsuranceFrom.Text = "Period Of Insurance From:";
            // 
            // txtClientName
            // 
            this.txtClientName.Location = new System.Drawing.Point(159, 19);
            this.txtClientName.Name = "txtClientName";
            this.txtClientName.Size = new System.Drawing.Size(299, 20);
            this.txtClientName.TabIndex = 1;
            // 
            // lblClientName
            // 
            this.lblClientName.AutoSize = true;
            this.lblClientName.Location = new System.Drawing.Point(11, 21);
            this.lblClientName.Name = "lblClientName";
            this.lblClientName.Size = new System.Drawing.Size(100, 13);
            this.lblClientName.TabIndex = 49;
            this.lblClientName.Text = "Client Name (in full):";
            // 
            // tabAccountExecutive
            // 
            this.tabAccountExecutive.Controls.Add(this.groupBox2);
            this.tabAccountExecutive.Controls.Add(this.groupBox3);
            this.tabAccountExecutive.Location = new System.Drawing.Point(4, 22);
            this.tabAccountExecutive.Name = "tabAccountExecutive";
            this.tabAccountExecutive.Size = new System.Drawing.Size(600, 584);
            this.tabAccountExecutive.TabIndex = 7;
            this.tabAccountExecutive.Text = "Account Exec";
            this.tabAccountExecutive.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox2.Controls.Add(this.lblDepartment);
            this.groupBox2.Controls.Add(this.txtExecutiveDepartment);
            this.groupBox2.Controls.Add(this.label5);
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
            this.groupBox2.Location = new System.Drawing.Point(2, 3);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(590, 400);
            this.groupBox2.TabIndex = 60;
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
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(7, 213);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(78, 13);
            this.label5.TabIndex = 53;
            this.label5.Text = "Branch Phone:";
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
            // groupBox3
            // 
            this.groupBox3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox3.Controls.Add(this.label13);
            this.groupBox3.Controls.Add(this.txtAssitantExecDepartment);
            this.groupBox3.Controls.Add(this.btnAssistantAccountExecutiveLookup);
            this.groupBox3.Controls.Add(this.label1);
            this.groupBox3.Controls.Add(this.txtAssistantExecutivePhone);
            this.groupBox3.Controls.Add(this.label2);
            this.groupBox3.Controls.Add(this.label3);
            this.groupBox3.Controls.Add(this.label4);
            this.groupBox3.Controls.Add(this.txtAssistantExecutiveEmail);
            this.groupBox3.Controls.Add(this.txtAssistantExecutiveTitle);
            this.groupBox3.Controls.Add(this.txtAssistantExecutiveName);
            this.groupBox3.Location = new System.Drawing.Point(3, 407);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(587, 174);
            this.groupBox3.TabIndex = 57;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Assistant Account Exec:";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(6, 144);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(65, 13);
            this.label13.TabIndex = 43;
            this.label13.Text = "Department:";
            // 
            // txtAssitantExecDepartment
            // 
            this.txtAssitantExecDepartment.Location = new System.Drawing.Point(135, 138);
            this.txtAssitantExecDepartment.Name = "txtAssitantExecDepartment";
            this.txtAssitantExecDepartment.Size = new System.Drawing.Size(300, 20);
            this.txtAssitantExecDepartment.TabIndex = 42;
            // 
            // btnAssistantAccountExecutiveLookup
            // 
            this.btnAssistantAccountExecutiveLookup.Location = new System.Drawing.Point(442, 12);
            this.btnAssistantAccountExecutiveLookup.Name = "btnAssistantAccountExecutiveLookup";
            this.btnAssistantAccountExecutiveLookup.Size = new System.Drawing.Size(75, 23);
            this.btnAssistantAccountExecutiveLookup.TabIndex = 2;
            this.btnAssistantAccountExecutiveLookup.Text = "Lookup";
            this.btnAssistantAccountExecutiveLookup.UseVisualStyleBackColor = true;
            this.btnAssistantAccountExecutiveLookup.Click += new System.EventHandler(this.btnAssistantAccountExecutiveLookup_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(7, 111);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(41, 13);
            this.label1.TabIndex = 37;
            this.label1.Text = "Phone:";
            // 
            // txtAssistantExecutivePhone
            // 
            this.txtAssistantExecutivePhone.Location = new System.Drawing.Point(136, 107);
            this.txtAssistantExecutivePhone.Name = "txtAssistantExecutivePhone";
            this.txtAssistantExecutivePhone.Size = new System.Drawing.Size(300, 20);
            this.txtAssistantExecutivePhone.TabIndex = 5;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 21);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(38, 13);
            this.label2.TabIndex = 35;
            this.label2.Text = "Name:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(6, 80);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(35, 13);
            this.label3.TabIndex = 34;
            this.label3.Text = "Email:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(6, 48);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(30, 13);
            this.label4.TabIndex = 33;
            this.label4.Text = "Title:";
            // 
            // txtAssistantExecutiveEmail
            // 
            this.txtAssistantExecutiveEmail.Location = new System.Drawing.Point(135, 76);
            this.txtAssistantExecutiveEmail.Name = "txtAssistantExecutiveEmail";
            this.txtAssistantExecutiveEmail.Size = new System.Drawing.Size(300, 20);
            this.txtAssistantExecutiveEmail.TabIndex = 4;
            // 
            // txtAssistantExecutiveTitle
            // 
            this.txtAssistantExecutiveTitle.Location = new System.Drawing.Point(135, 45);
            this.txtAssistantExecutiveTitle.Name = "txtAssistantExecutiveTitle";
            this.txtAssistantExecutiveTitle.Size = new System.Drawing.Size(300, 20);
            this.txtAssistantExecutiveTitle.TabIndex = 3;
            // 
            // txtAssistantExecutiveName
            // 
            this.txtAssistantExecutiveName.Location = new System.Drawing.Point(135, 14);
            this.txtAssistantExecutiveName.Name = "txtAssistantExecutiveName";
            this.txtAssistantExecutiveName.Size = new System.Drawing.Size(300, 20);
            this.txtAssistantExecutiveName.TabIndex = 1;
            // 
            // tabPolicyClass
            // 
            this.tabPolicyClass.Controls.Add(this.txtFindPolicyClass);
            this.tabPolicyClass.Controls.Add(this.label19);
            this.tabPolicyClass.Controls.Add(this.tvaPolicies);
            this.tabPolicyClass.Location = new System.Drawing.Point(4, 22);
            this.tabPolicyClass.Name = "tabPolicyClass";
            this.tabPolicyClass.Size = new System.Drawing.Size(600, 584);
            this.tabPolicyClass.TabIndex = 6;
            this.tabPolicyClass.Text = "Insurance Policy Class";
            this.tabPolicyClass.UseVisualStyleBackColor = true;
            // 
            // txtFindPolicyClass
            // 
            this.txtFindPolicyClass.Location = new System.Drawing.Point(38, 4);
            this.txtFindPolicyClass.Name = "txtFindPolicyClass";
            this.txtFindPolicyClass.Size = new System.Drawing.Size(535, 20);
            this.txtFindPolicyClass.TabIndex = 15;
            this.txtFindPolicyClass.TextChanged += new System.EventHandler(this.txtFindPolicyClass_TextChanged);
            // 
            // label19
            // 
            this.label19.AutoSize = true;
            this.label19.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label19.Location = new System.Drawing.Point(3, 7);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(35, 13);
            this.label19.TabIndex = 14;
            this.label19.Text = "Find:";
            // 
            // tvaPolicies
            // 
            this.tvaPolicies.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tvaPolicies.BackColor = System.Drawing.SystemColors.Window;
            this.tvaPolicies.Columns.Add(this.tvcPolicy);
            this.tvaPolicies.DefaultToolTipProvider = null;
            this.tvaPolicies.DragDropMarkColor = System.Drawing.Color.Black;
            this.tvaPolicies.FullRowSelect = true;
            this.tvaPolicies.LineColor = System.Drawing.SystemColors.ControlDark;
            this.tvaPolicies.Location = new System.Drawing.Point(-4, 30);
            this.tvaPolicies.Model = null;
            this.tvaPolicies.Name = "tvaPolicies";
            this.tvaPolicies.NodeControls.Add(this._checked);
            this.tvaPolicies.NodeControls.Add(this._name);
            this.tvaPolicies.SelectedNode = null;
            this.tvaPolicies.SelectionMode = Aga.Controls.Tree.TreeSelectionMode.Multi;
            this.tvaPolicies.Size = new System.Drawing.Size(608, 532);
            this.tvaPolicies.TabIndex = 6;
            this.tvaPolicies.Text = "treeViewAdv1";
            this.tvaPolicies.UseColumns = true;
            // 
            // tabQuestions
            // 
            this.tabQuestions.Controls.Add(this.label16);
            this.tabQuestions.Controls.Add(this.clbQustions);
            this.tabQuestions.Location = new System.Drawing.Point(4, 22);
            this.tabQuestions.Name = "tabQuestions";
            this.tabQuestions.Size = new System.Drawing.Size(600, 584);
            this.tabQuestions.TabIndex = 10;
            this.tabQuestions.Text = "Questions";
            this.tabQuestions.UseVisualStyleBackColor = true;
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(5, 10);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(144, 13);
            this.label16.TabIndex = 36;
            this.label16.Text = "Please select your questions:";
            // 
            // clbQustions
            // 
            this.clbQustions.CheckOnClick = true;
            this.clbQustions.FormattingEnabled = true;
            this.clbQustions.Location = new System.Drawing.Point(3, 33);
            this.clbQustions.Name = "clbQustions";
            this.clbQustions.Size = new System.Drawing.Size(594, 529);
            this.clbQustions.TabIndex = 0;
            // 
            // errorProvider
            // 
            this.errorProvider.BlinkStyle = System.Windows.Forms.ErrorBlinkStyle.NeverBlink;
            this.errorProvider.ContainerControl = this;
            // 
            // QuoteSlipWizard
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.ClientSize = new System.Drawing.Size(611, 711);
            this.Controls.Add(this.lblCoverPageTitle);
            this.Controls.Add(this.lblLogoTitle);
            this.Controls.Add(this.btnBack);
            this.Controls.Add(this.btnNext);
            this.Controls.Add(this.pbOampsLogoFull);
            this.Controls.Add(this.tbcWizardScreens);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "QuoteSlipWizard";
            this.Text = "Quotation Slip";
            this.Load += new System.EventHandler(this.QuoteSlipWizard_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pbOampsLogoFull)).EndInit();
            this.tbcWizardScreens.ResumeLayout(false);
            this.tabClientInfo.ResumeLayout(false);
            this.groupBox7.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.tabAccountExecutive.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.tabPolicyClass.ResumeLayout(false);
            this.tabPolicyClass.PerformLayout();
            this.tabQuestions.ResumeLayout(false);
            this.tabQuestions.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Aga.Controls.Tree.TreeColumn tvcPolicy;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.TextBox txtAssistantExecutivePhone;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnAssistantAccountExecutiveLookup;
        private System.Windows.Forms.TextBox txtAssitantExecDepartment;
        private Aga.Controls.Tree.TreeViewAdv tvaPolicies;
        private Aga.Controls.Tree.NodeControls.NodeCheckBox _checked;
        private Aga.Controls.Tree.NodeControls.NodeTextBox _name;
        private System.Windows.Forms.TabPage tabPolicyClass;
        
        private System.Windows.Forms.TextBox txtAssistantExecutiveName;
        private System.Windows.Forms.PictureBox pbOampsLogoFull;
        
        private System.Windows.Forms.TabPage tabClientInfo;
        private System.Windows.Forms.GroupBox groupBox1;
        private Helpers.Controls.RequiredFieldLabel requiredFieldLabel2;
        private System.Windows.Forms.TextBox txtClientCommonName;
        private System.Windows.Forms.Label lblClientCommonName;
        private System.Windows.Forms.DateTimePicker dtpPeriodOfInsuranceFrom;
        private System.Windows.Forms.Label lblPeriodOfInsuranceTo;
        private System.Windows.Forms.DateTimePicker dtpPeriodOfInsuranceTo;
        private System.Windows.Forms.Label lblPeriodOfInsuranceFrom;
        private System.Windows.Forms.TextBox txtClientName;
        private System.Windows.Forms.Label lblClientName;
        private System.Windows.Forms.TabControl tbcWizardScreens;
        private System.Windows.Forms.TabPage tabAccountExecutive;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtAssistantExecutiveEmail;
        private System.Windows.Forms.TextBox txtAssistantExecutiveTitle;
        private System.Windows.Forms.ErrorProvider errorProvider;
        private System.Windows.Forms.Label lblCoverPageTitle;
        private System.Windows.Forms.Label lblLogoTitle;
        private System.Windows.Forms.Button btnBack;
        public System.Windows.Forms.Button btnNext;
        private System.Windows.Forms.TextBox txtFindPolicyClass;
        private System.Windows.Forms.Label label19;
        private System.Windows.Forms.TabPage tabQuestions;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.GroupBox groupBox7;
        private System.Windows.Forms.CheckedListBox clbQuestions;
        private System.Windows.Forms.CheckedListBox clbQustions;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.DateTimePicker dtpDateRequired;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label lblDepartment;
        private System.Windows.Forms.TextBox txtExecutiveDepartment;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtBranchPhone;
        private System.Windows.Forms.Label lblPostalAddress2;
        private System.Windows.Forms.TextBox txtPostal2;
        private System.Windows.Forms.Label lblAddress2;
        private System.Windows.Forms.TextBox txtBranchAddress2;
        private System.Windows.Forms.Label lblFax;
        private System.Windows.Forms.Button btnAccountExecutiveLookup;
        private System.Windows.Forms.TextBox txtFax;
        private System.Windows.Forms.Label lblPostalAddress1;
        private System.Windows.Forms.TextBox txtPostal1;
        private System.Windows.Forms.Label lblAddress1;
        private System.Windows.Forms.TextBox txtBranchAddress1;
        private System.Windows.Forms.Label lblMobile;
        private System.Windows.Forms.TextBox txtExecutiveMobile;
        private System.Windows.Forms.Label lblOwnerPhone;
        private System.Windows.Forms.TextBox txtExecutivePhone;
        private System.Windows.Forms.Label lblOwnerName;
        private System.Windows.Forms.Label lblOwnerEmail;
        private System.Windows.Forms.Label lblOwnerTitle;
        private System.Windows.Forms.TextBox txtExecutiveEmail;
        private System.Windows.Forms.TextBox txtExecutiveTitle;
        private System.Windows.Forms.TextBox txtExecutiveName;
    }
}