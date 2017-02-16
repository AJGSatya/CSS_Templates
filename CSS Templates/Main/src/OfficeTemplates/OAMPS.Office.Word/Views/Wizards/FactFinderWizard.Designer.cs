namespace OAMPS.Office.Word.Views.Wizards
{
    partial class FactFinderWizard
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FactFinderWizard));
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
            this.groupBox9 = new System.Windows.Forms.GroupBox();
            this.grpWholesaleRetail = new System.Windows.Forms.GroupBox();
            this.chkSatutory = new System.Windows.Forms.CheckBox();
            this.chkFSG = new System.Windows.Forms.CheckBox();
            this.groupBox7 = new System.Windows.Forms.GroupBox();
            this.rdoApprovalNo = new System.Windows.Forms.RadioButton();
            this.label21 = new System.Windows.Forms.Label();
            this.rdoApprovalYes = new System.Windows.Forms.RadioButton();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.rdoWariningNo = new System.Windows.Forms.RadioButton();
            this.rdoWariningYes = new System.Windows.Forms.RadioButton();
            this.label20 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.requiredFieldLabel2 = new OAMPS.Office.Word.Helpers.Controls.RequiredFieldLabel();
            this.txtClientCommonName = new System.Windows.Forms.TextBox();
            this.lblClientCommonName = new System.Windows.Forms.Label();
            this.dtpPeriodOfInsuranceFrom = new System.Windows.Forms.DateTimePicker();
            this.lblPeriodOfInsuranceTo = new System.Windows.Forms.Label();
            this.dtpPeriodOfInsuranceTo = new System.Windows.Forms.DateTimePicker();
            this.lblPeriodOfInsuranceFrom = new System.Windows.Forms.Label();
            this.txtClientName = new System.Windows.Forms.TextBox();
            this.lblClientName = new System.Windows.Forms.Label();
            this.groupBox8 = new System.Windows.Forms.GroupBox();
            this.label18 = new System.Windows.Forms.Label();
            this.label22 = new System.Windows.Forms.Label();
            this.chkQuestionsOnly = new System.Windows.Forms.CheckBox();
            this.tabAccountExecutive = new System.Windows.Forms.TabPage();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.lblDepartment = new System.Windows.Forms.Label();
            this.txtExecutiveDepartment = new System.Windows.Forms.TextBox();
            this.lblAddress2 = new System.Windows.Forms.Label();
            this.txtBranchAddress2 = new System.Windows.Forms.TextBox();
            this.lblAddress1 = new System.Windows.Forms.Label();
            this.txtBranchAddress1 = new System.Windows.Forms.TextBox();
            this.label17 = new System.Windows.Forms.Label();
            this.txtExecutiveMobile = new System.Windows.Forms.TextBox();
            this.btnAccountExecutiveLookup = new System.Windows.Forms.Button();
            this.lblOwnerPhone = new System.Windows.Forms.Label();
            this.txtExecutivePhone = new System.Windows.Forms.TextBox();
            this.lblOwnerName = new System.Windows.Forms.Label();
            this.lblOwnerEmail = new System.Windows.Forms.Label();
            this.lblOwnerTitle = new System.Windows.Forms.Label();
            this.txtExecutiveEmail = new System.Windows.Forms.TextBox();
            this.txtExecutiveTitle = new System.Windows.Forms.TextBox();
            this.txtExecutiveName = new System.Windows.Forms.TextBox();
            this.tabPolicyClass = new System.Windows.Forms.TabPage();
            this.txtFindPolicyClass = new System.Windows.Forms.TextBox();
            this.label19 = new System.Windows.Forms.Label();
            this.tvaPolicies = new Aga.Controls.Tree.TreeViewAdv();
            this.tabQuestions = new System.Windows.Forms.TabPage();
            this.label16 = new System.Windows.Forms.Label();
            this.clbQustions = new System.Windows.Forms.CheckedListBox();
            this.errorProvider = new System.Windows.Forms.ErrorProvider(this.components);
            this.errorProvider1 = new System.Windows.Forms.ErrorProvider(this.components);
            this.lblSpeciality = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.pbOampsLogoFull)).BeginInit();
            this.tbcWizardScreens.SuspendLayout();
            this.tabClientInfo.SuspendLayout();
            this.groupBox9.SuspendLayout();
            this.grpWholesaleRetail.SuspendLayout();
            this.groupBox7.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.groupBox8.SuspendLayout();
            this.tabAccountExecutive.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.tabPolicyClass.SuspendLayout();
            this.tabQuestions.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider1)).BeginInit();
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
            this.lblCoverPageTitle.Size = new System.Drawing.Size(90, 13);
            this.lblCoverPageTitle.TabIndex = 42;
            this.lblCoverPageTitle.Text = "Cover Woodshed";
            this.lblCoverPageTitle.Visible = false;
            // 
            // lblLogoTitle
            // 
            this.lblLogoTitle.AutoSize = true;
            this.lblLogoTitle.Location = new System.Drawing.Point(144, 667);
            this.lblLogoTitle.Name = "lblLogoTitle";
            this.lblLogoTitle.Size = new System.Drawing.Size(182, 13);
            this.lblLogoTitle.TabIndex = 41;
            this.lblLogoTitle.Text = "Arthur J. Gallagher && Co (Aus) Limited";
            this.lblLogoTitle.Visible = false;
            // 
            // btnBack
            // 
            this.btnBack.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnBack.Enabled = false;
            this.btnBack.Location = new System.Drawing.Point(11, 656);
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
            this.btnNext.Location = new System.Drawing.Point(526, 657);
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
            this.pbOampsLogoFull.Image = global::OAMPS.Office.Word.Properties.Resources.Gallagher;
            this.pbOampsLogoFull.InitialImage = ((System.Drawing.Image)(resources.GetObject("pbOampsLogoFull.InitialImage")));
            this.pbOampsLogoFull.Location = new System.Drawing.Point(165, 2);
            this.pbOampsLogoFull.Name = "pbOampsLogoFull";
            this.pbOampsLogoFull.Size = new System.Drawing.Size(246, 58);
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
            this.tbcWizardScreens.Size = new System.Drawing.Size(608, 588);
            this.tbcWizardScreens.TabIndex = 33;
            // 
            // tabClientInfo
            // 
            this.tabClientInfo.Controls.Add(this.groupBox9);
            this.tabClientInfo.Controls.Add(this.groupBox8);
            this.tabClientInfo.Location = new System.Drawing.Point(4, 22);
            this.tabClientInfo.Name = "tabClientInfo";
            this.tabClientInfo.Padding = new System.Windows.Forms.Padding(3);
            this.tabClientInfo.Size = new System.Drawing.Size(600, 562);
            this.tabClientInfo.TabIndex = 5;
            this.tabClientInfo.Text = "Client Information";
            this.tabClientInfo.UseVisualStyleBackColor = true;
            // 
            // groupBox9
            // 
            this.groupBox9.Controls.Add(this.grpWholesaleRetail);
            this.groupBox9.Controls.Add(this.groupBox7);
            this.groupBox9.Controls.Add(this.groupBox4);
            this.groupBox9.Controls.Add(this.groupBox1);
            this.groupBox9.Location = new System.Drawing.Point(8, 109);
            this.groupBox9.Name = "groupBox9";
            this.groupBox9.Size = new System.Drawing.Size(584, 469);
            this.groupBox9.TabIndex = 60;
            this.groupBox9.TabStop = false;
            this.groupBox9.Text = "Full Insurance Fact Finder Document";
            // 
            // grpWholesaleRetail
            // 
            this.grpWholesaleRetail.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.grpWholesaleRetail.Controls.Add(this.chkSatutory);
            this.grpWholesaleRetail.Controls.Add(this.chkFSG);
            this.grpWholesaleRetail.Location = new System.Drawing.Point(7, 300);
            this.grpWholesaleRetail.Name = "grpWholesaleRetail";
            this.grpWholesaleRetail.Size = new System.Drawing.Size(565, 78);
            this.grpWholesaleRetail.TabIndex = 61;
            this.grpWholesaleRetail.TabStop = false;
            this.grpWholesaleRetail.Text = "Include the appropriate statutory information";
            // 
            // chkSatutory
            // 
            this.chkSatutory.AutoSize = true;
            this.chkSatutory.Checked = true;
            this.chkSatutory.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkSatutory.Enabled = false;
            this.chkSatutory.Location = new System.Drawing.Point(13, 19);
            this.chkSatutory.Name = "chkSatutory";
            this.chkSatutory.Size = new System.Drawing.Size(109, 17);
            this.chkSatutory.TabIndex = 69;
            this.chkSatutory.Text = "Important Notices";
            this.chkSatutory.UseVisualStyleBackColor = true;
            // 
            // chkFSG
            // 
            this.chkFSG.AutoSize = true;
            this.chkFSG.Location = new System.Drawing.Point(13, 42);
            this.chkFSG.Name = "chkFSG";
            this.chkFSG.Size = new System.Drawing.Size(143, 17);
            this.chkFSG.TabIndex = 68;
            this.chkFSG.Text = "Financial Services Guide";
            this.chkFSG.UseVisualStyleBackColor = true;
            // 
            // groupBox7
            // 
            this.groupBox7.Controls.Add(this.rdoApprovalNo);
            this.groupBox7.Controls.Add(this.label21);
            this.groupBox7.Controls.Add(this.rdoApprovalYes);
            this.groupBox7.Location = new System.Drawing.Point(7, 232);
            this.groupBox7.Name = "groupBox7";
            this.groupBox7.Size = new System.Drawing.Size(565, 53);
            this.groupBox7.TabIndex = 60;
            this.groupBox7.TabStop = false;
            this.groupBox7.Text = "Declaration / Proposal Forms";
            // 
            // rdoApprovalNo
            // 
            this.rdoApprovalNo.AutoSize = true;
            this.rdoApprovalNo.Location = new System.Drawing.Point(522, 22);
            this.rdoApprovalNo.Name = "rdoApprovalNo";
            this.rdoApprovalNo.Size = new System.Drawing.Size(39, 17);
            this.rdoApprovalNo.TabIndex = 5;
            this.rdoApprovalNo.Text = "No";
            this.rdoApprovalNo.UseVisualStyleBackColor = true;
            this.rdoApprovalNo.CheckedChanged += new System.EventHandler(this.rdoApprovalNo_CheckedChanged);
            // 
            // label21
            // 
            this.label21.AutoSize = true;
            this.label21.Location = new System.Drawing.Point(10, 24);
            this.label21.Name = "label21";
            this.label21.Size = new System.Drawing.Size(435, 13);
            this.label21.TabIndex = 1;
            this.label21.Text = "Do you have policies where declaration / proposal forms are to be completed by th" +
    "e client?";
            // 
            // rdoApprovalYes
            // 
            this.rdoApprovalYes.AutoSize = true;
            this.rdoApprovalYes.Checked = true;
            this.rdoApprovalYes.Location = new System.Drawing.Point(473, 22);
            this.rdoApprovalYes.Name = "rdoApprovalYes";
            this.rdoApprovalYes.Size = new System.Drawing.Size(43, 17);
            this.rdoApprovalYes.TabIndex = 4;
            this.rdoApprovalYes.TabStop = true;
            this.rdoApprovalYes.Text = "Yes";
            this.rdoApprovalYes.UseVisualStyleBackColor = true;
            this.rdoApprovalYes.CheckedChanged += new System.EventHandler(this.rdoApprovalYes_CheckedChanged);
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.rdoWariningNo);
            this.groupBox4.Controls.Add(this.rdoWariningYes);
            this.groupBox4.Controls.Add(this.label20);
            this.groupBox4.Location = new System.Drawing.Point(7, 171);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(565, 50);
            this.groupBox4.TabIndex = 59;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Claims Made Warning";
            // 
            // rdoWariningNo
            // 
            this.rdoWariningNo.AutoSize = true;
            this.rdoWariningNo.Location = new System.Drawing.Point(520, 19);
            this.rdoWariningNo.Name = "rdoWariningNo";
            this.rdoWariningNo.Size = new System.Drawing.Size(39, 17);
            this.rdoWariningNo.TabIndex = 3;
            this.rdoWariningNo.Text = "No";
            this.rdoWariningNo.UseVisualStyleBackColor = true;
            this.rdoWariningNo.CheckedChanged += new System.EventHandler(this.rdoWariningNo_CheckedChanged);
            // 
            // rdoWariningYes
            // 
            this.rdoWariningYes.AutoSize = true;
            this.rdoWariningYes.Checked = true;
            this.rdoWariningYes.Location = new System.Drawing.Point(473, 19);
            this.rdoWariningYes.Name = "rdoWariningYes";
            this.rdoWariningYes.Size = new System.Drawing.Size(43, 17);
            this.rdoWariningYes.TabIndex = 2;
            this.rdoWariningYes.TabStop = true;
            this.rdoWariningYes.Text = "Yes";
            this.rdoWariningYes.UseVisualStyleBackColor = true;
            this.rdoWariningYes.CheckedChanged += new System.EventHandler(this.rdoWariningYes_CheckedChanged);
            // 
            // label20
            // 
            this.label20.AutoSize = true;
            this.label20.Location = new System.Drawing.Point(10, 20);
            this.label20.Name = "label20";
            this.label20.Size = new System.Drawing.Size(207, 13);
            this.label20.TabIndex = 0;
            this.label20.Text = "Do you require the claims made warning?  ";
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.requiredFieldLabel2);
            this.groupBox1.Controls.Add(this.txtClientCommonName);
            this.groupBox1.Controls.Add(this.lblClientCommonName);
            this.groupBox1.Controls.Add(this.dtpPeriodOfInsuranceFrom);
            this.groupBox1.Controls.Add(this.lblPeriodOfInsuranceTo);
            this.groupBox1.Controls.Add(this.dtpPeriodOfInsuranceTo);
            this.groupBox1.Controls.Add(this.lblPeriodOfInsuranceFrom);
            this.groupBox1.Controls.Add(this.txtClientName);
            this.groupBox1.Controls.Add(this.lblClientName);
            this.groupBox1.Location = new System.Drawing.Point(6, 29);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(566, 132);
            this.groupBox1.TabIndex = 58;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Insurance Information";
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
            this.dtpPeriodOfInsuranceTo.ValueChanged += new System.EventHandler(this.dtpPeriodOfInsuranceTo_ValueChanged);
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
            // groupBox8
            // 
            this.groupBox8.Controls.Add(this.label18);
            this.groupBox8.Controls.Add(this.label22);
            this.groupBox8.Controls.Add(this.chkQuestionsOnly);
            this.groupBox8.Location = new System.Drawing.Point(8, 13);
            this.groupBox8.Name = "groupBox8";
            this.groupBox8.Size = new System.Drawing.Size(572, 53);
            this.groupBox8.TabIndex = 59;
            this.groupBox8.TabStop = false;
            this.groupBox8.Text = "Fact Finder Questions Only";
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label18.Location = new System.Drawing.Point(162, 23);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(41, 13);
            this.label18.TabIndex = 60;
            this.label18.Text = "without";
            // 
            // label22
            // 
            this.label22.AutoSize = true;
            this.label22.Location = new System.Drawing.Point(201, 23);
            this.label22.Name = "label22";
            this.label22.Size = new System.Drawing.Size(143, 13);
            this.label22.TabIndex = 61;
            this.label22.Text = "the encompassing document";
            // 
            // chkQuestionsOnly
            // 
            this.chkQuestionsOnly.AutoSize = true;
            this.chkQuestionsOnly.Location = new System.Drawing.Point(13, 22);
            this.chkQuestionsOnly.Name = "chkQuestionsOnly";
            this.chkQuestionsOnly.Size = new System.Drawing.Size(159, 17);
            this.chkQuestionsOnly.TabIndex = 59;
            this.chkQuestionsOnly.Text = "Create individual fact finders";
            this.chkQuestionsOnly.UseVisualStyleBackColor = true;
            this.chkQuestionsOnly.CheckedChanged += new System.EventHandler(this.chkQuestionsOnly_CheckedChanged);
            // 
            // tabAccountExecutive
            // 
            this.tabAccountExecutive.Controls.Add(this.groupBox2);
            this.tabAccountExecutive.Location = new System.Drawing.Point(4, 22);
            this.tabAccountExecutive.Name = "tabAccountExecutive";
            this.tabAccountExecutive.Size = new System.Drawing.Size(600, 562);
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
            this.groupBox2.Controls.Add(this.lblAddress2);
            this.groupBox2.Controls.Add(this.txtBranchAddress2);
            this.groupBox2.Controls.Add(this.lblAddress1);
            this.groupBox2.Controls.Add(this.txtBranchAddress1);
            this.groupBox2.Controls.Add(this.label17);
            this.groupBox2.Controls.Add(this.txtExecutiveMobile);
            this.groupBox2.Controls.Add(this.btnAccountExecutiveLookup);
            this.groupBox2.Controls.Add(this.lblOwnerPhone);
            this.groupBox2.Controls.Add(this.txtExecutivePhone);
            this.groupBox2.Controls.Add(this.lblOwnerName);
            this.groupBox2.Controls.Add(this.lblOwnerEmail);
            this.groupBox2.Controls.Add(this.lblOwnerTitle);
            this.groupBox2.Controls.Add(this.txtExecutiveEmail);
            this.groupBox2.Controls.Add(this.txtExecutiveTitle);
            this.groupBox2.Controls.Add(this.txtExecutiveName);
            this.groupBox2.Location = new System.Drawing.Point(3, 6);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(587, 266);
            this.groupBox2.TabIndex = 55;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Account Exec:";
            // 
            // lblDepartment
            // 
            this.lblDepartment.AutoSize = true;
            this.lblDepartment.Location = new System.Drawing.Point(7, 233);
            this.lblDepartment.Name = "lblDepartment";
            this.lblDepartment.Size = new System.Drawing.Size(65, 13);
            this.lblDepartment.TabIndex = 55;
            this.lblDepartment.Text = "Department:";
            // 
            // txtExecutiveDepartment
            // 
            this.txtExecutiveDepartment.Location = new System.Drawing.Point(136, 230);
            this.txtExecutiveDepartment.Name = "txtExecutiveDepartment";
            this.txtExecutiveDepartment.Size = new System.Drawing.Size(300, 20);
            this.txtExecutiveDepartment.TabIndex = 54;
            // 
            // lblAddress2
            // 
            this.lblAddress2.AutoSize = true;
            this.lblAddress2.Location = new System.Drawing.Point(7, 203);
            this.lblAddress2.Name = "lblAddress2";
            this.lblAddress2.Size = new System.Drawing.Size(91, 13);
            this.lblAddress2.TabIndex = 53;
            this.lblAddress2.Text = "Branch Address2:";
            // 
            // txtBranchAddress2
            // 
            this.txtBranchAddress2.Location = new System.Drawing.Point(136, 200);
            this.txtBranchAddress2.Name = "txtBranchAddress2";
            this.txtBranchAddress2.Size = new System.Drawing.Size(300, 20);
            this.txtBranchAddress2.TabIndex = 8;
            // 
            // lblAddress1
            // 
            this.lblAddress1.AutoSize = true;
            this.lblAddress1.Location = new System.Drawing.Point(7, 172);
            this.lblAddress1.Name = "lblAddress1";
            this.lblAddress1.Size = new System.Drawing.Size(91, 13);
            this.lblAddress1.TabIndex = 52;
            this.lblAddress1.Text = "Branch Address1:";
            // 
            // txtBranchAddress1
            // 
            this.txtBranchAddress1.Location = new System.Drawing.Point(136, 169);
            this.txtBranchAddress1.Name = "txtBranchAddress1";
            this.txtBranchAddress1.Size = new System.Drawing.Size(300, 20);
            this.txtBranchAddress1.TabIndex = 7;
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(6, 144);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(41, 13);
            this.label17.TabIndex = 39;
            this.label17.Text = "Mobile:";
            // 
            // txtExecutiveMobile
            // 
            this.txtExecutiveMobile.Location = new System.Drawing.Point(135, 141);
            this.txtExecutiveMobile.Name = "txtExecutiveMobile";
            this.txtExecutiveMobile.Size = new System.Drawing.Size(300, 20);
            this.txtExecutiveMobile.TabIndex = 6;
            // 
            // btnAccountExecutiveLookup
            // 
            this.btnAccountExecutiveLookup.Location = new System.Drawing.Point(442, 14);
            this.btnAccountExecutiveLookup.Name = "btnAccountExecutiveLookup";
            this.btnAccountExecutiveLookup.Size = new System.Drawing.Size(75, 23);
            this.btnAccountExecutiveLookup.TabIndex = 2;
            this.btnAccountExecutiveLookup.Text = "Lookup";
            this.btnAccountExecutiveLookup.UseVisualStyleBackColor = true;
            this.btnAccountExecutiveLookup.Click += new System.EventHandler(this.btnAccountExecutiveLookup_Click);
            // 
            // lblOwnerPhone
            // 
            this.lblOwnerPhone.AutoSize = true;
            this.lblOwnerPhone.Location = new System.Drawing.Point(7, 111);
            this.lblOwnerPhone.Name = "lblOwnerPhone";
            this.lblOwnerPhone.Size = new System.Drawing.Size(41, 13);
            this.lblOwnerPhone.TabIndex = 37;
            this.lblOwnerPhone.Text = "Phone:";
            // 
            // txtExecutivePhone
            // 
            this.txtExecutivePhone.Location = new System.Drawing.Point(136, 108);
            this.txtExecutivePhone.Name = "txtExecutivePhone";
            this.txtExecutivePhone.Size = new System.Drawing.Size(300, 20);
            this.txtExecutivePhone.TabIndex = 5;
            // 
            // lblOwnerName
            // 
            this.lblOwnerName.AutoSize = true;
            this.lblOwnerName.Location = new System.Drawing.Point(6, 21);
            this.lblOwnerName.Name = "lblOwnerName";
            this.lblOwnerName.Size = new System.Drawing.Size(38, 13);
            this.lblOwnerName.TabIndex = 35;
            this.lblOwnerName.Text = "Name:";
            // 
            // lblOwnerEmail
            // 
            this.lblOwnerEmail.AutoSize = true;
            this.lblOwnerEmail.Location = new System.Drawing.Point(6, 80);
            this.lblOwnerEmail.Name = "lblOwnerEmail";
            this.lblOwnerEmail.Size = new System.Drawing.Size(35, 13);
            this.lblOwnerEmail.TabIndex = 34;
            this.lblOwnerEmail.Text = "Email:";
            // 
            // lblOwnerTitle
            // 
            this.lblOwnerTitle.AutoSize = true;
            this.lblOwnerTitle.Location = new System.Drawing.Point(6, 48);
            this.lblOwnerTitle.Name = "lblOwnerTitle";
            this.lblOwnerTitle.Size = new System.Drawing.Size(30, 13);
            this.lblOwnerTitle.TabIndex = 33;
            this.lblOwnerTitle.Text = "Title:";
            // 
            // txtExecutiveEmail
            // 
            this.txtExecutiveEmail.Location = new System.Drawing.Point(135, 77);
            this.txtExecutiveEmail.Name = "txtExecutiveEmail";
            this.txtExecutiveEmail.Size = new System.Drawing.Size(300, 20);
            this.txtExecutiveEmail.TabIndex = 4;
            // 
            // txtExecutiveTitle
            // 
            this.txtExecutiveTitle.Location = new System.Drawing.Point(135, 45);
            this.txtExecutiveTitle.Name = "txtExecutiveTitle";
            this.txtExecutiveTitle.Size = new System.Drawing.Size(300, 20);
            this.txtExecutiveTitle.TabIndex = 3;
            // 
            // txtExecutiveName
            // 
            this.txtExecutiveName.Location = new System.Drawing.Point(135, 14);
            this.txtExecutiveName.Name = "txtExecutiveName";
            this.txtExecutiveName.Size = new System.Drawing.Size(300, 20);
            this.txtExecutiveName.TabIndex = 1;
            // 
            // tabPolicyClass
            // 
            this.tabPolicyClass.Controls.Add(this.txtFindPolicyClass);
            this.tabPolicyClass.Controls.Add(this.label19);
            this.tabPolicyClass.Controls.Add(this.tvaPolicies);
            this.tabPolicyClass.Location = new System.Drawing.Point(4, 22);
            this.tabPolicyClass.Name = "tabPolicyClass";
            this.tabPolicyClass.Size = new System.Drawing.Size(600, 562);
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
            this.tabQuestions.Size = new System.Drawing.Size(600, 562);
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
            // errorProvider1
            // 
            this.errorProvider1.BlinkStyle = System.Windows.Forms.ErrorBlinkStyle.NeverBlink;
            this.errorProvider1.ContainerControl = this;
            // 
            // lblSpeciality
            // 
            this.lblSpeciality.AutoSize = true;
            this.lblSpeciality.Location = new System.Drawing.Point(305, 338);
            this.lblSpeciality.Name = "lblSpeciality";
            this.lblSpeciality.Size = new System.Drawing.Size(0, 13);
            this.lblSpeciality.TabIndex = 43;
            this.lblSpeciality.Visible = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(313, 346);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(0, 13);
            this.label1.TabIndex = 44;
            // 
            // FactFinderWizard
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Window;
            this.ClientSize = new System.Drawing.Size(611, 689);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lblSpeciality);
            this.Controls.Add(this.lblCoverPageTitle);
            this.Controls.Add(this.lblLogoTitle);
            this.Controls.Add(this.btnBack);
            this.Controls.Add(this.btnNext);
            this.Controls.Add(this.pbOampsLogoFull);
            this.Controls.Add(this.tbcWizardScreens);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FactFinderWizard";
            this.Text = "Insurance Factfinder";
            this.Load += new System.EventHandler(this.PreRenewalQuestionareWizard_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pbOampsLogoFull)).EndInit();
            this.tbcWizardScreens.ResumeLayout(false);
            this.tabClientInfo.ResumeLayout(false);
            this.groupBox9.ResumeLayout(false);
            this.grpWholesaleRetail.ResumeLayout(false);
            this.grpWholesaleRetail.PerformLayout();
            this.groupBox7.ResumeLayout(false);
            this.groupBox7.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox8.ResumeLayout(false);
            this.groupBox8.PerformLayout();
            this.tabAccountExecutive.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.tabPolicyClass.ResumeLayout(false);
            this.tabPolicyClass.PerformLayout();
            this.tabQuestions.ResumeLayout(false);
            this.tabQuestions.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Aga.Controls.Tree.TreeColumn tvcPolicy;
        private System.Windows.Forms.Label lblAddress1;
        private System.Windows.Forms.TextBox txtBranchAddress1;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.Button btnAccountExecutiveLookup;
        private System.Windows.Forms.TextBox txtExecutiveDepartment;
        private System.Windows.Forms.TextBox txtBranchAddress2;
        private System.Windows.Forms.Label lblAddress2;
        private System.Windows.Forms.TextBox txtExecutiveMobile;
        private System.Windows.Forms.Label lblOwnerPhone;
        private System.Windows.Forms.TextBox txtExecutivePhone;
        private System.Windows.Forms.Label lblOwnerName;
        private System.Windows.Forms.Label lblOwnerEmail;
        private System.Windows.Forms.Label lblOwnerTitle;
        private System.Windows.Forms.TextBox txtExecutiveEmail;
        private System.Windows.Forms.TextBox txtExecutiveTitle;
        private System.Windows.Forms.TextBox txtExecutiveName;
        private Aga.Controls.Tree.TreeViewAdv tvaPolicies;
        private Aga.Controls.Tree.NodeControls.NodeCheckBox _checked;
        private Aga.Controls.Tree.NodeControls.NodeTextBox _name;
        private System.Windows.Forms.TabPage tabPolicyClass;
        private System.Windows.Forms.Label lblDepartment;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.PictureBox pbOampsLogoFull;

        private System.Windows.Forms.TabPage tabClientInfo;
        private System.Windows.Forms.TabControl tbcWizardScreens;
        private System.Windows.Forms.TabPage tabAccountExecutive;
        private System.Windows.Forms.ErrorProvider errorProvider;
        private System.Windows.Forms.Label lblCoverPageTitle;
        private System.Windows.Forms.Label lblLogoTitle;
        private System.Windows.Forms.Button btnBack;
        public System.Windows.Forms.Button btnNext;
        private System.Windows.Forms.TextBox txtFindPolicyClass;
        private System.Windows.Forms.Label label19;
        private System.Windows.Forms.TabPage tabQuestions;
        private System.Windows.Forms.CheckedListBox clbQustions;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.ErrorProvider errorProvider1;
        private System.Windows.Forms.GroupBox groupBox8;
        private System.Windows.Forms.CheckBox chkQuestionsOnly;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.Label label22;
        private System.Windows.Forms.GroupBox groupBox9;
        private System.Windows.Forms.GroupBox grpWholesaleRetail;
        private System.Windows.Forms.CheckBox chkSatutory;
        private System.Windows.Forms.CheckBox chkFSG;
        private System.Windows.Forms.GroupBox groupBox7;
        private System.Windows.Forms.RadioButton rdoApprovalNo;
        private System.Windows.Forms.Label label21;
        private System.Windows.Forms.RadioButton rdoApprovalYes;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.RadioButton rdoWariningNo;
        private System.Windows.Forms.RadioButton rdoWariningYes;
        private System.Windows.Forms.Label label20;
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
        private System.Windows.Forms.Label lblSpeciality;
        private System.Windows.Forms.Label label1;
    }
}