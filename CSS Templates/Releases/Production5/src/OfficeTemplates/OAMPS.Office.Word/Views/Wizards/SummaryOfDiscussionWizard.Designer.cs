namespace OAMPS.Office.Word.Views.Wizards
{
    partial class SummaryOfDiscussionWizard
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SummaryOfDiscussionWizard));
            this.tvcRecommended = new Aga.Controls.Tree.TreeColumn();
            this.tvcCurrent = new Aga.Controls.Tree.TreeColumn();
            this.tvcPolicy = new Aga.Controls.Tree.TreeColumn();
            this._checked = new Aga.Controls.Tree.NodeControls.NodeCheckBox();
            this._recomended = new Aga.Controls.Tree.NodeControls.NodeTextBox();
            this._current = new Aga.Controls.Tree.NodeControls.NodeTextBox();
            this._name = new Aga.Controls.Tree.NodeControls.NodeTextBox();
            this.btnBack = new System.Windows.Forms.Button();
            this.lblCoverPageTitle = new System.Windows.Forms.Label();
            this.lblLogoTitle = new System.Windows.Forms.Label();
            this.btnNext = new System.Windows.Forms.Button();
            this.tbcWizardScreens = new System.Windows.Forms.TabControl();
            this.tabClient = new System.Windows.Forms.TabPage();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.rdoOther = new System.Windows.Forms.RadioButton();
            this.rdoPhone = new System.Windows.Forms.RadioButton();
            this.rdoPerson = new System.Windows.Forms.RadioButton();
            this.lblNote = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.rdoUnderWriter = new System.Windows.Forms.RadioButton();
            this.rdoCaller = new System.Windows.Forms.RadioButton();
            this.rdoCustomer = new System.Windows.Forms.RadioButton();
            this.gbClient = new System.Windows.Forms.GroupBox();
            this.txtClientContactName = new System.Windows.Forms.TextBox();
            this.lblClientCompany = new System.Windows.Forms.Label();
            this.timeDiscussion = new System.Windows.Forms.DateTimePicker();
            this.lblTimeDiscussion = new System.Windows.Forms.Label();
            this.dateDiscussion = new System.Windows.Forms.DateTimePicker();
            this.lblDateDiscussion = new System.Windows.Forms.Label();
            this.txtClientCode = new System.Windows.Forms.TextBox();
            this.lblCltCode = new System.Windows.Forms.Label();
            this.txtClientName = new System.Windows.Forms.TextBox();
            this.lblClientName = new System.Windows.Forms.Label();
            this.tabExecutive = new System.Windows.Forms.TabPage();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.lblDepartment = new System.Windows.Forms.Label();
            this.txtExecutiveDepartment = new System.Windows.Forms.TextBox();
            this.btnAccountExecutiveLookup = new System.Windows.Forms.Button();
            this.lblMobile = new System.Windows.Forms.Label();
            this.txtExecutiveMobile = new System.Windows.Forms.TextBox();
            this.lblOwnerPhone = new System.Windows.Forms.Label();
            this.txtExecutivePhone = new System.Windows.Forms.TextBox();
            this.lblOwnerName = new System.Windows.Forms.Label();
            this.lblOwnerEmail = new System.Windows.Forms.Label();
            this.txtExecutiveEmail = new System.Windows.Forms.TextBox();
            this.txtExecutiveName = new System.Windows.Forms.TextBox();
            this.pbOampsLogoFull = new System.Windows.Forms.PictureBox();
            this.errorProvider1 = new System.Windows.Forms.ErrorProvider(this.components);
            this.tbcWizardScreens.SuspendLayout();
            this.tabClient.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.gbClient.SuspendLayout();
            this.tabExecutive.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pbOampsLogoFull)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider1)).BeginInit();
            this.SuspendLayout();
            // 
            // tvcRecommended
            // 
            this.tvcRecommended.Header = "Recommended Insurer";
            this.tvcRecommended.MinColumnWidth = 140;
            this.tvcRecommended.SortOrder = System.Windows.Forms.SortOrder.None;
            this.tvcRecommended.TooltipText = null;
            this.tvcRecommended.Width = 140;
            // 
            // tvcCurrent
            // 
            this.tvcCurrent.Header = "Current Insurer";
            this.tvcCurrent.MinColumnWidth = 140;
            this.tvcCurrent.SortOrder = System.Windows.Forms.SortOrder.None;
            this.tvcCurrent.TooltipText = null;
            this.tvcCurrent.Width = 140;
            // 
            // tvcPolicy
            // 
            this.tvcPolicy.Header = "Policy Class";
            this.tvcPolicy.MinColumnWidth = 240;
            this.tvcPolicy.SortOrder = System.Windows.Forms.SortOrder.Descending;
            this.tvcPolicy.TooltipText = null;
            this.tvcPolicy.Width = 240;
            // 
            // _checked
            // 
            this._checked.DataPropertyName = "Checked";
            this._checked.EditEnabled = true;
            this._checked.LeftMargin = 0;
            this._checked.ParentColumn = this.tvcPolicy;
            // 
            // _recomended
            // 
            this._recomended.DataPropertyName = "Recommended";
            this._recomended.IncrementalSearchEnabled = true;
            this._recomended.LeftMargin = 3;
            this._recomended.ParentColumn = this.tvcRecommended;
            // 
            // _current
            // 
            this._current.DataPropertyName = "Current";
            this._current.IncrementalSearchEnabled = true;
            this._current.LeftMargin = 3;
            this._current.ParentColumn = this.tvcCurrent;
            // 
            // _name
            // 
            this._name.DataPropertyName = "Text";
            this._name.IncrementalSearchEnabled = true;
            this._name.LeftMargin = 3;
            this._name.ParentColumn = this.tvcPolicy;
            // 
            // btnBack
            // 
            this.btnBack.Enabled = false;
            this.btnBack.Location = new System.Drawing.Point(17, 610);
            this.btnBack.Name = "btnBack";
            this.btnBack.Size = new System.Drawing.Size(75, 23);
            this.btnBack.TabIndex = 64;
            this.btnBack.Text = "&Back";
            this.btnBack.UseVisualStyleBackColor = true;
            this.btnBack.Click += new System.EventHandler(this.btnBack_Click);
            // 
            // lblCoverPageTitle
            // 
            this.lblCoverPageTitle.AutoSize = true;
            this.lblCoverPageTitle.Location = new System.Drawing.Point(153, 616);
            this.lblCoverPageTitle.Name = "lblCoverPageTitle";
            this.lblCoverPageTitle.Size = new System.Drawing.Size(54, 13);
            this.lblCoverPageTitle.TabIndex = 69;
            this.lblCoverPageTitle.Text = "Dark Blue";
            this.lblCoverPageTitle.Visible = false;
            // 
            // lblLogoTitle
            // 
            this.lblLogoTitle.AutoSize = true;
            this.lblLogoTitle.Location = new System.Drawing.Point(227, 615);
            this.lblLogoTitle.Name = "lblLogoTitle";
            this.lblLogoTitle.Size = new System.Drawing.Size(176, 13);
            this.lblLogoTitle.TabIndex = 68;
            this.lblLogoTitle.Text = "Arthur J. Gallagher & Co (Aus) Limited";
            this.lblLogoTitle.Visible = false;
            // 
            // btnNext
            // 
            this.btnNext.Location = new System.Drawing.Point(508, 612);
            this.btnNext.Name = "btnNext";
            this.btnNext.Size = new System.Drawing.Size(79, 23);
            this.btnNext.TabIndex = 65;
            this.btnNext.Text = "Next";
            this.btnNext.UseVisualStyleBackColor = true;
            this.btnNext.Click += new System.EventHandler(this.btnNext_Click);
            // 
            // tbcWizardScreens
            // 
            this.tbcWizardScreens.Controls.Add(this.tabClient);
            this.tbcWizardScreens.Controls.Add(this.tabExecutive);
            this.tbcWizardScreens.Location = new System.Drawing.Point(10, 104);
            this.tbcWizardScreens.Multiline = true;
            this.tbcWizardScreens.Name = "tbcWizardScreens";
            this.tbcWizardScreens.SelectedIndex = 0;
            this.tbcWizardScreens.Size = new System.Drawing.Size(569, 477);
            this.tbcWizardScreens.TabIndex = 66;
            this.tbcWizardScreens.Tag = "Letter Information";
            // 
            // tabClient
            // 
            this.tabClient.Controls.Add(this.groupBox3);
            this.tabClient.Controls.Add(this.lblNote);
            this.tabClient.Controls.Add(this.groupBox1);
            this.tabClient.Controls.Add(this.gbClient);
            this.tabClient.Location = new System.Drawing.Point(4, 22);
            this.tabClient.Name = "tabClient";
            this.tabClient.Padding = new System.Windows.Forms.Padding(3);
            this.tabClient.Size = new System.Drawing.Size(561, 451);
            this.tabClient.TabIndex = 0;
            this.tabClient.Text = "Client Information";
            this.tabClient.UseVisualStyleBackColor = true;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.rdoOther);
            this.groupBox3.Controls.Add(this.rdoPhone);
            this.groupBox3.Controls.Add(this.rdoPerson);
            this.groupBox3.Location = new System.Drawing.Point(7, 302);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(532, 106);
            this.groupBox3.TabIndex = 74;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Please select whether the discussion was: ";
            // 
            // rdoOther
            // 
            this.rdoOther.AutoSize = true;
            this.rdoOther.Location = new System.Drawing.Point(8, 74);
            this.rdoOther.Name = "rdoOther";
            this.rdoOther.Size = new System.Drawing.Size(51, 17);
            this.rdoOther.TabIndex = 6;
            this.rdoOther.TabStop = true;
            this.rdoOther.Text = "Other";
            this.rdoOther.UseVisualStyleBackColor = true;
            // 
            // rdoPhone
            // 
            this.rdoPhone.AutoSize = true;
            this.rdoPhone.Location = new System.Drawing.Point(8, 28);
            this.rdoPhone.Name = "rdoPhone";
            this.rdoPhone.Size = new System.Drawing.Size(71, 17);
            this.rdoPhone.TabIndex = 4;
            this.rdoPhone.TabStop = true;
            this.rdoPhone.Text = "By Phone";
            this.rdoPhone.UseVisualStyleBackColor = true;
            // 
            // rdoPerson
            // 
            this.rdoPerson.AutoSize = true;
            this.rdoPerson.Location = new System.Drawing.Point(8, 51);
            this.rdoPerson.Name = "rdoPerson";
            this.rdoPerson.Size = new System.Drawing.Size(70, 17);
            this.rdoPerson.TabIndex = 5;
            this.rdoPerson.TabStop = true;
            this.rdoPerson.Text = "In Person";
            this.rdoPerson.UseVisualStyleBackColor = true;
            // 
            // lblNote
            // 
            this.lblNote.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblNote.ForeColor = System.Drawing.Color.Red;
            this.lblNote.Location = new System.Drawing.Point(12, 411);
            this.lblNote.Name = "lblNote";
            this.lblNote.Size = new System.Drawing.Size(511, 33);
            this.lblNote.TabIndex = 57;
            this.lblNote.Text = "It is critical that you save and store this file note in pdf format on the client" +
    "s file. If the file note was emailed, you must store that email too.";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.rdoUnderWriter);
            this.groupBox1.Controls.Add(this.rdoCaller);
            this.groupBox1.Controls.Add(this.rdoCustomer);
            this.groupBox1.Location = new System.Drawing.Point(6, 194);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(533, 102);
            this.groupBox1.TabIndex = 56;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Who was the discussion with?";
            // 
            // rdoUnderWriter
            // 
            this.rdoUnderWriter.AutoSize = true;
            this.rdoUnderWriter.Location = new System.Drawing.Point(9, 65);
            this.rdoUnderWriter.Name = "rdoUnderWriter";
            this.rdoUnderWriter.Size = new System.Drawing.Size(79, 17);
            this.rdoUnderWriter.TabIndex = 3;
            this.rdoUnderWriter.TabStop = true;
            this.rdoUnderWriter.Text = "Underwriter";
            this.rdoUnderWriter.UseVisualStyleBackColor = true;
            // 
            // rdoCaller
            // 
            this.rdoCaller.AutoSize = true;
            this.rdoCaller.Location = new System.Drawing.Point(9, 42);
            this.rdoCaller.Name = "rdoCaller";
            this.rdoCaller.Size = new System.Drawing.Size(51, 17);
            this.rdoCaller.TabIndex = 2;
            this.rdoCaller.TabStop = true;
            this.rdoCaller.Text = "Caller";
            this.rdoCaller.UseVisualStyleBackColor = true;
            // 
            // rdoCustomer
            // 
            this.rdoCustomer.AutoSize = true;
            this.rdoCustomer.Location = new System.Drawing.Point(9, 19);
            this.rdoCustomer.Name = "rdoCustomer";
            this.rdoCustomer.Size = new System.Drawing.Size(51, 17);
            this.rdoCustomer.TabIndex = 1;
            this.rdoCustomer.TabStop = true;
            this.rdoCustomer.Text = "Client";
            this.rdoCustomer.UseVisualStyleBackColor = true;
            // 
            // gbClient
            // 
            this.gbClient.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gbClient.Controls.Add(this.txtClientContactName);
            this.gbClient.Controls.Add(this.lblClientCompany);
            this.gbClient.Controls.Add(this.timeDiscussion);
            this.gbClient.Controls.Add(this.lblTimeDiscussion);
            this.gbClient.Controls.Add(this.dateDiscussion);
            this.gbClient.Controls.Add(this.lblDateDiscussion);
            this.gbClient.Controls.Add(this.txtClientCode);
            this.gbClient.Controls.Add(this.lblCltCode);
            this.gbClient.Controls.Add(this.txtClientName);
            this.gbClient.Controls.Add(this.lblClientName);
            this.gbClient.Location = new System.Drawing.Point(6, 6);
            this.gbClient.Name = "gbClient";
            this.gbClient.Size = new System.Drawing.Size(534, 182);
            this.gbClient.TabIndex = 55;
            this.gbClient.TabStop = false;
            this.gbClient.Text = "Client Information";
            // 
            // txtClientContactName
            // 
            this.txtClientContactName.Location = new System.Drawing.Point(126, 53);
            this.txtClientContactName.Name = "txtClientContactName";
            this.txtClientContactName.Size = new System.Drawing.Size(215, 20);
            this.txtClientContactName.TabIndex = 2;
            // 
            // lblClientCompany
            // 
            this.lblClientCompany.AutoSize = true;
            this.lblClientCompany.Location = new System.Drawing.Point(6, 59);
            this.lblClientCompany.Name = "lblClientCompany";
            this.lblClientCompany.Size = new System.Drawing.Size(107, 13);
            this.lblClientCompany.TabIndex = 72;
            this.lblClientCompany.Text = "Client Contact Name:";
            // 
            // timeDiscussion
            // 
            this.timeDiscussion.CustomFormat = "h:mm tt";
            this.timeDiscussion.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.timeDiscussion.Location = new System.Drawing.Point(136, 142);
            this.timeDiscussion.Name = "timeDiscussion";
            this.timeDiscussion.ShowUpDown = true;
            this.timeDiscussion.Size = new System.Drawing.Size(111, 20);
            this.timeDiscussion.TabIndex = 5;
            // 
            // lblTimeDiscussion
            // 
            this.lblTimeDiscussion.AutoSize = true;
            this.lblTimeDiscussion.Location = new System.Drawing.Point(6, 142);
            this.lblTimeDiscussion.Name = "lblTimeDiscussion";
            this.lblTimeDiscussion.Size = new System.Drawing.Size(97, 13);
            this.lblTimeDiscussion.TabIndex = 70;
            this.lblTimeDiscussion.Text = "Time of discussion:";
            // 
            // dateDiscussion
            // 
            this.dateDiscussion.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dateDiscussion.Location = new System.Drawing.Point(136, 117);
            this.dateDiscussion.Name = "dateDiscussion";
            this.dateDiscussion.Size = new System.Drawing.Size(111, 20);
            this.dateDiscussion.TabIndex = 4;
            // 
            // lblDateDiscussion
            // 
            this.lblDateDiscussion.AutoSize = true;
            this.lblDateDiscussion.Location = new System.Drawing.Point(6, 117);
            this.lblDateDiscussion.Name = "lblDateDiscussion";
            this.lblDateDiscussion.Size = new System.Drawing.Size(97, 13);
            this.lblDateDiscussion.TabIndex = 67;
            this.lblDateDiscussion.Text = "Date of discussion:";
            // 
            // txtClientCode
            // 
            this.txtClientCode.Location = new System.Drawing.Point(126, 79);
            this.txtClientCode.Name = "txtClientCode";
            this.txtClientCode.Size = new System.Drawing.Size(215, 20);
            this.txtClientCode.TabIndex = 3;
            // 
            // lblCltCode
            // 
            this.lblCltCode.AutoSize = true;
            this.lblCltCode.Location = new System.Drawing.Point(6, 85);
            this.lblCltCode.Name = "lblCltCode";
            this.lblCltCode.Size = new System.Drawing.Size(64, 13);
            this.lblCltCode.TabIndex = 56;
            this.lblCltCode.Text = "Client Code:";
            // 
            // txtClientName
            // 
            this.txtClientName.Location = new System.Drawing.Point(126, 27);
            this.txtClientName.Name = "txtClientName";
            this.txtClientName.Size = new System.Drawing.Size(215, 20);
            this.txtClientName.TabIndex = 1;
            // 
            // lblClientName
            // 
            this.lblClientName.AutoSize = true;
            this.lblClientName.Location = new System.Drawing.Point(6, 30);
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
            this.tabExecutive.Size = new System.Drawing.Size(561, 451);
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
            this.groupBox2.Controls.Add(this.btnAccountExecutiveLookup);
            this.groupBox2.Controls.Add(this.lblMobile);
            this.groupBox2.Controls.Add(this.txtExecutiveMobile);
            this.groupBox2.Controls.Add(this.lblOwnerPhone);
            this.groupBox2.Controls.Add(this.txtExecutivePhone);
            this.groupBox2.Controls.Add(this.lblOwnerName);
            this.groupBox2.Controls.Add(this.lblOwnerEmail);
            this.groupBox2.Controls.Add(this.txtExecutiveEmail);
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
            this.lblDepartment.Location = new System.Drawing.Point(7, 141);
            this.lblDepartment.Name = "lblDepartment";
            this.lblDepartment.Size = new System.Drawing.Size(65, 13);
            this.lblDepartment.TabIndex = 41;
            this.lblDepartment.Text = "Department:";
            // 
            // txtExecutiveDepartment
            // 
            this.txtExecutiveDepartment.Location = new System.Drawing.Point(136, 138);
            this.txtExecutiveDepartment.Name = "txtExecutiveDepartment";
            this.txtExecutiveDepartment.Size = new System.Drawing.Size(300, 20);
            this.txtExecutiveDepartment.TabIndex = 40;
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
            // lblMobile
            // 
            this.lblMobile.AutoSize = true;
            this.lblMobile.Location = new System.Drawing.Point(6, 112);
            this.lblMobile.Name = "lblMobile";
            this.lblMobile.Size = new System.Drawing.Size(41, 13);
            this.lblMobile.TabIndex = 39;
            this.lblMobile.Text = "Mobile:";
            // 
            // txtExecutiveMobile
            // 
            this.txtExecutiveMobile.Location = new System.Drawing.Point(135, 109);
            this.txtExecutiveMobile.Name = "txtExecutiveMobile";
            this.txtExecutiveMobile.Size = new System.Drawing.Size(300, 20);
            this.txtExecutiveMobile.TabIndex = 6;
            // 
            // lblOwnerPhone
            // 
            this.lblOwnerPhone.AutoSize = true;
            this.lblOwnerPhone.Location = new System.Drawing.Point(7, 79);
            this.lblOwnerPhone.Name = "lblOwnerPhone";
            this.lblOwnerPhone.Size = new System.Drawing.Size(41, 13);
            this.lblOwnerPhone.TabIndex = 37;
            this.lblOwnerPhone.Text = "Phone:";
            // 
            // txtExecutivePhone
            // 
            this.txtExecutivePhone.Location = new System.Drawing.Point(136, 76);
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
            this.lblOwnerEmail.Location = new System.Drawing.Point(6, 48);
            this.lblOwnerEmail.Name = "lblOwnerEmail";
            this.lblOwnerEmail.Size = new System.Drawing.Size(35, 13);
            this.lblOwnerEmail.TabIndex = 34;
            this.lblOwnerEmail.Text = "Email:";
            // 
            // txtExecutiveEmail
            // 
            this.txtExecutiveEmail.Location = new System.Drawing.Point(135, 45);
            this.txtExecutiveEmail.Name = "txtExecutiveEmail";
            this.txtExecutiveEmail.Size = new System.Drawing.Size(300, 20);
            this.txtExecutiveEmail.TabIndex = 4;
            // 
            // txtExecutiveName
            // 
            this.txtExecutiveName.Location = new System.Drawing.Point(135, 14);
            this.txtExecutiveName.Name = "txtExecutiveName";
            this.txtExecutiveName.Size = new System.Drawing.Size(300, 20);
            this.txtExecutiveName.TabIndex = 1;
            // 
            // pbOampsLogoFull
            // 
            this.pbOampsLogoFull.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.pbOampsLogoFull.Image = global::OAMPS.Office.Word.Properties.Resources.Gallagher;
            this.pbOampsLogoFull.InitialImage = ((System.Drawing.Image)(resources.GetObject("pbOampsLogoFull.InitialImage")));
            this.pbOampsLogoFull.Location = new System.Drawing.Point(186, 8);
            this.pbOampsLogoFull.Name = "pbOampsLogoFull";
            this.pbOampsLogoFull.Size = new System.Drawing.Size(309, 90);
            this.pbOampsLogoFull.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pbOampsLogoFull.TabIndex = 67;
            this.pbOampsLogoFull.TabStop = false;
            // 
            // errorProvider1
            // 
            this.errorProvider1.BlinkStyle = System.Windows.Forms.ErrorBlinkStyle.NeverBlink;
            this.errorProvider1.ContainerControl = this;
            // 
            // SummaryOfDiscussionWizard
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(596, 643);
            this.Controls.Add(this.btnBack);
            this.Controls.Add(this.lblCoverPageTitle);
            this.Controls.Add(this.lblLogoTitle);
            this.Controls.Add(this.btnNext);
            this.Controls.Add(this.tbcWizardScreens);
            this.Controls.Add(this.pbOampsLogoFull);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SummaryOfDiscussionWizard";
            this.Text = "File Note";
            this.Load += new System.EventHandler(this.SummaryOfDiscussionWizard_Load);
            this.tbcWizardScreens.ResumeLayout(false);
            this.tabClient.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.gbClient.ResumeLayout(false);
            this.gbClient.PerformLayout();
            this.tabExecutive.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pbOampsLogoFull)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnBack;
        private Aga.Controls.Tree.TreeColumn tvcRecommended;
        private Aga.Controls.Tree.TreeColumn tvcCurrent;
        private Aga.Controls.Tree.TreeColumn tvcPolicy;
        private Aga.Controls.Tree.NodeControls.NodeCheckBox _checked;
        private System.Windows.Forms.Label lblCoverPageTitle;
        private System.Windows.Forms.Label lblLogoTitle;
        private Aga.Controls.Tree.NodeControls.NodeTextBox _recomended;
        private Aga.Controls.Tree.NodeControls.NodeTextBox _current;
        private Aga.Controls.Tree.NodeControls.NodeTextBox _name;
        private System.Windows.Forms.Button btnNext;
        private System.Windows.Forms.PictureBox pbOampsLogoFull;
        private System.Windows.Forms.ErrorProvider errorProvider1;
        private System.Windows.Forms.TabControl tbcWizardScreens;
        private System.Windows.Forms.TabPage tabClient;
        private System.Windows.Forms.GroupBox gbClient;
        private System.Windows.Forms.TextBox txtClientCode;
        private System.Windows.Forms.Label lblCltCode;
        private System.Windows.Forms.TextBox txtClientName;
        private System.Windows.Forms.Label lblClientName;
        private System.Windows.Forms.TabPage tabExecutive;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button btnAccountExecutiveLookup;
        private System.Windows.Forms.Label lblMobile;
        private System.Windows.Forms.TextBox txtExecutiveMobile;
        private System.Windows.Forms.Label lblOwnerPhone;
        private System.Windows.Forms.TextBox txtExecutivePhone;
        private System.Windows.Forms.Label lblOwnerName;
        private System.Windows.Forms.Label lblOwnerEmail;
        private System.Windows.Forms.TextBox txtExecutiveEmail;
        private System.Windows.Forms.TextBox txtExecutiveName;
        private System.Windows.Forms.DateTimePicker dateDiscussion;
        private System.Windows.Forms.Label lblDateDiscussion;
        private System.Windows.Forms.RadioButton rdoUnderWriter;
        private System.Windows.Forms.RadioButton rdoCaller;
        private System.Windows.Forms.RadioButton rdoCustomer;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton rdoOther;
        private System.Windows.Forms.RadioButton rdoPerson;
        private System.Windows.Forms.RadioButton rdoPhone;
        private System.Windows.Forms.DateTimePicker timeDiscussion;
        private System.Windows.Forms.Label lblTimeDiscussion;
        private System.Windows.Forms.TextBox txtClientContactName;
        private System.Windows.Forms.Label lblClientCompany;
        private System.Windows.Forms.Label lblNote;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Label lblDepartment;
        private System.Windows.Forms.TextBox txtExecutiveDepartment;

    }
}