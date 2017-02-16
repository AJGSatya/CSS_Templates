namespace OAMPS.Office.Word.Views.Wizards
{
    partial class ClientDiscoveryWizard
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ClientDiscoveryWizard));
            this.btnBack = new System.Windows.Forms.Button();
            this.btnNext = new System.Windows.Forms.Button();
            this.lblCoverPageTitle = new System.Windows.Forms.Label();
            this.lblLogoTitle = new System.Windows.Forms.Label();
            this.tbcWizardScreens = new System.Windows.Forms.TabControl();
            this.tbClient = new System.Windows.Forms.TabPage();
            this.gbClient = new System.Windows.Forms.GroupBox();
            this.datePrepared = new System.Windows.Forms.DateTimePicker();
            this.label1 = new System.Windows.Forms.Label();
            this.txtCltiBAIS = new System.Windows.Forms.TextBox();
            this.lblCltiBAIS = new System.Windows.Forms.Label();
            this.dateDiscussion = new System.Windows.Forms.DateTimePicker();
            this.lblDiscussion = new System.Windows.Forms.Label();
            this.lblCltContactName = new System.Windows.Forms.Label();
            this.txtClientContactName = new System.Windows.Forms.TextBox();
            this.txtClientName = new System.Windows.Forms.TextBox();
            this.lblClientName = new System.Windows.Forms.Label();
            this.tbClientExecutive = new System.Windows.Forms.TabPage();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label17 = new System.Windows.Forms.Label();
            this.txtExecutiveMobile = new System.Windows.Forms.TextBox();
            this.lblDepartment = new System.Windows.Forms.Label();
            this.txtExecutiveDepartment = new System.Windows.Forms.TextBox();
            this.lblAddress2 = new System.Windows.Forms.Label();
            this.txtBranchAddress2 = new System.Windows.Forms.TextBox();
            this.btnAccountExecutiveLookup = new System.Windows.Forms.Button();
            this.lblAddress1 = new System.Windows.Forms.Label();
            this.txtBranchAddress1 = new System.Windows.Forms.TextBox();
            this.lblOwnerPhone = new System.Windows.Forms.Label();
            this.txtExecutivePhone = new System.Windows.Forms.TextBox();
            this.lblOwnerName = new System.Windows.Forms.Label();
            this.lblOwnerEmail = new System.Windows.Forms.Label();
            this.lblOwnerTitle = new System.Windows.Forms.Label();
            this.txtExecutiveEmail = new System.Windows.Forms.TextBox();
            this.txtExecutiveTitle = new System.Windows.Forms.TextBox();
            this.txtExecutiveName = new System.Windows.Forms.TextBox();
            this.tbQuestions = new System.Windows.Forms.TabPage();
            this.chkSelectAll = new System.Windows.Forms.CheckBox();
            this.tvaQuestions = new Aga.Controls.Tree.TreeViewAdv();
            this._checked = new Aga.Controls.Tree.NodeControls.NodeCheckBox();
            this._name = new Aga.Controls.Tree.NodeControls.NodeTextBox();
            this.pbOampsLogoFull = new System.Windows.Forms.PictureBox();
            this.tbcWizardScreens.SuspendLayout();
            this.tbClient.SuspendLayout();
            this.gbClient.SuspendLayout();
            this.tbClientExecutive.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.tbQuestions.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pbOampsLogoFull)).BeginInit();
            this.SuspendLayout();
            // 
            // btnBack
            // 
            this.btnBack.Enabled = false;
            this.btnBack.Location = new System.Drawing.Point(3, 539);
            this.btnBack.Name = "btnBack";
            this.btnBack.Size = new System.Drawing.Size(75, 23);
            this.btnBack.TabIndex = 66;
            this.btnBack.Text = "&Back";
            this.btnBack.UseVisualStyleBackColor = true;
            this.btnBack.Click += new System.EventHandler(this.btnBack_Click);
            // 
            // btnNext
            // 
            this.btnNext.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.btnNext.Location = new System.Drawing.Point(494, 541);
            this.btnNext.Name = "btnNext";
            this.btnNext.Size = new System.Drawing.Size(79, 23);
            this.btnNext.TabIndex = 67;
            this.btnNext.Text = "Next";
            this.btnNext.UseVisualStyleBackColor = true;
            this.btnNext.Click += new System.EventHandler(this.btnNext_Click);
            // 
            // lblCoverPageTitle
            // 
            this.lblCoverPageTitle.AutoSize = true;
            this.lblCoverPageTitle.Location = new System.Drawing.Point(146, 549);
            this.lblCoverPageTitle.Name = "lblCoverPageTitle";
            this.lblCoverPageTitle.Size = new System.Drawing.Size(68, 13);
            this.lblCoverPageTitle.TabIndex = 65;
            this.lblCoverPageTitle.Text = "Boulder Opal";
            this.lblCoverPageTitle.Visible = false;
            // 
            // lblLogoTitle
            // 
            this.lblLogoTitle.AutoSize = true;
            this.lblLogoTitle.Location = new System.Drawing.Point(221, 549);
            this.lblLogoTitle.Name = "lblLogoTitle";
            this.lblLogoTitle.Size = new System.Drawing.Size(152, 13);
            this.lblLogoTitle.TabIndex = 64;
            this.lblLogoTitle.Text = "OAMPS Insurance Brokers Ltd";
            this.lblLogoTitle.Visible = false;
            // 
            // tbcWizardScreens
            // 
            this.tbcWizardScreens.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.tbcWizardScreens.Controls.Add(this.tbClient);
            this.tbcWizardScreens.Controls.Add(this.tbClientExecutive);
            this.tbcWizardScreens.Controls.Add(this.tbQuestions);
            this.tbcWizardScreens.Location = new System.Drawing.Point(3, 108);
            this.tbcWizardScreens.Name = "tbcWizardScreens";
            this.tbcWizardScreens.SelectedIndex = 0;
            this.tbcWizardScreens.Size = new System.Drawing.Size(578, 411);
            this.tbcWizardScreens.TabIndex = 61;
            // 
            // tbClient
            // 
            this.tbClient.Controls.Add(this.gbClient);
            this.tbClient.Location = new System.Drawing.Point(4, 22);
            this.tbClient.Name = "tbClient";
            this.tbClient.Padding = new System.Windows.Forms.Padding(3);
            this.tbClient.Size = new System.Drawing.Size(570, 385);
            this.tbClient.TabIndex = 0;
            this.tbClient.Text = "Client Information";
            this.tbClient.UseVisualStyleBackColor = true;
            // 
            // gbClient
            // 
            this.gbClient.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gbClient.Controls.Add(this.datePrepared);
            this.gbClient.Controls.Add(this.label1);
            this.gbClient.Controls.Add(this.txtCltiBAIS);
            this.gbClient.Controls.Add(this.lblCltiBAIS);
            this.gbClient.Controls.Add(this.dateDiscussion);
            this.gbClient.Controls.Add(this.lblDiscussion);
            this.gbClient.Controls.Add(this.lblCltContactName);
            this.gbClient.Controls.Add(this.txtClientContactName);
            this.gbClient.Controls.Add(this.txtClientName);
            this.gbClient.Controls.Add(this.lblClientName);
            this.gbClient.Location = new System.Drawing.Point(6, 17);
            this.gbClient.Name = "gbClient";
            this.gbClient.Size = new System.Drawing.Size(556, 196);
            this.gbClient.TabIndex = 56;
            this.gbClient.TabStop = false;
            this.gbClient.Text = "Client Information";
            // 
            // datePrepared
            // 
            this.datePrepared.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.datePrepared.Location = new System.Drawing.Point(166, 145);
            this.datePrepared.Name = "datePrepared";
            this.datePrepared.Size = new System.Drawing.Size(121, 20);
            this.datePrepared.TabIndex = 68;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 149);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(79, 13);
            this.label1.TabIndex = 69;
            this.label1.Text = "Date Prepared:";
            // 
            // txtCltiBAIS
            // 
            this.txtCltiBAIS.Location = new System.Drawing.Point(136, 19);
            this.txtCltiBAIS.Name = "txtCltiBAIS";
            this.txtCltiBAIS.Size = new System.Drawing.Size(299, 20);
            this.txtCltiBAIS.TabIndex = 66;
            // 
            // lblCltiBAIS
            // 
            this.lblCltiBAIS.AutoSize = true;
            this.lblCltiBAIS.Location = new System.Drawing.Point(6, 23);
            this.lblCltiBAIS.Name = "lblCltiBAIS";
            this.lblCltiBAIS.Size = new System.Drawing.Size(64, 13);
            this.lblCltiBAIS.TabIndex = 67;
            this.lblCltiBAIS.Text = "Client Code:";
            // 
            // dateDiscussion
            // 
            this.dateDiscussion.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dateDiscussion.Location = new System.Drawing.Point(166, 110);
            this.dateDiscussion.Name = "dateDiscussion";
            this.dateDiscussion.Size = new System.Drawing.Size(121, 20);
            this.dateDiscussion.TabIndex = 6;
            // 
            // lblDiscussion
            // 
            this.lblDiscussion.AutoSize = true;
            this.lblDiscussion.Location = new System.Drawing.Point(6, 114);
            this.lblDiscussion.Name = "lblDiscussion";
            this.lblDiscussion.Size = new System.Drawing.Size(149, 13);
            this.lblDiscussion.TabIndex = 65;
            this.lblDiscussion.Text = "Date of Discovery Discussion:";
            // 
            // lblCltContactName
            // 
            this.lblCltContactName.AutoSize = true;
            this.lblCltContactName.Location = new System.Drawing.Point(6, 75);
            this.lblCltContactName.Name = "lblCltContactName";
            this.lblCltContactName.Size = new System.Drawing.Size(107, 13);
            this.lblCltContactName.TabIndex = 60;
            this.lblCltContactName.Text = "Client Contact Name:";
            // 
            // txtClientContactName
            // 
            this.txtClientContactName.Location = new System.Drawing.Point(136, 71);
            this.txtClientContactName.Name = "txtClientContactName";
            this.txtClientContactName.Size = new System.Drawing.Size(299, 20);
            this.txtClientContactName.TabIndex = 2;
            // 
            // txtClientName
            // 
            this.txtClientName.Location = new System.Drawing.Point(136, 45);
            this.txtClientName.Name = "txtClientName";
            this.txtClientName.Size = new System.Drawing.Size(299, 20);
            this.txtClientName.TabIndex = 0;
            // 
            // lblClientName
            // 
            this.lblClientName.AutoSize = true;
            this.lblClientName.Location = new System.Drawing.Point(6, 49);
            this.lblClientName.Name = "lblClientName";
            this.lblClientName.Size = new System.Drawing.Size(100, 13);
            this.lblClientName.TabIndex = 49;
            this.lblClientName.Text = "Client Name (in full):";
            // 
            // tbClientExecutive
            // 
            this.tbClientExecutive.Controls.Add(this.groupBox1);
            this.tbClientExecutive.Location = new System.Drawing.Point(4, 22);
            this.tbClientExecutive.Name = "tbClientExecutive";
            this.tbClientExecutive.Padding = new System.Windows.Forms.Padding(3);
            this.tbClientExecutive.Size = new System.Drawing.Size(570, 385);
            this.tbClientExecutive.TabIndex = 1;
            this.tbClientExecutive.Text = "Account Exec";
            this.tbClientExecutive.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label17);
            this.groupBox1.Controls.Add(this.txtExecutiveMobile);
            this.groupBox1.Controls.Add(this.lblDepartment);
            this.groupBox1.Controls.Add(this.txtExecutiveDepartment);
            this.groupBox1.Controls.Add(this.lblAddress2);
            this.groupBox1.Controls.Add(this.txtBranchAddress2);
            this.groupBox1.Controls.Add(this.btnAccountExecutiveLookup);
            this.groupBox1.Controls.Add(this.lblAddress1);
            this.groupBox1.Controls.Add(this.txtBranchAddress1);
            this.groupBox1.Controls.Add(this.lblOwnerPhone);
            this.groupBox1.Controls.Add(this.txtExecutivePhone);
            this.groupBox1.Controls.Add(this.lblOwnerName);
            this.groupBox1.Controls.Add(this.lblOwnerEmail);
            this.groupBox1.Controls.Add(this.lblOwnerTitle);
            this.groupBox1.Controls.Add(this.txtExecutiveEmail);
            this.groupBox1.Controls.Add(this.txtExecutiveTitle);
            this.groupBox1.Controls.Add(this.txtExecutiveName);
            this.groupBox1.Location = new System.Drawing.Point(6, 6);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(558, 376);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Account Exec:";
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(20, 157);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(41, 13);
            this.label17.TabIndex = 77;
            this.label17.Text = "Mobile:";
            // 
            // txtExecutiveMobile
            // 
            this.txtExecutiveMobile.Location = new System.Drawing.Point(149, 154);
            this.txtExecutiveMobile.Name = "txtExecutiveMobile";
            this.txtExecutiveMobile.Size = new System.Drawing.Size(300, 20);
            this.txtExecutiveMobile.TabIndex = 76;
            // 
            // lblDepartment
            // 
            this.lblDepartment.AutoSize = true;
            this.lblDepartment.Location = new System.Drawing.Point(20, 258);
            this.lblDepartment.Name = "lblDepartment";
            this.lblDepartment.Size = new System.Drawing.Size(65, 13);
            this.lblDepartment.TabIndex = 75;
            this.lblDepartment.Text = "Department:";
            // 
            // txtExecutiveDepartment
            // 
            this.txtExecutiveDepartment.Location = new System.Drawing.Point(149, 255);
            this.txtExecutiveDepartment.Name = "txtExecutiveDepartment";
            this.txtExecutiveDepartment.Size = new System.Drawing.Size(300, 20);
            this.txtExecutiveDepartment.TabIndex = 74;
            // 
            // lblAddress2
            // 
            this.lblAddress2.AutoSize = true;
            this.lblAddress2.Location = new System.Drawing.Point(20, 220);
            this.lblAddress2.Name = "lblAddress2";
            this.lblAddress2.Size = new System.Drawing.Size(91, 13);
            this.lblAddress2.TabIndex = 73;
            this.lblAddress2.Text = "Branch Address2:";
            // 
            // txtBranchAddress2
            // 
            this.txtBranchAddress2.Location = new System.Drawing.Point(149, 217);
            this.txtBranchAddress2.Name = "txtBranchAddress2";
            this.txtBranchAddress2.Size = new System.Drawing.Size(300, 20);
            this.txtBranchAddress2.TabIndex = 9;
            // 
            // btnAccountExecutiveLookup
            // 
            this.btnAccountExecutiveLookup.Location = new System.Drawing.Point(454, 28);
            this.btnAccountExecutiveLookup.Name = "btnAccountExecutiveLookup";
            this.btnAccountExecutiveLookup.Size = new System.Drawing.Size(75, 23);
            this.btnAccountExecutiveLookup.TabIndex = 2;
            this.btnAccountExecutiveLookup.Text = "Lookup";
            this.btnAccountExecutiveLookup.UseVisualStyleBackColor = true;
            this.btnAccountExecutiveLookup.Click += new System.EventHandler(this.btnAccountExecutiveLookup_Click);
            // 
            // lblAddress1
            // 
            this.lblAddress1.AutoSize = true;
            this.lblAddress1.Location = new System.Drawing.Point(20, 189);
            this.lblAddress1.Name = "lblAddress1";
            this.lblAddress1.Size = new System.Drawing.Size(94, 13);
            this.lblAddress1.TabIndex = 70;
            this.lblAddress1.Text = "Branch Address 1:";
            // 
            // txtBranchAddress1
            // 
            this.txtBranchAddress1.Location = new System.Drawing.Point(149, 186);
            this.txtBranchAddress1.Name = "txtBranchAddress1";
            this.txtBranchAddress1.Size = new System.Drawing.Size(300, 20);
            this.txtBranchAddress1.TabIndex = 8;
            // 
            // lblOwnerPhone
            // 
            this.lblOwnerPhone.AutoSize = true;
            this.lblOwnerPhone.Location = new System.Drawing.Point(20, 127);
            this.lblOwnerPhone.Name = "lblOwnerPhone";
            this.lblOwnerPhone.Size = new System.Drawing.Size(41, 13);
            this.lblOwnerPhone.TabIndex = 68;
            this.lblOwnerPhone.Text = "Phone:";
            // 
            // txtExecutivePhone
            // 
            this.txtExecutivePhone.Location = new System.Drawing.Point(149, 124);
            this.txtExecutivePhone.Name = "txtExecutivePhone";
            this.txtExecutivePhone.Size = new System.Drawing.Size(300, 20);
            this.txtExecutivePhone.TabIndex = 5;
            // 
            // lblOwnerName
            // 
            this.lblOwnerName.AutoSize = true;
            this.lblOwnerName.Location = new System.Drawing.Point(19, 37);
            this.lblOwnerName.Name = "lblOwnerName";
            this.lblOwnerName.Size = new System.Drawing.Size(38, 13);
            this.lblOwnerName.TabIndex = 67;
            this.lblOwnerName.Text = "Name:";
            // 
            // lblOwnerEmail
            // 
            this.lblOwnerEmail.AutoSize = true;
            this.lblOwnerEmail.Location = new System.Drawing.Point(19, 96);
            this.lblOwnerEmail.Name = "lblOwnerEmail";
            this.lblOwnerEmail.Size = new System.Drawing.Size(35, 13);
            this.lblOwnerEmail.TabIndex = 66;
            this.lblOwnerEmail.Text = "Email:";
            // 
            // lblOwnerTitle
            // 
            this.lblOwnerTitle.AutoSize = true;
            this.lblOwnerTitle.Location = new System.Drawing.Point(19, 64);
            this.lblOwnerTitle.Name = "lblOwnerTitle";
            this.lblOwnerTitle.Size = new System.Drawing.Size(30, 13);
            this.lblOwnerTitle.TabIndex = 65;
            this.lblOwnerTitle.Text = "Title:";
            // 
            // txtExecutiveEmail
            // 
            this.txtExecutiveEmail.Location = new System.Drawing.Point(148, 93);
            this.txtExecutiveEmail.Name = "txtExecutiveEmail";
            this.txtExecutiveEmail.Size = new System.Drawing.Size(300, 20);
            this.txtExecutiveEmail.TabIndex = 4;
            // 
            // txtExecutiveTitle
            // 
            this.txtExecutiveTitle.Location = new System.Drawing.Point(148, 61);
            this.txtExecutiveTitle.Name = "txtExecutiveTitle";
            this.txtExecutiveTitle.Size = new System.Drawing.Size(300, 20);
            this.txtExecutiveTitle.TabIndex = 3;
            // 
            // txtExecutiveName
            // 
            this.txtExecutiveName.Location = new System.Drawing.Point(148, 30);
            this.txtExecutiveName.Name = "txtExecutiveName";
            this.txtExecutiveName.Size = new System.Drawing.Size(300, 20);
            this.txtExecutiveName.TabIndex = 1;
            // 
            // tbQuestions
            // 
            this.tbQuestions.Controls.Add(this.chkSelectAll);
            this.tbQuestions.Controls.Add(this.tvaQuestions);
            this.tbQuestions.Location = new System.Drawing.Point(4, 22);
            this.tbQuestions.Name = "tbQuestions";
            this.tbQuestions.Size = new System.Drawing.Size(570, 385);
            this.tbQuestions.TabIndex = 2;
            this.tbQuestions.Text = "Questions";
            this.tbQuestions.UseVisualStyleBackColor = true;
            // 
            // chkSelectAll
            // 
            this.chkSelectAll.AutoSize = true;
            this.chkSelectAll.Location = new System.Drawing.Point(3, 1);
            this.chkSelectAll.Name = "chkSelectAll";
            this.chkSelectAll.Size = new System.Drawing.Size(70, 17);
            this.chkSelectAll.TabIndex = 29;
            this.chkSelectAll.Text = "Select All";
            this.chkSelectAll.UseVisualStyleBackColor = true;
            this.chkSelectAll.CheckedChanged += new System.EventHandler(this.chkSelectAll_CheckedChanged);
            // 
            // tvaQuestions
            // 
            this.tvaQuestions.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.tvaQuestions.BackColor = System.Drawing.SystemColors.Window;
            this.tvaQuestions.DefaultToolTipProvider = null;
            this.tvaQuestions.DragDropMarkColor = System.Drawing.Color.Black;
            this.tvaQuestions.LineColor = System.Drawing.SystemColors.ControlDark;
            this.tvaQuestions.Location = new System.Drawing.Point(-4, 18);
            this.tvaQuestions.Model = null;
            this.tvaQuestions.Name = "tvaQuestions";
            this.tvaQuestions.NodeControls.Add(this._checked);
            this.tvaQuestions.NodeControls.Add(this._name);
            this.tvaQuestions.SelectedNode = null;
            this.tvaQuestions.Size = new System.Drawing.Size(574, 371);
            this.tvaQuestions.TabIndex = 28;
            this.tvaQuestions.Text = "treeViewAdv1";
            // 
            // _checked
            // 
            this._checked.DataPropertyName = "Checked";
            this._checked.EditEnabled = true;
            this._checked.IncrementalSearchEnabled = true;
            this._checked.IsVisibleAsParent = true;
            this._checked.LeftMargin = 3;
            this._checked.ParentColumn = null;
            // 
            // _name
            // 
            this._name.DataPropertyName = "Text";
            this._name.IncrementalSearchEnabled = true;
            this._name.LeftMargin = 3;
            this._name.ParentColumn = null;
            // 
            // pbOampsLogoFull
            // 
            this.pbOampsLogoFull.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.pbOampsLogoFull.Image = ((System.Drawing.Image)(resources.GetObject("pbOampsLogoFull.Image")));
            this.pbOampsLogoFull.InitialImage = ((System.Drawing.Image)(resources.GetObject("pbOampsLogoFull.InitialImage")));
            this.pbOampsLogoFull.Location = new System.Drawing.Point(170, 12);
            this.pbOampsLogoFull.Name = "pbOampsLogoFull";
            this.pbOampsLogoFull.Size = new System.Drawing.Size(185, 90);
            this.pbOampsLogoFull.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pbOampsLogoFull.TabIndex = 60;
            this.pbOampsLogoFull.TabStop = false;
            // 
            // ClientDiscoveryWizard
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Menu;
            this.ClientSize = new System.Drawing.Size(580, 570);
            this.Controls.Add(this.btnBack);
            this.Controls.Add(this.btnNext);
            this.Controls.Add(this.lblCoverPageTitle);
            this.Controls.Add(this.lblLogoTitle);
            this.Controls.Add(this.tbcWizardScreens);
            this.Controls.Add(this.pbOampsLogoFull);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ClientDiscoveryWizard";
            this.Text = "Client Discovery - Guide for Brokers";
            this.Load += new System.EventHandler(this.OnLoad_ClientDiscoveryWizard);
            this.tbcWizardScreens.ResumeLayout(false);
            this.tbClient.ResumeLayout(false);
            this.gbClient.ResumeLayout(false);
            this.gbClient.PerformLayout();
            this.tbClientExecutive.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.tbQuestions.ResumeLayout(false);
            this.tbQuestions.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pbOampsLogoFull)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.PictureBox pbOampsLogoFull;
        private System.Windows.Forms.TabControl tbcWizardScreens;
        private System.Windows.Forms.TabPage tbClient;
        private System.Windows.Forms.TabPage tbClientExecutive;
        private System.Windows.Forms.GroupBox gbClient;
        private System.Windows.Forms.TextBox txtClientContactName;
        private System.Windows.Forms.Label lblCltContactName;
        private System.Windows.Forms.TextBox txtClientName;
        private System.Windows.Forms.Label lblClientName;
        private System.Windows.Forms.DateTimePicker dateDiscussion;
        private System.Windows.Forms.Label lblDiscussion;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label lblAddress2;
        private System.Windows.Forms.TextBox txtBranchAddress2;
        private System.Windows.Forms.Button btnAccountExecutiveLookup;
        private System.Windows.Forms.Label lblAddress1;
        private System.Windows.Forms.TextBox txtBranchAddress1;
        private System.Windows.Forms.Label lblOwnerPhone;
        private System.Windows.Forms.TextBox txtExecutivePhone;
        private System.Windows.Forms.Label lblOwnerName;
        private System.Windows.Forms.Label lblOwnerEmail;
        private System.Windows.Forms.Label lblOwnerTitle;
        private System.Windows.Forms.TextBox txtExecutiveEmail;
        private System.Windows.Forms.TextBox txtExecutiveTitle;
        private System.Windows.Forms.TextBox txtExecutiveName;
        private System.Windows.Forms.Label lblCoverPageTitle;
        private System.Windows.Forms.Label lblLogoTitle;
        private System.Windows.Forms.Button btnBack;
        private System.Windows.Forms.Button btnNext;
        private System.Windows.Forms.TabPage tbQuestions;
        private Aga.Controls.Tree.TreeViewAdv tvaQuestions;
        private Aga.Controls.Tree.NodeControls.NodeCheckBox _checked;
        private Aga.Controls.Tree.NodeControls.NodeTextBox _name;
        private System.Windows.Forms.Label lblDepartment;
        private System.Windows.Forms.TextBox txtExecutiveDepartment;
        private System.Windows.Forms.TextBox txtCltiBAIS;
        private System.Windows.Forms.Label lblCltiBAIS;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.TextBox txtExecutiveMobile;
        private System.Windows.Forms.DateTimePicker datePrepared;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox chkSelectAll;
    }
}
