namespace OAMPS.Office.Word.Views.Wizards
{
    partial class PreRenewalAgendaWizard
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(PreRenewalAgendaWizard));
            this.lblCoverPageTitle = new System.Windows.Forms.Label();
            this.lblLogoTitle = new System.Windows.Forms.Label();
            this.btnBack = new System.Windows.Forms.Button();
            this.btnNext = new System.Windows.Forms.Button();
            this.tbcWizardScreens = new System.Windows.Forms.TabControl();
            this.tbpClientInfo = new System.Windows.Forms.TabPage();
            this.gbClient = new System.Windows.Forms.GroupBox();
            this.txtLocation = new System.Windows.Forms.TextBox();
            this.lblLocation = new System.Windows.Forms.Label();
            this.timeAgendaTo = new System.Windows.Forms.DateTimePicker();
            this.lblTimeTo = new System.Windows.Forms.Label();
            this.txtClientSubject = new System.Windows.Forms.TextBox();
            this.lblSubject = new System.Windows.Forms.Label();
            this.timeAgendaFrom = new System.Windows.Forms.DateTimePicker();
            this.lblTimeFrom = new System.Windows.Forms.Label();
            this.dateAgenda = new System.Windows.Forms.DateTimePicker();
            this.lblDateAgenda = new System.Windows.Forms.Label();
            this.txtClientName = new System.Windows.Forms.TextBox();
            this.lblClientName = new System.Windows.Forms.Label();
            this.tbpDocumentInfo = new System.Windows.Forms.TabPage();
            this.panQuestions = new System.Windows.Forms.Panel();
            this.lblQuestionDocumentType = new System.Windows.Forms.Label();
            this.rdoMeetingMinutes = new System.Windows.Forms.RadioButton();
            this.rdoMeetingAgenda = new System.Windows.Forms.RadioButton();
            this.pbOampsLogoFull = new System.Windows.Forms.PictureBox();
            this.errorProvider1 = new System.Windows.Forms.ErrorProvider(this.components);
            this.tbcWizardScreens.SuspendLayout();
            this.tbpClientInfo.SuspendLayout();
            this.gbClient.SuspendLayout();
            this.tbpDocumentInfo.SuspendLayout();
            this.panQuestions.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pbOampsLogoFull)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider1)).BeginInit();
            this.SuspendLayout();
            // 
            // lblCoverPageTitle
            // 
            this.lblCoverPageTitle.AutoSize = true;
            this.lblCoverPageTitle.Location = new System.Drawing.Point(152, 612);
            this.lblCoverPageTitle.Name = "lblCoverPageTitle";
            this.lblCoverPageTitle.Size = new System.Drawing.Size(54, 13);
            this.lblCoverPageTitle.TabIndex = 65;
            this.lblCoverPageTitle.Text = "Dark Blue";
            this.lblCoverPageTitle.Visible = false;
            // 
            // lblLogoTitle
            // 
            this.lblLogoTitle.AutoSize = true;
            this.lblLogoTitle.Location = new System.Drawing.Point(226, 611);
            this.lblLogoTitle.Name = "lblLogoTitle";
            this.lblLogoTitle.Size = new System.Drawing.Size(176, 13);
            this.lblLogoTitle.TabIndex = 64;
            this.lblLogoTitle.Text = "Arthur J. Gallagher & Co (Aus) Limited";
            this.lblLogoTitle.Visible = false;
            // 
            // btnBack
            // 
            this.btnBack.Enabled = false;
            this.btnBack.Location = new System.Drawing.Point(8, 604);
            this.btnBack.Name = "btnBack";
            this.btnBack.Size = new System.Drawing.Size(75, 26);
            this.btnBack.TabIndex = 23;
            this.btnBack.Text = "&Back";
            this.btnBack.UseVisualStyleBackColor = true;
            this.btnBack.Click += new System.EventHandler(this.btnBack_Click);
            // 
            // btnNext
            // 
            this.btnNext.Location = new System.Drawing.Point(507, 607);
            this.btnNext.Name = "btnNext";
            this.btnNext.Size = new System.Drawing.Size(75, 23);
            this.btnNext.TabIndex = 24;
            this.btnNext.Text = "&Next";
            this.btnNext.UseVisualStyleBackColor = true;
            this.btnNext.Click += new System.EventHandler(this.btnNext_Click);
            // 
            // tbcWizardScreens
            // 
            this.tbcWizardScreens.Controls.Add(this.tbpClientInfo);
            this.tbcWizardScreens.Controls.Add(this.tbpDocumentInfo);
            this.tbcWizardScreens.Location = new System.Drawing.Point(8, 98);
            this.tbcWizardScreens.Name = "tbcWizardScreens";
            this.tbcWizardScreens.SelectedIndex = 0;
            this.tbcWizardScreens.Size = new System.Drawing.Size(569, 500);
            this.tbcWizardScreens.TabIndex = 22;
            // 
            // tbpClientInfo
            // 
            this.tbpClientInfo.Controls.Add(this.gbClient);
            this.tbpClientInfo.Location = new System.Drawing.Point(4, 22);
            this.tbpClientInfo.Name = "tbpClientInfo";
            this.tbpClientInfo.Size = new System.Drawing.Size(561, 474);
            this.tbpClientInfo.TabIndex = 6;
            this.tbpClientInfo.Text = "Client Information";
            this.tbpClientInfo.UseVisualStyleBackColor = true;
            // 
            // gbClient
            // 
            this.gbClient.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gbClient.Controls.Add(this.txtLocation);
            this.gbClient.Controls.Add(this.lblLocation);
            this.gbClient.Controls.Add(this.timeAgendaTo);
            this.gbClient.Controls.Add(this.lblTimeTo);
            this.gbClient.Controls.Add(this.txtClientSubject);
            this.gbClient.Controls.Add(this.lblSubject);
            this.gbClient.Controls.Add(this.timeAgendaFrom);
            this.gbClient.Controls.Add(this.lblTimeFrom);
            this.gbClient.Controls.Add(this.dateAgenda);
            this.gbClient.Controls.Add(this.lblDateAgenda);
            this.gbClient.Controls.Add(this.txtClientName);
            this.gbClient.Controls.Add(this.lblClientName);
            this.gbClient.Location = new System.Drawing.Point(13, 15);
            this.gbClient.Name = "gbClient";
            this.gbClient.Size = new System.Drawing.Size(534, 265);
            this.gbClient.TabIndex = 56;
            this.gbClient.TabStop = false;
            this.gbClient.Text = "Client Information";
            // 
            // txtLocation
            // 
            this.txtLocation.Location = new System.Drawing.Point(148, 76);
            this.txtLocation.Name = "txtLocation";
            this.txtLocation.Size = new System.Drawing.Size(215, 20);
            this.txtLocation.TabIndex = 3;
            // 
            // lblLocation
            // 
            this.lblLocation.AutoSize = true;
            this.lblLocation.Location = new System.Drawing.Point(6, 83);
            this.lblLocation.Name = "lblLocation";
            this.lblLocation.Size = new System.Drawing.Size(139, 13);
            this.lblLocation.TabIndex = 76;
            this.lblLocation.Text = "Location(or telecom details):";
            // 
            // timeAgendaTo
            // 
            this.timeAgendaTo.CustomFormat = "h:mm tt";
            this.timeAgendaTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.timeAgendaTo.Location = new System.Drawing.Point(136, 168);
            this.timeAgendaTo.Name = "timeAgendaTo";
            this.timeAgendaTo.ShowUpDown = true;
            this.timeAgendaTo.Size = new System.Drawing.Size(111, 20);
            this.timeAgendaTo.TabIndex = 6;
            // 
            // lblTimeTo
            // 
            this.lblTimeTo.AutoSize = true;
            this.lblTimeTo.Location = new System.Drawing.Point(6, 168);
            this.lblTimeTo.Name = "lblTimeTo";
            this.lblTimeTo.Size = new System.Drawing.Size(49, 13);
            this.lblTimeTo.TabIndex = 74;
            this.lblTimeTo.Text = "Time To:";
            // 
            // txtClientSubject
            // 
            this.txtClientSubject.Location = new System.Drawing.Point(148, 50);
            this.txtClientSubject.Name = "txtClientSubject";
            this.txtClientSubject.Size = new System.Drawing.Size(215, 20);
            this.txtClientSubject.TabIndex = 2;
            // 
            // lblSubject
            // 
            this.lblSubject.AutoSize = true;
            this.lblSubject.Location = new System.Drawing.Point(6, 57);
            this.lblSubject.Name = "lblSubject";
            this.lblSubject.Size = new System.Drawing.Size(46, 13);
            this.lblSubject.TabIndex = 72;
            this.lblSubject.Text = "Subject:";
            // 
            // timeAgendaFrom
            // 
            this.timeAgendaFrom.CustomFormat = "h:mm tt";
            this.timeAgendaFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.timeAgendaFrom.Location = new System.Drawing.Point(136, 142);
            this.timeAgendaFrom.Name = "timeAgendaFrom";
            this.timeAgendaFrom.ShowUpDown = true;
            this.timeAgendaFrom.Size = new System.Drawing.Size(111, 20);
            this.timeAgendaFrom.TabIndex = 5;
            // 
            // lblTimeFrom
            // 
            this.lblTimeFrom.AutoSize = true;
            this.lblTimeFrom.Location = new System.Drawing.Point(6, 142);
            this.lblTimeFrom.Name = "lblTimeFrom";
            this.lblTimeFrom.Size = new System.Drawing.Size(59, 13);
            this.lblTimeFrom.TabIndex = 70;
            this.lblTimeFrom.Text = "Time From:";
            // 
            // dateAgenda
            // 
            this.dateAgenda.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dateAgenda.Location = new System.Drawing.Point(136, 117);
            this.dateAgenda.Name = "dateAgenda";
            this.dateAgenda.Size = new System.Drawing.Size(111, 20);
            this.dateAgenda.TabIndex = 4;
            // 
            // lblDateAgenda
            // 
            this.lblDateAgenda.AutoSize = true;
            this.lblDateAgenda.Location = new System.Drawing.Point(6, 117);
            this.lblDateAgenda.Name = "lblDateAgenda";
            this.lblDateAgenda.Size = new System.Drawing.Size(86, 13);
            this.lblDateAgenda.TabIndex = 67;
            this.lblDateAgenda.Text = "Date of Meeting:";
            // 
            // txtClientName
            // 
            this.txtClientName.Location = new System.Drawing.Point(148, 24);
            this.txtClientName.Name = "txtClientName";
            this.txtClientName.Size = new System.Drawing.Size(215, 20);
            this.txtClientName.TabIndex = 1;
            // 
            // lblClientName
            // 
            this.lblClientName.AutoSize = true;
            this.lblClientName.Location = new System.Drawing.Point(6, 28);
            this.lblClientName.Name = "lblClientName";
            this.lblClientName.Size = new System.Drawing.Size(67, 13);
            this.lblClientName.TabIndex = 49;
            this.lblClientName.Text = "Client Name:";
            // 
            // tbpDocumentInfo
            // 
            this.tbpDocumentInfo.Controls.Add(this.panQuestions);
            this.tbpDocumentInfo.Location = new System.Drawing.Point(4, 22);
            this.tbpDocumentInfo.Name = "tbpDocumentInfo";
            this.tbpDocumentInfo.Padding = new System.Windows.Forms.Padding(3);
            this.tbpDocumentInfo.Size = new System.Drawing.Size(561, 474);
            this.tbpDocumentInfo.TabIndex = 5;
            this.tbpDocumentInfo.Text = "Document Information";
            this.tbpDocumentInfo.UseVisualStyleBackColor = true;
            // 
            // panQuestions
            // 
            this.panQuestions.Controls.Add(this.lblQuestionDocumentType);
            this.panQuestions.Controls.Add(this.rdoMeetingMinutes);
            this.panQuestions.Controls.Add(this.rdoMeetingAgenda);
            this.panQuestions.Location = new System.Drawing.Point(6, 19);
            this.panQuestions.Name = "panQuestions";
            this.panQuestions.Size = new System.Drawing.Size(549, 155);
            this.panQuestions.TabIndex = 0;
            // 
            // lblQuestionDocumentType
            // 
            this.lblQuestionDocumentType.AutoSize = true;
            this.lblQuestionDocumentType.Location = new System.Drawing.Point(24, 9);
            this.lblQuestionDocumentType.Name = "lblQuestionDocumentType";
            this.lblQuestionDocumentType.Size = new System.Drawing.Size(185, 13);
            this.lblQuestionDocumentType.TabIndex = 2;
            this.lblQuestionDocumentType.Text = "What kind of document do you want?";
            // 
            // rdoMeetingMinutes
            // 
            this.rdoMeetingMinutes.AutoSize = true;
            this.rdoMeetingMinutes.Location = new System.Drawing.Point(27, 69);
            this.rdoMeetingMinutes.Name = "rdoMeetingMinutes";
            this.rdoMeetingMinutes.Size = new System.Drawing.Size(103, 17);
            this.rdoMeetingMinutes.TabIndex = 1;
            this.rdoMeetingMinutes.Text = "Meeting Minutes";
            this.rdoMeetingMinutes.UseVisualStyleBackColor = true;
            this.rdoMeetingMinutes.CheckedChanged += new System.EventHandler(this.rdoMeetingMinutes_CheckedChanged);
            // 
            // rdoMeetingAgenda
            // 
            this.rdoMeetingAgenda.AutoSize = true;
            this.rdoMeetingAgenda.Location = new System.Drawing.Point(27, 36);
            this.rdoMeetingAgenda.Name = "rdoMeetingAgenda";
            this.rdoMeetingAgenda.Size = new System.Drawing.Size(103, 17);
            this.rdoMeetingAgenda.TabIndex = 0;
            this.rdoMeetingAgenda.Text = "Meeting Agenda";
            this.rdoMeetingAgenda.UseVisualStyleBackColor = true;
            this.rdoMeetingAgenda.CheckedChanged += new System.EventHandler(this.rdoMeetingAgenda_CheckedChanged);
            // 
            // pbOampsLogoFull
            // 
            this.pbOampsLogoFull.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.pbOampsLogoFull.Image = global::OAMPS.Office.Word.Properties.Resources.Gallagher;
            this.pbOampsLogoFull.InitialImage = ((System.Drawing.Image)(resources.GetObject("pbOampsLogoFull.InitialImage")));
            this.pbOampsLogoFull.Location = new System.Drawing.Point(160, 2);
            this.pbOampsLogoFull.Name = "pbOampsLogoFull";
            this.pbOampsLogoFull.Size = new System.Drawing.Size(303, 90);
            this.pbOampsLogoFull.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pbOampsLogoFull.TabIndex = 25;
            this.pbOampsLogoFull.TabStop = false;
            // 
            // errorProvider1
            // 
            this.errorProvider1.ContainerControl = this;
            // 
            // PreRenewalAgendaWizard
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(596, 643);
            this.Controls.Add(this.lblCoverPageTitle);
            this.Controls.Add(this.lblLogoTitle);
            this.Controls.Add(this.btnBack);
            this.Controls.Add(this.btnNext);
            this.Controls.Add(this.tbcWizardScreens);
            this.Controls.Add(this.pbOampsLogoFull);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "PreRenewalAgendaWizard";
            this.Text = "Meeting Agenda/Minutes";
            this.Load += new System.EventHandler(this.SummaryOfDiscussionWizard_Load);
            this.tbcWizardScreens.ResumeLayout(false);
            this.tbpClientInfo.ResumeLayout(false);
            this.gbClient.ResumeLayout(false);
            this.gbClient.PerformLayout();
            this.tbpDocumentInfo.ResumeLayout(false);
            this.panQuestions.ResumeLayout(false);
            this.panQuestions.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pbOampsLogoFull)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnBack;
        private System.Windows.Forms.PictureBox pbOampsLogoFull;
        private System.Windows.Forms.TabPage tbpDocumentInfo;
        private System.Windows.Forms.Panel panQuestions;
        private System.Windows.Forms.Label lblQuestionDocumentType;
        private System.Windows.Forms.RadioButton rdoMeetingMinutes;
        private System.Windows.Forms.RadioButton rdoMeetingAgenda;
        private System.Windows.Forms.TabControl tbcWizardScreens;
        public System.Windows.Forms.Button btnNext;
        private System.Windows.Forms.TabPage tbpClientInfo;
        private System.Windows.Forms.GroupBox gbClient;
        private System.Windows.Forms.TextBox txtClientSubject;
        private System.Windows.Forms.Label lblSubject;
        private System.Windows.Forms.DateTimePicker dateAgenda;
        private System.Windows.Forms.Label lblDateAgenda;
        private System.Windows.Forms.TextBox txtClientName;
        private System.Windows.Forms.Label lblClientName;
        private System.Windows.Forms.TextBox txtLocation;
        private System.Windows.Forms.Label lblLocation;
        private System.Windows.Forms.DateTimePicker timeAgendaTo;
        private System.Windows.Forms.Label lblTimeTo;
        private System.Windows.Forms.DateTimePicker timeAgendaFrom;
        private System.Windows.Forms.Label lblTimeFrom;
        private System.Windows.Forms.ErrorProvider errorProvider1;
        private System.Windows.Forms.Label lblCoverPageTitle;
        private System.Windows.Forms.Label lblLogoTitle;
    }
}
