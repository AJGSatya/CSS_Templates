namespace OAMPS.Office.Word
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.grpOAMPS = this.Factory.CreateRibbonGroup();
            this.btnWizard = this.Factory.CreateRibbonButton();
            this.btnHelp = this.Factory.CreateRibbonButton();
            this.btnConvertPDF = this.Factory.CreateRibbonButton();
            this.btnUpdateFields = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.grpOAMPS.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.grpOAMPS);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // grpOAMPS
            // 
            this.grpOAMPS.Items.Add(this.btnWizard);
            this.grpOAMPS.Items.Add(this.btnHelp);
            this.grpOAMPS.Items.Add(this.btnConvertPDF);
            this.grpOAMPS.Items.Add(this.btnUpdateFields);
            this.grpOAMPS.Label = "OAMPS";
            this.grpOAMPS.Name = "grpOAMPS";
            // 
            // btnWizard
            // 
            this.btnWizard.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnWizard.Image = global::OAMPS.Office.Word.Properties.Resources.Wizard;
            this.btnWizard.Label = "Wizard";
            this.btnWizard.Name = "btnWizard";
            this.btnWizard.ShowImage = true;
            this.btnWizard.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // btnHelp
            // 
            this.btnHelp.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnHelp.Image = ((System.Drawing.Image)(resources.GetObject("btnHelp.Image")));
            this.btnHelp.Label = "Help";
            this.btnHelp.Name = "btnHelp";
            this.btnHelp.ShowImage = true;
            this.btnHelp.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnHelp_Click);
            // 
            // btnConvertPDF
            // 
            this.btnConvertPDF.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnConvertPDF.Image = global::OAMPS.Office.Word.Properties.Resources.pdf;
            this.btnConvertPDF.Label = "Convert PDF";
            this.btnConvertPDF.Name = "btnConvertPDF";
            this.btnConvertPDF.ShowImage = true;
            this.btnConvertPDF.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnConvertPDF_Click);
            // 
            // btnUpdateFields
            // 
            this.btnUpdateFields.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnUpdateFields.Image = global::OAMPS.Office.Word.Properties.Resources.calculator;
            this.btnUpdateFields.Label = "Refresh Summary Table";
            this.btnUpdateFields.Name = "btnUpdateFields";
            this.btnUpdateFields.ShowImage = true;
            this.btnUpdateFields.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdateFields_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.grpOAMPS.ResumeLayout(false);
            this.grpOAMPS.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpOAMPS;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnWizard;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnHelp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConvertPDF;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdateFields;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
