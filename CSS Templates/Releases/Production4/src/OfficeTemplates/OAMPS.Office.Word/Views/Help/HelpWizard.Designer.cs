﻿namespace OAMPS.Office.Word.Views.Help
{
    partial class HelpWizard
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(HelpWizard));
            this.webHelpWindow = new System.Windows.Forms.WebBrowser();
            this.SuspendLayout();
            // 
            // webHelpWindow
            // 
            this.webHelpWindow.Dock = System.Windows.Forms.DockStyle.Fill;
            this.webHelpWindow.Location = new System.Drawing.Point(0, 0);
            this.webHelpWindow.MinimumSize = new System.Drawing.Size(20, 20);
            this.webHelpWindow.Name = "webHelpWindow";
            this.webHelpWindow.Size = new System.Drawing.Size(631, 451);
            this.webHelpWindow.TabIndex = 0;
            this.webHelpWindow.DocumentCompleted += new System.Windows.Forms.WebBrowserDocumentCompletedEventHandler(this.webHelpWindow_DocumentCompleted);
            // 
            // HelpWizard
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(631, 451);
            this.Controls.Add(this.webHelpWindow);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "HelpWizard";
            this.Text = "Help Content";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.WebBrowser webHelpWindow;
    }
}