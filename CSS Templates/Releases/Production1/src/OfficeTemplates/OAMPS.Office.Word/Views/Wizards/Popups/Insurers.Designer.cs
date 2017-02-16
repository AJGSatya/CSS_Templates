﻿namespace OAMPS.Office.Word.Views.Wizards.Popups
{
    partial class Insurers
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
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOk = new System.Windows.Forms.Button();
            this.grpRecommended = new System.Windows.Forms.GroupBox();
            this.dgvRecommended = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.RecommendedId = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.txtFindRecommended = new System.Windows.Forms.TextBox();
            this.lblFindRecommended = new System.Windows.Forms.Label();
            this.grpCurrent = new System.Windows.Forms.GroupBox();
            this.dgvCurrent = new System.Windows.Forms.DataGridView();
            this.Insurer = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Category = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CurrentId = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.txtFindCurrent = new System.Windows.Forms.TextBox();
            this.lblFindCurrent = new System.Windows.Forms.Label();
            this.grpRecommended.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvRecommended)).BeginInit();
            this.grpCurrent.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvCurrent)).BeginInit();
            this.SuspendLayout();
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(601, 436);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 7;
            this.btnCancel.Text = "&Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnOk
            // 
            this.btnOk.Location = new System.Drawing.Point(520, 436);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(75, 23);
            this.btnOk.TabIndex = 6;
            this.btnOk.Text = "&OK";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // grpRecommended
            // 
            this.grpRecommended.Controls.Add(this.dgvRecommended);
            this.grpRecommended.Controls.Add(this.txtFindRecommended);
            this.grpRecommended.Controls.Add(this.lblFindRecommended);
            this.grpRecommended.Location = new System.Drawing.Point(409, 22);
            this.grpRecommended.Name = "grpRecommended";
            this.grpRecommended.Size = new System.Drawing.Size(388, 402);
            this.grpRecommended.TabIndex = 1;
            this.grpRecommended.TabStop = false;
            this.grpRecommended.Text = "Recommended";
            // 
            // dgvRecommended
            // 
            this.dgvRecommended.AllowUserToAddRows = false;
            this.dgvRecommended.AllowUserToDeleteRows = false;
            this.dgvRecommended.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvRecommended.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn1,
            this.dataGridViewTextBoxColumn2,
            this.RecommendedId});
            this.dgvRecommended.Location = new System.Drawing.Point(8, 58);
            this.dgvRecommended.MultiSelect = false;
            this.dgvRecommended.Name = "dgvRecommended";
            this.dgvRecommended.ReadOnly = true;
            this.dgvRecommended.RowHeadersVisible = false;
            this.dgvRecommended.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvRecommended.Size = new System.Drawing.Size(374, 338);
            this.dgvRecommended.TabIndex = 9;
            // 
            // dataGridViewTextBoxColumn1
            // 
            this.dataGridViewTextBoxColumn1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.dataGridViewTextBoxColumn1.DataPropertyName = "Title";
            this.dataGridViewTextBoxColumn1.HeaderText = "Insurer";
            this.dataGridViewTextBoxColumn1.Name = "dataGridViewTextBoxColumn1";
            this.dataGridViewTextBoxColumn1.ReadOnly = true;
            // 
            // dataGridViewTextBoxColumn2
            // 
            this.dataGridViewTextBoxColumn2.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.dataGridViewTextBoxColumn2.DataPropertyName = "Category";
            this.dataGridViewTextBoxColumn2.HeaderText = "Category";
            this.dataGridViewTextBoxColumn2.Name = "dataGridViewTextBoxColumn2";
            this.dataGridViewTextBoxColumn2.ReadOnly = true;
            this.dataGridViewTextBoxColumn2.Width = 74;
            // 
            // RecommendedId
            // 
            this.RecommendedId.DataPropertyName = "Id";
            this.RecommendedId.HeaderText = "Id";
            this.RecommendedId.Name = "RecommendedId";
            this.RecommendedId.ReadOnly = true;
            this.RecommendedId.Visible = false;
            // 
            // txtFindRecommended
            // 
            this.txtFindRecommended.Location = new System.Drawing.Point(9, 32);
            this.txtFindRecommended.Name = "txtFindRecommended";
            this.txtFindRecommended.Size = new System.Drawing.Size(373, 20);
            this.txtFindRecommended.TabIndex = 4;
            this.txtFindRecommended.TextChanged += new System.EventHandler(this.txtFindRecommended_TextChanged);
            // 
            // lblFindRecommended
            // 
            this.lblFindRecommended.AutoSize = true;
            this.lblFindRecommended.Location = new System.Drawing.Point(6, 16);
            this.lblFindRecommended.Name = "lblFindRecommended";
            this.lblFindRecommended.Size = new System.Drawing.Size(27, 13);
            this.lblFindRecommended.TabIndex = 1;
            this.lblFindRecommended.Text = "Find";
            // 
            // grpCurrent
            // 
            this.grpCurrent.Controls.Add(this.dgvCurrent);
            this.grpCurrent.Controls.Add(this.txtFindCurrent);
            this.grpCurrent.Controls.Add(this.lblFindCurrent);
            this.grpCurrent.Location = new System.Drawing.Point(12, 22);
            this.grpCurrent.Name = "grpCurrent";
            this.grpCurrent.Size = new System.Drawing.Size(388, 402);
            this.grpCurrent.TabIndex = 0;
            this.grpCurrent.TabStop = false;
            this.grpCurrent.Text = "Current";
            // 
            // dgvCurrent
            // 
            this.dgvCurrent.AllowUserToAddRows = false;
            this.dgvCurrent.AllowUserToDeleteRows = false;
            this.dgvCurrent.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvCurrent.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Insurer,
            this.Category,
            this.CurrentId});
            this.dgvCurrent.Location = new System.Drawing.Point(6, 58);
            this.dgvCurrent.MultiSelect = false;
            this.dgvCurrent.Name = "dgvCurrent";
            this.dgvCurrent.ReadOnly = true;
            this.dgvCurrent.RowHeadersVisible = false;
            this.dgvCurrent.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvCurrent.Size = new System.Drawing.Size(376, 338);
            this.dgvCurrent.TabIndex = 8;
            // 
            // Insurer
            // 
            this.Insurer.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.Insurer.DataPropertyName = "Title";
            this.Insurer.HeaderText = "Insurer";
            this.Insurer.Name = "Insurer";
            this.Insurer.ReadOnly = true;
            // 
            // Category
            // 
            this.Category.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.Category.DataPropertyName = "Category";
            this.Category.HeaderText = "Category";
            this.Category.Name = "Category";
            this.Category.ReadOnly = true;
            this.Category.Width = 74;
            // 
            // CurrentId
            // 
            this.CurrentId.DataPropertyName = "Id";
            this.CurrentId.HeaderText = "Id";
            this.CurrentId.Name = "CurrentId";
            this.CurrentId.ReadOnly = true;
            this.CurrentId.Visible = false;
            // 
            // txtFindCurrent
            // 
            this.txtFindCurrent.Location = new System.Drawing.Point(6, 32);
            this.txtFindCurrent.Name = "txtFindCurrent";
            this.txtFindCurrent.Size = new System.Drawing.Size(376, 20);
            this.txtFindCurrent.TabIndex = 2;
            this.txtFindCurrent.TextChanged += new System.EventHandler(this.txtFindCurrent_TextChanged);
            // 
            // lblFindCurrent
            // 
            this.lblFindCurrent.AutoSize = true;
            this.lblFindCurrent.Location = new System.Drawing.Point(6, 16);
            this.lblFindCurrent.Name = "lblFindCurrent";
            this.lblFindCurrent.Size = new System.Drawing.Size(27, 13);
            this.lblFindCurrent.TabIndex = 1;
            this.lblFindCurrent.Text = "Find";
            // 
            // Insurers
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(803, 473);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.grpRecommended);
            this.Controls.Add(this.grpCurrent);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "Insurers";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Select Insurers";
            this.Load += new System.EventHandler(this.Insurers_Load);
            this.grpRecommended.ResumeLayout(false);
            this.grpRecommended.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvRecommended)).EndInit();
            this.grpCurrent.ResumeLayout(false);
            this.grpCurrent.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvCurrent)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox grpCurrent;
        private System.Windows.Forms.TextBox txtFindCurrent;
        private System.Windows.Forms.Label lblFindCurrent;
        private System.Windows.Forms.GroupBox grpRecommended;
        private System.Windows.Forms.TextBox txtFindRecommended;
        private System.Windows.Forms.Label lblFindRecommended;
        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.DataGridView dgvCurrent;
        private System.Windows.Forms.DataGridView dgvRecommended;
        private System.Windows.Forms.DataGridViewTextBoxColumn Insurer;
        private System.Windows.Forms.DataGridViewTextBoxColumn Category;
        private System.Windows.Forms.DataGridViewTextBoxColumn CurrentId;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
        private System.Windows.Forms.DataGridViewTextBoxColumn RecommendedId;
    }
}