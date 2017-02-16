using System;
using System.ComponentModel;
using System.Windows.Forms;
using OAMPS.Office.Word.Helpers.Controls;

namespace OAMPS.Office.Word.Views.Wizards.Popups
{
    public partial class EnterInsurer : Form
    {
        public string EnteredInsurer;

        public EnterInsurer()
        {
            InitializeComponent();
            txtName.Validating += txtName_Validating;
        }

        private void txtName_Validating(object sender, CancelEventArgs e)
        {
            if (String.IsNullOrEmpty(txtName.Text))
            {
                errorProvider.SetError(txtName, "Required Field");
                e.Cancel = true;
            }
            else
            {
                errorProvider.SetError(txtName, String.Empty);
            }
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            if (Validation.HasValidationErrors(Controls))
            {
                MessageBox.Show(@"Please ensure all required fields are populated", @"Required fields are missing",
                                MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            EnteredInsurer = txtName.Text;
            Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void EnterInsurer_Load(object sender, EventArgs e)
        {
        }
    }
}