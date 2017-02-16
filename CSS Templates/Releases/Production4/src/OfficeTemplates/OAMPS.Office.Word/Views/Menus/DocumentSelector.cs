using System;
using System.Windows.Forms;

namespace OAMPS.Office.Word.Views.Menus
{
    public partial class DocumentSelector : Form
    {
        public DocumentSelector()
        {
            InitializeComponent();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
            //todo: should we also exit the doucment?
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
        }

        private void LoadData()
        {
        }
    }
}