using System;
using System.Windows.Forms;
using OAMPS.Office.BusinessLogic.Helpers;

namespace OAMPS.Office.Word.Views.Help
{
    public partial class About : Form
    {
        public About()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void About_Load(object sender, EventArgs e)
        {
            lblVersion.Text = Constants.Configuration.VersionNumber;
        }
    }
}