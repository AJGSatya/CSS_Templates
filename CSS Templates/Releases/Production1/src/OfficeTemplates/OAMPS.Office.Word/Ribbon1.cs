using System.Diagnostics;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using OAMPS.Office.Word.Helpers;
using OAMPS.Office.Word.Views.Help;

namespace OAMPS.Office.Word
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.ActiveDocument.LoadWizard(Enums.FormLoadType.RibbonClick);
        }

        private void btnHelp_Click(object sender, RibbonControlEventArgs e)
        {
            var wizard = new HelpWizard()
            {
                TopMost = true,
                StartPosition = FormStartPosition.CenterParent
            };
#if DEBUG
                            wizard.TopMost = false;
#endif
            wizard.Show();
        }

        private void btnConvertPDF_Click(object sender, RibbonControlEventArgs e)
        {
            Process.Start(@"C:\Program Files\Free PDF Solutions\FreePDFSolutions_PDF2Word.exe");
            

        }

        private void btnUpdateFields_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.ActiveDocument.Fields.Update();
        }
    }
}
