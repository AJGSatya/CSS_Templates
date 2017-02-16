using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace OAMPS.Office.Word.Views
{
    public partial class SystemThreadSync : Form
    {
        private Document _document;

        public SystemThreadSync(Document doc)
        {
            InitializeComponent();

            _document = doc;
        }

        private void _system_ThreadSync_Load(object sender, EventArgs e)
        {
            //    var uiScheduler = TaskScheduler.FromCurrentSynchronizationContext();
            //    System.Threading.Tasks.Task.Factory.StartNew(() => ThisAddIn.CheckStartupTasks(document), CancellationToken.None, TaskCreationOptions.None, uiScheduler);
            //    Close();
        }
    }
}