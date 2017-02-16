using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace OAMPS.Office.Word.Views
{
    public partial class _system_ThreadSync : Form
    {
        private Document document = null;
        public _system_ThreadSync(Document doc)
        {
            InitializeComponent();

            document = doc;
        }

        private void _system_ThreadSync_Load(object sender, EventArgs e)
        {
        //    var uiScheduler = TaskScheduler.FromCurrentSynchronizationContext();
        //    System.Threading.Tasks.Task.Factory.StartNew(() => ThisAddIn.CheckStartupTasks(document), CancellationToken.None, TaskCreationOptions.None, uiScheduler);
        //    Close();
        }
    }
}
