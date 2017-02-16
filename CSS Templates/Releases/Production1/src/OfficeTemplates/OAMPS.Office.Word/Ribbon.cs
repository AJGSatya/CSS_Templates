using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Core;
using OAMPS.Office.BusinessLogic.Helpers;
using OAMPS.Office.Word.Helpers;
using OAMPS.Office.Word.Views.Help;
using Enums = OAMPS.Office.Word.Helpers.Enums;
using Office = Microsoft.Office.Core;
using WordOM = Microsoft.Office.Interop.Word;

namespace OAMPS.Office.Word
{
    [ComVisible(true)]
    public class Ribbon : Microsoft.Office.Core.IRibbonExtensibility
    {

        private Microsoft.Office.Core.IRibbonUI ribbon;
        
        public Ribbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("OAMPS.Office.Word.Ribbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, select the Ribbon XML item in Solution Explorer and then press F1

        public void Ribbon_Load(Microsoft.Office.Core.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }


        public void btnUpdateFields_Click(Microsoft.Office.Core.IRibbonControl control)
        {
            Globals.ThisAddIn.Application.ActiveDocument.Fields.Update();

            CalculatePremSummaryTable();
        }

        public void btnHelp_Click(Microsoft.Office.Core.IRibbonControl control)
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
        public void btnConvertPDF_Click(Microsoft.Office.Core.IRibbonControl control)
        {
            Process.Start(@"C:\Program Files\Free PDF Solutions\FreePDFSolutions_PDF2Word.exe");
        }
        public void btnWizard_Click(Microsoft.Office.Core.IRibbonControl control)
        {
            Globals.ThisAddIn.Application.ActiveDocument.LoadWizard(Enums.FormLoadType.RibbonClick);
        }


        #endregion

        #region Helpers




        public Bitmap GetImage(Microsoft.Office.Core.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "btnHelp":
                    {
                        return new Bitmap(Properties.Resources.Help);
                    }

                case "btnWizard":
                    {
                        return new Bitmap(Properties.Resources.Wizard);
                    }
                case "btnConvertPDF":
                    {
                        return new Bitmap(Properties.Resources.pdf);
                    }
                case "btnUpdateFields":
                    {
                        return new Bitmap(Properties.Resources.calculator);
                    }
            }
            return null;

        }


        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion

        private void CalculatePremSummaryTable()
        {
            foreach (WordOM.Table table in Globals.ThisAddIn.Application.ActiveDocument.Tables)
            {
                if (String.Equals("Premium Summary", table.Title, StringComparison.OrdinalIgnoreCase))
                {
                    var totalRowCount = table.Rows.Count;
                    foreach (WordOM.Cell cell in table.Range.Cells)
                    {
                        if(cell.RowIndex == totalRowCount)
                        {
                            CalculateTotalsRow(table, cell, totalRowCount);
                        }
                    }

                    CheckForTableModifications(table);
                }
            }
        }

        private static void CheckForTableModifications(WordOM.Table table)
        {
            var includedCount =
                Globals.ThisAddIn.Application.ActiveDocument.GetPropertyValue(
                    Constants.WordDocumentProperties.IncludedPolicyTypesCount);

            int includedCountOut;

            if (int.TryParse(includedCount, out includedCountOut))
            {
                //the default table has 5 rows plus any row from the policy classes they insert
                if (table.Rows.Count - 5 != includedCountOut)
                {
                    MessageBox.Show(
                        @"Your premium summary table does not match the default table style.  Please check your calculated values are correct.",
                        @"Custom premium summary table", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private static void CalculateTotalsRow(WordOM.Table table, WordOM.Cell cell, int totalRowCount)
        {
            var m1 = String.Empty;
            var m2 = String.Empty;
            foreach (WordOM.Cell subCell in table.Range.Cells)
            {
                if (subCell.ColumnIndex == cell.ColumnIndex)
                {
                    if (subCell.RowIndex == totalRowCount - 1)
                    {
                        // m1 = Regex.Replace(subCell.Range.Text, @"[^\d]", "");

                        var m = Regex.Match(subCell.Range.Text, @"(\d+(\,\d+)+(\.\d+)?)|(\d+(\.\d+)?)");
                        if (m.Success) m1 = m.Value;
                    }
                    else if (subCell.RowIndex == totalRowCount - 2)
                    {
                        //m2 = Regex.Replace(subCell.Range.Text, @"[^\d]", "");
                        var m = Regex.Match(subCell.Range.Text, @"(\d+(\,\d+)+(\.\d+)?)|(\d+(\.\d+)?)");
                        if (m.Success) m2 = m.Value;
                    }
                }
            }
            decimal dm1;
            decimal dm2;
            if (decimal.TryParse(m1, out dm1) && decimal.TryParse(m2, out dm2))
            {
                cell.Range.Text = String.Format("{0:C}", dm1 + dm2);
            }
        }
    }
}
