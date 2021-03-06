﻿using System;
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
using OAMPS.Office.BusinessLogic.Interfaces.Wizards;
using OAMPS.Office.BusinessLogic.Loggers;
using OAMPS.Office.BusinessLogic.Models.Wizards;
using OAMPS.Office.Word.Helpers;
using OAMPS.Office.Word.Models.Word;
using OAMPS.Office.Word.Views.Help;
using Enums = OAMPS.Office.Word.Helpers.Enums;
using Office = Microsoft.Office.Core;
using WordOM = Microsoft.Office.Interop.Word;

namespace OAMPS.Office.Word
{
    [ComVisible(true)]
    public class Ribbon : Microsoft.Office.Core.IRibbonExtensibility
    {

        public static Microsoft.Office.Core.IRibbonUI ribbon;
        
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

        //public string btnTogleLockUnlock_GetLabel(Microsoft.Office.Core.IRibbonControl control)
        //{
        //    var storedMode = Globals.ThisAddIn.Application.ActiveDocument.GetPropertyValue(Constants.WordDocumentProperties.ToggleLockMode);
        //    return String.Equals(storedMode, Constants.WordDocumentPropertyValues.ToggleLockModeLocked, StringComparison.OrdinalIgnoreCase) ? "Unlock" : "Lock";
        //}

        //public void btnTogleLockUnlock_Click(Microsoft.Office.Core.IRibbonControl control, bool value)
        //{
        //    var storedMode = Globals.ThisAddIn.Application.ActiveDocument.GetPropertyValue(Constants.WordDocumentProperties.ToggleLockMode);

        //    var newMode = String.Equals(storedMode,Constants.WordDocumentPropertyValues.ToggleLockModeUnlocked) ?  Constants.WordDocumentPropertyValues.ToggleLockModeLocked : Constants.WordDocumentPropertyValues.ToggleLockModeUnlocked;
        //    Globals.ThisAddIn.Application.ActiveDocument.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.ToggleLockMode, newMode);
        //    ribbon.Invalidate();

        //    //and do some stuff here to modify document...
        //}

        //public bool btnTogleLockUnlock_GetPressed(Microsoft.Office.Core.IRibbonControl control)
        //{
        //    var storedMode = Globals.ThisAddIn.Application.ActiveDocument.GetPropertyValue(Constants.WordDocumentProperties.ToggleLockMode);
        //    return !String.Equals(storedMode, Constants.WordDocumentPropertyValues.ToggleLockModeLocked, StringComparison.OrdinalIgnoreCase);
        //}

        //public bool btnToggleLock_GetEnabled(Microsoft.Office.Core.IRibbonControl control)
        //{
        //    var storedMode = Globals.ThisAddIn.Application.ActiveDocument.GetPropertyValue(Constants.WordDocumentProperties.ToggleLockMode);
        //    return String.Equals(storedMode, Constants.WordDocumentPropertyValues.ToggleLockModeUnlocked, StringComparison.OrdinalIgnoreCase);
        //}

        //public bool btnToggleUnlock_GetEnabled(Microsoft.Office.Core.IRibbonControl control)
        //{
        //        var storedMode = Globals.ThisAddIn.Application.ActiveDocument.GetPropertyValue(Constants.WordDocumentProperties.ToggleLockMode);
        //        return String.Equals(storedMode, Constants.WordDocumentPropertyValues.ToggleLockModeLocked, StringComparison.OrdinalIgnoreCase) || String.Equals(storedMode, String.Empty, StringComparison.OrdinalIgnoreCase);
        //}

        public bool btnUpdateFields_GetEnabled(Microsoft.Office.Core.IRibbonControl control)
        {
            if (Globals.ThisAddIn.Application.Documents.Count < 1) return false;

            var templateName = ((DocumentProperties)(Globals.ThisAddIn.Application.ActiveDocument.BuiltInDocumentProperties))[WordOM.WdBuiltInProperty.wdPropertyTitle].Value.ToString();
            return String.Equals(templateName, Constants.TemplateNames.InsuranceRenewalReport);
        }

        public bool btnSyncTable_GetEnabled(Microsoft.Office.Core.IRibbonControl control)
        {
            if (Globals.ThisAddIn.Application.Documents.Count < 1) return false;

            var templateName = ((DocumentProperties)(Globals.ThisAddIn.Application.ActiveDocument.BuiltInDocumentProperties))[WordOM.WdBuiltInProperty.wdPropertyTitle].Value.ToString();
            return String.Equals(templateName, Constants.TemplateNames.PreRenewalQuestionnaire);
        }

        public void Ribbon_Load(Microsoft.Office.Core.IRibbonUI ribbonUI)
        {
            ribbon = ribbonUI;
        }

        public void btnAbout_Click(Microsoft.Office.Core.IRibbonControl control)
        {
            var about = new About();
            about.ShowDialog();
        }

        public void btnUnlock_Click(Microsoft.Office.Core.IRibbonControl control)
        {
            foreach (WordOM.Table table in Globals.ThisAddIn.Application.ActiveDocument.Tables)
            {
                try
                {

                    var controls = table.Range.ContentControls;
                    if (controls == null || controls.Count == 0) continue;
                    controls[1].Ungroup();
                    //Globals.ThisAddIn.Application.Selection.Range.ParentContentControl.Ungroup();    
                }

                catch (Exception e)
                {

#if DEBUG
                    MessageBox.Show(e.ToString(), @"sorry");
#endif
                }
            }

            var storedMode = Globals.ThisAddIn.Application.ActiveDocument.GetPropertyValue(Constants.WordDocumentProperties.ToggleLockMode);
            var newMode = String.Equals(storedMode, Constants.WordDocumentPropertyValues.ToggleLockModeUnlocked) ? Constants.WordDocumentPropertyValues.ToggleLockModeLocked : Constants.WordDocumentPropertyValues.ToggleLockModeUnlocked;
            Globals.ThisAddIn.Application.ActiveDocument.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.ToggleLockMode, newMode);
            ribbon.Invalidate();

        }

        public void btnLock_Click(Microsoft.Office.Core.IRibbonControl control)
        {

            foreach (WordOM.Table table in Globals.ThisAddIn.Application.ActiveDocument.Tables)
            {
                try
                {
                    var controls = table.Range.ContentControls;
                    if (controls == null || controls.Count == 0) continue;

                    table.Select();
                    Globals.ThisAddIn.Application.Selection.Range.ContentControls.Add(WordOM.WdContentControlType.wdContentControlGroup);
                }
                catch (Exception e)
                {

#if DEBUG
                    MessageBox.Show(e.ToString(), @"sorry");
#endif
                }
            }


            var storedMode = Globals.ThisAddIn.Application.ActiveDocument.GetPropertyValue(Constants.WordDocumentProperties.ToggleLockMode);
            var newMode = String.Equals(storedMode, Constants.WordDocumentPropertyValues.ToggleLockModeUnlocked) ? Constants.WordDocumentPropertyValues.ToggleLockModeLocked : Constants.WordDocumentPropertyValues.ToggleLockModeUnlocked;
            Globals.ThisAddIn.Application.ActiveDocument.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.ToggleLockMode, newMode);
            ribbon.Invalidate();

        }

        public void btnSyncTable_Click(Microsoft.Office.Core.IRibbonControl control)
        {
            var doc = new OfficeDocument(Globals.ThisAddIn.Application.ActiveDocument);
            var proptectionType = -1;
            try
            {

                //proptectionType = doc.TurnOffProtection(String.Empty);

                var tables = Globals.ThisAddIn.Application.Selection.Tables;
                if (tables == null || tables.Count == 0)
                {
                    MessageBox.Show(@"Please ensure your cursor is within a table", @"No Table Found", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                
                var table = tables[1];
                

                var contentcontrols = table.Range.ContentControls;
                if (contentcontrols == null) return;

                foreach (WordOM.ContentControl contentControl in contentcontrols)
                {
                    if (String.IsNullOrEmpty(contentControl.Tag)) continue;

                    var matchingControls =
                   Globals.ThisAddIn.Application.ActiveDocument.SelectContentControlsByTag(contentControl.Tag);
                    foreach (WordOM.ContentControl mc in matchingControls)
                    {
                        mc.Range.Text = contentControl.Range.Text;
                    }
                }
            }
            catch (Exception e)
            {
                var logger = new EventViewerLogger();
                logger.Log(e.ToString(), BusinessLogic.Interfaces.Logging.Type.Error);

#if DEBUG
                MessageBox.Show(e.ToString(), @"sorry");
#endif

            }
            finally
            {
             //   doc.TurnOnProtection(proptectionType, String.Empty);
            }
        }

        public void btnSyncField_Click(Microsoft.Office.Core.IRibbonControl control)
        {
            var doc = new OfficeDocument(Globals.ThisAddIn.Application.ActiveDocument);
            var proptectionType = -1;
            try
            {
               
               proptectionType =  doc.TurnOffProtection(String.Empty);
                var contentcontrol = Globals.ThisAddIn.Application.Selection.Range.ParentContentControl;
                if (contentcontrol == null) return;
                var matchingControls =
                    Globals.ThisAddIn.Application.ActiveDocument.SelectContentControlsByTag(contentcontrol.Tag);
                foreach (WordOM.ContentControl mc in matchingControls)
                {
                    mc.Range.Text = contentcontrol.Range.Text;
                }
            }
            catch (Exception e)
            {
                var logger = new EventViewerLogger();
                logger.Log(e.ToString(), BusinessLogic.Interfaces.Logging.Type.Error);

#if DEBUG
                MessageBox.Show(e.ToString(), @"sorry");
#endif

            }
            finally
            {
                doc.TurnOnProtection(proptectionType, String.Empty);
            }
        }

        public void btnUpdateFields_Click(Microsoft.Office.Core.IRibbonControl control)
        {
            var tableCount = SummaryTableCount();

            if (tableCount != 1)
            {
                MessageBox.Show(String.Format("There is {0} no {1} table in this document.", tableCount, Constants.WordTables.RenewalReportPremiumSummary), @"No changes made.", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Globals.ThisAddIn.Application.ActiveDocument.Fields.Update();
            CalculatePremSummaryTable();
            SyncPremiumSummaryTable();
        }

        public void btnHelp_Click(Microsoft.Office.Core.IRibbonControl control)
        {
            ShowHelp();
        }

        private static void ShowHelp()
        {
            var templateName =
                ((DocumentProperties) (Globals.ThisAddIn.Application.ActiveDocument.BuiltInDocumentProperties))[
                    WordOM.WdBuiltInProperty.wdPropertyTitle].Value.ToString();
            var wizard = new HelpWizard(templateName)
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

                case "btnToggleLock":
                    {
                        return new Bitmap(Properties.Resources.Lock);
                    }

                case "btnTogleLockUnlock":
                    {
                        var mode = Globals.ThisAddIn.Application.ActiveDocument.GetPropertyValue(Constants.WordDocumentProperties.ToggleLockMode);
                        return String.Equals(mode,Constants.WordDocumentPropertyValues.ToggleLockModeLocked,StringComparison.OrdinalIgnoreCase) ? new Bitmap(Properties.Resources.Unlock) : new Bitmap(Properties.Resources.Lock);
                    }

                case "btnTogleUnlock":
                    {
                        return new Bitmap(Properties.Resources.Unlock);
                    }

                case "btnHelpHome":
                    {
                        return new Bitmap(Properties.Resources.Help);
                    }

                case "btnSyncField":
                    {
                        return new Bitmap(Properties.Resources.SyncField);
                    }

                case "btnSyncTableHome":
                    {
                        return new Bitmap(Properties.Resources.SyncTableInfo);
                    }
                    

                case "btnSyncTable":
                    {
                        return new Bitmap(Properties.Resources.SyncTableInfo);
                    }

                case "btnWizard":
                    {
                        return new Bitmap(Properties.Resources.Wizard);
                    }
                case "btnWizardHome":
                    {
                        return new Bitmap(Properties.Resources.Wizard);
                    }

                case "btnConvertPDF":
                    {
                        return new Bitmap(Properties.Resources.pdf);
                    }
                case "btnConvertPDFHome":
                    {
                        return new Bitmap(Properties.Resources.pdf);
                    }
                case "btnUpdateFields":
                    {
                        return new Bitmap(Properties.Resources.calculator);
                    }
                case "btnUpdateFieldsHome":
                    {
                        return new Bitmap(Properties.Resources.calculator);
                    }

                case "btnAbout":
                    {
                        return new Bitmap(Properties.Resources.About);
                    }
                case "btnAboutHome":
                    {
                        return new Bitmap(Properties.Resources.About);
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


        private int SummaryTableCount()
        {
            var list = new List<int>();
            foreach (WordOM.Table table in Globals.ThisAddIn.Application.ActiveDocument.Tables)
            {
                if (String.Equals(Constants.WordTables.RenewalReportPremiumSummary, table.Title, StringComparison.OrdinalIgnoreCase)) list.Add(1);
            }
            return list.Count;
        }

        private void SyncPremiumSummaryTable()
        {
            var allRows = new List<IPremiumSummaryTableRow>();

            //store cell indexs
            var classOfInsuranceCIndex = -1;
            var recommendedInsurerCIndex = -1;
            var basePremiumCIndex = -1;
            var policyUnderwriterGstIndex = -1;
            var fireServicesLeviesCIndex = -1;
            var totalGstCIndex = -1;
            var stampDutyCIndex = -1;
            var otherTaxesChargesCIndex = -1;
            var brokerFeePerPolicyClassCIndex = -1;


            var headerRowIndex = -1;

            //first build table object
            foreach (WordOM.Table premiumSummary in Globals.ThisAddIn.Application.ActiveDocument.Tables)
            {
                if (premiumSummary.Title != null && String.Equals(Constants.WordTables.RenewalReportPremiumSummary, premiumSummary.Title, StringComparison.OrdinalIgnoreCase))
                {
                    foreach (WordOM.Cell headerCell in premiumSummary.Range.Cells)
                    {
                        switch (headerCell.Range.Text.Replace("\r\a", string.Empty))
                        {
                            case "Class of insurance":
                                classOfInsuranceCIndex = headerCell.ColumnIndex;
                                break;

                            case "Recommended insurer":
                                recommendedInsurerCIndex = headerCell.ColumnIndex;
                                break;

                            case "Base premium":
                                basePremiumCIndex = headerCell.ColumnIndex;
                                break;

                            case "Policy (Underwriter) GST":
                                policyUnderwriterGstIndex = headerCell.ColumnIndex;
                                break;

                            case "Fire services levies":
                                fireServicesLeviesCIndex = headerCell.ColumnIndex;
                                break;

                            case "Broker fee GST":
                                totalGstCIndex = headerCell.ColumnIndex;
                                break;

                            case "Stamp duties":
                                stampDutyCIndex = headerCell.ColumnIndex;
                                break;

                            case "Other taxes/ charges":
                                otherTaxesChargesCIndex = headerCell.ColumnIndex;
                                break;

                            case "Broker fee per policy class":
                                brokerFeePerPolicyClassCIndex = headerCell.ColumnIndex;
                                break;

                        }
                    }


                    var previous = -1;
                    PremiumSummaryTableRow pr = null;

                    foreach (WordOM.Cell cc in premiumSummary.Range.Cells)
                    {
                        if (cc.RowIndex == previous)
                        {
                            if (pr == null)
                                continue;

                            if (cc.ColumnIndex == classOfInsuranceCIndex)
                                pr.ClassOfInsurance = cc.Range.Text;

                            else if (cc.ColumnIndex == recommendedInsurerCIndex)
                                pr.RecommendedInsurer = cc.Range.Text;

                            else if (cc.ColumnIndex == basePremiumCIndex)
                                pr.BasePremium = cc.Range.Text;

                            else if (cc.ColumnIndex == fireServicesLeviesCIndex)
                                pr.FireServicesLevies = cc.Range.Text;

                            else if (cc.ColumnIndex == totalGstCIndex)
                                pr.TotalGST = cc.Range.Text;

                            else if (cc.ColumnIndex == policyUnderwriterGstIndex)
                                pr.PolicyUnderwriterGst = cc.Range.Text;

                            else if (cc.ColumnIndex == stampDutyCIndex)
                                pr.StampDuty = cc.Range.Text;

                            else if (cc.ColumnIndex == otherTaxesChargesCIndex)
                                pr.OtherTaxesCharges = cc.Range.Text;

                            else if (cc.ColumnIndex == brokerFeePerPolicyClassCIndex)
                                pr.BrokerFeePerPolicyClass = cc.Range.Text;
                        }
                        else
                        {
                            if (pr != null)
                                allRows.Add(pr);

                            pr = new PremiumSummaryTableRow();
                        }
                        previous = cc.RowIndex;
                    }
                }
            }

            foreach (WordOM.Table premiumCosts in Globals.ThisAddIn.Application.ActiveDocument.Tables)
            {
                if (premiumCosts.Title != null && premiumCosts.Title.Contains(Constants.WordTables.RenewalReportPremiumCosts + "_"))
                {
                    var classOfInsurance = premiumCosts.Title.Split('_');
                    if (classOfInsurance.Length == 2)
                    {
                        var premRow = allRows.FirstOrDefault(i => i.ClassOfInsurance != null && String.Equals(i.ClassOfInsurance.Replace("\r\a", string.Empty), classOfInsurance[1], StringComparison.OrdinalIgnoreCase));
                        if (premRow != null)
                        {
                            premiumCosts.Rows[basePremiumCIndex - 3].Cells[3].Range.Text = premRow.BasePremium.Replace("\r\a", string.Empty);
                            premiumCosts.Rows[fireServicesLeviesCIndex - 3].Cells[3].Range.Text = premRow.FireServicesLevies.Replace("\r\a", string.Empty);
                            premiumCosts.Rows[totalGstCIndex - 3].Cells[3].Range.Text = premRow.TotalGST.Replace("\r\a", string.Empty);

                            premiumCosts.Rows[policyUnderwriterGstIndex - 3].Cells[3].Range.Text = premRow.PolicyUnderwriterGst.Replace("\r\a", string.Empty);
                            premiumCosts.Rows[stampDutyCIndex - 3].Cells[3].Range.Text = premRow.StampDuty.Replace("\r\a", string.Empty);
                            premiumCosts.Rows[otherTaxesChargesCIndex - 3].Cells[3].Range.Text = premRow.OtherTaxesCharges.Replace("\r\a", string.Empty);
                            premiumCosts.Rows[brokerFeePerPolicyClassCIndex - 3].Cells[3].Range.Text = premRow.BrokerFeePerPolicyClass.Replace("\r\a", string.Empty);

                            //now do the total row

                            premiumCosts.Rows[8].Cells[3].Range.Text =  "$" + CalculateTotalFromPremiumSummaryRow(premRow).ToString();
                        }
                    }
                }
            }
        }

        private decimal CalculateTotalFromPremiumSummaryRow(IPremiumSummaryTableRow row)
        {
            bool anyFailed = false;

            //cleanup
            row.BasePremium = row.BasePremium.Replace("\r\a", string.Empty);
            row.BasePremium = row.BasePremium.Replace("$", string.Empty);

            row.BrokerFeePerPolicyClass = row.BrokerFeePerPolicyClass.Replace("\r\a", string.Empty);
            row.BrokerFeePerPolicyClass = row.BrokerFeePerPolicyClass.Replace("$", string.Empty);

            row.FireServicesLevies = row.FireServicesLevies.Replace("\r\a", string.Empty);
            row.FireServicesLevies = row.FireServicesLevies.Replace("$", string.Empty);

            row.OtherTaxesCharges = row.OtherTaxesCharges.Replace("\r\a", string.Empty);
            row.OtherTaxesCharges = row.OtherTaxesCharges.Replace("$", string.Empty);

            row.OtherTaxesCharges = row.OtherTaxesCharges.Replace("\r\a", string.Empty);
            row.OtherTaxesCharges = row.OtherTaxesCharges.Replace("$", string.Empty);

            row.PolicyUnderwriterGst = row.PolicyUnderwriterGst.Replace("\r\a", string.Empty);
            row.PolicyUnderwriterGst = row.PolicyUnderwriterGst.Replace("$", string.Empty);

            row.StampDuty = row.StampDuty.Replace("\r\a", string.Empty);
            row.StampDuty = row.StampDuty.Replace("$", string.Empty);

            row.TotalGST = row.TotalGST.Replace("\r\a", string.Empty);
            row.TotalGST = row.TotalGST.Replace("$", string.Empty);

            //convert to double
            decimal dBasePrem; if (!decimal.TryParse(row.BasePremium, out dBasePrem)) anyFailed = true;
            decimal dbrokerPolicyFee; if (!decimal.TryParse(row.BrokerFeePerPolicyClass, out dbrokerPolicyFee)) anyFailed = true;
            decimal dFireServices; if (!decimal.TryParse(row.FireServicesLevies, out dFireServices)) anyFailed = true;
            decimal dOtherTaxes; if (!decimal.TryParse(row.OtherTaxesCharges, out dOtherTaxes)) anyFailed = true;
            decimal dPolicyUnderwriterGst; if (!decimal.TryParse(row.PolicyUnderwriterGst, out dPolicyUnderwriterGst)) anyFailed = true;
            decimal dStampDuty; if (!decimal.TryParse(row.StampDuty, out dStampDuty)) anyFailed = true;
            decimal dTotalGst; if (!decimal.TryParse(row.TotalGST, out dTotalGst)) anyFailed = true;

            if (anyFailed)
                return 0;
            
            return dBasePrem + dbrokerPolicyFee + dFireServices + dOtherTaxes + dPolicyUnderwriterGst + dStampDuty + dTotalGst;
        }

        private void CalculatePremSummaryTable()
        {
            var found = false;

            foreach (WordOM.Table table in Globals.ThisAddIn.Application.ActiveDocument.Tables)
            {
                if (String.Equals(Constants.WordTables.RenewalReportPremiumSummary, table.Title, StringComparison.OrdinalIgnoreCase))
                {
                    found = true;

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

            if (!found)
                MessageBox.Show(String.Format(@"There is no {0} table in this document.",Constants.WordTables.RenewalReportPremiumSummary), @"No changes made.", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                    MessageBox.Show(@"Your " + Constants.WordTables.RenewalReportPremiumSummary + @" table does not match the default table style.  Please check your calculated values are correct.",
                        @"Custom " + Constants.WordTables.RenewalReportPremiumSummary +  @" table", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
