using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Caching;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Core;
using OAMPS.Office.BusinessLogic.Helpers;
using OAMPS.Office.BusinessLogic.Interfaces.Template;
using OAMPS.Office.BusinessLogic.Interfaces.Wizards;
using OAMPS.Office.BusinessLogic.Loggers;
using OAMPS.Office.BusinessLogic.Models.Template;
using OAMPS.Office.BusinessLogic.Models.Wizards;
using OAMPS.Office.BusinessLogic.Presenters.SharePoint;
using OAMPS.Office.BusinessLogic.Presenters.Wizards;
using OAMPS.Office.Word.Helpers;
using OAMPS.Office.Word.Helpers.LocalSharePoint;
using OAMPS.Office.Word.Models.SharePoint;
using OAMPS.Office.Word.Models.Word;
using OAMPS.Office.Word.Properties;
using OAMPS.Office.Word.Views.Help;
using OAMPS.Office.Word.Views.Wizards;
using OAMPS.Office.Word.Views.Word;
using Enums = OAMPS.Office.Word.Helpers.Enums;
using Type = OAMPS.Office.BusinessLogic.Interfaces.Logging.Type;
using WordOM = Microsoft.Office.Interop.Word;

namespace OAMPS.Office.Word
{
    [ComVisible(true)]
    public class Ribbon : IRibbonExtensibility
    {
        internal static IRibbonUI ribbon;

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonId)
        {
            return GetResourceText("OAMPS.Office.Word.Ribbon.xml");
        }

        #endregion

        private int SummaryTableCount()
        {
            var list = new List<int>();
// ReSharper disable LoopCanBeConvertedToQuery
            foreach (WordOM.Table table in Globals.ThisAddIn.Application.ActiveDocument.Tables)
// ReSharper restore LoopCanBeConvertedToQuery
            {
                if (string.Equals(Constants.WordTables.RenewalReportPremiumSummary, table.Title,
                    StringComparison.OrdinalIgnoreCase)) list.Add(1);
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


            //      int headerRowIndex = -1;

            //first build table object
            foreach (WordOM.Table premiumSummary in Globals.ThisAddIn.Application.ActiveDocument.Tables)
            {
                if (premiumSummary.Title != null &&
                    string.Equals(Constants.WordTables.RenewalReportPremiumSummary, premiumSummary.Title,
                        StringComparison.OrdinalIgnoreCase))
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
                                pr.TotalGst = cc.Range.Text;

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
                if (premiumCosts.Title != null &&
                    premiumCosts.Title.Contains(Constants.WordTables.RenewalReportPremiumCosts + "_"))
                {
                    var classOfInsurance = premiumCosts.Title.Split('_');
                    if (classOfInsurance.Length == 2)
                    {
                        var premRow =
                            allRows.FirstOrDefault(
                                i =>
                                    i.ClassOfInsurance != null &&
                                    string.Equals(i.ClassOfInsurance.Replace("\r\a", string.Empty), classOfInsurance[1],
                                        StringComparison.OrdinalIgnoreCase));
                        if (premRow != null)
                        {
                            premiumCosts.Rows[basePremiumCIndex - 3].Cells[3].Range.Text =
                                premRow.BasePremium.Replace("\r\a", string.Empty);
                            premiumCosts.Rows[fireServicesLeviesCIndex - 3].Cells[3].Range.Text =
                                premRow.FireServicesLevies.Replace("\r\a", string.Empty);
                            premiumCosts.Rows[totalGstCIndex - 3].Cells[3].Range.Text = premRow.TotalGst.Replace(
                                "\r\a", string.Empty);

                            premiumCosts.Rows[policyUnderwriterGstIndex - 3].Cells[3].Range.Text =
                                premRow.PolicyUnderwriterGst.Replace("\r\a", string.Empty);
                            premiumCosts.Rows[stampDutyCIndex - 3].Cells[3].Range.Text =
                                premRow.StampDuty.Replace("\r\a", string.Empty);
                            premiumCosts.Rows[otherTaxesChargesCIndex - 3].Cells[3].Range.Text =
                                premRow.OtherTaxesCharges.Replace("\r\a", string.Empty);
                            premiumCosts.Rows[brokerFeePerPolicyClassCIndex - 3].Cells[3].Range.Text =
                                premRow.BrokerFeePerPolicyClass.Replace("\r\a", string.Empty);

                            //now do the total row
                            var c = CalculateTotalFromPremiumSummaryRow(premRow);
                            premiumCosts.Rows[8].Cells[3].Range.Text = string.Format("{0:C}", c);
                        }
                    }
                }
            }
        }

        private decimal CalculateTotalFromPremiumSummaryRow(IPremiumSummaryTableRow row)
        {
            var anyFailed = false;

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

            row.TotalGst = row.TotalGst.Replace("\r\a", string.Empty);
            row.TotalGst = row.TotalGst.Replace("$", string.Empty);

            //convert to double
            decimal dBasePrem;
            if (!decimal.TryParse(row.BasePremium, out dBasePrem)) anyFailed = true;
            decimal dbrokerPolicyFee;
            if (!decimal.TryParse(row.BrokerFeePerPolicyClass, out dbrokerPolicyFee)) anyFailed = true;
            decimal dFireServices;
            if (!decimal.TryParse(row.FireServicesLevies, out dFireServices)) anyFailed = true;
            decimal dOtherTaxes;
            if (!decimal.TryParse(row.OtherTaxesCharges, out dOtherTaxes)) anyFailed = true;
            decimal dPolicyUnderwriterGst;
            if (!decimal.TryParse(row.PolicyUnderwriterGst, out dPolicyUnderwriterGst)) anyFailed = true;
            decimal dStampDuty;
            if (!decimal.TryParse(row.StampDuty, out dStampDuty)) anyFailed = true;
            decimal dTotalGst;
            if (!decimal.TryParse(row.TotalGst, out dTotalGst)) anyFailed = true;

            if (anyFailed)
                return 0;

            return dBasePrem + dbrokerPolicyFee + dFireServices + dOtherTaxes + dPolicyUnderwriterGst + dStampDuty +
                   dTotalGst;
        }

        private void CalculatePremSummaryTable()
        {
            var found = false;

            foreach (WordOM.Table table in Globals.ThisAddIn.Application.ActiveDocument.Tables)
            {
                if (string.Equals(Constants.WordTables.RenewalReportPremiumSummary, table.Title,
                    StringComparison.OrdinalIgnoreCase))
                {
                    found = true;

                    var totalRowCount = table.Rows.Count;
                    foreach (WordOM.Cell cell in table.Range.Cells)
                    {
                        if (cell.RowIndex == totalRowCount)
                        {
                            CalculateTotalsRow(table, cell, totalRowCount);
                        }
                    }
                    CheckForTableModifications(table);
                }
            }

            if (!found)
                MessageBox.Show(
                    string.Format(@"There is no {0} table in this document.",
                        Constants.WordTables.RenewalReportPremiumSummary), @"No changes made.", MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
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
                        @"Your " + Constants.WordTables.RenewalReportPremiumSummary +
                        @" table does not match the default table style.  Please check your calculated values are correct.",
                        @"Custom " + Constants.WordTables.RenewalReportPremiumSummary + @" table", MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }
            }
        }

        private static void CalculateTotalsRow(WordOM.Table table, WordOM.Cell cell, int totalRowCount)
        {
            var m1 = string.Empty;
            var m2 = string.Empty;
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
                cell.Range.Text = string.Format("{0:C}", dm1 + dm2);
            }
        }

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

        public bool btnUpdateFields_GetEnabled(IRibbonControl control)
        {
            if (Globals.ThisAddIn.Application.Documents.Count < 1) return false;

            dynamic templateName =
                ((DocumentProperties) (Globals.ThisAddIn.Application.ActiveDocument.BuiltInDocumentProperties))[
                    WordOM.WdBuiltInProperty.wdPropertyTitle].Value.ToString();
            return (string.Equals(templateName, Constants.TemplateNames.InsuranceRenewalReport) ||
                    string.Equals(templateName, Constants.TemplateNames.ShortFormProposal));
        }

        public bool btnFactFinderButton_GetEnabled(IRibbonControl control)
        {
            try
            {
                dynamic templateName =
                    ((DocumentProperties) (Globals.ThisAddIn.Application.ActiveDocument.BuiltInDocumentProperties))[
                        WordOM.WdBuiltInProperty.wdPropertyTitle].Value.ToString();
                return string.Equals(templateName, Constants.TemplateNames.PreRenewalQuestionnaire) ||
                       string.Equals(templateName, Constants.TemplateNames.BlankFragmentBaseOlly);
            }
            catch
            {
                return false;
            }
        }

        public bool btnRenewalReportButton_GetEnabled(IRibbonControl control)
        {
            try
            {
                dynamic templateName =
                    ((DocumentProperties) (Globals.ThisAddIn.Application.ActiveDocument.BuiltInDocumentProperties))[
                        WordOM.WdBuiltInProperty.wdPropertyTitle].Value.ToString();
                return string.Equals(templateName, Constants.TemplateNames.InsuranceRenewalReport);
            }
            catch
            {
                return false;
            }
        }


        public bool btnPlacementSlip_GetEnabled(IRibbonControl control)
        {
            try
            {
                dynamic templateName =
                    ((DocumentProperties) (Globals.ThisAddIn.Application.ActiveDocument.BuiltInDocumentProperties))[
                        WordOM.WdBuiltInProperty.wdPropertyTitle].Value.ToString();
                return string.Equals(templateName, Constants.TemplateNames.QuoteSlip);
            }
            catch
            {
                return false;
            }
        }

        public void Ribbon_Load(IRibbonUI ribbonUi)
        {
            ribbon = ribbonUi;
        }

        public void btnAbout_Click(IRibbonControl control)
        {
            var about = new About();
            about.ShowDialog();
        }

        public void btnInsertPortrait_Click(IRibbonControl control)
        {
            InsertPortaitPage();
        }

        public void btnInsertLandscape_Click(IRibbonControl control)
        {
            InsertLandscapePage();
        }

        public void InsertLandscapePage()
        {
            //todo: turn off protection

            //var p = reportDoc.TurnOffProtection(string.Empty);
            //presenterManualDoc.PopulateData(reportTemplate);

            //manualDoc.PopulateControl(Constants.WordContentControls.DocumentTitle, "Insurance Manual");
            //reportDoc.TurnOnProtection(p, string.Empty);

            var currentOrientation = Globals.ThisAddIn.Application.Selection.PageSetup.Orientation;
            Globals.ThisAddIn.Application.Selection.InsertBreak(WordOM.WdBreakType.wdPageBreak);
            Globals.ThisAddIn.Application.Selection.InsertParagraphAfter();
            Globals.ThisAddIn.Application.Selection.InsertBreak(WordOM.WdBreakType.wdSectionBreakNextPage);
            Globals.ThisAddIn.Application.Selection.PageSetup.Orientation = WordOM.WdOrientation.wdOrientLandscape;
            Globals.ThisAddIn.Application.Selection.InsertBreak(WordOM.WdBreakType.wdSectionBreakNextPage);

            if (currentOrientation == WordOM.WdOrientation.wdOrientPortrait)
                Globals.ThisAddIn.Application.Selection.PageSetup.Orientation = WordOM.WdOrientation.wdOrientPortrait;

            //todo: turn off protection
        }

        public void InsertPortaitPage()
        {
            var currentOrientation = Globals.ThisAddIn.Application.Selection.PageSetup.Orientation;
            Globals.ThisAddIn.Application.Selection.InsertBreak(WordOM.WdBreakType.wdPageBreak);
            Globals.ThisAddIn.Application.Selection.InsertParagraphAfter();
            Globals.ThisAddIn.Application.Selection.InsertBreak(WordOM.WdBreakType.wdSectionBreakNextPage);
            Globals.ThisAddIn.Application.Selection.PageSetup.Orientation = WordOM.WdOrientation.wdOrientPortrait;
            Globals.ThisAddIn.Application.Selection.InsertBreak(WordOM.WdBreakType.wdSectionBreakNextPage);

            if (currentOrientation == WordOM.WdOrientation.wdOrientLandscape)
                Globals.ThisAddIn.Application.Selection.PageSetup.Orientation = WordOM.WdOrientation.wdOrientLandscape;
        }

        public void btnUnlock_Click(IRibbonControl control)
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

            var storedMode =
                Globals.ThisAddIn.Application.ActiveDocument.GetPropertyValue(
                    Constants.WordDocumentProperties.ToggleLockMode);
            var newMode = string.Equals(storedMode, Constants.WordDocumentPropertyValues.ToggleLockModeUnlocked)
                ? Constants.WordDocumentPropertyValues.ToggleLockModeLocked
                : Constants.WordDocumentPropertyValues.ToggleLockModeUnlocked;
            Globals.ThisAddIn.Application.ActiveDocument.UpdateOrCreatePropertyValue(
                Constants.WordDocumentProperties.ToggleLockMode, newMode);
            ribbon.Invalidate();
        }

        public void btnOpenBICalculator_Click(IRibbonControl control)
        {
            var list = ListFactory.Create("Configuration");
            var presenter = new SharePointListPresenter(list, null);
            var item = presenter.GetItemByTitle("BICalculator.HttpAddress");
            var address = item.GetFieldValue("Value");
            if (string.IsNullOrEmpty(address)) return;
            Process.Start(address);
        }

        public void btnLock_Click(IRibbonControl control)
        {
            foreach (WordOM.Table table in Globals.ThisAddIn.Application.ActiveDocument.Tables)
            {
                try
                {
                    var controls = table.Range.ContentControls;
                    if (controls == null || controls.Count == 0) continue;

                    table.Select();
                    Globals.ThisAddIn.Application.Selection.Range.ContentControls.Add(
                        WordOM.WdContentControlType.wdContentControlGroup);
                }
                catch (Exception e)
                {
#if DEBUG
                    MessageBox.Show(e.ToString(), @"sorry");
#endif
                }
            }


            var storedMode =
                Globals.ThisAddIn.Application.ActiveDocument.GetPropertyValue(
                    Constants.WordDocumentProperties.ToggleLockMode);
            var newMode = string.Equals(storedMode, Constants.WordDocumentPropertyValues.ToggleLockModeUnlocked)
                ? Constants.WordDocumentPropertyValues.ToggleLockModeLocked
                : Constants.WordDocumentPropertyValues.ToggleLockModeUnlocked;
            Globals.ThisAddIn.Application.ActiveDocument.UpdateOrCreatePropertyValue(
                Constants.WordDocumentProperties.ToggleLockMode, newMode);
            ribbon.Invalidate();
        }

        private class PlacementSlip : IBaseTemplate
        {
            public string ClientName { get; set; }
            public string ClientCommonName { get; set; }
            public string OAMPSBranchAddress { get; set; }
            public string OAMPSBranchAddressLine2 { get; set; }
            public string Fax { get; set; }
            public string DocumentTitle { get; set; }
            public string DocumentSubTitle { get; set; }
            public string CoverPageImageUrl { get; set; }
            public string LogoImageUrl { get; set; }
            public string CoverPageTitle { get; set; }
            public string LogoTitle { get; set; }
            public string OAMPSCompanyName { get; set; }
            public string OAMPSAbnNumber { get; set; }
            public string OAMPSBranchPhone { get; set; }
            public string OAMPSAfsl { get; set; }
            public string WebSite { get; set; }
            public string ExecutiveName { get; set; }
            public string ExecutiveTitle { get; set; }
            public string ExecutivePhone { get; set; }
            public string ExecutiveMobile { get; set; }
            public string ExecutiveEmail { get; set; }
            public string ExecutiveDepartment { get; set; }
            public string LongBrandingDescription { get; set; }
        }

        public void btnConvertToPlacementSlip_Click(IRibbonControl control)
        {
            var docTitle =
                ((DocumentProperties) (Globals.ThisAddIn.Application.ActiveDocument.BuiltInDocumentProperties))[
                    WordOM.WdBuiltInProperty.wdPropertyTitle].Value;

            if (Constants.TemplateNames.PlacementSlip.Equals(docTitle.ToString(), StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show(@"Document has already been converted into a Placement Slip",
                    "Document is a Placement Slip", MessageBoxButtons.OK, MessageBoxIcon.Information);
                ribbon.Invalidate();
                return;
            }

            var doc = new OfficeDocument(Globals.ThisAddIn.Application.ActiveDocument);
            var d = Globals.ThisAddIn.Application.ActiveDocument;
            doc.PopulateControl(Constants.WordContentControls.DocumentTitle, Constants.TemplateNames.PlacementSlip);
            doc.MoveCursorToStartOfControl(Constants.WordContentControls.Instructions);
            doc.DeleteControl(Constants.WordContentControls.Instructions);
            doc.DeleteCharacter(1);
            doc.InsertFile(Settings.Default.PlacementSlipConditionsFragement);

            ((DocumentProperties) (Globals.ThisAddIn.Application.ActiveDocument.BuiltInDocumentProperties))[
                WordOM.WdBuiltInProperty.wdPropertyTitle].Value = Constants.TemplateNames.PlacementSlip;
            doc.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.BuiltInTitle,
                Constants.TemplateNames.PlacementSlip);
            ribbon.Invalidate();

            var s = doc.GetBookmarkStartRange(Constants.WordBookmarks.QuotationDetailsStart);
            var e = doc.GetBookmarkEndRange(Constants.WordBookmarks.QuotationDetailsEnd);
            doc.DeleteRange(s, e);
            if (doc.MoveCursorToStartOfBookmark(Constants.WordBookmarks.QuotationDetailsEnd))
            {
                doc.DeletePage();
                doc.DeleteCharacter(1);
            }

            var w = new BaseWizardForm();
            var t = new PlacementSlip();
            t.DocumentTitle = "Placement Slip";
            w.LogUsage(t, Enums.UsageTrackingType.ConvertDocument);
        }


        public void btnConvertToManual_Click(IRibbonControl control)
        {
            var reportDoc = new OfficeDocument(Globals.ThisAddIn.Application.ActiveDocument);
            var presenterreportDoc = new InsuranceRenewealReportWizardPresenter(reportDoc, null);

            var reportTemplate = new InsuranceRenewalReport();
            reportTemplate = (InsuranceRenewalReport) presenterreportDoc.LoadData(reportTemplate);

            reportTemplate = (InsuranceRenewalReport) presenterreportDoc.LoadIncludedPolicyClasses(reportTemplate);


            ObjectCache cache = MemoryCache.Default;
            cache.Add(Constants.CacheNames.ConvertWizard, reportTemplate, new CacheItemPolicy());

            var manualDoc = reportDoc.OpenFile(Settings.Default.TemplateInsuranceManual);


            var allTables = reportDoc.GetTablesByNamePrefix(string.Empty);
            foreach (var tableName in allTables)
            {
                if (!manualDoc.TableExists(tableName)) continue;

                if (tableName.Contains("Policy Wording"))
                {
                    var text = reportDoc.GetTableCellText(tableName, 0, 2);
                    manualDoc.TypeTableCellText(tableName, 0, 2, text);
                }
                else
                {
                    reportDoc.CopyTable(tableName);
                    manualDoc.SelectTable(tableName);
                    manualDoc.DeleteTable(tableName);
                    manualDoc.InsertParagraphBreak();
                    manualDoc.PasteClipboardOriginalFormatting();
                }
            }

            //if (reportDoc.TableExists(Constants.WordTables.PolicyWording) && manualDoc.TableExists(Constants.WordTables.PolicyWording))
            //{
            //    reportDoc.CopyTable(Constants.WordTables.PolicyWording);
            //    manualDoc.DeleteTable(Constants.WordTables.PolicyWording);
            //    manualDoc.PasteClipboard();
            //}
            var allBookmarkedSections = reportDoc.GetBookmarksByPartialName(Constants.WordBookmarks.SummaryOfCoverStart);
            foreach (var bookmark in allBookmarkedSections)
            {
                var id = bookmark.Replace(Constants.WordBookmarks.SummaryOfCoverStart, string.Empty);
                if (reportDoc.HasBookmark(Constants.WordBookmarks.SummaryOfCoverStart + id) &&
                    reportDoc.HasBookmark(Constants.WordBookmarks.SummaryOfCoverEnd + id))
                {
                    var rs = reportDoc.GetBookmarkStartRange(Constants.WordBookmarks.SummaryOfCoverStart + id);
                    var re = reportDoc.GetBookmarkStartRange(Constants.WordBookmarks.SummaryOfCoverEnd + id);
                    reportDoc.CopyRange(rs, re);

                    if (manualDoc.HasBookmark(Constants.WordBookmarks.SummaryOfCoverStart + id))
                    {
                        //var ms = manualDoc.GetBookmarkStartRange(Constants.WordBookmarks.SummaryOfCoverStart + id);
                        manualDoc.MoveCursorToStartOfBookmark(Constants.WordBookmarks.SummaryOfCoverStart + id);
                        manualDoc.PasteClipboard();
                    }
                }
            }


            if (reportDoc.HasBookmark(Constants.WordBookmarks.InsertClientProfile) &&
                reportDoc.HasBookmark(Constants.WordBookmarks.InsertClientProfileEnd))
            {
                var rs = reportDoc.GetBookmarkStartRange(Constants.WordBookmarks.InsertClientProfile);
                var re = reportDoc.GetBookmarkStartRange(Constants.WordBookmarks.InsertClientProfileEnd);
                reportDoc.CopyRange(rs, re);
                if (manualDoc.HasBookmark(Constants.WordBookmarks.InsertClientProfile))
                {
                    var ms = manualDoc.GetBookmarkStartRange(Constants.WordBookmarks.InsertClientProfile);
                    var me = manualDoc.GetBookmarkStartRange(Constants.WordBookmarks.InsertClientProfileEnd);
                    manualDoc.DeleteRange(ms, me);

                    manualDoc.MoveCursorToStartOfBookmark(Constants.WordBookmarks.InsertClientProfile);

                    manualDoc.PasteClipboard();
                }
            }

            manualDoc.MoveToStartOfDocument();

            //var presenterManualDoc = new RibbonWizardPresenter(manualDoc, null);

            //var p = reportDoc.TurnOffProtection(string.Empty);
            //presenterManualDoc.PopulateData(reportTemplate);

            //manualDoc.PopulateControl(Constants.WordContentControls.DocumentTitle, "Insurance Manual");
            //reportDoc.TurnOnProtection(p, string.Empty);

            //var w = new BaseWizardForm();
            //var t = new InsuranceManual();
            //t.DocumentTitle = "Insurance Manual zzzz";
            //w.LogUsage(t, Enums.UsageTrackingType.ConvertDocument);
        }

        public void btnConvertToQuoteSlip_Click(IRibbonControl control)
        {
            #region Not Used - delete after 24th Sep 2013 once testing passes

            //var doc = new OfficeDocument(Globals.ThisAddIn.Application.ActiveDocument);
            //WordOM.Document d = Globals.ThisAddIn.Application.ActiveDocument;
            //IDocument quoteSlipDoc = doc.OpenFile(Constants.CacheNames.GenerateQuoteSlip, Settings.Default.TemplateQuoteSlip);

            //var ribbonPresenter = new RibbonPresenter(quoteSlipDoc, null);

            //var quoteSlipPresenter = new QuoteSlipPresenter(quoteSlipDoc, null);

            //int startRange = doc.GetBookmarkStartRange("FactFinderStart");
            //int endRange = doc.GetBookmarkEndRange("FactFinderEnd");

            //d.Range(startRange, endRange).Select();
            //Globals.ThisAddIn.Application.Selection.Copy();
            //quoteSlipDoc.PasteClipboard();


            //string includedPolicys = quoteSlipDoc.GetPropertyValue(Constants.WordDocumentProperties.IncludedPolicyTypes);

            //var qs = includedPolicys.Split(';').Select(p => new QuestionClass
            //    {
            //        Url = "/testing/Quote%20Slip%20Schedules/Industrial%20Special%20Risks.docx"
            //    }).Cast<IQuestionClass>().ToList();

            //quoteSlipDoc.Activate();
            //quoteSlipPresenter.InsertPolicySchedule(qs, false);

            //quoteSlipPresenter.MoveToStartOfDocument();

            //quoteSlipPresenter.CloseInformationPanel(true); 

            #endregion

            var doc = new OfficeDocument(Globals.ThisAddIn.Application.ActiveDocument);

            if (!doc.FileExistsInSharePoint(Settings.Default.TemplateQuoteSlip))
            {
                MessageBox.Show(@"Quote slips are not ready for use.", @"Quote slip not ready", MessageBoxButtons.OK,
                    MessageBoxIcon.Stop);
                return;
            }

            var convertForm = new ConvertFactFinderToQuoteSlip(doc);
            convertForm.ShowDialog();
        }

        public void btnSyncTable_Click(IRibbonControl control)
        {
            SyncTable();
        }

        public void SyncTable()
        {
            // var doc = new OfficeDocument(Globals.ThisAddIn.Application.ActiveDocument);
            //var proptectionType = -1;
            try
            {
                //proptectionType = doc.TurnOffProtection(String.Empty);


                var tables = Globals.ThisAddIn.Application.Selection.Tables;
                if (tables == null || tables.Count == 0)
                {
                    MessageBox.Show(@"Please ensure your cursor is within a table", @"No Table Found",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var table = tables[1];


                var contentcontrols = table.Range.ContentControls;
                if (contentcontrols == null) return;

                foreach (WordOM.ContentControl contentControl in contentcontrols)
                {
                    if (string.IsNullOrEmpty(contentControl.Tag)) continue;

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
                logger.Log(e.ToString(), Type.Error);

#if DEBUG
                MessageBox.Show(e.ToString(), @"sorry");
#endif
            }
        }

        public void btnSyncField_Click(IRibbonControl control)
        {
            var doc = new OfficeDocument(Globals.ThisAddIn.Application.ActiveDocument);
            var proptectionType = -1;
            try
            {
                proptectionType = doc.TurnOffProtection(string.Empty);
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
                logger.Log(e.ToString(), Type.Error);

#if DEBUG
                MessageBox.Show(e.ToString(), @"sorry");
#endif
            }
            finally
            {
                doc.TurnOnProtection(proptectionType, string.Empty);
            }
        }

        public void btnUpdateFields_Click(IRibbonControl control)
        {
            var tableCount = SummaryTableCount();

            if (tableCount != 1)
            {
                MessageBox.Show(
                    string.Format("There is {0} no {1} table in this document.", tableCount,
                        Constants.WordTables.RenewalReportPremiumSummary), @"No changes made.", MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            Globals.ThisAddIn.Application.ActiveDocument.Fields.Update();
            CalculatePremSummaryTable();
            SyncPremiumSummaryTable();
        }

        public void btnHelp_Click(IRibbonControl control)
        {
            ShowHelp();
        }

        private static void ShowHelp()
        {
            dynamic templateName =
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

        public void btnConvertPDF_Click(IRibbonControl control)
        {
            Process.Start(@"C:\Program Files\Free PDF Solutions\FreePDFSolutions_PDF2Word.exe");
        }

        public void btnWizard_Click(IRibbonControl control)
        {
            Globals.ThisAddIn.Application.ActiveDocument.LoadWizard(Enums.FormLoadType.RibbonClick);
        }

        #endregion

        #region Helpers

        public Bitmap GetImage(IRibbonControl control)
        {
            switch (control.Id)
            {
                case "btnHelp":
                {
                    return new Bitmap(Resources.Help);
                }

                case "btnOpenBICalculator":
                {
                    return new Bitmap(Resources.BICalculator);
                }

                case "btnToggleLock":
                {
                    return new Bitmap(Resources.Lock);
                }

                case "btnConvertToManual":
                {
                    return new Bitmap(Resources.InsuranceManual);
                }

                case "btnConvertToPlacementSlip":
                {
                    return new Bitmap(Resources.PlacementSlip);
                }

                case "btnConvertToPlacementSlipHome":
                {
                    return new Bitmap(Resources.PlacementSlip);
                }

                case "btnTogleLockUnlock":
                {
                    var mode =
                        Globals.ThisAddIn.Application.ActiveDocument.GetPropertyValue(
                            Constants.WordDocumentProperties.ToggleLockMode);
                    return string.Equals(mode, Constants.WordDocumentPropertyValues.ToggleLockModeLocked,
                        StringComparison.OrdinalIgnoreCase)
                        ? new Bitmap(Resources.Unlock)
                        : new Bitmap(Resources.Lock);
                }

                case "btnConvertToQuoteSlip":
                {
                    return new Bitmap(Resources.QuoteSlip1);
                }

                case "btnTogleUnlock":
                {
                    return new Bitmap(Resources.Unlock);
                }

                case "btnHelpHome":
                {
                    return new Bitmap(Resources.Help);
                }

                case "btnSyncField":
                {
                    return new Bitmap(Resources.SyncField);
                }

                case "btnSyncTableHome":
                {
                    return new Bitmap(Resources.SyncTableInfo);
                }


                case "btnSyncTable":
                {
                    return new Bitmap(Resources.SyncTableInfo);
                }

                case "btnWizard":
                {
                    return new Bitmap(Resources.Wizard);
                }
                case "btnWizardHome":
                {
                    return new Bitmap(Resources.Wizard);
                }

                case "btnConvertPDF":
                {
                    return new Bitmap(Resources.pdf);
                }
                case "btnConvertPDFHome":
                {
                    return new Bitmap(Resources.pdf);
                }
                case "btnUpdateFields":
                {
                    return new Bitmap(Resources.calculator);
                }
                case "btnUpdateFieldsHome":
                {
                    return new Bitmap(Resources.calculator);
                }

                case "btnAbout":
                {
                    return new Bitmap(Resources.About);
                }
                case "btnAboutHome":
                {
                    return new Bitmap(Resources.About);
                }

                case "btnInsertPortrait":
                {
                    return new Bitmap(Resources.portratPage);
                }

                case "btnInsertLandscape":
                {
                    return new Bitmap(Resources.landscapePage);
                }
                case "btnInsertPortraitHome":
                {
                    return new Bitmap(Resources.portratPage);
                }

                case "btnInsertLandscapeHome":
                {
                    return new Bitmap(Resources.landscapePage);
                }
            }
            return null;
        }


        private static string GetResourceText(string resourceName)
        {
            var asm = Assembly.GetExecutingAssembly();
            var resourceNames = asm.GetManifestResourceNames();
            foreach (var t in resourceNames)
            {
                if (string.Compare(resourceName, t, StringComparison.OrdinalIgnoreCase) != 0) continue;
                if (t == null) continue;
                using (var resourceReader = new StreamReader(asm.GetManifestResourceStream(t)))
                {
                    return resourceReader.ReadToEnd();
                }
            }
            return null;
        }

        #endregion
    }
}