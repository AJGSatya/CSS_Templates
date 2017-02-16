using System;
using System.Collections.Generic;
using System.Globalization;
using OAMPS.Office.BusinessLogic.Helpers;
using OAMPS.Office.BusinessLogic.Interfaces;
using OAMPS.Office.BusinessLogic.Interfaces.Template;
using OAMPS.Office.BusinessLogic.Interfaces.Wizards;
using OAMPS.Office.BusinessLogic.Interfaces.Word;
using OAMPS.Office.BusinessLogic.Models.Wizards;
using OAMPS.Office.BusinessLogic.Presenters.Wizards;

namespace OAMPS.Office.BusinessLogic.Presenters
{
    public class InsuranceManualWizardPresenter : BaseWizardPresenter
    {
        public InsuranceManualWizardPresenter(IDocument document, IBaseView view)
            : base(document, view)
        {
        }

        public void PopulateImportantNotices()
        {
        }

        public void PopulateProgramSummarys(List<IPolicyClass> policies, string fragmentUrl)
        {
            var c = 1;
            for (var i = policies.Count -1; i > -1; i--)
            {
                var policy = policies[i];
                Document.MoveCursorToStartOfBookmark(Constants.WordBookmarks.BasisOfCoverPrevious);

                Document.InsertFile(fragmentUrl);

                if (c > 1) //dont include pagebreak on the first page
                {
                    Document.InsertPageBreak();
                }

                if (Document.HasBookmark("coInsuranceText"))
                {
                    var curInsurers = GenerateInsurers(policy.CurrentInsurer);

                    Document.RenameBookmark("coInsuranceText", "coInsuranceText" + policy.Id);
                    //If co-insurance is recommended, then include a reference to the “Average / Co-insurance Clause” in details of quotations by policy. It should not appear if co-insurance is not recommened
                    if (curInsurers.Split('%').Length > 2)
                    {
                        if (Document.MoveCursorToStartOfBookmark("coInsuranceText" + policy.Id))
                        {
                            Document.TypeText("Please refer to the Average / Co-insurance clause in the 'Important notes' section.");
                        }
                    }
                }
              

                Document.RenameControl("ClassOfInsuranceTitle", "ClassOfInsuranceTitle" + policy.Id);
                Document.RenameControl("Insurer", "Insurer" + policy.Id);
                Document.RenameControl("PolicyNumber", "PolicyNumber" + policy.Id);

                Document.PopulateControl("ClassOfInsuranceTitle" + policy.Id, policy.Title);
                Document.PopulateControl("PolicyNumber" + policy.Id, policy.PolicyNumber);
                Document.PopulateControl("Insurer" + policy.Id, GenerateInsurers(policy.CurrentInsurer));  //dev comment: this is actually the recommended insurer.

                Document.RenameTable(Constants.WordTables.PolicyWording, Constants.WordTables.PolicyWording + policy.Id + policy.Order);
                Document.RenameTable(Constants.WordTables.Liabilityrisksandexposures, Constants.WordTables.Liabilityrisksandexposures + policy.Id + policy.Order);
                Document.RenameTable(Constants.WordTables.AssetRiskProtection, Constants.WordTables.AssetRiskProtection + policy.Id + policy.Order);
                Document.RenameTable(Constants.WordTables.IncomeandotherFinancialExposures, Constants.WordTables.IncomeandotherFinancialExposures + policy.Id + policy.Order);

                Document.RenameBookmark(Constants.WordBookmarks.SummaryOfCoverStart, Constants.WordBookmarks.SummaryOfCoverStart + policy.Id + policy.Order);
                Document.RenameBookmark(Constants.WordBookmarks.SummaryOfCoverEnd, Constants.WordBookmarks.SummaryOfCoverEnd + policy.Id + policy.Order);

                c++;
            }
        }

        public string GenerateInsurers(string insurers)
        {
            return insurers.Replace(Constants.Seperators.Lineseperator, Environment.NewLine).Replace(Constants.Seperators.Spaceseperator, " ");
        }

        public void ForceUpdateToc()
        {
            Document.UpdateToc();
        }

        public void PopulateClaimsProcedures(List<IPolicyClass> policies)
        {
            for (var i = policies.Count -1; i > -1; i--)
            {
                var policy = policies[i];
                Document.MoveCursorToStartOfBookmark(Constants.WordBookmarks.PreviousClaimsProcedure);

               // Document.InsertPageBreak();
                Document.InsertFile(policy.FragmentPolicyUrl);
            }
        }

        public void PopulatePolicyTable(List<IPolicyClass> policies)
        {
            if (policies == null || policies.Count <= 0) return;
            
            //policies.Sort((x,y) => x.Order.CompareTo(y.Order));

            var r = 2; //start at second row, vsto arrays start at 1 + the 1 row for the table header.

            foreach (var policy in policies)
            {
                var rowCount = Document.TableRowOrColumnCount(false, Constants.WordTables.InsuranceManualProgramSummary);

                if (r > rowCount)
                {
                    Document.InsertTableRowNonCellLoop(rowCount, Constants.WordTables.InsuranceManualProgramSummary);
                }
                Document.PopulateTableCell(r, 1, policy.Title, Constants.WordTables.InsuranceManualProgramSummary);
                Document.PopulateTableCell(r, 2, GenerateInsurers(policy.CurrentInsurer), Constants.WordTables.InsuranceManualProgramSummary);
                Document.PopulateTableCell(r, 3, policy.PolicyNumber, Constants.WordTables.InsuranceManualProgramSummary);
                r++;
            }
        }

        public void PopulateClaimsProcedures(List<string> urls)
        {
            //var c = 1;
            foreach (var u in urls)
            {
                if (!Document.MoveCursorToStartOfBookmark(Constants.WordBookmarks.ClaimsProcedures)) return;
                Document.InsertFile(u);

                //if (c > 1 && c < urls.Count) //dont include pagebreak on first page
                //{
                //    Document.InsertPageBreak();
                //}
                //if (c == urls.Count -1) //nclude pagebreak on last page
                //{
                //    Document.InsertPageBreak();
                //}

                //c++;
            }
        }

        public void PopulateclientProfile(string url)
        {
            if (!Document.MoveCursorToStartOfBookmark(Constants.WordBookmarks.InsertClientProfile)) return;

            Document.InsertFile(url);
            Document.InsertRealPageBreak();
            Document.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.ClientProfile, "true");
        }

        public void PopulateContractingProcedure(string url)
        {
            if (!Document.MoveCursorToStartOfBookmark(Constants.WordBookmarks.InsertContractingProcedure)) return;

            Document.InsertFile(url);
            Document.InsertRealPageBreak();
            Document.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.ContractingProcedure, "true");
        }

        public IInsuranceManual LoadIncludedPolicyClasses(IInsuranceManual template)
        {
            //load selected policy classes
            template.SelectedPolicyClasses = new List<IPolicyClass>();
            string items = Document.GetPropertyValue(Constants.WordDocumentProperties.IncludedPolicyTypes);
           // var regex = new Regex(Constants.Seperators.Lineseperator);
            //  var split = regex.Split(items);
            foreach (string i in items.Split(';'))
                //foreach (var i in split)
            {
                string[] d = i.Split('_');
                if (d.Length == 3)
                {
                    var p = new PolicyClass
                        {
                            Id = (d.Length == 0) ? string.Empty : d[0].ToString(CultureInfo.InvariantCulture),
                            RecommendedInsurer = Document.ReadContentControlValue(d[2]),
                            CurrentInsurer = Document.ReadContentControlValue(d[1]),
                            RecommendedInsurerId = d[2].Substring(0, d[2].IndexOf("r", StringComparison.Ordinal)),
                            CurrentInsurerId = d[1].Substring(0, d[1].IndexOf("c", StringComparison.Ordinal)),
                            Order = int.Parse(d[2].Substring(d[2].IndexOf("r", StringComparison.Ordinal) + 1))
                        };
                    template.SelectedPolicyClasses.Add(p);
                }
            }
            return template;
        }
    }
}