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
            for (var i = policies.Count -1; i > -1; i--)
            {
                var policy = policies[i];
                Document.MoveCursorToStartOfBookmark(Constants.WordBookmarks.BasisOfCoverPrevious);
                Document.InsertPageBreak();
                Document.InsertFile(fragmentUrl);

                Document.RenameControl("ClassOfInsuranceTitle", "ClassOfInsuranceTitle" + policy.Id);
                Document.RenameControl("Insurer", "Insurer" + policy.Id);
                Document.RenameControl("PolicyNumber", "PolicyNumber" + policy.Id);

                Document.PopulateControl("ClassOfInsuranceTitle" + policy.Id, policy.Title);
                Document.PopulateControl("PolicyNumber" + policy.Id, policy.PolicyNumber);
                Document.PopulateControl("Insurer" + policy.Id, GenerateInsurers(policy.CurrentInsurer));
            }
        }

        public string GenerateInsurers(string insurers)
        {
            return insurers.Replace(Constants.Seperators.Lineseperator, Environment.NewLine).Replace(Constants.Seperators.Spaceseperator, " ");
        }

        public void PopulateClaimsProcedures(List<IPolicyClass> policies)
        {
            for (var i = policies.Count -1; i > -1; i--)
            {
                var policy = policies[i];
                Document.MoveCursorToStartOfBookmark(Constants.WordBookmarks.PreviousClaimsProcedure);

                Document.InsertPageBreak();
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
                Document.PopulateTableCell(r, 2, policy.PolicyNumber, Constants.WordTables.InsuranceManualProgramSummary);
                r++;
            }
        }

        public void PopulateclientProfile(string url)
        {
            if (!Document.MoveCursorToStartOfBookmark(Constants.WordBookmarks.InsertClientProfile)) return;

            Document.InsertFile(url);
            Document.InsertRealPageBreak();
            Document.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.ClientProfile, "true");
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