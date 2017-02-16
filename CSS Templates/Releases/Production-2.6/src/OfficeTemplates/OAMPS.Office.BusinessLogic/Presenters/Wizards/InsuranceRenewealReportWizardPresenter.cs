using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using OAMPS.Office.BusinessLogic.Helpers;
using OAMPS.Office.BusinessLogic.Interfaces;
using OAMPS.Office.BusinessLogic.Interfaces.Template;
using OAMPS.Office.BusinessLogic.Interfaces.Wizards;
using OAMPS.Office.BusinessLogic.Interfaces.Word;
using OAMPS.Office.BusinessLogic.Models.Template;
using OAMPS.Office.BusinessLogic.Models.Wizards;
using OAMPS.Office.BusinessLogic.Presenters.Wizards;

namespace OAMPS.Office.BusinessLogic.Presenters
{
    public class InsuranceRenewealReportWizardPresenter : BaseWizardPresenter
    {
        private const int FirstrowPremiumsummary = 3;

        public InsuranceRenewealReportWizardPresenter(IDocument document, IBaseView view)
            : base(document, view)
        {
        }

        public void PopulateclientProfile(string url)
        {
            if (Document.MoveCursorToStartOfBookmark(Constants.WordBookmarks.InsertClientProfile))
            {
                //Document.InsertPageBreak();
                Document.InsertFile(url);
                Document.InsertRealPageBreak();
                Document.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.ClientProfile, "true");
            }
        }

        public void PopulatePremiumSummary(List<IPolicyClass> fragements)
        {
            Document.SelectTable(Constants.WordTables.RenewalReportPremiumSummary);
            var rowCount = Document.TableRowOrColumnCount(false);
            var columnCount = Document.TableRowOrColumnCount(true);
            if (rowCount <= 0) return;

            var position = FirstrowPremiumsummary;

            //word arrays start at 1 not 0.  also we start at 2 as this table has merged rows for the header.
            int insertRowCount = 0;

            foreach (var f in fragements)
            {
                if (insertRowCount > 0)
                {
                    Document.InsertTableRow(position);
                    position++;
                }

                for (var y = 2; y < columnCount; y++)
                {
                    switch (y)
                    {
                        case 2:
                            Document.PopulateTableCell(position, y, f.Title);
                            break;
                        case 3:
                            {
                                var regex = new Regex(Constants.Seperators.Lineseperator);
                                string[] split = regex.Split(f.RecommendedInsurer);
                                // var split =  f.RecommendedInsurer.Split(Constants.Seperators.Lineseperator);
                                string all = string.Empty;
                                foreach (var a in split)
                                {
                                    var regex2 = new Regex(Constants.Seperators.Spaceseperator);
                                    var split2 = regex2.Split(a);

                                    all += split2[0] + ", ";
                                }

                                while (all.EndsWith(", "))
                                {
                                    all = all.Remove(all.Length - 2, 2);
                                }

                                Document.PopulateTableCell(position, y, all);
                            }
                            break;
                        default:
                            Document.PopulateTableCell(position, y, "$0.00");
                            break;
                    }
                }

                insertRowCount++;
            }
        }

        public void PopulateExecutiveSummary(Enums.Remuneration remuneration, string pathCommision, string pathFeeAndCombination)
        {
            if (Document.MoveCursorToStartOfBookmark(Constants.WordBookmarks.ExecutiveSummary))
            {
                if (remuneration == Enums.Remuneration.Commission)
                {
                    Document.InsertFile(pathCommision);
                }
                else if (remuneration == Enums.Remuneration.Combined || remuneration == Enums.Remuneration.Fee)
                {
                    Document.InsertFile(pathFeeAndCombination);
                }
            }
        }

        public void PopulatePurposeOfReport(Enums.Segment segment, string path23, string path45, string path)
        {
            if (Document.MoveCursorToStartOfBookmark(Constants.WordBookmarks.PurposeOfReport))
            {
                if (segment == Enums.Segment.Four || segment == Enums.Segment.Five)
                {
                    Document.InsertFile(path45);
                }
                else if (segment == Enums.Segment.Two || segment == Enums.Segment.Three)
                {
                    Document.InsertFile(path23);
                }
            }


            if (Document.MoveCursorToStartOfBookmark(Constants.WordBookmarks.PurposeOfReportGeneric))
            {
                Document.InsertFile(path);
            }
        }

        public void PopulateServiceLineAgrement(Enums.Segment segment, Dictionary<Enums.Segment, string> segmentDocuments, DateTime insuranceStartDate)
        {
            string filePath = segmentDocuments[segment];
            if (String.IsNullOrEmpty(filePath)) return;

            Document.MoveCursorPastStartOfBookmark("ServiceDeliveryPlan", 1);
            Document.InsertFile(filePath);
            Document.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.Segment, segment.ToString());

            //fix this after POC demo with rob
            Document.PopulateControl("ctr-1", insuranceStartDate.AddMonths(-1).ToString(CultureInfo.CurrentCulture));
            Document.PopulateControl("ctr-2", insuranceStartDate.AddMonths(-2).ToString(CultureInfo.CurrentCulture));
            Document.PopulateControl("ctr-3", insuranceStartDate.AddMonths(-3).ToString(CultureInfo.CurrentCulture));
            Document.PopulateControl("ctr-4", insuranceStartDate.AddMonths(-4).ToString(CultureInfo.CurrentCulture));
            Document.PopulateControl("ctr-5", insuranceStartDate.AddMonths(-5).ToString(CultureInfo.CurrentCulture));
            Document.PopulateControl("ctr-6", insuranceStartDate.AddMonths(-6).ToString(CultureInfo.CurrentCulture));
            Document.PopulateControl("ctr-7", insuranceStartDate.AddMonths(-7).ToString(CultureInfo.CurrentCulture));
            Document.PopulateControl("ctr-8", insuranceStartDate.AddMonths(-8).ToString(CultureInfo.CurrentCulture));
            Document.PopulateControl("ctr-9", insuranceStartDate.AddMonths(-9).ToString(CultureInfo.CurrentCulture));
            Document.PopulateControl("ctr-10", insuranceStartDate.AddMonths(-10).ToString(CultureInfo.CurrentCulture));
            Document.PopulateControl("ctr-11", insuranceStartDate.AddMonths(-11).ToString(CultureInfo.CurrentCulture));
            Document.PopulateControl("ctr-12", insuranceStartDate.AddMonths(-12).ToString(CultureInfo.CurrentCulture));

            Document.PopulateControl("ctr0", insuranceStartDate.ToString(CultureInfo.CurrentCulture));
            Document.PopulateControl("ctr1", insuranceStartDate.AddMonths(1).ToString(CultureInfo.CurrentCulture));
            Document.PopulateControl("ctr2", insuranceStartDate.AddMonths(2).ToString(CultureInfo.CurrentCulture));
            Document.PopulateControl("ctr3", insuranceStartDate.AddMonths(3).ToString(CultureInfo.CurrentCulture));
            Document.PopulateControl("ctr4", insuranceStartDate.AddMonths(4).ToString(CultureInfo.CurrentCulture));
            Document.PopulateControl("ctr5", insuranceStartDate.AddMonths(5).ToString(CultureInfo.CurrentCulture));
            Document.PopulateControl("ctr6", insuranceStartDate.AddMonths(6).ToString(CultureInfo.CurrentCulture));
            Document.PopulateControl("ctr7", insuranceStartDate.AddMonths(7).ToString(CultureInfo.CurrentCulture));
            Document.PopulateControl("ctr8", insuranceStartDate.AddMonths(8).ToString(CultureInfo.CurrentCulture));
            Document.PopulateControl("ctr9", insuranceStartDate.AddMonths(9).ToString(CultureInfo.CurrentCulture));
            Document.PopulateControl("ctr10", insuranceStartDate.AddMonths(10).ToString(CultureInfo.CurrentCulture));
            Document.PopulateControl("ctr11", insuranceStartDate.AddMonths(11).ToString(CultureInfo.CurrentCulture));
            Document.PopulateControl("ctr12", insuranceStartDate.AddMonths(12).ToString(CultureInfo.CurrentCulture));
        }

        public string GenerateBasisOfCoverInsurers(string insurers)
        {
            return insurers.Replace(Constants.Seperators.Lineseperator, Environment.NewLine).Replace(Constants.Seperators.Spaceseperator, " ");
        }

        public void PopulateBasisOfCover(List<IPolicyClass> fragements, string fragmentUrl)
        {
            var c = 1;
            fragements.Sort((x, y) => x.Order.CompareTo(y.Order));


            //fragements.Sort(delegate(IPolicyClass p1, IPolicyClass p2)
            //{
            //    int compare = p1.Order.CompareTo(p2.Order);
            //    return compare;
            //});

            const string bookmarkTemplate = "BasisOfCoverPrevious";
            for (var index = fragements.Count - 1; index >= 0; index--)
            {
                var f = fragements[index];
                var order = f.Order;
                if (Document.MoveCursorToStartOfBookmark(bookmarkTemplate))
                {
                    if (c > 0 && c < fragements.Count) //dont include pagebreak on first page
                    {
                        Document.InsertPageBreak();
                    }


                    var recInsurers = GenerateBasisOfCoverInsurers(f.RecommendedInsurer);
                    var curInsurers = GenerateBasisOfCoverInsurers(f.CurrentInsurer);

                    Document.InsertFile(fragmentUrl);

                    if (Document.HasBookmark("coInsuranceText"))
                    {
                        Document.RenameBookmark("coInsuranceText", "coInsuranceText" + f.Id);
                        //If co-insurance is recommended, then include a reference to the “Average / Co-insurance Clause” in details of quotations by policy. It should not appear if co-insurance is not recommened
                        if (recInsurers.Split('%').Length > 2)
                        {
                            if (Document.MoveCursorToStartOfBookmark("coInsuranceText" + f.Id))
                            {
                                Document.TypeText("Please refer to the Average / Co-insurance clause in the 'Important notes' section.");
                            }
                        }
                    }
              

                    Document.RenameControl("ClassOfInsuranceTitle", f.Id + order + "ClassOfInsuranceTitle");
                    Document.RenameControl("ClassOfInsuranceCurrentInsurer", f.CurrentInsurerId + "c" + order);
                    Document.RenameControl("ClassOfInsuranceRecommendedInsurer", f.RecommendedInsurerId + "r" + order);
                    Document.PopulateControl(f.Id + order + "ClassOfInsuranceTitle", f.Title);
                    Document.PopulateControl(f.CurrentInsurerId + "c" + order, curInsurers);
                    Document.PopulateControl(f.RecommendedInsurerId + "r" + order, recInsurers);
                    string currentValue = Document.GetPropertyValue(Constants.WordDocumentProperties.IncludedPolicyTypes);
                    Document.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.IncludedPolicyTypes,
                                                         currentValue + ";" + f.Id + "_" + f.CurrentInsurerId + "c" +
                                                         order + "_" + f.RecommendedInsurerId + "r" + order);

                    Document.RenameTable(Constants.WordTables.RenewalReportPremiumCosts, Constants.WordTables.RenewalReportPremiumCosts + "_" + f.Title);


                  
                }
                c++;
            }

            Document.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.IncludedPolicyTypesCount,
                                                 fragements.Count.ToString(CultureInfo.InvariantCulture));
        }

        public void PopulateUFI(bool populateUFI, string ufiUrl)
        {
            if (!populateUFI) return;
            if (Document.MoveCursorToStartOfBookmark(Constants.WordBookmarks.UFIBookmark))
            {
                Document.InsertPageBreak();
                Document.InsertFile(ufiUrl);
                Document.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.UFI, "true");
            }
        }

        public void SendUFIMessage(InsuranceRenewalReport template, string recipients)
        {
            var textInfo = new CultureInfo("en-AU", false).TextInfo;
            var name = textInfo.ToTitleCase(Environment.UserName.Replace(".", " "));
            var bodywithValues =
                string.Format(@"Hello, 
                                <br/> <br/> This is an automated email generated by the CSS Template wizard.
                                <br/><br/> {0} has just created an Insurance Renewal Report recommending an unauthorised foreign insurer.
                                 <br/><br/> Details are below.  Please action as appropriate.
                                <br/>
                                    Account Executive: {1}<br/>
                                    Client Name: {2}<br/>", name, template.ExecutiveName, template.ClientName);

            Document.SendEmail("UFI recommended", bodywithValues, recipients);
        }

        public void SendSeg4or5Message(InsuranceRenewalReport template, string recipients)
        {
            var textInfo = new CultureInfo("en-AU", false).TextInfo;
            var name = textInfo.ToTitleCase(Environment.UserName.Replace(".", " "));
            var bodywithValues =
                string.Format(@"Hello, 
                                <br/> <br/> This is an automated email generated by the CSS Template wizard.
                                <br/><br/> {0} has just created an Insurance Renewal Report recommending an unauthorised foreign insurer.
                                 <br/><br/> Details are below.  Please action as appropriate.
                                <br/>
                                    Account Executive: {1}<br/>
                                    Client Name: {2}<br/>", name, template.ExecutiveName, template.ClientName);

            Document.SendEmail("UFI recommended", bodywithValues, recipients);
        }

        private string ConstructTableColumn(string name)
        {
            return "<td>" + name + "</td>";
        }

        private string BuildTableRow(params string[] value)
        {
            string row = "<tr>";
            foreach (string s in value)
            {
                row += ConstructTableColumn(s);
            }
            row += "</tr>";
            return row;
        }

        private string BuildTable(List<IPolicyClass> fragments)
        {
            //do heading
            string table = BuildTableRow(new[] {"policy class", "Current Insurer", "Recommended Insurer"});
            foreach (IPolicyClass policyClass in fragments)
            {
                table += BuildTableRow(new[] {policyClass.Title, policyClass.CurrentInsurer, policyClass.RecommendedInsurer});
            }
            return table;
        }

        public IInsuranceRenewalReport LoadIncludedPolicyClasses(IInsuranceRenewalReport template)
        {
            //load selected policy classes
            template.SelectedDocumentFragments = new List<IPolicyClass>();
            string items = Document.GetPropertyValue(Constants.WordDocumentProperties.IncludedPolicyTypes);
            //var regex = new Regex(Constants.Seperators.Lineseperator);
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
                    template.SelectedDocumentFragments.Add(p);
                }
            }
            return template;
        }

        public void PopulateDocumentFragements(List<IPolicyClass> fragements)
        {
            foreach (IPolicyClass fragement in fragements)
            {
                Document.SetCurrentRangeText(fragement.Url);
            }
        }

        public void DeletePage(int pageNumber)
        {
            if (Document.PageCount > 1)
            {
                Document.DeletePage(pageNumber);
            }
            else
            {
                View.DisplayMessage(
                    String.Format("Unable to delete page {0}, there is only {1} page in this document", pageNumber,
                                  Document.PageCount), "Cannot Delete Page");
            }
        }

        public void PopulateImportantNotices(Enums.Statutory statutory, string importantNoticesUrl, string privacyStatementUrl, string fsgUrl, string termsOfEngagementUrl)
        {

            if (!Document.MoveCursorToStartOfBookmark(Constants.WordBookmarks.ImportantNotes)) return;

            Document.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.StatutoryInformation, statutory.ToString());

            //insert stat notices
            Document.InsertFile(importantNoticesUrl); //in settings
            Document.InsertPageBreak();



            switch (statutory)
            {
                case Enums.Statutory.Retail:
                    {
                        Document.InsertFile(fsgUrl); //in settings
                        break;
                    }
                case Enums.Statutory.Wholesale:
                    {
                        //insert pivacy statement
                        //Document.InsertFile(privacyStatementUrl); //compliance rule change, privacy not required ever.
                       //Document.InsertPageBreak();

                        Document.InsertFile(termsOfEngagementUrl); //in settings
                        break;
                    }
                case Enums.Statutory.WholesaleWithRetail:
                    {
                        //insert pivacy statement
                        //Document.InsertFile(privacyStatementUrl); //compliance rule change, privacy not required ever.
                        //Document.InsertPageBreak();

                        Document.InsertFile(fsgUrl); //in settings

                        Document.InsertPageBreak();
                        Document.InsertFile(termsOfEngagementUrl); //in settings
                        break;
                    }
            }


        //    if (Document.MoveCursorToStartOfBookmark(Constants.WordBookmarks.ImportantNotes))
        //    {
        //        Document.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.StatutoryInformation,
        //                                             statutory.ToString());
        //        //insert stat notices
        //        Document.InsertFile(importantNoticesUrl); //in settings
        //        Document.InsertPageBreak();

        //        //insert pivacy statement
        //        Document.InsertFile(privacyStatementUrl); //in settings

        //        switch (statutory)
        //        {
        //            case Enums.Statutory.Retail:
        //                {
        //                    Document.InsertPageBreak();
        //                    Document.InsertFile(fsgUrl); //in settings
        //                    break;
        //                }
        //            case Enums.Statutory.Wholesale:
        //                {
        //                    Document.InsertPageBreak();
        //                    Document.InsertFile(termsOfEngagementUrl); //in settings
        //                    break;
        //                }
        //            case Enums.Statutory.WholesaleWithRetail:
        //                {
        //                    Document.InsertPageBreak();
        //                    Document.InsertFile(fsgUrl); //in settings
        //                    Document.InsertPageBreak();
        //                    Document.InsertFile(termsOfEngagementUrl); //in settings
        //                    break;
        //                }
        //        }
        //    }
        }

        public void PopulateRemuneration(Enums.Remuneration remuneration, List<DocumentFragment> documents)
        {
            DocumentFragment frag = documents.Find(i => i.Title.Equals(remuneration.ToString(), StringComparison.OrdinalIgnoreCase));
            Document.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.Remuneration, remuneration.ToString());

            if (!Document.HasBookmark(Constants.WordBookmarks.AddRenumeration)) return;

            if (Document.HasBookmark(Constants.WordBookmarks.Renumeration))
            {
                if (Document.MoveCursorToStartOfBookmark(Constants.WordBookmarks.Renumeration))
                {
                    Document.DeletePage();
                    Document.AddBookmarkToCurrentLocation(Constants.WordBookmarks.AddRenumeration);
                }
            }
            if (String.IsNullOrEmpty(frag.Url)) return;

            if (!Document.MoveCursorToStartOfBookmark(Constants.WordBookmarks.AddRenumeration)) return;

            Document.MoveCursorUp(1);
            Document.InsertPageBreak();
            Document.ChangePageOrientToPortrait();
            Document.InsertFile(frag.Url);
        }
    }
}