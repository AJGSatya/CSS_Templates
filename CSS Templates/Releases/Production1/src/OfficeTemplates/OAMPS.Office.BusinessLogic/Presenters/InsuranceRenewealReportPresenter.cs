using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Net;
using System.Net.Mail;
using System.Reflection;
using OAMPS.Office.BusinessLogic.Helpers;
using OAMPS.Office.BusinessLogic.Interfaces;
using OAMPS.Office.BusinessLogic.Interfaces.Wizards;
using OAMPS.Office.BusinessLogic.Interfaces.Word;
using OAMPS.Office.BusinessLogic.Interfaces.Template;
using OAMPS.Office.BusinessLogic.Models.Wizards;

namespace OAMPS.Office.BusinessLogic.Presenters
{
    public class InsuranceRenewealReportPresenter : BasePresenter
    {
        private const int FIRSTROW_PREMIUMSUMMARY = 3;
        public InsuranceRenewealReportPresenter(IDocument document, IBaseView view)
            : base(document, view)
        {

        }

        //method no longer required as business rules have changed.
        //public void RemoveBasisOfCover(List<IPolicyClass> fragements)
        //{
        //    var curerntIncludes = Document.GetPropertyValue(Helpers.Constants.WordDocumentProperties.IncludedPolicyTypes);
        //    foreach (var f in fragements)
        //    {
        //        foreach (var c in curerntIncludes.Split(';'))
        //        {
        //            if (String.Equals(c, f.Title, StringComparison.OrdinalIgnoreCase))
        //                curerntIncludes = c.Replace(c.ToUpper(), f.Title.ToUpper());
        //        }
        //    }

        //    Document.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.IncludedPolicyTypes, curerntIncludes);
        //}

        public void PopulatePremiumSummary(List<IPolicyClass> fragements)
        {
            Document.SelectTable("Premium Summary");
            var rowCount = Document.TableRowCount();

            if (rowCount > 0)
            {
                var position = FIRSTROW_PREMIUMSUMMARY;
                
                    //word arrays start at 1 not 0.  also we start at 2 as this table has merged rows for the header.
                var insertRowCount = 0;
                for (var i = fragements.Count - 1; i > -1; i--)
                {
                    var f = fragements[i];

                    if (insertRowCount == 0)
                    {
                        Document.PopulateTableCell(position, 1, f.Title);
                        Document.PopulateTableCell(position, 2, f.RecommendedInsurer);
                    }
                    else
                    {
                        Document.InsertTableRow(position);
                        position++;
                        Document.PopulateTableCell(position, 1, f.Title);
                        Document.PopulateTableCell(position, 2, f.RecommendedInsurer);
                    }
                                       

                    insertRowCount++;
                    //if (position >= rowCount - 2)
                    //    //we -4 as there are 4 rows in the table with static content (header, and row totals)
                    //{
                    //    Document.InsertTableRow(position);
                    //}


                }
            }
        }

        public void PopulateExecutiveSummary(Enums.Remuneration remuneration, string pathCommision, string pathFeeAndCombination)
        {
            Document.MoveCursorToStartOfBookmark(Constants.WordBookmarks.ExecutiveSummary);
            if (remuneration ==Enums.Remuneration.Commission)
            {
                Document.InsertFile(pathCommision);
            }
            else if (remuneration ==Enums.Remuneration.Combined || remuneration == Enums.Remuneration.Fee)
            {
                Document.InsertFile(pathFeeAndCombination);
            }
        }

        public void PopulatePurposeOfReport( Enums.Segment segment,string path23, string path45 )
        {
            Document.MoveCursorToStartOfBookmark(Constants.WordBookmarks.PurposeOfReport);
            if (segment == Enums.Segment.Four || segment == Enums.Segment.Five)
            {                
                Document.InsertFile(path45);
            }
            else if (segment == Enums.Segment.Two || segment == Enums.Segment.Three)
            {                
                Document.InsertFile(path23);
            }
        }

        public void PopulateServiceLineAgrement(Enums.Segment segment,
                                                Dictionary<Enums.Segment, string> segmentDocuments,
                                                DateTime insuranceStartDate)
        {
            var filePath = segmentDocuments[segment];
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

        public void PopulateBasisOfCover(List<IPolicyClass> fragements, string fragmentUrl)
        {
            var c = 1;

            const string bookmarkTemplate = "BasisOfCoverPrevious";
            foreach (var f in fragements)
            {

                Document.MoveCursorToStartOfBookmark(bookmarkTemplate);

                if (c > 0 && c < fragements.Count) //dont include pagebreak on first page
                {
                    Document.InsertPageBreak();
                }

                Document.InsertFile(fragmentUrl);
                Document.RenameControl("ClassOfInsuranceTitle", f.TitleNoWhiteSpace + "ClassOfInsuranceTitle");
                Document.RenameControl("ClassOfInsuranceCurrentInsurer",
                                       f.CurrentInsurerId + "ClassOfInsuranceCurrentInsurer");
                Document.RenameControl("ClassOfInsuranceRecommendedInsurer",
                                       f.RecommendedInsurerId + "ClassOfInsuranceRecommendedInsurer");
                Document.PopulateControl(f.TitleNoWhiteSpace + "ClassOfInsuranceTitle", f.Title);
                Document.PopulateControl(f.CurrentInsurerId + "ClassOfInsuranceCurrentInsurer", f.CurrentInsurer);
                Document.PopulateControl(f.RecommendedInsurerId + "ClassOfInsuranceRecommendedInsurer",
                                         f.RecommendedInsurer);

                var currentValue = Document.GetPropertyValue(Constants.WordDocumentProperties.IncludedPolicyTypes);
                Document.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.IncludedPolicyTypes,
                                                     currentValue + ";" + f.Id);

                c++;
            }

            Document.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.IncludedPolicyTypesCount,
                                                 fragements.Count.ToString(CultureInfo.InvariantCulture));
        }


        public void PopulateUFI(bool populateUFI, string ufiUrl)
        {
            if (!populateUFI) return;
            Document.MoveCursorToStartOfBookmark(Constants.WordBookmarks.UFIBookmark);

            Document.InsertPageBreak();
            Document.InsertFile(ufiUrl);
            Document.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.UFI, "true");
        }

        public void SendUFIMessage(bool sendEmail, string to, string from)
        {
            if (!sendEmail) return;


            //string fromEmail = "dan.colarossi@oamps.com.au";//sending email from...
            //string ToEmail = "dan.colarossi@oamps.com.au";//destination email             
            //string body = "test b";
            //string subject = "test s";

            //try
            //{
            //    SmtpClient sMail = new SmtpClient("mail.wesins.wiroot.internal");//exchange or smtp server goes here.
            //    sMail.DeliveryMethod = SmtpDeliveryMethod.Network;
            //    //sMail.Credentials = new NetworkCredential("username","password");this line most likely wont be needed if you are already an         authenticated user.
            //    sMail.Send(fromEmail, ToEmail, subject, body);
            //}
            //catch (Exception ex)
            //{
            //    //do something after error if there is one
            //}


            // Document.SendEmail();

            //var message = new MailMessage();
            //var smtp = new SmtpClient();

            //message.From = new MailAddress("dan.colarossi@oamps.com.au");
            //message.To.Add(new MailAddress("dan.colarossi@oamps.com.au"));
            //message.Subject = "Test";
            //message.Body = "Content";

            ////smtp.Port = 587;
            //smtp.Host = "mail.wesins.wiroot.internal";
            //smtp.EnableSsl = true;
            //smtp.UseDefaultCredentials = true;
            ////smtp.Credentials =   NetworkCredential("from@gmail.com", "pwd");
            //smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
            //smtp.Send(message);
        }

        public IInsuranceRenewalReport LoadIncludedPolicyClasses(IInsuranceRenewalReport template)
        {
            //load selected policy classes
            template.SelectedDocumentFragments = new List<IPolicyClass>();
            var items = Document.GetPropertyValue(Helpers.Constants.WordDocumentProperties.IncludedPolicyTypes);
            foreach (var i in items.Split(';'))
            {
                if (!string.IsNullOrEmpty(i))
                {
                    var p = new PolicyClass
                    {

                        //Title = i.ToString(CultureInfo.InvariantCulture)
                        Id = i
                    };
                    template.SelectedDocumentFragments.Add(p);    
                }
                
            }
            return template;
        }

        public void PopulateDocumentFragements(List<IPolicyClass> fragements)
        {
            foreach (var fragement in fragements)
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

        public void PopulateFragments()
        {
            //_document.MoveCursorToStartOfBookmark();
        }

        public void PopulateImportantNotices(Enums.Statutory statutory, string importantNoticesUrl, string privacyStatementUrl, string fsgUrl, string termsOfEngagementUrl)
        {
            Document.MoveCursorToStartOfBookmark(Constants.WordBookmarks.ImportantNotes);
            Document.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.StatutoryInformation,statutory.ToString());
            //insert stat notices
            Document.InsertFile(importantNoticesUrl); //in settings
            Document.InsertPageBreak();

            //insert pivacy statement
            Document.InsertFile(privacyStatementUrl);//in settings


            switch (statutory)
            {
                case Enums.Statutory.Retail:
                    {
                        Document.InsertPageBreak();
                        Document.InsertFile(fsgUrl);//in settings
                        break;
                    }
                case Enums.Statutory.Wholesale:
                    {
                        Document.InsertPageBreak();
                        Document.InsertFile(termsOfEngagementUrl); //in settings
                        break;
                    }
                case Enums.Statutory.WholesaleWithResale:
                    {
                        Document.InsertPageBreak();
                        Document.InsertFile(fsgUrl); //in settings
                        Document.InsertPageBreak();
                        Document.InsertFile(termsOfEngagementUrl); //in settings
                        break;
                    }
            }
        }

        public void PopulateRemuneration(Enums.Remuneration remuneration, List<DocumentFragment> documents)
        {
            var frag = documents.Find(i => i.Title.Equals(remuneration.ToString(), StringComparison.OrdinalIgnoreCase));
            Document.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.Remuneration, remuneration.ToString());
            
            if (!Document.HasBookmark(Constants.WordBookmarks.AddRenumeration)) return;

            if (Document.HasBookmark(Constants.WordBookmarks.Renumeration))
            {
                Document.MoveCursorToStartOfBookmark(Constants.WordBookmarks.Renumeration);
                Document.DeletePage();
            }

            if (String.IsNullOrEmpty(frag.Url)) return;
            Document.MoveCursorToStartOfBookmark(Constants.WordBookmarks.AddRenumeration);
            Document.MoveCursorUp(1);
            Document.InsertPageBreak();
            Document.ChangePageOrientToPortrait();
            Document.InsertFile(frag.Url);
        }
    }
}
