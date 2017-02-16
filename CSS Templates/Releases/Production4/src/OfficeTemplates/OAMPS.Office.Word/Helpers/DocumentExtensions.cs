using System;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using OAMPS.Office.BusinessLogic.Helpers;
using OAMPS.Office.BusinessLogic.Presenters.SharePoint;
using OAMPS.Office.Word.Models.SharePoint;
using OAMPS.Office.Word.Models.Word;
using OAMPS.Office.Word.Views.Wizards;

namespace OAMPS.Office.Word.Helpers
{
    public static class DocumentExtensions
    {
        public static string GetPropertyValue(this Document document, string propertyName, bool trySharePointProperties = true)
        {
            //there are two types of properties metadata and custom.  
            //if not found in custom then search the metatdata properties.

            string returnValue = string.Empty;

            string value;
            if (GetBuiltInProperty(document, propertyName, out value)) return value;
            if (GetCustomLocalProperty(document, propertyName, out value)) return value;
            if (trySharePointProperties) if (GetCustomSharePointProperty(document, propertyName, out value)) return value;

            return string.Empty;
        }

        private static bool GetCustomSharePointProperty(Document document, string propertyName, out string returnValue)
        {
            try
            {
                MetaProperties metadataProps = document.ContentTypeProperties;
                //this will error if there are no sharepoint properties. No way around this.

                if (((metadataProps != null)))
                {
                    foreach (MetaProperty prop in metadataProps)
                    {
                        if (String.Equals(prop.Name, propertyName, StringComparison.OrdinalIgnoreCase))
                        {
                            if (((prop.Value != null)))
                            {
                                returnValue = prop.Value.ToString().ToLower();
                                return true;
                            }
                        }
                    }
                }
            }
// ReSharper disable EmptyGeneralCatchClause
            catch //supress all errors as ContentTypeProperties generates too many errors on startup.
// ReSharper restore EmptyGeneralCatchClause
            {
            }
            returnValue = string.Empty;
            return false;
        }

        //perfomacance issue here as we are doing two trips to sharepoint
        public static string GetPostalAddressFragment(this string office, bool isPart1)
        {
            if (office == null)
                return string.Empty;

            var list = new SharePointList(Properties.Settings.Default.SharePointContextUrl, "Office Addresses", Constants.SharePointQueries.GetItemByTitleQuery);
            var presenter = new SharePointListPresenter(list, null);
            var item = presenter.GetItemByTitle(office);

            if (item == null)
                return string.Empty;

            return isPart1 ? item.GetFieldValue("PostalAddressLine1") : item.GetFieldValue("PostalAddressLine2");
        }

        private static bool GetCustomLocalProperty(Document document, string propertyName, out string returnValue)
        {
            dynamic customProperties = document.CustomDocumentProperties;
            if (customProperties != null)
            {
                foreach (dynamic prop in customProperties)
                {
                    if (String.Equals(prop.Name, propertyName, StringComparison.OrdinalIgnoreCase))
                    {
                        if (((prop.Value != null)))
                        {
                            returnValue = prop.Value.ToString().ToLower();
                            return true;
                        }
                    }
                }
            }
            returnValue = string.Empty;
            return false;
        }

        private static bool GetBuiltInProperty(Document document, string propertyName, out string returnValue)
        {
            var builtInProperties = (DocumentProperties) document.BuiltInDocumentProperties;
            foreach (DocumentProperty p in builtInProperties)
            {
                if (p.Name == propertyName)
                {
                    {
                        returnValue = p.Value;
                        return true;
                    }
                }
            }
            returnValue = string.Empty;
            return false;
        }

        public static void LoadWizard(this Document document, Enums.FormLoadType loadType)
        {
            if (document.Path.Contains("http://") || (ThisAddIn.IsWizzardRunning && loadType != Enums.FormLoadType.RegenerateTemplate)) //do not load on sharepoint documents (or any document loaded via http).
                return;

            var officeDoc = new OfficeDocument(Globals.ThisAddIn.Application.ActiveDocument);

            dynamic templateName = ((DocumentProperties) (Globals.ThisAddIn.Application.ActiveDocument.BuiltInDocumentProperties))[WdBuiltInProperty.wdPropertyTitle].Value.ToString();

            if (!String.IsNullOrEmpty(templateName))
            {
                if (String.Equals(templateName, Constants.TemplateNames.PlacementSlip, StringComparison.OrdinalIgnoreCase))
                {
                }
                else if (String.Equals(templateName, Constants.TemplateNames.FileNote, StringComparison.OrdinalIgnoreCase)) //we do not show this wizard on docload.
                {
                    var wizard = new SummaryOfDiscussionWizard(officeDoc)
                        {
                            TopMost = true,
                            StartPosition = FormStartPosition.CenterParent,
                            Reload = (loadType != Enums.FormLoadType.OpenDocument)
                        };
#if DEBUG
                    wizard.TopMost = false;
#endif
                    wizard.Show();
                }
                else if (String.Equals(templateName, Constants.TemplateNames.PreRenewalAgenda, StringComparison.OrdinalIgnoreCase))
                {
                    var wizard = new PreRenewalAgendaWizard(officeDoc)
                        {
                            TopMost = true,
                            StartPosition = FormStartPosition.CenterParent,
                            Reload = (loadType != Enums.FormLoadType.OpenDocument)
                        };
#if DEBUG
                    wizard.TopMost = false;
#endif
                    wizard.Show();
                }
                else if (String.Equals(templateName, Constants.TemplateNames.InsuranceRenewalReport, StringComparison.OrdinalIgnoreCase))
                {
                    var wizard = new InsuranceRenewalReportWizard(officeDoc, loadType)
                        {
                            TopMost = true,
                            StartPosition = FormStartPosition.CenterParent,
                            Reload = (loadType != Enums.FormLoadType.OpenDocument)
                        };
#if DEBUG
                    wizard.TopMost = false;
#endif
                    wizard.Show();
                }
                else if (String.Equals(templateName, Constants.TemplateNames.RenewalLetter, StringComparison.OrdinalIgnoreCase))
                {
                    var wizard = new RenewalLetterWizard(officeDoc)
                        {
                            TopMost = true,
                            StartPosition = FormStartPosition.CenterParent,
                            Reload = (loadType != Enums.FormLoadType.OpenDocument)
                        };
#if DEBUG
                    wizard.TopMost = false;
#endif
                    wizard.Show();
                }
                else if (String.Equals(templateName, Constants.TemplateNames.ClientDiscovery, StringComparison.OrdinalIgnoreCase))
                {
                    var wizard = new ClientDiscoveryWizard(officeDoc, loadType)
                        {
                            TopMost = true,
                            StartPosition = FormStartPosition.CenterParent,
                            Reload = (loadType != Enums.FormLoadType.OpenDocument)
                        };
#if DEBUG
                    wizard.TopMost = false;
#endif
                    wizard.Show();
                }

                else if (String.Equals(templateName, Constants.TemplateNames.PreRenewalQuestionnaire, StringComparison.OrdinalIgnoreCase))
                {
                    var wizard = new FactFinderWizard(officeDoc, loadType)
                        {
                            TopMost = true,
                            StartPosition = FormStartPosition.CenterParent,
                            Reload = (loadType != Enums.FormLoadType.OpenDocument)
                        };
#if DEBUG
                    wizard.TopMost = false;
#endif
                    wizard.Show();
                }

                else if (String.Equals(templateName, Constants.TemplateNames.GenericLetter, StringComparison.OrdinalIgnoreCase))
                {
                    var wizard = new GenericLetterWizard(officeDoc)
                        {
                            TopMost = true,
                            StartPosition = FormStartPosition.CenterParent,
                            Reload = (loadType != Enums.FormLoadType.OpenDocument)
                        };
#if DEBUG
                    wizard.TopMost = false;
#endif
                    wizard.Show();
                }

                else if (String.Equals(templateName, Constants.TemplateNames.QuoteSlip, StringComparison.OrdinalIgnoreCase))
                {
                    var wizard = new QuoteSlipWizard(officeDoc, loadType)
                        {
                            TopMost = true,
                            StartPosition = FormStartPosition.CenterParent,
                            Reload = (loadType != Enums.FormLoadType.OpenDocument)
                        };
#if DEBUG
                    wizard.TopMost = false;
#endif
                    wizard.Show();
                }

                else if (String.Equals(templateName, Constants.TemplateNames.InsuranceManual, StringComparison.OrdinalIgnoreCase))
                {
                    var wizard = new InsuranceManualWizard(officeDoc, loadType)
                        {
                            TopMost = true,
                            StartPosition = FormStartPosition.CenterParent,
                            Reload = (loadType != Enums.FormLoadType.OpenDocument)
                        };
#if DEBUG
                    wizard.TopMost = false;
#endif
                    wizard.Show();

                }
            }
        }

        public static bool CheckPropertyExists(this Document document, string propertyName)
        {
            var customProperties = (DocumentProperties) document.CustomDocumentProperties;
            return customProperties.Cast<DocumentProperty>().Any(prop => String.Equals(prop.Name.ToLower(), propertyName, StringComparison.OrdinalIgnoreCase));
        }

        public static void UpdateOrCreatePropertyValue(this Document document, string propertyName, string value)
        {
            var customProperties = (DocumentProperties) document.CustomDocumentProperties;
            if (CheckPropertyExists(document, propertyName))
            {
                customProperties[propertyName].Value = value;
            }
            else
            {
                customProperties.Add(propertyName, false, MsoDocProperties.msoPropertyTypeString, value);
            }
        }
    }
}