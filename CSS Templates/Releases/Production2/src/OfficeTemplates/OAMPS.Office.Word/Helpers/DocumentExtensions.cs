
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Core;
using System;
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
                else if (GetCustomLocalProperty(document, propertyName, out value)) return value;
                else if(trySharePointProperties) if(GetCustomSharePointProperty(document, propertyName, out value)) return value;

            return string.Empty;
        }

        private static bool GetCustomSharePointProperty(Document document, string propertyName, out string returnValue)
        {
            try
            {
                var metadataProps = document.ContentTypeProperties;
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

        //TODO FIX THIS
        public static string GetPostalAddressFragment(this string address, bool isPart1)
        {
            if (address == null)
                return null;

            var lowercase = address.ToLowerInvariant().Trim();
          
            switch (lowercase)
            {
                case "po box 10016 adelaide bc sa 5000":
                    {
                        if (isPart1)
                        {
                            return "PO Box 10016";
                        }
                        else
                        {
                            return "ADELAIDE BC SA 5000";
                        }
                        
                    }
               
                case "po box 852 east melbourne vic 8002":
                    {
                        if (isPart1)
                        {
                            return "PO Box 852";
                        }
                        else
                        {
                            return "EAST MELBOURNE VIC 8002";
                        }

                    }
                case "po box 1454 hobart tas 7001":
                    {
                        if (isPart1)
                        {
                            return "PO Box 1454";
                        }
                        else
                        {
                            return "HOBART TAS 7001";
                        }
                        
                    }
                case "po box 1898 north sydney nsw 2059":
                    {
                        if (isPart1)
                        {
                            return "PO Box 1898";
                        }
                        else
                        {
                            return "NORTH SYDNEY NSW 2059";
                        }
                        
                    }
                case "po box 3036 parramatta nsw 2124":
                    {
                        if (isPart1)
                        {
                            return "PO Box 3036";
                        }
                        else
                        {
                            return "PARRAMATTA NSW 2124";
                        }
                        
                    }
                default:
                    {
                        return null;
                    }


            }
            
        }

        private static bool GetCustomLocalProperty(Document document, string propertyName, out string returnValue)
        {
            var customProperties = document.CustomDocumentProperties;
            if (customProperties != null)
            {
                foreach (var prop in customProperties)
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
            if (document.Path.Contains("http://")|| (ThisAddIn.IsWizzardRunning && loadType!= Enums.FormLoadType.RegenerateTemplate)) //do not load on sharepoint documents (or any document loaded via http).
                return;

            var officeDoc = new OfficeDocument(Globals.ThisAddIn.Application.ActiveDocument);
            
            var templateName =
                ((DocumentProperties) (Globals.ThisAddIn.Application.ActiveDocument.BuiltInDocumentProperties))[
                    WdBuiltInProperty.wdPropertyTitle].Value.ToString();

            if (!String.IsNullOrEmpty(templateName))
            {
                if (String.Equals(templateName, BusinessLogic.Helpers.Constants.TemplateNames.PlacementSlip,
                                  StringComparison.OrdinalIgnoreCase))
                {

                }
                else if (String.Equals(templateName, BusinessLogic.Helpers.Constants.TemplateNames.FileNote,
                                  StringComparison.OrdinalIgnoreCase)) //we do not show this wizard on docload.
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
                else if (String.Equals(templateName, BusinessLogic.Helpers.Constants.TemplateNames.PreRenewalAgenda,
                                       StringComparison.OrdinalIgnoreCase))
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
                else if (String.Equals(templateName,
                                       BusinessLogic.Helpers.Constants.TemplateNames.InsuranceRenewalReport,
                                       StringComparison.OrdinalIgnoreCase))
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
                else if (String.Equals(templateName,
                                       BusinessLogic.Helpers.Constants.TemplateNames.RenewalLetter,
                                       StringComparison.OrdinalIgnoreCase))
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
                else if (String.Equals(templateName,
                                       BusinessLogic.Helpers.Constants.TemplateNames.ClientDiscovery,
                                       StringComparison.OrdinalIgnoreCase))
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
