using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OAMPS.Office.BusinessLogic.Interfaces;
using OAMPS.Office.BusinessLogic.Interfaces.Template;
using OAMPS.Office.BusinessLogic.Interfaces.Word;
using OAMPS.Office.BusinessLogic.Helpers;
namespace OAMPS.Office.BusinessLogic.Presenters
{
    public class BasePresenter
    {
        internal readonly IDocument Document;
        internal readonly IBaseView View;

        public BasePresenter()
        {
            
        }

        public void GenerateNewTemplate(string cacheName, string url= null)
        {
            Document.OpenFile(cacheName, url);
        }

        public BasePresenter(IDocument document, IBaseView view)
        {
              Document = document;
            View = view;
        }
        
        public void UpdateFields()
        {
            Document.UpdateFields();
        }

        public void ActivateDocument()
        {
            Document.Activate();
        }

        public int OnDocumentLoaded()
        {
            if (Document.IsRangeReadOnly())
            {
                return Document.TurnOffProtection(string.Empty);
            }
            return 0;
        }

        public void OnDocumentLoadCompleted(int protectionType)
        {
            if (protectionType != 0)
            {
                Document.TurnOnProtection(protectionType, string.Empty);
            }            
        }

        public void PopulateData(IBaseTemplate template)
        {
            //populate field controls.
            foreach (var pInfo in template.GetType().GetProperties())
            {
                if (pInfo == null)
                    break;

                var value = String.Empty;
                if (pInfo.GetValue(template, null) != null)
                    value = pInfo.GetValue(template, null).ToString();

                    Document.PopulateControl(pInfo.Name, value);
            }

            Document.UpdateFields();
            CloseInformationPanel();
        }

        public object LoadData(IBaseTemplate template)
        {
            //read content controls and set template values.
            foreach (var pInfo in template.GetType().GetProperties())
            {
                if (pInfo == null)
                    break;

                if (!String.Equals(pInfo.PropertyType.Name, "string", StringComparison.OrdinalIgnoreCase)) continue;
                var value = Document.ReadContentControlValue(pInfo.Name);
                pInfo.SetValue(template, value, null);
            }

            //read cover page and logo property values and set template values
            template.CoverPageTitle = Document.GetPropertyValue(Helpers.Constants.WordDocumentProperties.CoverPageTitle);
            template.LogoTitle = Document.GetPropertyValue(Helpers.Constants.WordDocumentProperties.LogoTitle);
            return template;
        }

        public void CloseInformationPanel(bool delay = false)
        {
            Document.CloseInformationPanel(delay);
        }

        public void PopulateGraphics(IBaseTemplate template, string previousCoverPageTitle, string previousLogoTitle)
        {
            var names = new string[2];
            var values = new string[2];

            //do not populate array with images if their values have not changed.
            if (!String.Equals(template.CoverPageTitle, previousCoverPageTitle, StringComparison.OrdinalIgnoreCase))
            {
                if (template.CoverPageImageUrl != null)
                {
                    names[0] = Constants.ImageProperties.Theme;
                    values[0] = template.CoverPageImageUrl;
                }
            }

            if (!String.Equals(template.LogoTitle, previousLogoTitle, StringComparison.OrdinalIgnoreCase))
            {
                if (template.LogoImageUrl != null)
                {
                    names[1] = Constants.ImageProperties.CompanyLogo;
                    values[1] = template.LogoImageUrl;
                }
            }
          
            Document.ChangeDocumentImages(names, values);
            Document.UpdateOrCreatePropertyValue(Helpers.Constants.WordDocumentProperties.CoverPageTitle, template.CoverPageTitle);
            Document.UpdateOrCreatePropertyValue(Helpers.Constants.WordDocumentProperties.LogoTitle, template.LogoTitle);

        }

        public void SwitchScreenUpdating(bool show)
        {
            Document.SwitchScreenUpdating(show);
        }

        public bool CheckName(string name)
        {
            return String.Equals(Document.Name, name, StringComparison.OrdinalIgnoreCase);
        }

        public void CloseDocument(bool saveChanges)
        {
            Document.CloseMe(saveChanges);
        }

        public void DeleteAllComments()
        {
            Document.DeleteAllComments();
        }

        public void MoveToStartOfDocument()
        {
            Document.MoveToStartOfDocument();
        }

        public string ReadDocumentProperty(string property)
        {
            if (string.IsNullOrEmpty(property)) return String.Empty;
            return Document.GetPropertyValue(property);
        }

        public void CreateOrUpdateDocumentProperty(string property, string value )
        {
            if (string.IsNullOrEmpty(property)) return;
            
            Document.UpdateOrCreatePropertyValue(property, value);
        }

        public static void LogUsage()
        {

        }
    }
}
