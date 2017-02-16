using System;
using System.Linq;
using System.Windows.Forms;
using OAMPS.Office.BusinessLogic.Helpers;
using OAMPS.Office.BusinessLogic.Interfaces.SharePoint;
using OAMPS.Office.BusinessLogic.Presenters;
using OAMPS.Office.BusinessLogic.Presenters.SharePoint;
using OAMPS.Office.BusinessLogic.Presenters.Word;
using OAMPS.Office.Word.Models.SharePoint;
using OAMPS.Office.Word.Models.Word;
using OAMPS.Office.Word.Properties;
using OAMPS.Office.Word.Views.Wizards;

namespace OAMPS.Office.Word.Views.Help
{
    public partial class HelpWizard : BaseWizardForm
    {
        private readonly HelpContentWizardPresenter _wizardPresenter;

        public HelpWizard(string templateName)
        {
            InitializeComponent();


            var officeDoc = new OfficeDocument(Globals.ThisAddIn.Application.ActiveDocument);
            _wizardPresenter = new HelpContentWizardPresenter(officeDoc, this);

            BaseWizardPresenter = _wizardPresenter;


//            var find = Globals.ThisAddIn.Application.Selection.Find;

//            var r = Globals.ThisAddIn.Application.Application.Selection.Move();


//            //error handling needed.
//// ReSharper disable UseIndexedProperty
//            find.set_Style(Globals.ThisAddIn.Application.ActiveDocument.Styles[BusinessLogic.Helpers.Constants.WordStyles.Heading1]);
//// ReSharper restore UseIndexedProperty
//            find.Text = String.Empty;
//            find.Forward = false;
//            find.MatchWildcards = true;
//            find.Execute();

//            BusinessLogic.Helpers.Constants.WordStyles.Heading1
//            var heading = Globals.ThisAddIn.Application.Selection.Text;

            string heading = _wizardPresenter.FindHeadingTextForCurrentDocument();

            webHelpWindow.DocumentText = GetHelpContentFromSharePoint(heading, Settings.Default.SharePointContextUrl, Settings.Default.HelpContentListName, templateName);

            //    Globals.ThisAddIn.Application.Application.Selection.GoTo(WdGoToItem.wdGoToTable)
        }


        public string GetHelpContentFromSharePoint(string heading, string sPContext, string helpListName, string templateName)
        {
            if (!String.IsNullOrEmpty(heading))
            {
                heading = heading.Replace("\r", string.Empty).Trim();
                var query = string.Format(Constants.SharePointQueries.HelpGetItemQuery, templateName, heading);

                //todo  max move the list to a settings
                var helpList = new SharePointList(sPContext, helpListName, query);
                var presenter = new SharePointListPresenter(helpList, this);

                var fitem = presenter.GetHelpItems().FirstOrDefault();

                if (fitem != null)
                    return fitem.GetFieldValue(Constants.SharePointFields.Content);
                
                
                var generalHelp = string.Format(Constants.SharePointQueries.HelpGetItemQuery, templateName, Constants.SharePointFields.WizardHelp);
                helpList.UpdateCamlQuery(generalHelp);
                var gitem = presenter.GetHelpItems().FirstOrDefault();
                return gitem != null ? gitem.GetFieldValue(Constants.SharePointFields.Content) : "Unable to find the help content for this document";
            }
            return "Unable to find the help content for this document";
        }


        private void webHelpWindow_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
        }
    }
}