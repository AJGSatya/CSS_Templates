using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using OAMPS.Office.BusinessLogic.Helpers;
using OAMPS.Office.BusinessLogic.Interfaces;
using OAMPS.Office.BusinessLogic.Presenters;
using OAMPS.Office.BusinessLogic.Presenters.SharePoint;
using OAMPS.Office.Word.Helpers;

using OAMPS.Office.Word.Models.SharePoint;
using OAMPS.Office.Word.Models.Word;
using OAMPS.Office.Word.Properties;
using OAMPS.Office.Word.Views.Wizards;

namespace OAMPS.Office.Word.Views.Help
{
    public partial class HelpWizard : BaseForm
    {

        private HelpContentPresenter _presenter;

        public HelpWizard(string templateName)
        {
            InitializeComponent();


            var officeDoc = new OfficeDocument(Globals.ThisAddIn.Application.ActiveDocument);
            _presenter = new HelpContentPresenter(officeDoc, this);

            base.BasePresenter = _presenter;


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

            var heading = _presenter.FindHeadingTextForCurrentDocument();

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
                else
                {
                    var generalHelp = string.Format(Constants.SharePointQueries.HelpGetItemQuery, templateName, Constants.SharePointFields.WizardHelp);
                    helpList.UpdateCamlQuery(generalHelp);
                    var gitem = presenter.GetHelpItems().FirstOrDefault();
                    if (gitem != null)
                    {
                        return gitem.GetFieldValue(Constants.SharePointFields.Content);
                    }

                }
            }
            return "Unable to find the help content for this document";
        }
      

        private void webHelpWindow_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {

        }
    }
}
