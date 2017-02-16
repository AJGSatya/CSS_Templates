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
using OAMPS.Office.BusinessLogic.Interfaces;
using OAMPS.Office.BusinessLogic.Presenters.SharePoint;
using OAMPS.Office.Word.Helpers;
using OAMPS.Office.Word.Models.SharePoint;
using OAMPS.Office.Word.Properties;
using OAMPS.Office.Word.Views.Wizards;

namespace OAMPS.Office.Word.Views.Help
{
    public partial class HelpWizard : BaseForm
    {
        public HelpWizard()
        {
            InitializeComponent();

            //todo: max we need to move this off accessing the word document directly.
            //todo: move this to the presenter for Help.

            var find = Globals.ThisAddIn.Application.Selection.Find;

            var r = Globals.ThisAddIn.Application.Application.Selection.Move();
            

            //error handling needed.
// ReSharper disable UseIndexedProperty
            find.set_Style(Globals.ThisAddIn.Application.ActiveDocument.Styles[BusinessLogic.Helpers.Constants.WordStyles.Heading1]);
// ReSharper restore UseIndexedProperty
            find.Text = String.Empty;
            find.Forward = false;
            find.MatchWildcards = true;
            find.Execute();


            var heading = Globals.ThisAddIn.Application.Selection.Text;

            if (!String.IsNullOrEmpty(heading))
            {
                heading = heading.Replace("\r",string.Empty).Trim();
                var templateName =  
                    ((DocumentProperties) (Globals.ThisAddIn.Application.ActiveDocument.BuiltInDocumentProperties))[
                        WdBuiltInProperty.wdPropertyTitle].Value.ToString(); //todo max not able to use anthing off globals.thisaddin in wizard screens.  move this to a helpPresenter.
                
                //todo  max move these to constants, there is a caml query constants class

                var query = "<View>" +
                                "<Query>" +
                                    "<Where>" +
                                        "<And>" +
                                            "<Eq><FieldRef Name='Template' /><Value Type='Lookup'>" + templateName + "</Value></Eq>" +
                                            "<Eq><FieldRef Name='Title' /><Value Type='Text'>" + heading + "</Value></Eq>" +
                                        "</And>" +
                                    "</Where>" +
                                "</Query>" +
                            "</View>";

                //todo  max move the list to a settings
                var helpList = new SharePointList(Settings.Default.SharePointContextUrl, "Word Help Content", query);
                var presenter = new SharePointListPresenter(helpList, this);
                
                var fitem = presenter.GetHelpItems().FirstOrDefault();
                
                if(fitem!=null)
                webHelpWindow.DocumentText = fitem.GetFieldValue("Content"); //todo: max move the fieldname to a constants
                else
                {
                    //todo  max move these to constants, there is a caml query constants class
                    
                    var generalHelp = "<View>" +
                                "<Query>" +
                                    "<Where>" +
                                        "<And>" +
                                            "<Eq><FieldRef Name='Template' /><Value Type='Lookup'>" + templateName + "</Value></Eq>" +
                                            "<Eq><FieldRef Name='Title' /><Value Type='Text'>" + "Wizard Help" + "</Value></Eq>" +
                                        "</And>" +
                                    "</Where>" +
                                "</Query>" +
                            "</View>"; 
                    helpList.UpdateCamlQuery(generalHelp);
                    var gitem = presenter.GetHelpItems().FirstOrDefault();
                    webHelpWindow.DocumentText = gitem != null ? gitem.GetFieldValue("Content") : "Unable to find the help content for this document";
                    
                }
            }        

        //    Globals.ThisAddIn.Application.Application.Selection.GoTo(WdGoToItem.wdGoToTable)
        }

        public void DisplayMessage(string text, string caption)
        {
            throw new NotImplementedException();
        }

        private void webHelpWindow_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {

        }
    }
}
