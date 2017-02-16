using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using OAMPS.Office.BusinessLogic.Helpers;
using OAMPS.Office.BusinessLogic.Interfaces;
using OAMPS.Office.BusinessLogic.Interfaces.Template;
using OAMPS.Office.BusinessLogic.Interfaces.Wizards;
using OAMPS.Office.BusinessLogic.Interfaces.Word;
using OAMPS.Office.BusinessLogic.Models.Template;
using OAMPS.Office.BusinessLogic.Models.Wizards;
using OAMPS.Office.BusinessLogic.Presenters.SharePoint;
using OAMPS.Office.BusinessLogic.Presenters.Word;
using OAMPS.Office.Word.Models.SharePoint;
using OAMPS.Office.Word.Models.Word;
using OAMPS.Office.Word.Properties;
using OAMPS.Office.Word.Views.Wizards;
using Enums = OAMPS.Office.Word.Helpers.Enums;


namespace OAMPS.Office.Word.Views.Word
{
    public partial class ConvertFactFinderToQuoteSlip : Form , IBaseView
    {
        private readonly ConvertFactFinderToQuoteSlipPresenter _presenter;
        private readonly IDocument _document;
        private List<IDocument> _convertedDocuments = new List<IDocument>();
        public ConvertFactFinderToQuoteSlip()
        {
            InitializeComponent();
        }

        public ConvertFactFinderToQuoteSlip(IDocument document)
        {
            InitializeComponent();
            _presenter = new ConvertFactFinderToQuoteSlipPresenter(document, this);
            _document = document;
        }

        public void DisplayMessage(string text, string caption)
        {
            MessageBox.Show(text, caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            try
            {
                SetupProgressStyle();
                Task.Factory.StartNew(() => Thread.Sleep(100)).Wait();
                Convert();
                _presenter.CloseInfoPanel();

                var w = new BaseWizardForm();
                var t = new QuoteSlip();
                t.DocumentTitle = "Quote Slip";
                w.LogUsage(t, Enums.UsageTrackingType.ConvertDocument);


                ThisAddIn.IsWizzardRunning = false;
            }
            catch (Exception)
            {

                //todo: implement logger
            }
            finally
            {
                Close();
            }
        }



        private void SetupProgressStyle()
        {
            Height = Height - clbQuestions.Height;
            Width = Width - 150;
            lblMessage.Text = @"Please wait while we generate the Quote Slips";
            FormBorderStyle = FormBorderStyle.None;
            //System.Threading.SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
            Visible = false;
            //backgroundWorker.RunWorkerAsync();
            progressBar.Visible = true;
            btnOk.Visible = false;
            btnCancel.Visible = false;

            lblResults.Visible = true;
            clbQuestions.Visible = false;
            CenterToScreen();
            CenterToParent();
            Refresh();
            Visible = true;
        }

        private void Convert()
        {
            // var selectedItems = new List<IQuestionClass>();
            var totalItemCount = clbQuestions.CheckedItems.Count;
            var processingIndex = 1;
            var template = LoadTemplateData();

            foreach (var i in clbQuestions.CheckedItems)
            {
                var converted = DoConvert(i as QuoteSlipSchedules, template);
                _convertedDocuments.Add(converted);
                processingIndex = ReportProgress(totalItemCount, processingIndex);
            }

            var t = new System.Windows.Forms.Timer {Interval = 2000, Enabled = true};
            t.Start();
            t.Tick += new EventHandler(t_Tick);
            
        }
       

        private IBaseTemplate LoadTemplateData()
        {
            var document = new OfficeDocument(Globals.ThisAddIn.Application.ActiveDocument);

            var template = new FactFinder();
            //read content controls and set template values.
            foreach (PropertyInfo pInfo in template.GetType().GetProperties())
            {
                if (pInfo == null)
                    break;

                if (!String.Equals(pInfo.PropertyType.Name, "string", StringComparison.OrdinalIgnoreCase)) continue;
                string value = document.ReadContentControlValue(pInfo.Name);
                pInfo.SetValue(template, value, null);
            }

            //read cover page and logo property values and set template values
            template.CoverPageTitle = document.GetPropertyValue(Constants.WordDocumentProperties.CoverPageTitle);
            template.LogoTitle = document.GetPropertyValue(Constants.WordDocumentProperties.LogoTitle);
            return template;
        }






        void t_Tick(object sender, EventArgs e)
        {
            var timer = sender as System.Windows.Forms.Timer;
            if (timer != null) timer.Enabled = false;

            foreach (var doc in _convertedDocuments)
            {
                doc.CloseInformationPanel(false);
            }
        }

        private int ReportProgress(int totalItemCount, int processingIndex)
        {
            //Task.Factory.StartNew(() => Thread.Sleep(4000)).Wait();

            if (clbQuestions.CheckedItems.Count > 1)
            {
                Thread.Sleep(4000);
            }

            var percent = (100/totalItemCount)*processingIndex;
            progressBar.Value = percent;
            progressBar.Refresh();
            lblResults.Text = (percent.ToString(CultureInfo.InvariantCulture) + @"%");
            lblResults.Refresh();
            //worker.ReportProgress((100 / totalItemCount) * processingIndex);
            processingIndex++;
            return processingIndex;
        }

        private IDocument DoConvert(QuoteSlipSchedules i, IBaseTemplate template)
        {
            var quoteSlipDoc = _document.OpenFile(Constants.CacheNames.GenerateQuoteSlip, Settings.Default.TemplateQuoteSlip);
            _presenter.ConvertToQuoteSlip(Settings.Default.TemplateQuoteSlip, quoteSlipDoc, i as QuoteSlipSchedules, template);
            return quoteSlipDoc;
        }

        private List<IQuoteSlipSchedules> GetQuestionsForIncludedPolicies(string includedPolicies)
        {
            var spList = new SharePointList(Settings.Default.SharePointContextUrl, "Quote Slip Schedules", Constants.SharePointQueries.AllItemsSortByTitle); //todo: move listname to settings
            var presenter = new SharePointListPresenter(spList, this);
            var items = presenter.GetItems();


            var spListQuestions = new SharePointList(Settings.Default.SharePointContextUrl, Settings.Default.PreRenewalQuestionareMappingsListName, Constants.SharePointQueries.AllItemsSortByTitle); //todo: move listname to settings
            var presenterQuestions = new SharePointListPresenter(spListQuestions, this);
            var itemsQuestions = presenterQuestions.GetItems();

          
            var qs = includedPolicies.Split(';').Select(delegate(string p)
                {
                    if (String.IsNullOrEmpty(p)) p = "-123"; //saftey net incase there are any blank lookup values in sharepoint, better not to display them in the popup

                    var g = items.FirstOrDefault(i => i.GetLookupFieldValueIdArray(Constants.SharePointFields.QuoteSlipSchedulesFfLookupId).Contains(p));//presenter.GetQuoteSlipSchedule(p);
                    var title = string.Empty;
                  
                    if (g != null)
                    {
                        title = g.Title;

                        if (g.Title.ToLower().Contains("workers"))
                        {
                            var question = itemsQuestions.FirstOrDefault(i => i.GetFieldValue(Constants.SharePointFields.FieldId) == p);//presenter.GetQuoteSlipSchedule(p);    

                            title = question.Title;
                        }
                        
                        
                    }

                    return g == null ? null : new QuoteSlipSchedules
                        {
                            Title = title,
                            Url = g.FileRef, 
                            Id = g.GetFieldValue(Constants.SharePointFields.FieldId),
                            LinkedQuestionId = g.GetLookupFieldValueIdArray(Constants.SharePointFields.QuoteSlipSchedulesFfLookupId)
                            
                        };
                }).Cast<IQuoteSlipSchedules>().ToList();

            return qs;
        }

        private void ConvertFactFinderToQuoteSlip_Load(object sender, EventArgs e)
        {
            CenterToScreen();
            CenterToParent();
            
            var includedPolicies = _presenter.GetIncludedPolicies(_document);
            var questions = GetQuestionsForIncludedPolicies(includedPolicies);

            questions.RemoveAll(i => i == null);

            clbQuestions.DataSource = questions;
            clbQuestions.DisplayMember = "Title";
            clbQuestions.ValueMember = "Title";

            for (var i = 0; i < clbQuestions.Items.Count; i++)
            {
                clbQuestions.SetItemChecked(i,true);
            }
        }
    }
}
