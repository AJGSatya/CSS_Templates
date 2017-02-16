using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using OAMPS.Office.BusinessLogic.Helpers;
using OAMPS.Office.BusinessLogic.Interfaces;
using OAMPS.Office.BusinessLogic.Interfaces.Template;
using OAMPS.Office.BusinessLogic.Interfaces.Wizards;
using OAMPS.Office.Word.Helpers;
using OAMPS.Office.Word.Helpers.Controls;
using OAMPS.Office.Word.Models.SharePoint;
using OAMPS.Office.Word.Models.Word;
using OAMPS.Office.Word.Properties;
using OAMPS.Office.BusinessLogic.Presenters.SharePoint;
using OAMPS.Office.BusinessLogic.Presenters;
using OAMPS.Office.BusinessLogic.Models.Template;
using System.Threading.Tasks;
using System.Threading;

namespace OAMPS.Office.Word.Views.Wizards
{
    public partial class SummaryOfDiscussionWizard : BaseForm
    {
        private readonly SummaryOfDiscussionPresenter _presenter;
        public SummaryOfDiscussionWizard(OfficeDocument document)
        {
            InitializeComponent();

            tbcWizardScreens.SelectedIndexChanged += new EventHandler(tbcWizardScreens_SelectedIndexChanged);

            //send marketing template to the presenter
            
            _presenter = new SummaryOfDiscussionPresenter(document, this);
            BasePresenter = _presenter;
        }

        private void SummaryOfDiscussionWizard_Load(object sender, EventArgs e)
        {
            //pbOampsLogoFull.Left = (ClientSize.Width - pbOampsLogoFull.Width) / 2;
            txtClientName.Focus();
            txtClientName.Select();

            var uiScheduler = TaskScheduler.FromCurrentSynchronizationContext();            
            base.LoadGenericImageTabs(uiScheduler, tbcWizardScreens, lblCoverPageTitle.Text, lblLogoTitle.Text);

            if (Reload) // this happens if they click the button on the ribbon.
            {
                var template = new SummaryOfDiscussions();
                var values = _presenter.LoadData(template);
                var v = ((ISummaryOfDiscussions)values);
                ReloadFields(v);
            }
        }

        private void ReloadFields(ISummaryOfDiscussions v)
        {
            txtClientName.Text = v.ClientName;
            txtClientCode.Text = v.ClientCode;
            txtExecutiveEmail.Text = v.ExecutiveEmail;
            txtExecutiveName.Text = v.ExecutiveName;
            txtExecutivePhone.Text = v.ExecutivePhone;
            txtExecutiveMobile.Text = v.ExecutiveMobile;
            
            txtExecutiveDepartment.Text = v.ExecutiveDepartment;
            lblLogoTitle.Text = v.LogoTitle;
            lblCoverPageTitle.Text = v.CoverPageTitle;
            txtClientContactName.Text = v.ClientContactName;


            bool outValue = false;
             if (Boolean.TryParse(v.IsDiscussedWithCaller, out outValue))
            {
                rdoCaller.Checked = outValue;
            }

             outValue = false;
             if (Boolean.TryParse(v.IsDiscussedInPerson, out outValue))
             {
                 rdoPerson.Checked = outValue;
             }

             outValue = false;
             if (Boolean.TryParse(v.IsDiscussedWithCustomer, out outValue))
             {
                 rdoCustomer.Checked = outValue;
             }

             outValue = false;
             if (Boolean.TryParse(v.IsOther, out outValue))
             {
                 rdoOther.Checked = outValue;
             }

             outValue = false;
             if (Boolean.TryParse(v.IsDiscussedWithUnderWriter, out outValue))
             {
                 rdoUnderWriter.Checked = outValue;
             }

             outValue = false;
             if (Boolean.TryParse(v.IsDiscussedByPhone, out outValue))
             {
                 rdoPhone.Checked = outValue;
             }

          


            //bool outValue = false;
            //var propVal = _presenter.ReadDocumentProperty(Constants.WordDocumentProperties.FileNoteIsCaller);
            //if (Boolean.TryParse(propVal, out outValue))
            //{
            //    rdoCaller.Checked = outValue;
            //}

            //propVal = _presenter.ReadDocumentProperty(Constants.WordDocumentProperties.FileNoteIsClient);
            //if (Boolean.TryParse(propVal, out outValue))
            //{
            //    rdoCustomer.Checked = outValue;
            //}

            //propVal = _presenter.ReadDocumentProperty(Constants.WordDocumentProperties.FileNoteIsUnderwriter);
            //if (Boolean.TryParse(propVal, out outValue))
            //{
            //    rdoUnderWriter.Checked = outValue;
            //}

            //propVal = _presenter.ReadDocumentProperty(Constants.WordDocumentProperties.FileNoteByPhone);
            //if (Boolean.TryParse(propVal, out outValue))
            //{
            //    rdoPhone.Checked = outValue;
            //}

            //propVal = _presenter.ReadDocumentProperty(Constants.WordDocumentProperties.FileNoteInPerson);
            //if (Boolean.TryParse(propVal, out outValue))
            //{
            //    rdoPerson.Checked = outValue;
            //}

            //propVal = _presenter.ReadDocumentProperty(Constants.WordDocumentProperties.FileNoteOther);
            //if (Boolean.TryParse(propVal, out outValue))
            //{
            //    rdoOther.Checked = outValue;
            //}
        }

        void tbcWizardScreens_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnNext.Text = tbcWizardScreens.SelectedIndex == tbcWizardScreens.TabCount - 1 ? "&Finish" : "&Next";
            btnBack.Enabled = tbcWizardScreens.SelectedIndex != 0;
            btnNext.Enabled = btnNext.Text != @"&Finish" || LoadComplete;
        }

        private void SwitchTab(int index)
        {
            tbcWizardScreens.SelectedIndex = index;
        }

        private void PopulateDocument()
        {
            //buid the marketing template
            var template = new SummaryOfDiscussions
                {
                    DocumentTitle = BasePresenter.ReadDocumentProperty("Title"), //Constants.TemplateNames.FileNote,
                    DocumentSubTitle = string.Empty,

                    ClientName = txtClientName.Text,
                    ClientCode = txtClientCode.Text,
                    ClientContactName = txtClientContactName.Text,

                    DateDiscussion = dateDiscussion.Text,
                    TimeDiscussion = timeDiscussion.Text,

                    ExecutiveEmail = txtExecutiveEmail.Text,
                    ExecutiveMobile = txtExecutiveMobile.Text,
                    ExecutiveName = txtExecutiveName.Text,
                    ExecutivePhone = txtExecutivePhone.Text,


                    ExecutiveDepartment = txtExecutiveDepartment.Text,

                    IsDiscussedByPhone = rdoPhone.Checked.ToString(),
                    IsDiscussedInPerson = rdoPerson.Checked.ToString(),
                    IsOther = rdoOther.Checked.ToString(),

                    IsDiscussedWithCaller = rdoCaller.Checked.ToString(),
                    IsDiscussedWithCustomer = rdoCustomer.Checked.ToString(),
                    IsDiscussedWithUnderWriter = rdoUnderWriter.Checked.ToString()
                };

            var baseTemplate = (BaseTemplate) template;

            var logoTab = tbcWizardScreens.TabPages[Constants.ControlNames.TabPageLogosName];

            PopulateLogosToTemplate(logoTab, ref baseTemplate);

            //foreach (Control c in logoTab.Controls)
            //{
            //    if (c.GetType() == typeof(ValueRadioButton))
            //    {
            //        var v = ((ValueRadioButton)c);
            //        if (v.Checked)
            //        {
            //            template.LogoImageUrl = v.GetValue(Constants.RadioButtonValues.ImageUrl);
            //        }
            //    }
            //}

            var covberTab = tbcWizardScreens.TabPages[Constants.ControlNames.TabPageCoverPagesName];
            //foreach (Control c in covberTab.Controls)
            //{
            //    if (c.GetType() == typeof(ValueRadioButton))
            //    {
            //        var v = ((ValueRadioButton)c);
            //        if (v.Checked)
            //        {
            //            template.CoverPageImageUrl = v.GetValue(Constants.RadioButtonValues.ImageUrl);
            //        }
            //    }
            //}
            PopulateCoversToTemplate(covberTab, ref baseTemplate);

            //populate the content controls
            _presenter.PopulateData(template);
            _presenter.CreatePropertiesForRadioButtons(rdoCustomer.Checked, rdoCaller.Checked, rdoUnderWriter.Checked,
                                                       rdoPhone.Checked, rdoPerson.Checked, rdoOther.Checked);

            //change the graphics selected
            //if (Streams == null) return;
            _presenter.PopulateGraphics(template, String.Empty, string.Empty);

            //tracking
            LogUsage(template,
                     Reload == true
                         ? Helpers.Enums.UsageTrackingType.UpdateData
                         : Helpers.Enums.UsageTrackingType.NewDocument);
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            if (String.Equals(btnNext.Text, "&Finish", StringComparison.CurrentCultureIgnoreCase))
            {
                try
                {
                    this.Cursor = Cursors.WaitCursor;
                    BasePresenter.SwitchScreenUpdating(false);
                    //call presenter to populate
                    PopulateDocument();
                }
                catch (Exception ex)
                {

                    OnError(ex);
                }
                finally
                {
                    this.Cursor = Cursors.Default;
                    BasePresenter.SwitchScreenUpdating(true);
                    //thie information panel loads when a document is in sharePoint that has metadata
                    //clients don't wish to see this so force the close of the panel once the wizard completes.
                    _presenter.CloseInformationPanel();
                }
            

       

                Close();
            }
            else
            {
                SwitchTab(tbcWizardScreens.SelectedIndex + 1);
            }
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            SwitchTab(tbcWizardScreens.SelectedIndex - 1);
        }

        private void btnAccountExecutiveLookup_Click(object sender, EventArgs e)
        {
            var peoplePicker = new Popups.PeoplePicker(txtExecutiveName.Text,this);
            if (peoplePicker.SelectedUser == null)
            {
                this.TopMost = false;

                peoplePicker.ShowDialog();
                if (peoplePicker.SelectedUser == null) return;
            }

            

            txtExecutiveName.Text = peoplePicker.SelectedUser.DisplayName;            
            txtExecutiveEmail.Text = peoplePicker.SelectedUser.EmailAddress;
            txtExecutivePhone.Text = peoplePicker.SelectedUser.VoiceTelephoneNumber;
            txtExecutiveMobile.Text = peoplePicker.SelectedUser.Mobile;
            txtExecutiveDepartment.Text = peoplePicker.SelectedUser.Branch;
        }

    }
}
