using System;
using System.Threading.Tasks;
using System.Windows.Forms;
using OAMPS.Office.BusinessLogic.Helpers;
using OAMPS.Office.BusinessLogic.Interfaces.Template;
using OAMPS.Office.BusinessLogic.Models.Template;
using OAMPS.Office.BusinessLogic.Presenters.Wizards;
using OAMPS.Office.Word.Models.Word;
using OAMPS.Office.Word.Views.Wizards.Popups;
using Enums = OAMPS.Office.Word.Helpers.Enums;
using OAMPS.Office.Word.Properties;

namespace OAMPS.Office.Word.Views.Wizards
{
    public partial class SummaryOfDiscussionWizard : BaseWizardForm
    {
        private readonly SummaryOfDiscussionWizardPresenter _wizardPresenter;

        public SummaryOfDiscussionWizard(OfficeDocument document)
        {
            InitializeComponent();

            tbcWizardScreens.SelectedIndexChanged += tbcWizardScreens_SelectedIndexChanged;

            //send marketing template to the presenter

            _wizardPresenter = new SummaryOfDiscussionWizardPresenter(document, this);
            BaseWizardPresenter = _wizardPresenter;

            //ShouldUpdateTemplate(Settings.Default.TemplateLibraryName, "File Note.docx");
        }

        private void SummaryOfDiscussionWizard_Load(object sender, EventArgs e)
        {
            //pbOampsLogoFull.Left = (ClientSize.Width - pbOampsLogoFull.Width) / 2;
            txtClientName.Focus();
            txtClientName.Select();

            var uiScheduler = TaskScheduler.FromCurrentSynchronizationContext();
            LoadCompanyLogoImagesTab(uiScheduler, tbcWizardScreens, lblLogoTitle.Text);
            LoadBrandingImagesTab(null, tbcWizardScreens, lblCoverPageTitle.Text, lblSpeciality.Text);

            if (Reload) // this happens if they click the button on the ribbon.
            {
                var template = new SummaryOfDiscussions();
                var values = _wizardPresenter.LoadData(template);
                var v = ((ISummaryOfDiscussions) values);
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


            var outValue = false;
            if (bool.TryParse(v.IsDiscussedWithCaller, out outValue))
            {
                rdoCaller.Checked = outValue;
            }

            outValue = false;
            if (bool.TryParse(v.IsDiscussedInPerson, out outValue))
            {
                rdoPerson.Checked = outValue;
            }

            outValue = false;
            if (bool.TryParse(v.IsDiscussedWithCustomer, out outValue))
            {
                rdoCustomer.Checked = outValue;
            }

            outValue = false;
            if (bool.TryParse(v.IsOther, out outValue))
            {
                rdoOther.Checked = outValue;
            }

            outValue = false;
            if (bool.TryParse(v.IsDiscussedWithUnderWriter, out outValue))
            {
                rdoUnderWriter.Checked = outValue;
            }

            outValue = false;
            if (bool.TryParse(v.IsDiscussedByPhone, out outValue))
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

        private void tbcWizardScreens_SelectedIndexChanged(object sender, EventArgs e)
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
                DocumentTitle = BaseWizardPresenter.ReadDocumentProperty("Title"), //Constants.TemplateNames.FileNote,
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
            _wizardPresenter.PopulateData(template);
            _wizardPresenter.CreatePropertiesForRadioButtons(rdoCustomer.Checked, rdoCaller.Checked,
                rdoUnderWriter.Checked,
                rdoPhone.Checked, rdoPerson.Checked, rdoOther.Checked);

            //change the graphics selected
            //if (Streams == null) return;
            _wizardPresenter.PopulateGraphics(template, string.Empty, string.Empty);

            //tracking
            LogUsage(template,
                Reload
                    ? Enums.UsageTrackingType.UpdateData
                    : Enums.UsageTrackingType.NewDocument);
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            if (string.Equals(btnNext.Text, "&Finish", StringComparison.CurrentCultureIgnoreCase))
            {
                try
                {
                    Cursor = Cursors.WaitCursor;
                    BaseWizardPresenter.SwitchScreenUpdating(false);
                    //call presenter to populate
                    PopulateDocument();
                }
                catch (Exception ex)
                {
                    OnError(ex);
                }
                finally
                {
                    Cursor = Cursors.Default;
                    BaseWizardPresenter.SwitchScreenUpdating(true);
                    //thie information panel loads when a document is in sharePoint that has metadata
                    //clients don't wish to see this so force the close of the panel once the wizard completes.
                    _wizardPresenter.CloseInformationPanel();
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
            var peoplePicker = new PeoplePicker(txtExecutiveName.Text, this);
            if (peoplePicker.SelectedUser == null)
            {
                TopMost = false;

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