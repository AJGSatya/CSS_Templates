using System;
using System.Runtime.Caching;
using System.Windows.Forms;
using OAMPS.Office.Word.Helpers.Controls;
using OAMPS.Office.Word.Models.Word;
using OAMPS.Office.BusinessLogic.Presenters;
using OAMPS.Office.BusinessLogic.Models.Template;
using OAMPS.Office.BusinessLogic.Helpers;
using System.Threading.Tasks;
using OAMPS.Office.Word.Properties;

namespace OAMPS.Office.Word.Views.Wizards
{
    public partial class PreRenewalAgendaWizard :  BaseForm
    {
        private static string _meetingType = "Meeting Agenda";
        private AgendaPresenter _presenter;

        public PreRenewalAgendaWizard(OfficeDocument document)
        {
            InitializeComponent();
            tbcWizardScreens.SelectedIndexChanged += new EventHandler(tbcWizardScreens_SelectedIndexChanged);

            //send marketing template to the presenter
            
            _presenter = new AgendaPresenter(document, this);
            BasePresenter = _presenter;

            rdoMeetingAgenda.Validating += new System.ComponentModel.CancelEventHandler(rdoMeetingAgenda_Validating);
            rdoMeetingMinutes.Validating += new System.ComponentModel.CancelEventHandler(rdoMeetingMinutes_Validating);
        }

        void rdoMeetingMinutes_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
          
        }

        void rdoMeetingAgenda_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            
        }

        private void SummaryOfDiscussionWizard_Load(object sender, EventArgs e)
        {
          //  pbOampsLogoFull.Left = (ClientSize.Width - pbOampsLogoFull.Width) / 2;
            WizardBeingUpdated = true;
            txtClientName.Focus();
            txtClientName.Select();

            bool isRegen = Cache.Contains(Constants.CacheNames.RegenerateTemplate);
            var template = new Agenda();

            if (isRegen)
            {
                //MessageBox.Show(BusinessLogic.Helpers.Constants.Miscellaneous.RegenerateOnLoadMsg, String.Empty,
                //                MessageBoxButtons.OK, MessageBoxIcon.Information);

                template = GetCachedTempalteObject<Agenda>();
                ReloadWizardFields(template);
                base.LoadGenericImageTabs(null, tbcWizardScreens, lblCoverPageTitle.Text, lblLogoTitle.Text);
                base.StartTimer();
            }

            else if (Reload)
            {
                template = (Agenda)_presenter.LoadData(template);
                ReloadWizardFields(template);

              

            }

            if (!isRegen)
            {
                var uiScheduler = TaskScheduler.FromCurrentSynchronizationContext();
                base.LoadGenericImageTabs(uiScheduler, tbcWizardScreens, lblCoverPageTitle.Text, lblLogoTitle.Text);
            }
         
            WizardBeingUpdated = false;
        }

        private void ReloadWizardFields(Agenda template)
        {
            txtClientName.Text = template.ClientName;
            txtClientSubject.Text = template.Subject;
            txtLocation.Text = template.Location;

            dateAgenda.Text = template.AgendaDate;
            timeAgendaFrom.Text = template.AgendaTimeFrom;
            timeAgendaTo.Text = template.AgendaTimeTo;

            lblLogoTitle.Text = template.LogoTitle;
            lblCoverPageTitle.Text = template.CoverPageTitle;

            if (template.DocumentTitle == BusinessLogic.Helpers.Constants.TemplateNames.PreRenewalAgenda)
            {
                rdoMeetingAgenda.Checked = true;
            }
            else
            {
                rdoMeetingMinutes.Checked = true;
            }
        }

        private void ValidateMeetingOrAgenda(EventArgs e)
        {
            //if (rdoMeetingAgenda.Checked == false && rdoMeetingMinutes.Checked == false )
            //{
            //    errorProvider1.SetError(rdoMeetingAgenda, "You must choose what kind of document you want.");
            //    e.c = true;
            //}
            //else
            //{
            //    errorProvider.SetError(grpFees, string.Empty);
            //}
        }


        private void tbcWizardScreens_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnNext.Text = tbcWizardScreens.SelectedIndex == tbcWizardScreens.TabCount - 1 ? "&Finish" : "&Next";
            btnBack.Enabled = tbcWizardScreens.SelectedIndex != 0;
          //  btnNext.Enabled = btnNext.Text != @"&Finish" || LoadComplete;
            btnNext.Enabled = true;
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            if (String.Equals(btnNext.Text, "&Finish", StringComparison.CurrentCultureIgnoreCase))
            {
                if (Validation.HasValidationErrors(this.Controls))
                {
                    MessageBox.Show(@"Please ensure all required fields are populated", @"Required fields are missing",
                                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                //call presenter to populate
                //buid the marketing template
                var template = new Agenda
                {
                    DocumentTitle = _meetingType,
                    DocumentSubTitle = string.Empty,
                    AgendaDate = dateAgenda.Text,
                    AgendaTimeFrom = timeAgendaFrom.Text,
                    AgendaTimeTo = timeAgendaTo.Text,
                    Subject = txtClientSubject.Text,
                    ClientName = txtClientName.Text,
                    Location = txtLocation.Text
                    
                };

                var baseTemplate = (BaseTemplate) template;

                var logoTab = tbcWizardScreens.TabPages[BusinessLogic.Helpers.Constants.ControlNames.TabPageLogosName];

                PopulateLogosToTemplate(logoTab, ref baseTemplate);
                //foreach (Control c in logoTab.Controls)
                //{
                //    if (c is ValueRadioButton)
                //    {
                //        var v = ((ValueRadioButton)c);
                //        if (v.Checked)
                //        {
                //            template.LogoImageUrl = v.GetValue(Constants.RadioButtonValues.ImageUrl);
                //            template.LogoTitle = v.Text;
                                
                //        }
                //    }
                //}

                var covberTab = tbcWizardScreens.TabPages[BusinessLogic.Helpers.Constants.ControlNames.TabPageCoverPagesName];

                PopulateCoversToTemplate(covberTab, ref baseTemplate);
                //foreach (Control c in covberTab.Controls)
                //{
                //    if (c is ValueRadioButton)
                //    {
                //        var v = ((ValueRadioButton)c);
                //        if (v.Checked)
                //        {
                //            template.CoverPageImageUrl = v.GetValue(Constants.RadioButtonValues.ImageUrl);
                //            template.CoverPageTitle = v.Text;
                //        }
                //    }
                //}



                if (_generateNewTemplate)
                {
                    Cache.Add(Constants.CacheNames.RegenerateTemplate, template, new CacheItemPolicy());

                    _presenter.GenerateNewTemplate(Constants.CacheNames.RegenerateTemplate,
                                                   Settings.Default.TemplateAgenda);
                }
                else
                {
                    if (rdoMeetingMinutes.Checked && !Reload)
                    {
                        _presenter.InsertMinutesFragement(Settings.Default.FragmentAGMinutes);
                    }


                    base.PopulateDocument(template);

                    //tracking
                    LogUsage(template,
                             Reload == true
                                 ? Helpers.Enums.UsageTrackingType.UpdateData
                                 : Helpers.Enums.UsageTrackingType.NewDocument);
                }

                //thie information panel loads when a document is in sharePoint that has metadata
                //clients don't wish to see this so force the close of the panel once the wizard completes.
                _presenter.CloseInformationPanel();

                Close();
            }
            else
            {
                base.SwitchTab(tbcWizardScreens, tbcWizardScreens.SelectedIndex + 1);
            }
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            base.SwitchTab(tbcWizardScreens, tbcWizardScreens.SelectedIndex - 1);
        }

     
        private void rdoMeetingAgenda_CheckedChanged(object sender, EventArgs e)
        {
            bool isChecked = rdoMeetingAgenda.Checked;

            if (Reload && !WizardBeingUpdated && isChecked)
            {
                

                if (ContinueWithSignificantChange(sender, false))
                {
                    _generateNewTemplate = true;
                }
                else
                {
                    return;
                }
            }

            _meetingType = "Meeting Agenda";

        }

        private void rdoMeetingMinutes_CheckedChanged(object sender, EventArgs e)
        {

            bool isChecked = rdoMeetingMinutes.Checked;

            if (Reload && !WizardBeingUpdated && isChecked)
            {
                

                if (ContinueWithSignificantChange(sender, false))
                {
                    _generateNewTemplate = true;
                }
                else
                {
                    return;
                }
            }
            _meetingType = "Meeting Minutes";
        }

     
    }
}