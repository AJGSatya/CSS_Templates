using System;
using System.ComponentModel;
using System.Runtime.Caching;
using System.Threading.Tasks;
using System.Windows.Forms;
using OAMPS.Office.BusinessLogic.Helpers;
using OAMPS.Office.BusinessLogic.Models.Template;
using OAMPS.Office.BusinessLogic.Presenters;
using OAMPS.Office.Word.Helpers.Controls;
using OAMPS.Office.Word.Models.Word;
using OAMPS.Office.Word.Properties;
using Enums = OAMPS.Office.Word.Helpers.Enums;

namespace OAMPS.Office.Word.Views.Wizards
{
    public partial class PreRenewalAgendaWizard : BaseWizardForm
    {
        private static string _meetingType = "Meeting Agenda";
        private readonly AgendaWizardPresenter _wizardPresenter;

        public PreRenewalAgendaWizard(OfficeDocument document)
        {
            InitializeComponent();
            tbcWizardScreens.SelectedIndexChanged += tbcWizardScreens_SelectedIndexChanged;

            //send marketing template to the presenter

            _wizardPresenter = new AgendaWizardPresenter(document, this);
            BaseWizardPresenter = _wizardPresenter;

            rdoMeetingAgenda.Validating += rdoMeetingAgenda_Validating;
            rdoMeetingMinutes.Validating += rdoMeetingMinutes_Validating;
        }

        private void rdoMeetingMinutes_Validating(object sender, CancelEventArgs e)
        {
        }

        private void rdoMeetingAgenda_Validating(object sender, CancelEventArgs e)
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
                template = (Agenda) _wizardPresenter.LoadData(template);
                ReloadWizardFields(template);
            }

            if (!isRegen)
            {
                TaskScheduler uiScheduler = TaskScheduler.FromCurrentSynchronizationContext();
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

            if (template.DocumentTitle == Constants.TemplateNames.PreRenewalAgenda)
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
                if (Validation.HasValidationErrors(Controls))
                {
                    MessageBox.Show(@"Please ensure all required fields are populated", @"Required fields are missing",
                                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                try
                {
                    Cursor = Cursors.WaitCursor;
                    BaseWizardPresenter.SwitchScreenUpdating(false);
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
                    TabPage logoTab = tbcWizardScreens.TabPages[Constants.ControlNames.TabPageLogosName];
                    PopulateLogosToTemplate(logoTab, ref baseTemplate);
                    TabPage covberTab = tbcWizardScreens.TabPages[Constants.ControlNames.TabPageCoverPagesName];
                    PopulateCoversToTemplate(covberTab, ref baseTemplate);
                    if (GenerateNewTemplate)
                    {
                        Cache.Add(Constants.CacheNames.RegenerateTemplate, template, new CacheItemPolicy());
                        _wizardPresenter.GenerateNewTemplate(Constants.CacheNames.RegenerateTemplate,
                                                       Settings.Default.TemplateAgenda);
                    }
                    else
                    {
                        if (rdoMeetingMinutes.Checked && !Reload)
                        {
                            _wizardPresenter.InsertMinutesFragement(Settings.Default.FragmentAGMinutes);
                        }

                        if (rdoMeetingMinutes.Checked && AutoComplete)
                        {
                            _wizardPresenter.InsertMinutesFragement(Settings.Default.FragmentAGMinutes);
                        }
                        base.PopulateDocument(template, lblCoverPageTitle.Text,lblLogoTitle.Text);

                        //tracking
                        LogUsage(template,
                                 Reload
                                     ? Enums.UsageTrackingType.UpdateData
                                     : Enums.UsageTrackingType.NewDocument);
                    }
                    //thie information panel loads when a document is in sharePoint that has metadata
                    //clients don't wish to see this so force the close of the panel once the wizard completes.
                    _wizardPresenter.CloseInformationPanel();
                    Close();
                }
                catch (Exception ex)
                {
                    OnError(ex);
                }
                finally
                {
                    Cursor = Cursors.Default;
                    BaseWizardPresenter.SwitchScreenUpdating(true);
                }
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
                    GenerateNewTemplate = true;
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
                    GenerateNewTemplate = true;
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