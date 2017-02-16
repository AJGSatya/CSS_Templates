using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Runtime.Caching;
using System.Text;
using System.Windows.Forms;
using OAMPS.Office.BusinessLogic.Interfaces.SharePoint;
using OAMPS.Office.BusinessLogic.Interfaces.Wizards;
using OAMPS.Office.BusinessLogic.Models.Template;
using OAMPS.Office.BusinessLogic.Helpers;
using OAMPS.Office.BusinessLogic.Presenters;
using OAMPS.Office.BusinessLogic.Presenters.SharePoint;
using OAMPS.Office.Word.Helpers.Controls;
using OAMPS.Office.Word.Models.SharePoint;
using OAMPS.Office.Word.Models.Word;
using OAMPS.Office.Word.Properties;
using OAMPS.Office.BusinessLogic.Models.Wizards;
using System.Threading.Tasks;
using Aga.Controls.Tree;
using Aga.Controls.Tree.NodeControls;
using OAMPS.Office.Word.Helpers;

namespace OAMPS.Office.Word.Views.Wizards
{
    public partial class GenericLetterWizard : BaseForm
    {
        private GenericLetterPresenter _presenter;

        private IPolicyClass _selectedPolicy;

        private Dictionary<string, DocumentFragment> _availableAttachments;
        //private bool isPrePrintedStationary;

        public GenericLetterWizard(OfficeDocument document)
        {
            InitializeComponent();
            tbcWizardScreens.SelectedIndexChanged += new EventHandler(tbcWizardScreens_SelectedIndexChanged);

            _presenter = new GenericLetterPresenter(document, this);
            base.BasePresenter = _presenter;            

            txtClientName.Validating += new CancelEventHandler(ClientNameValidating);
        }

        private void ClientNameValidating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (string.IsNullOrEmpty(txtClientName.Text))
            {
                errorProvider1.SetError(txtClientName, "You must enter a client name");
                e.Cancel = true;
            }
            else
            {
                errorProvider1.SetError(txtClientName, string.Empty);
            }
        }

        private void GetFragements()
        {
            _availableAttachments = new Dictionary<string, DocumentFragment>();
            List<ISharePointListItem> fragments = null;
            if (Cache.Contains(Constants.CacheNames.RenLtrFragments))
            {
                fragments = (List<ISharePointListItem>)Cache.Get(Constants.CacheNames.RenLtrFragments);
            }
            else
            {

                var list = new SharePointList(Settings.Default.SharePointContextUrl, Settings.Default.GeneralFragmentsListName, Constants.SharePointQueries.RenewalLetterFragmentsByKey);
                var presenter = new SharePointListPresenter(list, this);
                fragments = presenter.GetItems();
            }




            foreach (ISharePointListItem i in fragments)
            {
                var key = i.GetFieldValue("Key");
                string txtTitle = i.Title + " " + i.GetFieldValue("OAMPS_x0020_Version");
                switch (key)
                {
                    case Constants.FragmentKeys.FinancialServicesGuideLetter:
                        {
                            chkFSG.Text = txtTitle;
                            _availableAttachments.Add(key, new DocumentFragment
                            {
                                Title = txtTitle,
                                Key = key,
                                Url = Settings.Default.FragmentFSG
                            });

                            break;
                        }


                    case Constants.FragmentKeys.GeneralAdviceWarning:
                        {
                            chkWarning.Text = txtTitle;
                            _availableAttachments.Add(key, new DocumentFragment
                            {
                                Title = txtTitle,
                                Key = key,
                                Url = Settings.Default.FragmentWarning
                            });
                            break;
                        }


                    case Constants.FragmentKeys.PrivacyStatement:
                        {
                            chkPrivacy.Text = txtTitle;
                            _availableAttachments.Add(key, new DocumentFragment
                            {
                                Title = txtTitle,
                                Key = key,
                                Url = Settings.Default.FragmentPrivacy
                            });

                            break;
                        }


                    case Constants.FragmentKeys.StatutoryNotices:
                        {
                            chkSatutory.Text = txtTitle;
                            _availableAttachments.Add(key, new DocumentFragment
                            {
                                Title = txtTitle,
                                Key = key,
                                Url = Settings.Default.FragmentStatutory
                            });

                            break;
                        }


                    case Constants.FragmentKeys.UninsuredRisksReviewList:
                        {
                            chkRisks.Text = txtTitle;
                            _availableAttachments.Add(key, new DocumentFragment
                            {
                                Title = txtTitle,
                                Key = key,
                                Url = Settings.Default.FragmentUninsuredRisks
                            });
                            break;
                        }
                }
            }
        }


        private void RenewalLetter_Load(object sender, EventArgs e)
        {
            txtAddressee.Focus();
            txtAddressee.Select();

            GetFragements();

            var uiScheduler = TaskScheduler.FromCurrentSynchronizationContext();

            if (Reload)
            {
                ReloadFields();
            }
            else
            {
                Task.Factory.StartNew(() => base.LoadTreeViewClasses(uiScheduler));
                _presenter.CloseInformationPanel(true);
            }

            base.LoadGenericImageTabs(uiScheduler, tbcWizardScreens, lblCoverPageTitle.Text, lblLogoTitle.Text);
        }

        private void ReloadFields()
        {

            var auto = false;

            WizardBeingUpdated = true;
            var template = new GenericLetter();
            string policyReference = null;

            if (Cache.Contains(Constants.CacheNames.RegenerateTemplate))
            {
                template = (GenericLetter)Cache.Get(Constants.CacheNames.RegenerateTemplate);
                Cache.Remove(Constants.CacheNames.RegenerateTemplate);

                if (template.PolicyType != null)
                    policyReference = template.PolicyType.Title;

                chkWarning.Checked = template.IsAdviceWarningSelected;
                chkFSG.Checked = template.IsFSGSelected;
                chkPrivacy.Checked = template.IsPrivacySelected;
                chkRisks.Checked = template.IsRisksSelected;
                chkSatutory.Checked = template.IsSatutorySelected;
                
                if (template.IsPrePrintSelected)
                {
                    chkPrePrint.Checked = template.IsPrePrintSelected;
                }
                Reload = false; //behaviour like a new from this point
                auto = true;
            }
            else
            {
                template = (GenericLetter)_presenter.LoadData(template);
                UpdatePrePrintedCheckBoxChecked();

                var fgKeys = _presenter.ReadDocumentProperty(Constants.WordDocumentProperties.RlSubFragments);
                if (fgKeys != null)
                {
                    var keys = fgKeys.Split(';');
                    foreach (string key in keys)
                    {

                        if (key.Equals(Constants.FragmentKeys.FinancialServicesGuide, StringComparison.InvariantCultureIgnoreCase))
                        {
                            chkFSG.Checked = true;
                        }
                        else if (key.Equals(Constants.FragmentKeys.GeneralAdviceWarning, StringComparison.InvariantCultureIgnoreCase))
                        {
                            chkWarning.Checked = true;
                        }
                        else if (key.Equals(Constants.FragmentKeys.PrivacyStatement, StringComparison.InvariantCultureIgnoreCase))
                        {
                            chkPrivacy.Checked = true;
                        }
                        else if (key.Equals(Constants.FragmentKeys.StatutoryNotices, StringComparison.InvariantCultureIgnoreCase))
                        {
                            chkSatutory.Checked = true;
                        }
                        else if (key.Equals(Constants.FragmentKeys.UninsuredRisksReviewList, StringComparison.InvariantCultureIgnoreCase))
                        {
                            chkRisks.Checked = true;
                        }

                    }
                }
            }

            txtAddressee.Text = template.Addressee;
            txtSalutation.Text = template.Salutation;
            txtClientAddress.Text = template.ClientAddress;
            txtClientAddress2.Text = template.ClientAddressLine2;
            txtClientAddress3.Text = template.ClientAddressLine3;
            txtClientName.Text = template.ClientName;

            txtSubject.Text = template.Subject;
            txtPDSVersion.Text = template.PDSDescVersion;

            txtExecutiveName.Text = template.ExecutiveName;
            txtExecutiveEmail.Text = template.ExecutiveEmail;
            txtExecutiveMobile.Text = template.ExecutiveMobile;
            txtExecutivePhone.Text = template.ExecutivePhone;
            txtExecutiveTitle.Text = template.ExecutiveTitle;
            txtExecutiveDepartment.Text = template.ExecutiveDepartment;

            txtBranchAddress1.Text = template.OAMPSBranchAddress;
            txtBranchAddress2.Text = template.OAMPSBranchAddressLine2;

            txtPostal1.Text = template.OAMPSPostalAddress;
            txtPostal2.Text = template.OAMPSPostalAddressLine2;

            txtBranchPhone.Text = template.OAMPSBranchPhone;

            lblLogoTitle.Text = template.LogoTitle;
            lblCoverPageTitle.Text = template.CoverPageTitle;

            txtReference.Text = template.Reference;
            txtFax.Text = template.Fax;

            DateTime outDate;
           
            dateReport.Value = DateTime.TryParse(template.DatePrepared, out outDate)
                                            ? outDate
                                            : DateTime.Today;
            WizardBeingUpdated = false;

            if (auto) base.StartTimer();
        }

        private void UpdatePrePrintedCheckBoxChecked()
        {
            var value = _presenter.ReadDocumentProperty(Constants.WordDocumentProperties.RlRdoPreprint);
            bool isSelected;
            if (Boolean.TryParse(value, out isSelected))
            {
                chkPrePrint.Checked = isSelected;
            }
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

        private void Next_Click(object sender, EventArgs e)
        {
            try
            {
                this.Cursor = Cursors.WaitCursor;
            //    BasePresenter.SwitchScreenUpdating(false);

                if (String.Equals(btnNext.Text, "&Finish", StringComparison.CurrentCultureIgnoreCase))
                {
                    if (Validation.HasValidationErrors(this.Controls))
                    {
                        MessageBox.Show(@"Please ensure all required fields are populated", @"Required fields are missing",
                                        MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }

                  //  _presenter.CloseInformationPanel(); //has to be here to close as the presenter opens new word documents taht become active.

                    var letter = new GenericLetter
                        {
                            Addressee = txtAddressee.Text,
                            Salutation = txtSalutation.Text,
                            ClientAddress = txtClientAddress.Text,
                            ClientAddressLine2 = txtClientAddress2.Text,
                            ClientAddressLine3 = txtClientAddress3.Text,

                            ClientName = txtClientName.Text,
                            DatePrepared = dateReport.Value.ToString(),
                            ExecutiveName = txtExecutiveName.Text,
                            ExecutiveEmail = txtExecutiveEmail.Text,
                            ExecutiveMobile = txtExecutiveMobile.Text,
                            ExecutivePhone = txtExecutivePhone.Text,
                            ExecutiveTitle = txtExecutiveTitle.Text,

                            ExecutiveDepartment = txtExecutiveDepartment.Text,
                            OAMPSBranchPhone = txtBranchPhone.Text,
                            Reference = txtReference.Text,
                            Subject = txtSubject.Text,
                            Fax = txtFax.Text,
                            PolicyType = _selectedPolicy,

                            IsFSGSelected = chkFSG.Checked,
                            IsPrivacySelected = chkPrivacy.Checked,
                            IsRisksSelected = chkRisks.Checked,
                            IsSatutorySelected = chkSatutory.Checked,
                            IsAdviceWarningSelected = chkWarning.Checked,
                            OAMPSBranchAddress = txtBranchAddress1.Text,
                            OAMPSBranchAddressLine2 = txtBranchAddress2.Text,
                            PDSDescVersion = txtPDSVersion.Text,
                            OAMPSPostalAddress = txtPostal1.Text,
                            OAMPSPostalAddressLine2 = txtPostal2.Text,
                            IsPrePrintSelected = chkPrePrint.Checked
                        };

                    _presenter.CreateOrUpdateDocumentProperty(Constants.WordDocumentProperties.RlRdoPreprint, chkPrePrint.Checked.ToString());
                   
                    TabPage logoTab = tbcWizardScreens.TabPages[BusinessLogic.Helpers.Constants.ControlNames.TabPageLogosName];
                    var baseTemplate = (BaseTemplate)letter;
                    PopulateLogosToTemplate(logoTab, ref baseTemplate);

                    if (_generateNewTemplate)
                    {
                        Cache.Add(Constants.CacheNames.RegenerateTemplate, letter, new CacheItemPolicy());
                        _presenter.GenerateNewTemplate(Constants.CacheNames.RegenerateTemplate, Settings.Default.TemplateGenericLetter);
                    }
                    else
                    {
                        PopulateAddress(letter);
                        //call presenter to populate
                        PopulateDocument(letter, lblCoverPageTitle.Text, lblLogoTitle.Text);

                        if (chkPrePrint.Checked)
                        {
                            _presenter.DeleteDocumentHeaderAndFooter();
                        }

                        if (!Reload)
                            AddAttachments(letter);

                        //tracking
                        LogUsage(letter,
                                 _loadType == Helpers.Enums.FormLoadType.RegenerateTemplate
                                     ? Helpers.Enums.UsageTrackingType.RegenerateDocument
                                     : Helpers.Enums.UsageTrackingType.NewDocument);
                    }
                //    _presenter.ActivateDocument();
                    Close();
                }
                else
                {
                    SwitchTab(tbcWizardScreens.SelectedIndex + 1);
                }
            }
            catch (Exception ex)
            {
                OnError(ex);
            }
            finally
            {
                this.Cursor = Cursors.Default;
              //  BasePresenter.SwitchScreenUpdating(true);
            }
        }

        private void AddAttachments(GenericLetter template)
        {
            var attachmentFragments = new List<DocumentFragment>();

            if (template.IsSatutorySelected)
            {
                attachmentFragments.Add(_availableAttachments[Constants.FragmentKeys.StatutoryNotices]);
            }

            if (template.IsPrivacySelected)
            {
                attachmentFragments.Add(_availableAttachments[Constants.FragmentKeys.PrivacyStatement]);
            }

            if (template.IsFSGSelected)
            {
                attachmentFragments.Add(_availableAttachments[Constants.FragmentKeys.FinancialServicesGuideLetter]);
            }

            if (template.IsAdviceWarningSelected)
            {
                attachmentFragments.Add(_availableAttachments[Constants.FragmentKeys.GeneralAdviceWarning]);
            }

            if (template.IsRisksSelected)
            {
                attachmentFragments.Add(_availableAttachments[Constants.FragmentKeys.UninsuredRisksReviewList]);
            }

            if (attachmentFragments.Count > 0)
                _presenter.InsertSubFragments(Constants.WordBookmarks.RenewalLetterSub, Constants.WordBookmarks.RenewalLetterMain, attachmentFragments, Settings.Default.BlankFSGFragement);
        }

        private void PopulateAddress(GenericLetter letter)
        {
            if (!String.IsNullOrEmpty(txtBranchAddress1.Text))
            {
                letter.OAMPSBranchAddress = txtBranchAddress1.Text;
                letter.OAMPSBranchAddressLine2 = txtBranchAddress2.Text;
            }

            if (!String.IsNullOrEmpty(txtPostal1.Text))
            {
                letter.OAMPSPostalAddress = txtPostal1.Text;
                letter.OAMPSPostalAddressLine2 = txtPostal2.Text;
            }
        }

        private void btnAccountExecutiveLookup_Click(object sender, EventArgs e)
        {
            var peoplePicker = new Popups.PeoplePicker(txtExecutiveName.Text, this);

            if (peoplePicker.SelectedUser == null)
            {
                this.TopMost = false;

                peoplePicker.ShowDialog();
                if (peoplePicker.SelectedUser == null) return;
            }

            txtExecutiveName.Text = peoplePicker.SelectedUser.DisplayName;
            txtExecutiveTitle.Text = peoplePicker.SelectedUser.Title;
            txtExecutiveEmail.Text = peoplePicker.SelectedUser.EmailAddress;
            txtExecutivePhone.Text = peoplePicker.SelectedUser.VoiceTelephoneNumber;
            txtBranchAddress1.Text = peoplePicker.SelectedUser.BranchAddressLine1;
            txtExecutiveMobile.Text = peoplePicker.SelectedUser.Mobile;
            txtBranchAddress2.Text = peoplePicker.SelectedUser.BranchAddressLine2;

            txtExecutiveDepartment.Text = peoplePicker.SelectedUser.Branch;

            txtBranchPhone.Text = peoplePicker.SelectedUser.BranchPhone;
            txtFax.Text = peoplePicker.SelectedUser.Fax;

            txtPostal1.Text = peoplePicker.SelectedUser.Suburb.GetPostalAddressFragment(true);
            txtPostal2.Text = peoplePicker.SelectedUser.Suburb.GetPostalAddressFragment(false);
        }



        private void btnBack_Click(object sender, EventArgs e)
        {
            SwitchTab(tbcWizardScreens.SelectedIndex - 1);
        }

       

        private void chkWarning_CheckedChanged(object sender, EventArgs e)
        {
            if (Reload && !WizardBeingUpdated)
            {
                var previous = !chkWarning.Checked;

                if (ContinueWithSignificantChange(sender, previous)) _generateNewTemplate = true;
            }
        }

        private void chkFSG_CheckedChanged(object sender, EventArgs e)
        {
            if (Reload && !WizardBeingUpdated)
            {
                var previous = !chkFSG.Checked;

                if (ContinueWithSignificantChange(sender, previous)) _generateNewTemplate = true;
            }
        }

        private void chkPrivacy_CheckedChanged(object sender, EventArgs e)
        {
            if (Reload && !WizardBeingUpdated)
            {
                var previous = !chkPrivacy.Checked;

                if (ContinueWithSignificantChange(sender, previous)) _generateNewTemplate = true;
            }
        }

        private void chkRisks_CheckedChanged(object sender, EventArgs e)
        {
            if (Reload && !WizardBeingUpdated)
            {
                var previous = !chkRisks.Checked;

                if (ContinueWithSignificantChange(sender, previous)) _generateNewTemplate = true;
            }
        }

        private void chkSatutory_CheckedChanged(object sender, EventArgs e)
        {
            if (Reload && !WizardBeingUpdated)
            {
                var previous = !chkSatutory.Checked;

                if (ContinueWithSignificantChange(sender, previous)) _generateNewTemplate = true;
            }
        }
    
       
        private void chkPrePrint_CheckedChanged(object sender, EventArgs e)
        {
            if (Reload && !WizardBeingUpdated)
            {
                var previous = !chkPrePrint.Checked;

                if (ContinueWithSignificantChange(sender, previous)) _generateNewTemplate = true;
            }
        }

        
        

    }
}
