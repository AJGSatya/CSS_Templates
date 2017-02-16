﻿using System;
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
    public partial class RenewalLetterWizard : BaseForm
    {
        private RenwalLetterPresenter _presenter;

        private IPolicyClass _selectedPolicy;

        private Dictionary<string, DocumentFragment> _availableAttachments;
        private bool isPrePrintedStationary;
        public RenewalLetterWizard(OfficeDocument document)
        {                        
            InitializeComponent();
            tbcWizardScreens.SelectedIndexChanged += new EventHandler(tbcWizardScreens_SelectedIndexChanged);
            
            _presenter = new RenwalLetterPresenter(document, this);
            
            _checked.CheckStateChanged += new EventHandler<TreePathEventArgs>(_checked_CheckStateChanged);
            
            base.BasePresenter =_presenter;
            rdoPrePrintNo.Checked = true;
            tvaPolicies.Validating += new CancelEventHandler(PolicyValidating);
             

            txtClientName.Validating += new CancelEventHandler(ClientNameValidating);
        }

        private void ClientNameValidating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (string.IsNullOrEmpty(txtClientName.Text))
            {
                errorProvider1.SetError(txtClientName, "You must select the policy class");
                e.Cancel = true;
            }
            else
            {
                errorProvider1.SetError(txtClientName,string.Empty);
            }
        }

        private void PolicyValidating(object sender, CancelEventArgs e)
        {
            if (_selectedPolicy == null)
            {
                errorProvider1.SetError(lblRequiredPolicyClass, "You must select the policy class");
                e.Cancel = true;
            }
            else
            {
                errorProvider1.SetError(lblRequiredPolicyClass,string.Empty);
            }
        }

        private void _checked_CheckStateChanged(object sender, TreePathEventArgs e)
        {


            var node = ((AdvancedTreeNode) e.Path.LastNode);
            var isChecked = node.Checked;

            if (isChecked)
            {
                var item = MinorItems.Find(i => i.Title == node.Text);
                if (item != null)
                {
                    tvaPolicies.AllNodes.ToList().ForEach(
                        (x) =>
                            {
                                var nodeControl = tvaPolicies.GetNodeControls(x);
                                if (nodeControl != null)
                                {
                                    var checkbox = nodeControl.FirstOrDefault(y => (y.Control is NodeCheckBox));
                                    //checkbox found
                                    var dCheckBox = (NodeCheckBox) checkbox.Control;
                                    if(dCheckBox!=null)
                                    dCheckBox.SetValue(x, false);
                                }
                            }
                        );
                    node.Checked = true;
                    _selectedPolicy = item;
                }
            }

        }

        private void RenewalLetter_Load(object sender, EventArgs e)
        {

            txtAddressee.Focus();
            txtAddressee.Select();
            _availableAttachments = new Dictionary<string, DocumentFragment>();

            var list = new SharePointList(Settings.Default.SharePointContextUrl, Settings.Default.GeneralFragmentsListName,Constants.SharePointQueries.RenewalLetterFragmentsByKey);
            var presenter = new SharePointListPresenter(list, this);
            var fragments = presenter.GetItems();         

            foreach (ISharePointListItem i in fragments)
            {
                var key = i.GetFieldValue("Key");
                string txtTitle = i.Title + " " + i.GetFieldValue("OAMPS_x0020_Version");
                switch (key)
                {
                    case Constants.FragmentKeys.FinancialServicesGuide:
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
                                    Title   = txtTitle,
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

            var uiScheduler = TaskScheduler.FromCurrentSynchronizationContext();
            
            if (Reload)
            {
                base.LoadTreeViewClasses(null);                

                ReloadFields();
            }
            else
            {
                datePayment.Value = DateTime.Today.AddDays(14);

                Task.Factory.StartNew(() => base.LoadTreeViewClasses(uiScheduler));
            }
            base.LoadGenericImageTabs(uiScheduler, tbcWizardScreens, lblCoverPageTitle.Text, lblLogoTitle.Text);

            
        }

        private void ReloadFields()
        {
            WizardBeingUpdated = true;


            var template = new RenewalLetter();
            try
            {
                string policyReference = null;
                
                if (Cache.Contains(Constants.CacheNames.RegenerateTemplate))
                {

                    MessageBox.Show(BusinessLogic.Helpers.Constants.Miscellaneous.RegenerateOnLoadMsg, String.Empty,
                                MessageBoxButtons.OK, MessageBoxIcon.Information);

                    template = (RenewalLetter) Cache.Get(Constants.CacheNames.RegenerateTemplate);
                    Cache.Remove(Constants.CacheNames.RegenerateTemplate);

                    if(template.PolicyType!=null)
                    policyReference = template.PolicyType.Title;

                    chkWarning.Checked = template.IsWarningSelected;
                    chkFSG.Checked = template.IsFSGSelected;
                    chkPrivacy.Checked = template.IsPrivacySelected;
                    chkRisks.Checked = template.IsRisksSelected;
                    chkSatutory.Checked = template.IsSatutorySelected;

                    chkContacted.Checked = template.IsContactSelected;
                    chkNewInsurer.Checked = template.IsNewClientSelected;
                    chkFunding.Checked = template.IsFundingSelected;
                    Reload = false; //behaviour like a new from this point
                }
                else
                {
                    template = (RenewalLetter)_presenter.LoadData(template);
                    

                    policyReference = _presenter.ReadPolicyReference();

                    UpdateMainFragmentCheckBox(Constants.WordDocumentProperties.RlChkContacted, chkContacted);
                    UpdateMainFragmentCheckBox(Constants.WordDocumentProperties.RlChkFunding, chkFunding);
                    UpdateMainFragmentCheckBox(Constants.WordDocumentProperties.RlChkNewClient, chkNewInsurer);

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
                txtClientName.Text = template.ClientName;

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
                
                datePayment.Text = template.PaymentDate;
                datePolicy.Text = template.DatePolicyExpiry;
                dateReport.Text = template.DatePrepared;                

                foreach (var no in tvaPolicies.AllNodes)
                {
                    foreach (var cno in no.Children)
                    {
                        if (String.Equals(cno.Tag.ToString(), policyReference, StringComparison.OrdinalIgnoreCase))
                        {
                            no.Expand();
                            var path = tvaPolicies.GetPath(cno);

                            var node = ((AdvancedTreeNode)path.LastNode);
                            node.CheckState = CheckState.Checked;
                            node.Checked = true;
                            _selectedPolicy = MinorItems.FirstOrDefault(i => i.Title == node.Text);
                        }
                    }
                }

                
                
            }
            catch (Exception)
            {


            }
            

           

            WizardBeingUpdated = false;
        }

        

        private void UpdateMainFragmentCheckBox(string proName,CheckBox chk)
        {
            var value = _presenter.ReadDocumentProperty(proName);
            bool isSelected;
            if (Boolean.TryParse(value, out isSelected))
            {
                chk.Checked = isSelected;
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
     

            if (String.Equals(btnNext.Text, "&Finish", StringComparison.CurrentCultureIgnoreCase))
            {
                if (Validation.HasValidationErrors(this.Controls))
                {
                    MessageBox.Show(@"Please ensure all required fields are populated", @"Required fields are missing",
                                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                var letter = new RenewalLetter
                {
                    Addressee = txtAddressee.Text,
                    Salutation = txtSalutation.Text,
                    ClientAddress = txtClientAddress.Text,
                    ClientAddressLine2 = txtClientAddress2.Text,
                    
                    ClientName = txtClientName.Text,
                    DatePolicyExpiry = datePolicy.Text,
                    DatePrepared = dateReport.Text,
                    ExecutiveName = txtExecutiveName.Text,
                    ExecutiveEmail = txtExecutiveEmail.Text,
                    ExecutiveMobile = txtExecutiveMobile.Text,
                    ExecutivePhone = txtExecutivePhone.Text,
                    ExecutiveTitle = txtExecutiveTitle.Text,

                    ExecutiveDepartment = txtExecutiveDepartment.Text,
                    PaymentDate = datePayment.Text,
                    OAMPSBranchPhone = txtBranchPhone.Text,
                    Reference = txtReference.Text,
                    Fax = txtFax.Text,
                    PolicyType = _selectedPolicy,
                    IsContactSelected = chkContacted.Checked,
                    IsFundingSelected = chkFunding.Checked,
                    IsNewClientSelected = chkNewInsurer.Checked,

                    IsFSGSelected = chkFSG.Checked,
                    IsPrivacySelected = chkPrivacy.Checked,
                    IsRisksSelected = chkRisks.Checked,
                    IsSatutorySelected = chkSatutory.Checked,
                    IsWarningSelected = chkWarning.Checked,
                    OAMPSBranchAddress = txtBranchAddress1.Text,
                    OAMPSBranchAddressLine2 = txtBranchAddress2.Text,
                    PDSDescVersion = txtPDSVersion.Text
                    
                };
                TabPage logoTab = tbcWizardScreens.TabPages[BusinessLogic.Helpers.Constants.ControlNames.TabPageLogosName];
                var baseTemplate = (BaseTemplate) letter;
                PopulateLogosToTemplate(logoTab, ref baseTemplate);


                if (_generateNewTemplate)
                {
                    Cache.Add(Constants.CacheNames.RegenerateTemplate, letter, new CacheItemPolicy());

                    _presenter.GenerateNewTemplate(Constants.CacheNames.RegenerateTemplate, Settings.Default.TemplateInsuranceRenewalLetter);
                }
                else
                {
                   
                    
                    _presenter.PopulatePolicy(letter.PolicyType);

                    PopulateAddress(letter);

                    if (!Reload)
                        AddMainFragment(letter);

                    //call presenter to populate
                    PopulateDocument(letter);

                    //thie information panel loads when a document is in sharePoint that has metadata
                    //clients don't wish to see this so force the close of the panel once the wizard completes.
                    //Presenter.CloseInformationPanel();
                    if (!Reload)
                        AddAttachments(letter);

                    if (isPrePrintedStationary)
                    {
                        _presenter.DeleteDocumentHeaderAndFooter();
                    }
                }

                
                Close();
            }
            else
            {
                SwitchTab(tbcWizardScreens.SelectedIndex + 1);
            }
        }

        

        private void PopulateAddress(RenewalLetter letter)
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
        

        private void AddMainFragment(RenewalLetter template)
        {
            if (!chkContacted.Checked && !chkNewInsurer.Checked && !chkGAW.Checked && !chkFunding.Checked)
            {
                _presenter.InsertMainFragment(Constants.WordBookmarks.RenewalLetterMain, Settings.Default.FragmentRLNoContact_ExistingInsr_NoGAW_NoFunding, template);
            }
            else if (chkContacted.Checked && !chkNewInsurer.Checked && !chkGAW.Checked && !chkFunding.Checked)
            {
                _presenter.InsertMainFragment(Constants.WordBookmarks.RenewalLetterMain, Settings.Default.FragmentRLYesContact_ExistingInsr_NoGAW_NoFunding, template);
            }
            else if (!chkContacted.Checked && !chkNewInsurer.Checked && chkGAW.Checked && !chkFunding.Checked)
            {
                _presenter.InsertMainFragment(Constants.WordBookmarks.RenewalLetterMain, Settings.Default.FragmentRLNoContact_ExistingInsr_YesGAW_NoFunding, template);
            }
            else if (chkContacted.Checked && !chkNewInsurer.Checked && chkGAW.Checked && !chkFunding.Checked)
            {
                _presenter.InsertMainFragment(Constants.WordBookmarks.RenewalLetterMain, Settings.Default.FragmentRLYesContact_ExistingInsr_YesGAW_NoFunding, template);
            }
            else if (!chkContacted.Checked && !chkNewInsurer.Checked && chkGAW.Checked && chkFunding.Checked)
            {
                _presenter.InsertMainFragment(Constants.WordBookmarks.RenewalLetterMain, Settings.Default.FragmentRLNoContact_ExistingInsr_YesGAW_YesFunding, template);
            }
            else if (chkContacted.Checked && !chkNewInsurer.Checked && chkGAW.Checked && chkFunding.Checked)
            {
                _presenter.InsertMainFragment(Constants.WordBookmarks.RenewalLetterMain, Settings.Default.FragmentRLYesContact_ExistingInsr_YesGAW_YesFunding, template);
            }
            else if (chkContacted.Checked && !chkNewInsurer.Checked && !chkGAW.Checked && chkFunding.Checked)
            {
                _presenter.InsertMainFragment(Constants.WordBookmarks.RenewalLetterMain, Settings.Default.FragmentRLYesContact_ExistingInsr_NoGAW_YesFunding, template);
            }
            else if (!chkContacted.Checked && chkNewInsurer.Checked && !chkGAW.Checked && !chkFunding.Checked)
            {
                _presenter.InsertMainFragment(Constants.WordBookmarks.RenewalLetterMain,  Settings.Default.FragmentRLNoContact_NewInsr_NoGAW_NoFunding, template);
            }
            else if (chkContacted.Checked && chkNewInsurer.Checked && !chkGAW.Checked && !chkFunding.Checked)
            {
                _presenter.InsertMainFragment(Constants.WordBookmarks.RenewalLetterMain, Settings.Default.FragmentRLYesContact_NewInsr_NoGAW_NoFunding, template);
            }
            else if (!chkContacted.Checked && chkNewInsurer.Checked && chkGAW.Checked && !chkFunding.Checked)
            {
                _presenter.InsertMainFragment(Constants.WordBookmarks.RenewalLetterMain, Settings.Default.FragmentRLNoContact_NewInsr_YesGAW_NoFunding, template);
            }
            else if (chkContacted.Checked && chkNewInsurer.Checked && chkGAW.Checked && !chkFunding.Checked)
            {
                _presenter.InsertMainFragment(Constants.WordBookmarks.RenewalLetterMain, Settings.Default.FragmentRLYesContact_NewInsr_YesGAW_NoFunding, template);
            }
            else if (!chkContacted.Checked && chkNewInsurer.Checked && chkGAW.Checked && chkFunding.Checked)
            {
                _presenter.InsertMainFragment(Constants.WordBookmarks.RenewalLetterMain, Settings.Default.FragmentRLNoContact_NewInsr_YesGAW_YesFunding, template);
            }
            else if (chkContacted.Checked && chkNewInsurer.Checked && chkGAW.Checked && chkFunding.Checked)
            {
                _presenter.InsertMainFragment(Constants.WordBookmarks.RenewalLetterMain, Settings.Default.FragmentRLYesContact_NewInsr_YesGAW_YesFunding, template);
            }
            else if (!chkContacted.Checked && chkNewInsurer.Checked && !chkGAW.Checked && chkFunding.Checked)
            {
                _presenter.InsertMainFragment(Constants.WordBookmarks.RenewalLetterMain, Settings.Default.FragmentRLNoContact_NewInsr_NoGAW_YesFunding, template);
            }
            else if (chkContacted.Checked && chkNewInsurer.Checked && !chkGAW.Checked && chkFunding.Checked)
            {
                _presenter.InsertMainFragment(Constants.WordBookmarks.RenewalLetterMain, Settings.Default.FragmentRLYesContact_NewInsr_NoGAW_YesFunding, template);
            }
              else if (!chkContacted.Checked && !chkNewInsurer.Checked && !chkGAW.Checked && chkFunding.Checked)
            {
                _presenter.InsertMainFragment(Constants.WordBookmarks.RenewalLetterMain, Settings.Default.FragmentRLNoContact_ExistingInsr_NoGAW_YesFunding, template);
            }
            

        }

        private void AddAttachments(RenewalLetter template)
        {          
            var attachmentFragments = new List<DocumentFragment>();

            //if (chkWarning.Checked)
            if(template.IsWarningSelected)
            {
                attachmentFragments.Add(_availableAttachments[Constants.FragmentKeys.GeneralAdviceWarning]);
            }
            //if (chkRisks.Checked)
            if(template.IsRisksSelected)
            {
                attachmentFragments.Add(_availableAttachments[Constants.FragmentKeys.UninsuredRisksReviewList]);                
            }
            //if (chkPrivacy.Checked)
            if(template.IsPrivacySelected)
            {
                attachmentFragments.Add(_availableAttachments[Constants.FragmentKeys.PrivacyStatement]);
            }
            //if (chkSatutory.Checked)
            if(template.IsSatutorySelected)
            {
                attachmentFragments.Add(_availableAttachments[Constants.FragmentKeys.StatutoryNotices]);
            }
            //if (chkFSG.Checked)
            if(template.IsFSGSelected)
            {
                attachmentFragments.Add(_availableAttachments[Constants.FragmentKeys.FinancialServicesGuide]);
            }
           
            _presenter.InsertSubFragments(Constants.WordBookmarks.RenewalLetterSub, Constants.WordBookmarks.RenewalLetterMain, attachmentFragments);
        }


        private void btnAccountExecutiveLookup_Click(object sender, EventArgs e)
        {
            var peoplePicker = new Popups.PeoplePicker { txtFind = { Text = txtExecutiveName.Text } };
                       
            this.TopMost = false;
            
            peoplePicker.ShowDialog();
            if (peoplePicker.SelectedUser == null) return;

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
            
            //TODO fix this (we have no branch postal addresses)
            txtPostal1.Text = peoplePicker.SelectedUser.BranchPostalAddress.GetPostalAddressFragment(true);
            txtPostal2.Text = peoplePicker.SelectedUser.BranchPostalAddress.GetPostalAddressFragment(false);
        }

        private void datePolicy_ValueChanged(object sender, EventArgs e)
        {
            if (datePayment != null) datePayment.Value = datePolicy.Value.AddDays(14);
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            SwitchTab(tbcWizardScreens.SelectedIndex - 1);
        }

        private void chkContacted_CheckedChanged(object sender, EventArgs e)
        {
            if (Reload && !WizardBeingUpdated)
            {
                var previous = !chkContacted.Checked;

                if (ContinueWithSignificantChange(sender, previous)) _generateNewTemplate = true;
            }
        }

        private void chkNewClient_CheckedChanged(object sender, EventArgs e)
        {
            if (Reload && !WizardBeingUpdated)
            {
                var previous = !chkNewInsurer.Checked;

                if (ContinueWithSignificantChange(sender, previous)) _generateNewTemplate = true;
            }
        }

        private void chkFunding_CheckedChanged(object sender, EventArgs e)
        {
            if (Reload && !WizardBeingUpdated)
            {
                var previous = !chkFunding.Checked;

                if (ContinueWithSignificantChange(sender, previous)) _generateNewTemplate = true;
            }
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

        private void rdoPrePrintYes_CheckedChanged(object sender, EventArgs e)
        {
            isPrePrintedStationary = true;
        }


        private void rdoPrePrintNo_CheckedChanged(object sender, EventArgs e)
        {
            isPrePrintedStationary = false;
        }
       
    }
}
