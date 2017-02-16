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
    public partial class RenewalLetterWizard : BaseForm
    {
        private readonly RenwalLetterPresenter _presenter;

        private List<IPolicyClass> _selectedPolicies;

        private Dictionary<string, DocumentFragment> _availableAttachments;
        //private bool isPrePrintedStationary;

        public RenewalLetterWizard(OfficeDocument document)
        {
            InitializeComponent();
            tbcWizardScreens.SelectedIndexChanged += new EventHandler(tbcWizardScreens_SelectedIndexChanged);

            _presenter = new RenwalLetterPresenter(document, this);

            _checked.CheckStateChanged += new EventHandler<TreePathEventArgs>(_checked_CheckStateChanged);

            _name.DrawText += new EventHandler<DrawEventArgs>(_name_DrawText);
            base.BasePresenter = _presenter;            
            tvaPolicies.Validating += new CancelEventHandler(PolicyValidating);

            txtClientName.Validating += new CancelEventHandler(ClientNameValidating);
        }

        private void _name_DrawText(object sender, DrawEventArgs e)
        {
            if (String.IsNullOrEmpty(txtFindPolicyClass.Text)) return;
            if (e.Node.Parent == null) return;

            if (e.Text.ToUpper().Contains(txtFindPolicyClass.Text.ToUpper()))
            {
                e.TextColor = System.Drawing.Color.Black;
            }
            else
            {
                e.TextColor = System.Drawing.Color.Gray;
            }

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
                errorProvider1.SetError(txtClientName, string.Empty);
            }
        }

        private void PolicyValidating(object sender, CancelEventArgs e)
        {
            if (_selectedPolicies == null || _selectedPolicies.Count == 0)
            {
                errorProvider1.SetError(lblRequiredPolicyClass, "You must select the policy class");
                e.Cancel = true;
            }
            else
            {
                errorProvider1.SetError(lblRequiredPolicyClass, string.Empty);
            }
        }

        private void _checked_CheckStateChanged(object sender, TreePathEventArgs e)
        {
            var node = ((AdvancedTreeNode)e.Path.LastNode);
            var isChecked = node.Checked;

            if (Reload && !WizardBeingUpdated)
            {
                var previous = !isChecked;

                if (ContinueWithSignificantChange(sender, previous))
                {
                    _generateNewTemplate = true;
                }
                else
                {
                    node.Checked = previous;
                }
            }

            //var node = ((AdvancedTreeNode)e.Path.LastNode);
            //var isChecked = node.Checked;

            //if (isChecked)
            //{
            //    var item = MinorItems.Find(i => i.Title == node.Text);
            //    if (item != null)
            //    {
            //        tvaPolicies.AllNodes.ToList().ForEach(
            //            (x) =>
            //            {
            //                var nodeControl = tvaPolicies.GetNodeControls(x);
            //                if (nodeControl != null)
            //                {
            //                    var checkbox = nodeControl.FirstOrDefault(y => (y.Control is NodeCheckBox));
            //                    //checkbox found
            //                    var dCheckBox = (NodeCheckBox)checkbox.Control;
            //                    if (dCheckBox != null)
            //                        dCheckBox.SetValue(x, false);
            //                }
            //            }
            //            );
            //        node.Checked = true;
            //        _selectedPolicy = item;
            //    }
            //}

        }

        private void RenewalLetter_Load(object sender, EventArgs e)
        {
            try
            {

                txtAddressee.Focus();
                txtAddressee.Select();

                GetFragements();

                var uiScheduler = TaskScheduler.FromCurrentSynchronizationContext();

                if (Reload)
                {
                    base.LoadTreeViewClasses(null);
                    ReloadFields();
                }
                else
                {
                    datePolicy.Value = DateTime.Today;
                    dateReport.Value = DateTime.Today;

                    datePayment.Value = DateTime.Today.AddDays(14);
                    Task.Factory.StartNew(() => base.LoadTreeViewClasses(uiScheduler));

                }
                base.LoadGenericImageTabs(uiScheduler, tbcWizardScreens, lblCoverPageTitle.Text, lblLogoTitle.Text);


                // _presenter.CloseInformationPanel(true);
            }
            catch (Exception ex)
            {
                OnError(ex);
            }
        }

        private void GetFragements()
        {
            _availableAttachments = new Dictionary<string, DocumentFragment>();
            List<ISharePointListItem> fragments = null;
            if (Cache.Contains(Constants.CacheNames.RenLtrFragments))
            {
                fragments = (List<ISharePointListItem>) Cache.Get(Constants.CacheNames.RenLtrFragments);
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
                                    Url = Settings.Default.FragmentFSGLetter
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

        private void ReloadFields()
        {

            var auto = false;

            WizardBeingUpdated = true;
            var template = new RenewalLetter();
            List<IPolicyClass> policiesReference = null;

            _selectedPolicies = new List<IPolicyClass>();

            if (Cache.Contains(Constants.CacheNames.RegenerateTemplate))
            {
                template = (RenewalLetter)Cache.Get(Constants.CacheNames.RegenerateTemplate);
                Cache.Remove(Constants.CacheNames.RegenerateTemplate);

                if (template.Policies != null)
                    policiesReference = template.Policies;

                chkWarning.Checked = template.IsAdviceWarningSelected;
                chkFSG.Checked = template.IsFSGSelected;
                chkPrivacy.Checked = template.IsPrivacySelected;
                chkRisks.Checked = template.IsRisksSelected;
                chkSatutory.Checked = template.IsSatutorySelected;

                chkGAW.Checked = template.IsGAWSelected;
                chkContacted.Checked = template.IsContactSelected;
                chkNewInsurer.Checked = template.IsNewClientSelected;
                chkFunding.Checked = template.IsFundingSelected;

                
                if (template.IsPrePrintSelected)
                {
                    chkPrePrint.Checked = template.IsPrePrintSelected;
                }


                Reload = false; //behaviour like a new from this point

                auto = true;
            }
            else
            {
                template = (RenewalLetter)_presenter.LoadData(template);
                policiesReference = _presenter.ReadPoliciesInDocument();

                UpdateMainFragmentCheckBox(Constants.WordDocumentProperties.RlChkContacted, chkContacted);
                UpdateMainFragmentCheckBox(Constants.WordDocumentProperties.RlChkFunding, chkFunding);
                UpdateMainFragmentCheckBox(Constants.WordDocumentProperties.RlChkNewClient, chkNewInsurer);
                UpdateMainFragmentCheckBox(Constants.WordDocumentProperties.RlChkGAW, chkGAW);



                //var value = _presenter.ReadDocumentProperty(Constants.WordDocumentProperties.RlRdoPreprint);
                //bool isSelected;
                //if (Boolean.TryParse(value, out isSelected))
                //{
                //    rdoPrePrintYes.Checked = isSelected;
                //}


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


            var processedTitles = new List<string>();
            foreach (var no in tvaPolicies.AllNodes)
            {
                foreach (var cno in no.Children)
                {
                    if (policiesReference != null && policiesReference.Count > 0)
                    {
                        var p = policiesReference.Find(i => i.Title.Replace("\r\a", string.Empty).Equals(cno.Tag.ToString().Replace("\r\a", string.Empty), StringComparison.OrdinalIgnoreCase));
                        if (p != null)
                        {
                            if (!processedTitles.Contains(cno.Tag.ToString()))
                            {
                                if (String.Equals(cno.Tag.ToString(), p.Title.Replace("\r\a", string.Empty), StringComparison.OrdinalIgnoreCase))
                                {
                                    processedTitles.Add(cno.Tag.ToString());
                                    no.Expand();
                                    var path = tvaPolicies.GetPath(cno);

                                    var node = ((AdvancedTreeNode) path.LastNode);
                                    node.CheckState = CheckState.Checked;
                                    node.Checked = true;

                                    if ((_selectedPolicies.Find(j => j.Title == node.Text) == null))
                                        _selectedPolicies.Add(MinorItems.FirstOrDefault(i => i.Title == node.Text));
                                }
                            }
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
            datePolicy.Value = DateTime.TryParse(template.DatePolicyExpiry, out outDate)
                                   ? outDate
                                   : DateTime.Today;

            datePayment.Value = DateTime.TryParse(template.PaymentDate, out outDate)
                                            ? outDate
                                            : DateTime.Today;

            dateReport.Value = DateTime.TryParse(template.DatePrepared, out outDate)
                                            ? outDate
                                            : DateTime.Today;

            

            WizardBeingUpdated = false;

            if (auto)
            {
                base.StartTimer();
            }
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


        private void UpdateMainFragmentCheckBox(string proName, CheckBox chk)
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

        private List<IPolicyClass> GetSelectedPolicies()
        {
            var policies = new List<IPolicyClass>();

            foreach (var no in tvaPolicies.AllNodes)
            {
                foreach (var cno in no.Children)
                {
                    var path = tvaPolicies.GetPath(cno);

                    var node = ((AdvancedTreeNode) path.LastNode);
                    if (node.Checked)
                    {
                        policies.Add(MinorItems.FirstOrDefault(i => i.Title == node.Text));
                    }
                }
            }

            return policies;
        }

        private void Next_Click(object sender, EventArgs e)
        {
            try
            {
                this.Cursor = Cursors.WaitCursor;
                //   BasePresenter.SwitchScreenUpdating(false);

                if (String.Equals(btnNext.Text, "&Finish", StringComparison.CurrentCultureIgnoreCase))
                {
                    _selectedPolicies = GetSelectedPolicies();

                    if (Validation.HasValidationErrors(this.Controls))
                    {
                        MessageBox.Show(@"Please ensure all required fields are populated", @"Required fields are missing", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }

                    _presenter.CloseInformationPanel(true); //has to be here to close as the presenter opens new word documents taht become active.

                    var letter = new RenewalLetter
                        {
                            Addressee = txtAddressee.Text,
                            Salutation = txtSalutation.Text,
                            ClientAddress = txtClientAddress.Text,
                            ClientAddressLine2 = txtClientAddress2.Text,
                            ClientAddressLine3 = txtClientAddress3.Text,

                            ClientName = txtClientName.Text,
                            DatePolicyExpiry = datePolicy.Value.ToShortDateString(),
                            DatePrepared = dateReport.Value.ToShortDateString(),
                            ExecutiveName = txtExecutiveName.Text,
                            ExecutiveEmail = txtExecutiveEmail.Text,
                            ExecutiveMobile = txtExecutiveMobile.Text,
                            ExecutivePhone = txtExecutivePhone.Text,
                            ExecutiveTitle = txtExecutiveTitle.Text,

                            ExecutiveDepartment = txtExecutiveDepartment.Text,
                            PaymentDate = datePayment.Value.ToShortDateString(),
                            OAMPSBranchPhone = txtBranchPhone.Text,
                            Reference = txtReference.Text,
                            Fax = txtFax.Text,
                            Policies = _selectedPolicies,
                            IsContactSelected = chkContacted.Checked,
                            IsFundingSelected = chkFunding.Checked,
                            IsNewClientSelected = chkNewInsurer.Checked,

                            IsFSGSelected = chkFSG.Checked,
                            IsPrivacySelected = chkPrivacy.Checked,
                            IsRisksSelected = chkRisks.Checked,
                            IsSatutorySelected = chkSatutory.Checked,
                            IsAdviceWarningSelected = chkWarning.Checked,
                            IsGAWSelected = chkGAW.Checked,
                            OAMPSBranchAddress = txtBranchAddress1.Text,
                            OAMPSBranchAddressLine2 = txtBranchAddress2.Text,
                            PDSDescVersion = txtPDSVersion.Text,
                            OAMPSPostalAddress = txtPostal1.Text,
                            OAMPSPostalAddressLine2 = txtPostal2.Text,
                            IsPrePrintSelected = chkPrePrint.Checked
                        };

                    _presenter.CreateOrUpdateDocumentProperty(Constants.WordDocumentProperties.RlRdoPreprint, chkPrePrint.Checked.ToString());

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
                        if (!Reload)
                            _presenter.PopulatePolicy(letter.Policies, letter.DatePolicyExpiry, chkNewInsurer.Checked);

                        PopulateAddress(letter);

                        if (!Reload)
                            AddMainFragment(letter);

                        //call presenter to populate
                        PopulateDocument(letter, lblCoverPageTitle.Text, lblLogoTitle.Text);

                        //thie information panel loads when a document is in sharePoint that has metadata
                        //clients don't wish to see this so force the close of the panel once the wizard completes.
                        //Presenter.CloseInformationPanel();
                        if (!Reload)
                            AddAttachments(letter);

                        if (chkPrePrint.Checked)
                        {
                            _presenter.DeleteDocumentHeaderAndFooter();
                        }

                        //tracking
                        LogUsage(letter,
                                 _loadType == Helpers.Enums.FormLoadType.RegenerateTemplate
                                     ? Helpers.Enums.UsageTrackingType.RegenerateDocument
                                     : Helpers.Enums.UsageTrackingType.NewDocument);
                    }

                    //_presenter.CloseInformationPanel(true);
                    Close();

                    //foreach (Microsoft.Office.Interop.Word.Document doc in Globals.ThisAddIn.Application.Documents)
                    //{
                    //    doc.Application.DisplayDocumentInformationPanel = false;
                    //}
                }
                else
                {
                    SwitchTab(tbcWizardScreens.SelectedIndex + 1);
                }
            }

            catch (AggregateException aggregateException)
            {
                OnError(aggregateException);
            }
            catch (Exception ex)
            {
                OnError(ex);
            }
            finally
            {
                this.Cursor = Cursors.Default;
                // BasePresenter.SwitchScreenUpdating(true);
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
                _presenter.InsertMainFragment(Constants.WordBookmarks.RenewalLetterMain, Settings.Default.FragmentRLNoContact_NewInsr_NoGAW_NoFunding, template);
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

        private void datePolicy_ValueChanged(object sender, EventArgs e)
        {
            if (datePayment != null) datePayment.Value = datePolicy.Value;//.AddDays(14);
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
    
        private void txtFindPolicyClass_TextChanged(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(txtFindPolicyClass.Text))
                return;

            tvaPolicies.CollapseAll();

            foreach (var n in tvaPolicies.AllNodes)
            {
                if (n.Tag.ToString().ToUpper().Contains(txtFindPolicyClass.Text.ToUpper()))
                {
                    if (n.Parent != null)
                    {
                        if (n.Parent.IsExpanded == false)
                            n.Parent.Expand(false);
                    }
                }
            }
            tvaPolicies.Refresh();
        }

        private void chkGAW_CheckedChanged(object sender, EventArgs e)
        {
            if (Reload && !WizardBeingUpdated)
            {
                var previous = !chkGAW.Checked;

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
