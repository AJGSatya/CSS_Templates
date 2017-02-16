using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.Caching;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Aga.Controls.Tree;
using Aga.Controls.Tree.NodeControls;
using OAMPS.Office.BusinessLogic.Helpers;
using OAMPS.Office.BusinessLogic.Interfaces.Wizards;
using OAMPS.Office.BusinessLogic.Models.Template;
using OAMPS.Office.BusinessLogic.Presenters.SharePoint;
using OAMPS.Office.BusinessLogic.Presenters.Wizards;
using OAMPS.Office.Word.Helpers;
using OAMPS.Office.Word.Helpers.Controls;
using OAMPS.Office.Word.Models.SharePoint;
using OAMPS.Office.Word.Models.Word;
using OAMPS.Office.Word.Properties;
using OAMPS.Office.Word.Views.Wizards.Popups;
using Enums = OAMPS.Office.Word.Helpers.Enums;

namespace OAMPS.Office.Word.Views.Wizards
{
    public partial class QuoteSlipWizard : BaseWizardForm
    {
        protected static List<IQuestionClass> Questions = null;
        private readonly QuoteSlipWizardPresenter _wizardPresenter;
        private bool _populateApprovalForm;
        private bool _populateClaimMadeWarning;
        private List<IPolicyClass> _selectedQuestionnaireFragments = new List<IPolicyClass>();

        public QuoteSlipWizard(OfficeDocument document, Enums.FormLoadType loadType)
        {
            InitializeComponent();
            tbcWizardScreens.SelectedIndexChanged += tbcWizardScreens_SelectedIndexChanged;
            _wizardPresenter = new QuoteSlipWizardPresenter(document, this);
            base.BaseWizardPresenter = _wizardPresenter;
            LoadType = loadType;

            _checked.CheckStateChanged += _checked_CheckStateChanged;

            _name.DrawText += OnDrawPolicyNode;

            clbQuestions.ItemCheck += clbQuestions_ItemCheck;

            clbQuestions.Validating += new System.ComponentModel.CancelEventHandler(clbQuestions_Validating);
            txtClientName.Validating += new System.ComponentModel.CancelEventHandler(txtClientName_Validating);
        }

        void clbQuestions_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (clbQuestions.CheckedItems.Count < 1)
            {
                errorProvider.SetError(clbQuestions, "Field is required");
                e.Cancel = true;
            }
            else
            {
                errorProvider.SetError(clbQuestions, string.Empty);
            }
        }

        void txtClientName_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (String.IsNullOrEmpty(txtClientName.Text))
            {
                errorProvider.SetError(txtClientName, "Field is required");
                e.Cancel = true;
            }
            else
            {
                errorProvider.SetError(txtClientName, string.Empty);
            }
        }



        private void clbQuestions_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            if (clbQuestions.CheckedItems.Count == 1)
            {
                Boolean isCheckedItemBeingUnchecked = (e.CurrentValue == CheckState.Checked);
                if (isCheckedItemBeingUnchecked)
                {
                    e.NewValue = CheckState.Checked;
                }
                else
                {
                    Int32 checkedItemIndex = clbQuestions.CheckedIndices[0];
                    clbQuestions.ItemCheck -= clbQuestions_ItemCheck;
                    clbQuestions.SetItemChecked(checkedItemIndex, false);
                    clbQuestions.ItemCheck += clbQuestions_ItemCheck;
                }

                return;
            }
        }

        private void OnDrawPolicyNode(object sender, DrawEventArgs e)
        {
            ColorFindMatches(e);
            ColorMappings(e);
        }

        private void ColorMappings(DrawEventArgs e)
        {
            if (MinorItems.Any(x => x.FragmentPolicyUrl != null && x.Title == e.Text))
            {
                e.TextColor = Color.Green;
            }
            else
            {
                if (e.Node.Children.Count > 0 || !String.IsNullOrEmpty(txtFindPolicyClass.Text)) return;
                e.TextColor = Color.DimGray;
            }
        }

        private void ColorFindMatches(DrawEventArgs e)
        {
            if (!String.IsNullOrEmpty(txtFindPolicyClass.Text))
            {
                if (e.Node.Parent != null)
                {
                    e.TextColor = e.Text.ToUpper().Contains(txtFindPolicyClass.Text.ToUpper()) ? Color.Black : Color.Gray;
                }
            }
        }

        private void tbcWizardScreens_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnNext.Text = tbcWizardScreens.SelectedIndex == tbcWizardScreens.TabCount - 1 ? "&Finish" : "&Next";
            btnBack.Enabled = tbcWizardScreens.SelectedIndex != 0;
            btnNext.Enabled = btnNext.Text != @"&Finish" || LoadComplete;
        }

        private void _checked_CheckStateChanged(object sender, TreePathEventArgs e)
        {
            var node = ((AdvancedTreeNode) e.Path.LastNode);
            bool isChecked = node.Checked;


            if (Reload && !WizardBeingUpdated)
            {
                bool previous = !isChecked;

                if (ContinueWithSignificantChange(sender, previous))
                {
                    GenerateNewTemplate = true;
                }
                else
                {
                    TreeNodeAdv nodeAdv = tvaPolicies.AllNodes.FirstOrDefault(x => x.ToString() == node.Text);

                    Type type = sender.GetType();
                    if (type == typeof (NodeCheckBox))
                    {
                        ((NodeCheckBox) sender).SetValue(nodeAdv, previous);
                    }
                    return;
                }
            }

            IPolicyClass item = MinorItems.Find(i => i.Title == node.Text);

            if (isChecked)
            {
                if (item != null)
                {
                    //does it already exist in the selected fragemnts.  
                    //multiple policy classes are mapped to the same fragement, we only want to insert one copy of the fragement in the template
                    //also alert the user that they have selected a policy class already mapped to a selected policy class
                    IPolicyClass mappingAlreadySelected = _selectedQuestionnaireFragments.Find(i => i.FragmentPolicyUrl == item.FragmentPolicyUrl);
                    if (mappingAlreadySelected == null) //no existing mapping found
                    {
                        _selectedQuestionnaireFragments.Add(item);
                    }
                    else
                    {
                        MessageBox.Show(String.Format("The selected policy class '{0}' is already mapped to the same questionnaire as a previously selected policy class '{1}'. \n You will only receive one questionnaire for these policy classes.", item.Title, mappingAlreadySelected.Title));
                    }
                }
            }
            else
            {
                _selectedQuestionnaireFragments.Remove(item);
            }
        }

        protected override List<IPolicyClass> LoadMinorPolicyTypes()
        {
            List<IPolicyClass> policies = base.LoadMinorPolicyTypes();

            var spList = new SharePointList(Settings.Default.SharePointContextUrl, Settings.Default.PreRenewalQuestionareMappingsListName);
            var presenter = new SharePointListPresenter(spList, this);

            List<IPreRenewalQuestionareMappings> mappings = presenter.GetPreRenewalQuestionareMappings();
            policies.ForEach((x) =>
                {
                    IPreRenewalQuestionareMappings w = mappings.FirstOrDefault(f => f.PolicyType.Contains(x.Title));
                    if (w != null)
                    {
                        x.FragmentPolicyUrl = w.FragmentUrl;
                    }
                });
            return policies;
        }

        private QuoteSlip GenerateTempalteObject()
        {
            //buid the marketing template
            var template = new QuoteSlip
                {
                    DocumentTitle = "Quotation Slip", //BaseWizardPresenter.ReadDocumentProperty("Title"), //Constants.TemplateNames.InsuranceRenewalReport,
                    DocumentSubTitle = string.Empty,
                    ClientName = txtClientName.Text,
                    ClientCommonName = txtClientCommonName.Text,
                    PeriodOfInsuranceFrom = dtpPeriodOfInsuranceFrom.Text,
                    PeriodOfInsuranceTo = dtpPeriodOfInsuranceTo.Text,
                    
                    ExecutiveName = txtExecutiveName.Text,
                    ExecutiveEmail = txtExecutiveEmail.Text,
                    ExecutivePhone = txtExecutivePhone.Text,
                    ExecutiveTitle = txtExecutiveTitle.Text,
                    ExecutiveMobile = txtExecutiveMobile.Text,
                    ExecutiveDepartment = txtExecutiveDepartment.Text,
                    OAMPSBranchPhone = txtBranchPhone.Text,
                    Fax = txtFax.Text,
                    OAMPSPostalAddress = txtPostal1.Text,
                    OAMPSPostalAddressLine2 = txtPostal2.Text,

                    AssistantExecutiveName = txtAssistantExecutiveName.Text,
                    AssistantExecutiveTitle = txtAssistantExecutiveTitle.Text,
                    AssistantExecutivePhone = txtAssistantExecutivePhone.Text,
                    AssistantExecutiveEmail = txtAssistantExecutiveEmail.Text,
                    AssistantExecDepartment = txtAssitantExecDepartment.Text,
                   
                    OAMPSBranchAddress = txtBranchAddress1.Text,
                    OAMPSBranchAddressLine2 = txtBranchAddress2.Text,

                    DatePrepared = DateTime.Now.ToString(@"dd/MM/yyyy"),
                    DateRequiredBy = dtpDateRequired.Value.ToString(@"dd/MM/yyyy"),
                    SelectedDocumentFragments = _selectedQuestionnaireFragments,
                    PopulateApprovalForm = _populateApprovalForm,
                    PopulateClaimMadeWarning = _populateClaimMadeWarning
                };


            var baseTemplate = (BaseTemplate) template;
            TabPage logoTab = tbcWizardScreens.TabPages[Constants.ControlNames.TabPageLogosName];

            PopulateLogosToTemplate(logoTab, ref baseTemplate);

            TabPage covberTab =
                tbcWizardScreens.TabPages[Constants.ControlNames.TabPageCoverPagesName];

            PopulateCoversToTemplate(covberTab, ref baseTemplate);

            return template;
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
                    if (GenerateNewTemplate)
                    {
                        var template = GenerateTempalteObject();

                        Cache.Add(Constants.CacheNames.RegenerateTemplate, template, new CacheItemPolicy());

                        _wizardPresenter.GenerateNewTemplate(Constants.CacheNames.RegenerateTemplate, Settings.Default.TemplatePrerenewalQuestionare);
                    }
                    else
                    {
                        //    //call presenter to populate
                        var tempalte = GenerateTempalteObject();

                        if (!Reload)
                        {
                            //_presenter.PopulateClaimMadeWarningFragment(_populateClaimMadeWarning, Settings.Default.FragmentPRQClaimsMadeWarning);
                            //_presenter.PopulateApprovalFormFragment(_populateApprovalForm, Settings.Default.FragmentPRQApprovalForm);

                            PopulateQuestionFragments();
                            //_presenter.InsertQuestionnaireFragement(_selectedQuestionnaireFragments);
                        }

                        PopulateDocument(tempalte, lblCoverPageTitle.Text, lblLogoTitle.Text);

                        
                        if (!Reload)
                        {
                            //    //thie information panel loads when a document is in sharePoint that has metadata
                            //    //clients don't wish to see this so force the close of the panel once the wizard completes.
                            _wizardPresenter.CloseInformationPanel();
                            _wizardPresenter.MoveToStartOfDocument();
                        }
                    }


                    Close();
                }
                catch (Exception ex)
                {
                    OnError(ex);
                }
                finally
                {
                    Cursor = Cursors.Default;
                    //    BasePresenter.SwitchScreenUpdating(true);
                }
            }
            else
            {
                SwitchTab(tbcWizardScreens.SelectedIndex + 1);
            }
        }

        private void PopulateQuestionFragments()
        {
            CheckedListBox.CheckedItemCollection c = clbQuestions.CheckedItems;
            var questionItems = new List<IQuestionClass>();
            var scheduleItems = new List<IQuestionClass>();

            foreach (object i in c)
            {
                questionItems.Add((IQuestionClass) i);

                var t = (IQuestionClass) i;
                IQuestionClass sh = LoadSchedule(Settings.Default.SharePointContextUrl, Settings.Default.PolicySchedulesListName, t.Title);

                if (sh == null) continue;

                scheduleItems.Add(sh);
            }

            _wizardPresenter.InsertPolicySchedule(scheduleItems, true);

           // _wizardPresenter.InsertQuestionnaireFragement(questionItems);

        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            SwitchTab(tbcWizardScreens.SelectedIndex - 1);
        }

        private void SwitchTab(int index)
        {
            tbcWizardScreens.SelectedIndex = index;
        }

        private void rdoWariningYes_CheckedChanged(object sender, EventArgs e)
        {
            if (Reload && !WizardBeingUpdated)
            {
                if (ContinueWithSignificantChange(sender, true)) GenerateNewTemplate = true;
            }
            _populateClaimMadeWarning = true;
        }

        private void rdoWariningNo_CheckedChanged(object sender, EventArgs e)
        {
            if (Reload && !WizardBeingUpdated)
            {
                if (ContinueWithSignificantChange(sender, true)) GenerateNewTemplate = true;
            }
            _populateClaimMadeWarning = false;
        }

        private void rdoApprovalYes_CheckedChanged(object sender, EventArgs e)
        {
            if (Reload && !WizardBeingUpdated)
            {
                if (ContinueWithSignificantChange(sender, true)) GenerateNewTemplate = true;
            }
            _populateApprovalForm = true;
        }

        private void rdoApprovalNo_CheckedChanged(object sender, EventArgs e)
        {
            if (Reload && !WizardBeingUpdated)
            {
                if (ContinueWithSignificantChange(sender, true)) GenerateNewTemplate = true;
            }
            _populateApprovalForm = false;
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

        private void btnAssistantAccountExecutiveLookup_Click(object sender, EventArgs e)
        {
            var peoplePicker = new PeoplePicker(txtAssistantExecutiveName.Text, this);
            if (peoplePicker.SelectedUser == null)
            {
                TopMost = false;

                peoplePicker.ShowDialog();
                if (peoplePicker.SelectedUser == null) return;
            }

            txtAssistantExecutiveName.Text = peoplePicker.SelectedUser.DisplayName;
            txtAssistantExecutiveTitle.Text = peoplePicker.SelectedUser.Title;
            txtAssistantExecutiveEmail.Text = peoplePicker.SelectedUser.EmailAddress;
            txtAssistantExecutivePhone.Text = peoplePicker.SelectedUser.VoiceTelephoneNumber;
            txtAssitantExecDepartment.Text = peoplePicker.SelectedUser.Branch;
        }

        private void txtFindPolicyClass_TextChanged(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(txtFindPolicyClass.Text))
                return;

            tvaPolicies.CollapseAll();

            foreach (TreeNodeAdv n in tvaPolicies.AllNodes)
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

        private void QuoteSlipWizard_Load(object sender, EventArgs e)
        {
            tbcWizardScreens.TabPages.Remove(tabPolicyClass);
            tbcWizardScreens.TabPages.Remove(tabQuestions);
            txtClientName.Focus();
            txtClientName.Select();

            //bool auto = false;
            WizardBeingUpdated = true;

            TaskScheduler uiScheduler = TaskScheduler.FromCurrentSynchronizationContext();

            //if (Cache.Contains(Constants.CacheNames.RegenerateTemplate))
            //{
            //    Reload = false;

            //    var template = GetCachedTempalteObject<PreRenewalQuestionare>();
            //    lblCoverPageTitle.Text = template.CoverPageTitle;
            //    lblLogoTitle.Text = template.LogoTitle;

            //    _selectedQuestionnaireFragments = template.SelectedDocumentFragments;

            //    //LoadDataSources(null); //dont use new thread

            //    ReloadAllFields(template);

            //    //ReloadPolicyClasses(true);

            //    //rdoWariningYes.Checked = template.PopulateClaimMadeWarning;
            //    //rdoApprovalYes.Checked  = template.PopulateApprovalForm;

            //    auto = true;
            //}
            //else
            //{
            if (Reload)
            {
                MessageBox.Show(@"You cannot make any changes to this quote slip through the wizard", @"Cannot make any changes", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Close();
                //LoadDataSources(null); //dont use new thread

                //var template = new PreRenewalQuestionare();
                //template = (PreRenewalQuestionare) _presenter.LoadData(template);

                //ReloadAllFields(template);

                //ReloadPolicyClasses(false);

                //var claimsMade = _presenter.ReadDocumentProperty(Constants.WordDocumentProperties.ClaimsMadeWarning);
                //var approvalForms = _presenter.ReadDocumentProperty(Constants.WordDocumentProperties.ApprovalForm);

                //if (String.Equals(claimsMade, "true", StringComparison.OrdinalIgnoreCase)) rdoWariningYes.Checked = true;

                //if (String.Equals(approvalForms, "true", StringComparison.OrdinalIgnoreCase)) rdoApprovalYes.Checked = true;
            }
            else
            {
                Task.Factory.StartNew(() =>
                    {
                        // LoadDataSources(uiScheduler);
                        LoadQuestions(uiScheduler);
                    });
            }
            //}

            base.LoadGenericImageTabs(uiScheduler, tbcWizardScreens, lblCoverPageTitle.Text, lblLogoTitle.Text);

            WizardBeingUpdated = false;

           // if (auto) base.StartTimer();
        }

        private void LoadQuestions(TaskScheduler uiScheduler)
        {
            if (Cache.Contains(Constants.CacheNames.QuoteSlipSchedules))
            {
                Questions =
                    ((List<IQuestionClass>)
                     Cache.Get(Constants.CacheNames.QuoteSlipSchedules));
            }
            else
            {
                Questions = LoadQuestionsFromSharePoint(Settings.Default.SharePointContextUrl, Settings.Default.PolicySchedulesListName);
                Cache.Add(Constants.CacheNames.QuoteSlipSchedules, Questions,
                          new CacheItemPolicy());
            }

            Task.Factory.StartNew(() =>
                {
                    clbQuestions.DataSource = Questions;
                    clbQuestions.DisplayMember = "Title";
                    clbQuestions.ValueMember = "Id";
                }, CancellationToken.None, TaskCreationOptions.None, uiScheduler);
        }

        protected List<IQuestionClass> LoadQuestionsFromSharePoint(string contextUrl, string listName)
        {
            var list = new SharePointList(contextUrl, listName, Constants.SharePointQueries.AllItemsSortBySortOrder);
            var presenter = new SharePointListPresenter(list, this);
            return presenter.GetPreRenewalQuestionaireQuestions();
        }

        protected IQuestionClass LoadSchedule(string contextUrl, string listName, string title)
        {
            var list = new SharePointList(contextUrl, listName, String.Format(Constants.SharePointQueries.GetItemByTitleQuery, title));
            var presenter = new SharePointListPresenter(list, this);
            List<IQuestionClass> t = presenter.GetPreRenewalQuestionaireQuestions();

            return (t != null && t.Count > 0) ? t[0] : null;
        }

        private void ReloadPolicyClasses(bool regenProcess)
        {
            string ids = regenProcess ? _selectedQuestionnaireFragments.Aggregate(string.Empty, (current, p) => current + p.Id + ";")
                             : _wizardPresenter.ReadDocumentProperty(Constants.WordDocumentProperties.IncludedPolicyTypes);

            if (String.IsNullOrEmpty(ids)) return;
            string[] l = ids.Split(';');
            foreach (string id in l)
            {
                IPolicyClass found = MinorItems.FirstOrDefault(x => x.Id == id);
                if (found == null) continue;

                foreach (TreeNodeAdv no in tvaPolicies.AllNodes)
                {
                    if (String.Equals(no.Tag.ToString(), found.Title, StringComparison.OrdinalIgnoreCase))
                    {
                        if (!regenProcess)
                            _selectedQuestionnaireFragments.Add(found);


                        TreePath path = tvaPolicies.GetPath(no);
                        var node = ((AdvancedTreeNode) path.LastNode);
                        node.CheckState = CheckState.Checked;
                        node.Checked = true;
                        if (no.Parent != null) no.Parent.ExpandAll();
                    }
                }
            }
        }

        private void ReloadAllFields(FactFinder template)
        {
            txtAssistantExecutiveEmail.Text = template.AssistantExecutiveEmail;
            txtAssistantExecutiveName.Text = template.AssistantExecutiveName;
            txtAssistantExecutivePhone.Text = template.AssistantExecutivePhone;
            txtAssistantExecutiveTitle.Text = template.AssistantExecutiveTitle;
            txtAssitantExecDepartment.Text = template.AssistantExecDepartment;

            txtBranchAddress1.Text = template.OAMPSBranchAddress;
            txtBranchAddress2.Text = template.OAMPSBranchAddressLine2;

            txtClientCommonName.Text = template.ClientCommonName;
            txtClientName.Text = template.ClientName;

            txtExecutiveDepartment.Text = template.ExecutiveDepartment;
            txtExecutiveEmail.Text = template.ExecutiveEmail;
            txtExecutiveName.Text = template.ExecutiveName;
            txtExecutivePhone.Text = template.ExecutivePhone;
            txtExecutiveTitle.Text = template.ExecutiveTitle;

        }

        private void dtpPeriodOfInsuranceFrom_ValueChanged(object sender, EventArgs e)
        {
            dtpPeriodOfInsuranceTo.Value = dtpPeriodOfInsuranceFrom.Value.AddMonths(12);
        }



    
    }
}