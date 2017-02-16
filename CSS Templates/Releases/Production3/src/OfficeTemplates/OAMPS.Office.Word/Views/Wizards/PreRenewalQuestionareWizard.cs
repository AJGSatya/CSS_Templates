using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.Caching;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Aga.Controls.Tree;
using Aga.Controls.Tree.NodeControls;
using OAMPS.Office.BusinessLogic.Helpers;
using OAMPS.Office.BusinessLogic.Interfaces.SharePoint;
using OAMPS.Office.BusinessLogic.Interfaces.Wizards;
using OAMPS.Office.BusinessLogic.Models.Template;
using OAMPS.Office.BusinessLogic.Models.Wizards;
using OAMPS.Office.BusinessLogic.Presenters;
using OAMPS.Office.BusinessLogic.Presenters.SharePoint;
using OAMPS.Office.Word.Helpers.Controls;
using OAMPS.Office.Word.Models.SharePoint;
using OAMPS.Office.Word.Models.Word;
using OAMPS.Office.Word.Properties;

namespace OAMPS.Office.Word.Views.Wizards
{
    public partial class PreRenewalQuestionareWizard : BaseForm
    {
        private List<IPolicyClass> _selectedQuestionnaireFragments = new List<IPolicyClass>();
        private bool _populateClaimMadeWarning;
        private bool _populateApprovalForm;
        private readonly PreRenewalQuestionarePresenter _presenter;
        protected static List<IQuestionClass> Questions = null;

        public PreRenewalQuestionareWizard(OfficeDocument document, Helpers.Enums.FormLoadType loadType)
        {
            InitializeComponent();
            tbcWizardScreens.SelectedIndexChanged += new EventHandler(tbcWizardScreens_SelectedIndexChanged);
            _presenter = new PreRenewalQuestionarePresenter(document, this);
            base.BasePresenter = _presenter;
            _loadType = loadType;

            _checked.CheckStateChanged += new EventHandler<TreePathEventArgs>(_checked_CheckStateChanged);

            _name.DrawText += new EventHandler<DrawEventArgs>(OnDrawPolicyNode);

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
                    if (e.Text.ToUpper().Contains(txtFindPolicyClass.Text.ToUpper()))
                    {
                        e.TextColor = System.Drawing.Color.Black;
                    }
                    else
                    {
                        e.TextColor = System.Drawing.Color.Gray;
                    }
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
                    TreeNodeAdv nodeAdv = tvaPolicies.AllNodes.FirstOrDefault(x => x.ToString() == node.Text);

                    var type = sender.GetType();
                    if (type == typeof(NodeCheckBox))
                    {
                        ((NodeCheckBox)sender).SetValue(nodeAdv, previous);
                    }
                    return;
                }
            }

            var item = MinorItems.Find(i => i.Title == node.Text);
            
            if (isChecked)
            {                
                if (item != null)
                {
                    //does it already exist in the selected fragemnts.  
                    //multiple policy classes are mapped to the same fragement, we only want to insert one copy of the fragement in the template
                    //also alert the user that they have selected a policy class already mapped to a selected policy class
                    var mappingAlreadySelected = _selectedQuestionnaireFragments.Find(i => i.FragmentPolicyUrl == item.FragmentPolicyUrl);
                    if (mappingAlreadySelected == null) //no existing mapping found
                    {
                        _selectedQuestionnaireFragments.Add(item);    
                    }
                    else
                    {
                        MessageBox.Show(String.Format("The selected policy class '{0}' is already mapped to the same questionnaire as a previously selected policy class '{1}'. \n You will only receive one questionnaire for these policy classes." , item.Title, mappingAlreadySelected.Title));
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

            var policies = base.LoadMinorPolicyTypes();

            var spList = new SharePointList(Settings.Default.SharePointContextUrl, Settings.Default.PreRenewalQuestionareMappingsListName);
            var presenter = new SharePointListPresenter(spList, this);
            
            var mappings = presenter.GetPreRenewalQuestionareMappings();
            policies.ForEach((x) =>
                {
                    var w = mappings.FirstOrDefault(f => f.PolicyType.Contains(x.Title ));
                    if (w != null)
                    {
                        x.FragmentPolicyUrl = w.FragmentUrl;
                    }
                });
            return policies;
        }

        private PreRenewalQuestionare GenerateTempalteObject()
        {
            //buid the marketing template
            var template = new PreRenewalQuestionare
            {
                DocumentTitle = BasePresenter.ReadDocumentProperty("Title"),//Constants.TemplateNames.InsuranceRenewalReport,
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

                AssistantExecutiveName = txtAssistantExecutiveName.Text,
                AssistantExecutiveTitle = txtAssistantExecutiveTitle.Text,
                AssistantExecutivePhone = txtAssistantExecutivePhone.Text,
                AssistantExecutiveEmail = txtAssistantExecutiveEmail.Text,
                AssistantExecDepartment = txtAssitantExecDepartment.Text,


                ClaimsExecutiveName = txtClaimsExecutiveName.Text,
                ClaimsExecutiveTitle = txtClaimsExecutiveTitle.Text,
                ClaimsExecutivePhone = txtClaimsExecutivePhone.Text,
                ClaimsExecutiveEmail = txtClaimsExecutiveEmail.Text,
                ClaimsExecDepartment = txtClaimExecDepartment.Text,

                OtherContactName = txtOtherContactName.Text,
                OtherContactTitle = txtOtherContactTitle.Text,
                OtherContactPhone = txtOtherContactPhone.Text,
                OtherContactEmail = txtOtherContactEmail.Text,
                OtherExecDepartment = txtOtherExecDepartment.Text,

                OAMPSBranchAddress = txtBranchAddress1.Text,
                OAMPSBranchAddressLine2 = txtBranchAddress2.Text,
                DatePrepared = DateTime.Now.ToString(@"dd/MM/yyyy"),

                SelectedDocumentFragments = _selectedQuestionnaireFragments,
                PopulateApprovalForm = _populateApprovalForm,
                PopulateClaimMadeWarning = _populateClaimMadeWarning
                
            };


            var baseTemplate = (BaseTemplate)template;
            var logoTab = tbcWizardScreens.TabPages[Constants.ControlNames.TabPageLogosName];

            PopulateLogosToTemplate(logoTab, ref baseTemplate);

            var covberTab =
                tbcWizardScreens.TabPages[Constants.ControlNames.TabPageCoverPagesName];

            PopulateCoversToTemplate(covberTab, ref baseTemplate);

            return template;
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

                try
                {
                    this.Cursor = Cursors.WaitCursor;
                    if (_generateNewTemplate)
                    {
                        var template = GenerateTempalteObject();

                        Cache.Add(Constants.CacheNames.RegenerateTemplate, template, new CacheItemPolicy());

                        _presenter.GenerateNewTemplate(Constants.CacheNames.RegenerateTemplate, Settings.Default.TemplatePrerenewalQuestionare);
                    }
                    else
                    {
                        //    //call presenter to populate
                        var tempalte = GenerateTempalteObject();

                        if (!Reload)
                        {
                            _presenter.PopulateClaimMadeWarningFragment(_populateClaimMadeWarning, Settings.Default.FragmentPRQClaimsMadeWarning);
                            _presenter.PopulateApprovalFormFragment(_populateApprovalForm, Settings.Default.FragmentPRQApprovalForm);
                            
                            var c = clbQustions.CheckedItems;
                            var items = new List<IQuestionClass>();
                            foreach (var i in c)
                            {
                                items.Add((IQuestionClass)i);
                            }
                            _presenter.InsertQuestionnaireFragement(items);
                            //_presenter.InsertQuestionnaireFragement(_selectedQuestionnaireFragments);
                        }

                        PopulateDocument(tempalte, lblCoverPageTitle.Text, lblLogoTitle.Text);


                        if (!Reload)
                        {
                            //    //thie information panel loads when a document is in sharePoint that has metadata
                            //    //clients don't wish to see this so force the close of the panel once the wizard completes.
                            _presenter.CloseInformationPanel();
                            _presenter.MoveToStartOfDocument();
                        }


                        //tracking
                        LogUsage(tempalte,
                         _loadType == Helpers.Enums.FormLoadType.RegenerateTemplate
                             ? Helpers.Enums.UsageTrackingType.RegenerateDocument
                             : Helpers.Enums.UsageTrackingType.NewDocument);
                    }


                    Close();
                }
                     catch (Exception ex)
                {
                    OnError(ex);
                }
                finally
                {
                    this.Cursor = Cursors.Default;
                //    BasePresenter.SwitchScreenUpdating(true);
                }
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

        private void SwitchTab(int index)
        {
            tbcWizardScreens.SelectedIndex = index;
        }

        private void rdoWariningYes_CheckedChanged(object sender, EventArgs e)
        {
            if (Reload && !WizardBeingUpdated)
            {
                if (ContinueWithSignificantChange(sender, true)) _generateNewTemplate = true;
            }
            _populateClaimMadeWarning = true;
        }

        private void rdoWariningNo_CheckedChanged(object sender, EventArgs e)
        {
            if (Reload && !WizardBeingUpdated)
            {
                if (ContinueWithSignificantChange(sender, true)) _generateNewTemplate = true;
            }
            _populateClaimMadeWarning = false;
        }

        private void rdoApprovalYes_CheckedChanged(object sender, EventArgs e)
        {
            if (Reload && !WizardBeingUpdated)
            {
                if (ContinueWithSignificantChange(sender, true)) _generateNewTemplate = true;
            }
            _populateApprovalForm = true;
        }

        private void rdoApprovalNo_CheckedChanged(object sender, EventArgs e)
        {
            if (Reload && !WizardBeingUpdated)
            {
                if (ContinueWithSignificantChange(sender, true)) _generateNewTemplate = true;
            }
            _populateApprovalForm = false;
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
            txtExecutiveMobile.Text = peoplePicker.SelectedUser.Mobile;
            txtBranchAddress1.Text = peoplePicker.SelectedUser.BranchAddressLine1;
            txtBranchAddress2.Text = peoplePicker.SelectedUser.BranchAddressLine2;
            txtExecutiveDepartment.Text = peoplePicker.SelectedUser.Branch;
        }

        private void btnAssistantAccountExecutiveLookup_Click(object sender, EventArgs e)
        {
            var peoplePicker = new Popups.PeoplePicker(txtAssistantExecutiveName.Text, this);
            if (peoplePicker.SelectedUser == null)
            {
                this.TopMost = false;

                peoplePicker.ShowDialog();
                if (peoplePicker.SelectedUser == null) return;
            }

            txtAssistantExecutiveName.Text = peoplePicker.SelectedUser.DisplayName;
            txtAssistantExecutiveTitle.Text = peoplePicker.SelectedUser.Title;
            txtAssistantExecutiveEmail.Text = peoplePicker.SelectedUser.EmailAddress;
            txtAssistantExecutivePhone.Text = peoplePicker.SelectedUser.VoiceTelephoneNumber;
            txtAssitantExecDepartment.Text = peoplePicker.SelectedUser.Branch;
        }

        private void btnClaimsExecutiveLookup_Click(object sender, EventArgs e)
        {
            var peoplePicker = new Popups.PeoplePicker(txtClaimsExecutiveName.Text, this);
            if (peoplePicker.SelectedUser == null)
            {
                this.TopMost = false;

                peoplePicker.ShowDialog();
                if (peoplePicker.SelectedUser == null) return;
            }

            txtClaimsExecutiveName.Text = peoplePicker.SelectedUser.DisplayName;
            txtClaimsExecutiveTitle.Text = peoplePicker.SelectedUser.Title;
            txtClaimsExecutiveEmail.Text = peoplePicker.SelectedUser.EmailAddress;
            txtClaimsExecutivePhone.Text = peoplePicker.SelectedUser.VoiceTelephoneNumber;
            txtClaimExecDepartment.Text = peoplePicker.SelectedUser.Branch;
        }

        private void btnOtherContactLookup_Click(object sender, EventArgs e)
        {
            var peoplePicker = new Popups.PeoplePicker(txtOtherContactName.Text, this);
            if (peoplePicker.SelectedUser == null)
            {
                this.TopMost = false;

                peoplePicker.ShowDialog();
                if (peoplePicker.SelectedUser == null) return;
            }

            txtOtherContactName.Text = peoplePicker.SelectedUser.DisplayName;
            txtOtherContactTitle.Text = peoplePicker.SelectedUser.Title;
            txtOtherContactEmail.Text = peoplePicker.SelectedUser.EmailAddress;
            txtOtherContactPhone.Text = peoplePicker.SelectedUser.VoiceTelephoneNumber;
            txtOtherExecDepartment.Text = peoplePicker.SelectedUser.Branch;
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

        private void PreRenewalQuestionareWizard_Load(object sender, EventArgs e)
        {
            tbcWizardScreens.TabPages.Remove(tabPolicyClass);
            txtClientName.Focus();
            txtClientName.Select();

            var auto = false;
            WizardBeingUpdated = true;

            var uiScheduler = TaskScheduler.FromCurrentSynchronizationContext();

            if (Cache.Contains(Constants.CacheNames.RegenerateTemplate))
            {
                Reload = false;

                var template = GetCachedTempalteObject<PreRenewalQuestionare>();
                lblCoverPageTitle.Text = template.CoverPageTitle;
                lblLogoTitle.Text = template.LogoTitle;

                _selectedQuestionnaireFragments = template.SelectedDocumentFragments;
                LoadDataSources(null); //dont use new thread

                ReloadAllFields(template);
                ReloadPolicyClasses(true);

                rdoWariningYes.Checked = template.PopulateClaimMadeWarning;
                rdoApprovalYes.Checked  = template.PopulateApprovalForm;

                auto = true;

            }
            else
            {
                if (Reload)
                {
                    LoadDataSources(null); //dont use new thread

                    var template = new PreRenewalQuestionare();
                    template = (PreRenewalQuestionare) _presenter.LoadData(template);

                    ReloadAllFields(template);
                    ReloadPolicyClasses(false);

                    var claimsMade = _presenter.ReadDocumentProperty(Constants.WordDocumentProperties.ClaimsMadeWarning);
                    var approvalForms = _presenter.ReadDocumentProperty(Constants.WordDocumentProperties.ApprovalForm);

                    if (String.Equals(claimsMade, "true", StringComparison.OrdinalIgnoreCase)) rdoWariningYes.Checked = true;

                    if (String.Equals(approvalForms, "true", StringComparison.OrdinalIgnoreCase)) rdoApprovalYes.Checked = true;
                }
                else
                {
                    Task.Factory.StartNew(() =>
                    {
                        LoadDataSources(uiScheduler);
                        LoadQuestions(uiScheduler);
                    });
                }
            }

            base.LoadGenericImageTabs(uiScheduler, tbcWizardScreens, lblCoverPageTitle.Text, lblLogoTitle.Text);

            WizardBeingUpdated = false; 

            if(auto)base.StartTimer();
        }

        private void LoadQuestions(TaskScheduler uiScheduler)
        {
            if (Cache.Contains(Constants.CacheNames.PreRenewalQuestionareQuestions))
            {
                Questions =
                    ((List<IQuestionClass>)
                     Cache.Get(Constants.CacheNames.PreRenewalQuestionareQuestions));
            }
            else
            {
                Questions = LoadQuestionsFromSharePoint(Settings.Default.SharePointContextUrl, Settings.Default.PreRenewalQuestionareMappingsListName);
                Cache.Add(Constants.CacheNames.MajorPolicyClassItems, MajorItems,
                          new CacheItemPolicy());
            }

            Task.Factory.StartNew(() =>
            {
                clbQustions.DataSource = Questions;
                clbQustions.DisplayMember = "Title";
                clbQustions.ValueMember = "Id";
            }, CancellationToken.None, TaskCreationOptions.None, uiScheduler);
        }

        protected List<IQuestionClass> LoadQuestionsFromSharePoint(string contextUrl, string listName)
        {
            var list = new SharePointList(contextUrl, listName,Constants.SharePointQueries.AllItemsSortBySortOrder);
            var presenter = new SharePointListPresenter(list, this);
            return presenter.GetPreRenewalQuestionaireQuestions();
        }

        private void ReloadPolicyClasses(bool regenProcess)
        {
            var ids = regenProcess ? _selectedQuestionnaireFragments.Aggregate(string.Empty, (current, p) => current + p.Id + ";")
                          : _presenter.ReadDocumentProperty(Constants.WordDocumentProperties.IncludedPolicyTypes);

            if (String.IsNullOrEmpty(ids)) return;
            var l = ids.Split(';');
            foreach (var id in l)
            {
                var found = MinorItems.FirstOrDefault(x => x.Id == id);
                if (found == null) continue;

                foreach (var no in tvaPolicies.AllNodes)
                {
                    if (String.Equals(no.Tag.ToString(), found.Title, StringComparison.OrdinalIgnoreCase))
                    {
                        if(!regenProcess)
                            _selectedQuestionnaireFragments.Add(found);


                        var path = tvaPolicies.GetPath(no);
                        var node = ((AdvancedTreeNode) path.LastNode);
                        node.CheckState = CheckState.Checked;
                        node.Checked = true;
                        if (no.Parent != null) no.Parent.ExpandAll();
                    }
                }
            }
        }

        private void ReloadAllFields(PreRenewalQuestionare template)
        {
            txtAssistantExecutiveEmail.Text = template.AssistantExecutiveEmail;
            txtAssistantExecutiveName.Text = template.AssistantExecutiveName;
            txtAssistantExecutivePhone.Text = template.AssistantExecutivePhone;
            txtAssistantExecutiveTitle.Text = template.AssistantExecutiveTitle;
            txtAssitantExecDepartment.Text = template.AssistantExecDepartment;

            txtBranchAddress1.Text = template.OAMPSBranchAddress;
            txtBranchAddress2.Text = template.OAMPSBranchAddressLine2;

            txtClaimExecDepartment.Text = template.ClaimsExecDepartment;
            txtClaimsExecutiveEmail.Text = template.ClaimsExecutiveEmail;
            txtClaimsExecutiveName.Text = template.ClaimsExecutiveName;
            txtClaimsExecutivePhone.Text = template.ClaimsExecutivePhone;
            txtClaimsExecutiveTitle.Text = template.ClaimsExecutiveTitle;

            txtClientCommonName.Text = template.ClientCommonName;
            txtClientName.Text = template.ClientName;

            txtExecutiveDepartment.Text = template.ExecutiveDepartment;
            txtExecutiveEmail.Text = template.ExecutiveEmail;
            txtExecutiveName.Text = template.ExecutiveName;
            txtExecutivePhone.Text = template.ExecutivePhone;
            txtExecutiveTitle.Text = template.ExecutiveTitle;
            
            txtOtherExecDepartment.Text = template.OtherExecDepartment;
            txtOtherContactEmail.Text = template.OtherContactEmail;
            txtOtherContactName.Text = template.OtherContactName;
            txtOtherContactPhone.Text = template.OtherContactPhone;
            txtOtherContactTitle.Text = template.OtherContactTitle;

        }

        private void dtpPeriodOfInsuranceFrom_ValueChanged(object sender, EventArgs e)
        {
             dtpPeriodOfInsuranceTo.Value = dtpPeriodOfInsuranceFrom.Value.AddMonths(12);
        }

    }
}
