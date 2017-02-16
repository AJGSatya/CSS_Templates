using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Runtime.Caching;
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
using OAMPS.Office.BusinessLogic.Presenters.SharePoint;
using OAMPS.Office.BusinessLogic.Presenters.Wizards;
using OAMPS.Office.Word.Helpers.Controls;
using OAMPS.Office.Word.Models.SharePoint;
using OAMPS.Office.Word.Models.Word;
using OAMPS.Office.Word.Properties;
using OAMPS.Office.Word.Views.Wizards.Popups;
using Enums = OAMPS.Office.Word.Helpers.Enums;

namespace OAMPS.Office.Word.Views.Wizards
{
    public partial class FactFinderWizard : BaseWizardForm
    {
        protected static List<IQuestionClass> Questions = null;
        private readonly FactFinderWizardPresenter _wizardPresenter;
        private bool _populateApprovalForm = true;
        private bool _populateClaimMadeWarning = true;
        private List<IPolicyClass> _selectedQuestionnaireFragments = new List<IPolicyClass>();
        private List<IQuestionClass> _selectedQuestions = new List<IQuestionClass>();
        public BusinessLogic.Helpers.Enums.Statutory SelectedStatutory;
        private string _storedStatutory;
        private Dictionary<string, DocumentFragment> _availableAttachments;

        public FactFinderWizard(OfficeDocument document, Enums.FormLoadType loadType)
        {
            InitializeComponent();
            tbcWizardScreens.SelectedIndexChanged += tbcWizardScreens_SelectedIndexChanged;
            _wizardPresenter = new FactFinderWizardPresenter(document, this);
            base.BaseWizardPresenter = _wizardPresenter;
            LoadType = loadType;

            _checked.CheckStateChanged += _checked_CheckStateChanged;

            clbQustions.ItemCheck += new ItemCheckEventHandler(clbQustions_ItemCheck);

            _name.DrawText += OnDrawPolicyNode;

            txtClientName.Validating += ClientNameValidating;
            //rdoWholesale.Validating += rdoRetailWholesale_Validating;
            //rdoWholesaleWithRetail.Validating += rdoRetailWholesale_Validating;
            //rdoRetailFSG.Validating += rdoRetailWholesale_Validating;
        }

        private void clbQustions_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            if (Reload && !WizardBeingUpdated)
            {
                if (ContinueWithSignificantChange(sender, true)) GenerateNewTemplate = true;
            }
        }

        //private void rdoRetailWholesale_Validating(object sender, CancelEventArgs e)
        //{
        //    if (rdoWholesale.Enabled == false) return;

        //    if (rdoRetailFSG.Checked == false && rdoWholesale.Checked == false &&
        //        rdoWholesaleWithRetail.Checked == false)
        //    {
        //        errorProvider.SetError(grpWholesaleRetail, "You must select from retail or wholsale");
        //        e.Cancel = true;
        //    }
        //    else
        //    {
        //        errorProvider.SetError(grpWholesaleRetail, string.Empty);
        //    }
        //}


        private void ClientNameValidating(object sender, CancelEventArgs e)
        {
            if (txtClientName.Enabled == false) return;

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


        private void GetFragements()
        {
            _availableAttachments = new Dictionary<string, DocumentFragment>();
            List<ISharePointListItem> fragments = null;
            if (Cache.Contains(Constants.CacheNames.QuoteSlipFragments))
            {
                fragments = (List<ISharePointListItem>)Cache.Get(Constants.CacheNames.QuoteSlipFragments);
            }
            else
            {
                var list = new SharePointList(Settings.Default.SharePointContextUrl, Settings.Default.GeneralFragmentsListName, Constants.SharePointQueries.FactFinderFragmentsByKey);
                var presenter = new SharePointListPresenter(list, this);
                fragments = presenter.GetItems();
            }


            foreach (var i in fragments)
            {
                string key = i.GetFieldValue("Key");
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
                                Url = Settings.Default.FragmentFSG,
                                Locked = i.GetFieldValue("Locked")
                            });

                            break;
                        }


                    case Constants.FragmentKeys.GeneralAdviceWarning:
                        {
                            _availableAttachments.Add(key, new DocumentFragment
                            {
                                Title = txtTitle,
                                Key = key,
                                Url = Settings.Default.FragmentWarning,
                                Locked = i.GetFieldValue("Locked")
                            });
                            break;
                        }


                    case Constants.FragmentKeys.PrivacyStatement:
                        {
                            _availableAttachments.Add(key, new DocumentFragment
                            {
                                Title = txtTitle,
                                Key = key,
                                Url = Settings.Default.FragmentPrivacy,
                                Locked = i.GetFieldValue("Locked")
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
                                Url = Settings.Default.FragmentStatutory,
                                Locked = i.GetFieldValue("Locked")
                            });

                            break;
                        }


                    case Constants.FragmentKeys.UninsuredRisksReviewList:
                        {
                            _availableAttachments.Add(key, new DocumentFragment
                            {
                                Title = txtTitle,
                                Key = key,
                                Url = Settings.Default.FragmentUninsuredRisks,
                                Locked = i.GetFieldValue("Locked")
                            });
                            break;
                        }
                }
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

        private FactFinder GenerateTempalteObject()
        {
            //buid the marketing template
            var template = new FactFinder
                {
                    DocumentTitle = BaseWizardPresenter.ReadDocumentProperty("Title"), //Constants.TemplateNames.InsuranceRenewalReport,
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
                    
                    OAMPSBranchAddress = txtBranchAddress1.Text,
                    OAMPSBranchAddressLine2 = txtBranchAddress2.Text,
                    DatePrepared = DateTime.Now.ToString(@"dd/MM/yyyy"),
                    SelectedDocumentFragments = _selectedQuestionnaireFragments,
                    PopulateApprovalForm = _populateApprovalForm,
                    PopulateClaimMadeWarning = _populateClaimMadeWarning,
                    Statutory = SelectedStatutory.ToString()
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
                    if (chkQuestionsOnly.Checked)
                    {
                        RunQuestionsOnly();
                        _wizardPresenter.CloseDocument(false,true);
                    }
                       
                    else
                        RunFullWizard();


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

        private void RunQuestionsOnly()
        {
            CheckedListBox.CheckedItemCollection c = clbQustions.CheckedItems;
            var items = c.Cast<IQuestionClass>().ToList();
            _wizardPresenter.GenerateQuestionsOnly(items, Settings.Default.BlankFragmentBaseOlly);
        }

        private void RunFullWizard()
        {
            Cursor = Cursors.WaitCursor;
            if (GenerateNewTemplate)
            {
                var template = GenerateTempalteObject();
                _selectedQuestions.Clear();

                var c = clbQustions.CheckedItems;
                var items = c.Cast<IQuestionClass>().ToList();
                _selectedQuestions.AddRange(items);

                template.SelectedQuestions = _selectedQuestions;

                Cache.Add(Constants.CacheNames.RegenerateTemplate, template, new CacheItemPolicy());

                _wizardPresenter.GenerateNewTemplate(Constants.CacheNames.RegenerateTemplate, Settings.Default.TemplatePrerenewalQuestionare);
            }
            else
            {
                //    //call presenter to populate
                var tempalte = GenerateTempalteObject();

                if (!Reload)
                {
                    _wizardPresenter.PopulateClaimMadeWarningFragment(_populateClaimMadeWarning, Settings.Default.FragmentPRQClaimsMadeWarning);
                    _wizardPresenter.PopulateApprovalFormFragment(_populateApprovalForm, Settings.Default.FragmentPRQApprovalForm);

                    var c = clbQustions.CheckedItems;
                    var items = c.Cast<IQuestionClass>().ToList();
                    _wizardPresenter.InsertQuestionnaireFragement(items, Settings.Default.InformationForPolicyFragment);
                    //_presenter.InsertQuestionnaireFragement(_selectedQuestionnaireFragments);


                    var attachmentFragments = new List<DocumentFragment>();

                    if(chkSatutory.Checked)
                        attachmentFragments.Add(_availableAttachments[Constants.FragmentKeys.StatutoryNotices]);

                    if(chkFSG.Checked)
                        attachmentFragments.Add(_availableAttachments[Constants.FragmentKeys.FinancialServicesGuide]);

                    if (attachmentFragments.Count > 0)
                    {
                        _wizardPresenter.PopulateImportantNotices(attachmentFragments, Settings.Default.FragmentStatutory,
                                                                  Settings.Default.FragmentPrivacy, Settings.Default.FragmentFSG,
                                                                  Settings.Default.FragmentTermsOfEngagement);
                    }
                }

                PopulateDocument(tempalte, lblCoverPageTitle.Text, lblLogoTitle.Text);


                if (!Reload)
                {
                    //    //thie information panel loads when a document is in sharePoint that has metadata
                    //    //clients don't wish to see this so force the close of the panel once the wizard completes.
                    _wizardPresenter.CloseInformationPanel();
                    _wizardPresenter.MoveToStartOfDocument();
                }

                //tracking
                LogUsage(tempalte,
                         LoadType == Enums.FormLoadType.RegenerateTemplate
                             ? Enums.UsageTrackingType.RegenerateDocument
                             : Enums.UsageTrackingType.NewDocument);
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
            txtExecutiveMobile.Text = peoplePicker.SelectedUser.Mobile;
            txtBranchAddress1.Text = peoplePicker.SelectedUser.BranchAddressLine1;
            txtBranchAddress2.Text = peoplePicker.SelectedUser.BranchAddressLine2;
            txtExecutiveDepartment.Text = peoplePicker.SelectedUser.Branch;
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

        private void PreRenewalQuestionareWizard_Load(object sender, EventArgs e)
        {
            tbcWizardScreens.TabPages.Remove(tabPolicyClass);
            txtClientName.Focus();
            txtClientName.Select();

            GetFragements();

            var auto = false;
            WizardBeingUpdated = true;

            TaskScheduler uiScheduler = TaskScheduler.FromCurrentSynchronizationContext();

            if (Cache.Contains(Constants.CacheNames.RegenerateTemplate))
            {
                Reload = false;

                var template = GetCachedTempalteObject<FactFinder>();
                lblCoverPageTitle.Text = template.CoverPageTitle;
                lblLogoTitle.Text = template.LogoTitle;

                _selectedQuestionnaireFragments = template.SelectedDocumentFragments;
                //  LoadDataSources(null); //dont use new thread

                ReloadAllFields(template);
                //ReloadPolicyClasses(true);
                _selectedQuestions = template.SelectedQuestions;

                LoadQuestions(null);
                RetickQuestionsOnReload(true);
                rdoWariningYes.Checked = template.PopulateClaimMadeWarning;
                rdoApprovalYes.Checked = template.PopulateApprovalForm;
                ReloadStatutory(true, template.Statutory);



                auto = true;
            }
            else
            {
                if (Reload)
                {
                    MessageBox.Show(@"You cannot make any changes to this document through the wizard", @"Cannot make any changes", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Close();

                    LoadDataSources(null); //dont use new thread

                    var template = new FactFinder();
                    template = (FactFinder) _wizardPresenter.LoadData(template);

                    ReloadAllFields(template);
                    // ReloadPolicyClasses(false);
                    LoadQuestions(null);

                    var claimsMade = _wizardPresenter.ReadDocumentProperty(Constants.WordDocumentProperties.ClaimsMadeWarning);
                    var approvalForms = _wizardPresenter.ReadDocumentProperty(Constants.WordDocumentProperties.ApprovalForm);
                    ReloadStatutory(false);

                    if (String.Equals(claimsMade, "true", StringComparison.OrdinalIgnoreCase)) rdoWariningYes.Checked = true;

                    if (String.Equals(approvalForms, "true", StringComparison.OrdinalIgnoreCase)) rdoApprovalYes.Checked = true;



                    RetickQuestionsOnReload(false);
                }
                else
                {
                    Task.Factory.StartNew(() =>
                        {
                            //  LoadDataSources(uiScheduler);
                            LoadQuestions(uiScheduler);
                        });
                }
            }

            base.LoadGenericImageTabs(uiScheduler, tbcWizardScreens, lblCoverPageTitle.Text, lblLogoTitle.Text);

            WizardBeingUpdated = false;

            if (auto) base.StartTimer();
        }

        private void ReloadStatutory(bool fromRegen, string statutory = "")
        {
            _storedStatutory = fromRegen ? statutory : _wizardPresenter.ReadDocumentProperty(Constants.WordDocumentProperties.StatutoryInformation);


            //if (string.Equals(_storedStatutory, BusinessLogic.Helpers.Enums.Statutory.Wholesale.ToString(), StringComparison.OrdinalIgnoreCase))
            //{
            //    rdoWholesale.Checked = true;
            //}
            //else if (string.Equals(_storedStatutory, BusinessLogic.Helpers.Enums.Statutory.Retail.ToString(),
            //                       StringComparison.OrdinalIgnoreCase))
            //{
            //    rdoRetailFSG.Checked = true;
            //}
            //else if (string.Equals(_storedStatutory, BusinessLogic.Helpers.Enums.Statutory.WholesaleWithRetail.ToString(),
            //                       StringComparison.OrdinalIgnoreCase))
            //{
            //    rdoWholesaleWithRetail.Checked = true;
            //}
        }

        private void RetickQuestionsOnReload(bool isRegenProcess)
        {
            var selectedPolicies = string.Empty;

            if (isRegenProcess)
            {
                selectedPolicies = _selectedQuestions.Aggregate(selectedPolicies, (current, i) => current + i.Id + ";");
            }
            else
            {
                selectedPolicies = _wizardPresenter.ReadDocumentProperty(Constants.WordDocumentProperties.IncludedPolicyTypes);
            }


            for (int i = 0; i < clbQustions.Items.Count; i++)
            {
                var item = clbQustions.Items[i] as IQuestionClass;

                if (item != null && selectedPolicies.Contains(item.Id))
                {
                    clbQustions.SetItemChecked(i, true);
                }

            }
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
                Cache.Add(Constants.CacheNames.PreRenewalQuestionareQuestions, Questions,
                          new CacheItemPolicy());
            }

            if (uiScheduler == null)
            {
                clbQustions.DataSource = Questions;
                clbQustions.DisplayMember = "Title";
                clbQustions.ValueMember = "Id";
            }
            else
            {
                Task.Factory.StartNew(() =>
                    {
                        clbQustions.DataSource = Questions;
                        clbQustions.DisplayMember = "Title";
                        clbQustions.ValueMember = "Id";
                    }, CancellationToken.None, TaskCreationOptions.None, uiScheduler);
            }
        }

        protected List<IQuestionClass> LoadQuestionsFromSharePoint(string contextUrl, string listName)
        {
            var list = new SharePointList(contextUrl, listName, Constants.SharePointQueries.AllItemsSortBySortOrder);
            var presenter = new SharePointListPresenter(list, this);
            return presenter.GetPreRenewalQuestionaireQuestions();
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

        //private void rdoWholesale_CheckedChanged(object sender, EventArgs e)
        //{
        //    SelectedStatutory = BusinessLogic.Helpers.Enums.Statutory.Wholesale;

        //    if (Reload && string.Equals(_storedStatutory, BusinessLogic.Helpers.Enums.Statutory.Wholesale.ToString(), StringComparison.OrdinalIgnoreCase))
        //    {
        //        if (rdoWholesale.Checked == false)
        //        {
        //            if (ContinueWithSignificantChange(sender, true)) GenerateNewTemplate = true;
        //        }
        //    }
        //}

        //private void rdoRetailFSG_CheckedChanged(object sender, EventArgs e)
        //{
        //    SelectedStatutory = BusinessLogic.Helpers.Enums.Statutory.Retail;

        //    if (Reload && string.Equals(_storedStatutory, BusinessLogic.Helpers.Enums.Statutory.Retail.ToString(), StringComparison.OrdinalIgnoreCase))
        //    {
        //        if (rdoRetailFSG.Checked == false)
        //        {
        //            if (ContinueWithSignificantChange(sender, true)) GenerateNewTemplate = true;
        //        }
        //    }
        //}

        //private void rdoWholesaleWithRetail_CheckedChanged(object sender, EventArgs e)
        //{
        //    SelectedStatutory = BusinessLogic.Helpers.Enums.Statutory.WholesaleWithRetail;

        //    if (Reload && string.Equals(_storedStatutory, BusinessLogic.Helpers.Enums.Statutory.WholesaleWithRetail.ToString(), StringComparison.OrdinalIgnoreCase))
        //    {
        //        if (rdoWholesaleWithRetail.Checked == false)
        //        {
        //            if (ContinueWithSignificantChange(sender, true)) GenerateNewTemplate = true;
        //        }
        //    }
        //}

        private void chkQuestionsOnly_CheckedChanged(object sender, EventArgs e)
        {
            groupBox1.Enabled = !chkQuestionsOnly.Checked;
            groupBox4.Enabled = !chkQuestionsOnly.Checked;
            groupBox7.Enabled = !chkQuestionsOnly.Checked;
            grpWholesaleRetail.Enabled = !chkQuestionsOnly.Checked;
            tabAccountExecutive.Enabled = !chkQuestionsOnly.Checked;
            tabPolicyClass.Enabled = !chkQuestionsOnly.Checked;
        }

        private void dtpPeriodOfInsuranceTo_ValueChanged(object sender, EventArgs e)
        {

        }

 

      
    }
}