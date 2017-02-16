using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.Caching;
using System.Windows.Forms;
using Aga.Controls.Tree;
using OAMPS.Office.BusinessLogic.Helpers;
using OAMPS.Office.BusinessLogic.Interfaces.Template;
using OAMPS.Office.BusinessLogic.Interfaces.Wizards;
using OAMPS.Office.BusinessLogic.Models.Template;
using OAMPS.Office.BusinessLogic.Models.Wizards;
using OAMPS.Office.BusinessLogic.Presenters;
using OAMPS.Office.Word.Helpers.Controls;
using OAMPS.Office.Word.Models.Word;
using System.Threading.Tasks;
using OAMPS.Office.Word.Properties;
using Aga.Controls.Tree.NodeControls;

namespace OAMPS.Office.Word.Views.Wizards
{
    public partial class InsuranceRenewalReportWizard : BaseForm
    {
        private List<IPolicyClass> _selectedDocumentFragments = new List<IPolicyClass>();
        //private readonly List<IPolicyClass> _unslectedDocumentFragements = new List<IPolicyClass>();
        private readonly InsuranceRenewealReportPresenter _presenter;
        private Enums.Segment _selectedSegment;
        private Enums.Remuneration _selectedRemuneration;
        private Enums.Statutory _selectedStatutory;
        private bool _populateUFI;
        private string _storedSegment;
        private string _storedStatutory;
        private string _storedRemuneration;
        private string _storedUFI;


        public InsuranceRenewalReportWizard(OfficeDocument document, Helpers.Enums.FormLoadType loadType)
        {
            InitializeComponent();

            tbcWizardScreens.SelectedIndexChanged += new EventHandler(tbcWizardScreens_SelectedIndexChanged);
            _checked.CheckStateChanged += new EventHandler<TreePathEventArgs>(_checked_CheckStateChanged);
            txtClientName.Validating += txtClientName_Validating;
            txtClientCommonName.Validating += txtClientCommonName_Validating;
            rdoSegment2.Validating += rdoSegment_Validating;
            rdoSegment3.Validating += rdoSegment_Validating;
            rdoSegment4.Validating += rdoSegment_Validating;
            rdoSegment5.Validating += rdoSegment_Validating;
            rdoWholesale.Validating += rdoRetailWholesale_Validating;
            rdoWholesaleWithRetail.Validating += rdoRetailWholesale_Validating;
            rdoRetailFSG.Validating += rdoRetailWholesale_Validating;
            rdoFeeOnly.Validating += rdoFee_Validating;
            rdoCombination.Validating += rdoFee_Validating;
            rdoCombination.Validating += rdoFee_Validating;
            rdoUFIYes.Validating += rdoUFI_Validating;
            rdoUFINo.Validating += rdoUFI_Validating;
            //send marketing template to the presenter

            _presenter = new InsuranceRenewealReportPresenter(document, this);
            base.BasePresenter = _presenter;
            _loadType = loadType;

            tvaPolicies.Resize += new EventHandler(tvaPolicies_Resize);
        }

        void tvaPolicies_Resize(object sender, EventArgs e)
        {
            tvcCurrent.Width = (tvaPolicies.Width / 5) + 150;
            tvcReccomended.Width = (tvaPolicies.Width / 5) + 150;
        }

        private void rdoUFI_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (rdoUFIYes.Checked == false && rdoUFINo.Checked == false)
            {
                errorProvider.SetError(rdoUFINo, "You must decide if you need a UFI");
                e.Cancel = true;
            }
            else
            {
                errorProvider.SetError(rdoUFINo, string.Empty);
            }
        }

        private void rdoFee_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (rdoFeeOnly.Checked == false && rdoCombination.Checked == false && rdoCommissionOnly.Checked == false)
            {
                errorProvider.SetError(grpFees, "You must select a fee");
                e.Cancel = true;
            }
            else
            {
                errorProvider.SetError(grpFees, string.Empty);
            }
        }

        private void rdoRetailWholesale_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (rdoRetailFSG.Checked == false && rdoWholesale.Checked == false &&
                rdoWholesaleWithRetail.Checked == false)
            {
                errorProvider.SetError(grpWholesaleRetail, "You must select from retail or wholsale");
                e.Cancel = true;
            }
            else
            {
                errorProvider.SetError(grpWholesaleRetail, string.Empty);
            }
        }

        private void rdoSegment_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (rdoSegment2.Checked == false && rdoSegment3.Checked == false && rdoSegment4.Checked == false &&
                rdoSegment5.Checked == false)
            {
                errorProvider.SetError(grpServicePlan, "You must select a segment");
                e.Cancel = true;
            }
            else
            {
                errorProvider.SetError(grpServicePlan, string.Empty);
            }
        }

        private void txtClientCommonName_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (String.IsNullOrEmpty(txtClientCommonName.Text))
            {
                errorProvider.SetError(txtClientCommonName, "Field is required");
                e.Cancel = true;
            }
            else
            {
                errorProvider.SetError(txtClientCommonName, string.Empty);
            }
        }

        private void txtClientName_Validating(object sender, System.ComponentModel.CancelEventArgs e)
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

            if (isChecked)
            {
                
                var parent = ((AdvancedTreeNode)e.Path.FirstNode);
                if (parent != null)
                {
                    var item = MinorItems.Find(i => i.Title == node.Text && i.MajorClass == parent.Text);
                    // var item  = fragementPresenter.
                    if (item != null)
                    {
                        //var d = _treeModel.Nodes[2];
                        var insurers = new Popups.Insurers();
                        this.TopMost = false;
                        //insurers.CurrentInsurers = _insurers;
                        insurers.ShowDialog();


                        if (insurers.RecommendedInsurer != null && insurers.CurrentInsurer != null)
                        {
                            item.CurrentInsurer = insurers.CurrentInsurer;
                            item.RecommendedInsurer = insurers.RecommendedInsurer;
                            item.CurrentInsurerId = insurers.CurrentInsurerId;
                            item.RecommendedInsurerId = insurers.RecommendedInsurerId;
                            _selectedDocumentFragments.Add(item);

                            node.Reccommended = insurers.RecommendedInsurer;
                            node.Current = insurers.CurrentInsurer;
                        }
                        else
                        {
                            node.Checked = false;
                        }

                    }
                }
                
            }
            else
            {
                var item = MinorItems.Find(i => i.Title == node.Text);
                // var item  = fragementPresenter.
                if (item != null)
                {
                    _selectedDocumentFragments.Remove(item);

                }
            }
        }

        private void InsuranceRenewalReport_Load(object sender, EventArgs e)
        {
            txtClientName.Focus();
            txtClientName.Select();

            if (_loadType == Helpers.Enums.FormLoadType.RegenerateTemplate)
            {
                RegenerateFormLoad();

                MessageBox.Show(BusinessLogic.Helpers.Constants.Miscellaneous.RegenerateOnLoadMsg, String.Empty,
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                //this occurs when a user changes data in a generated form that requires a new template to be created

                ////call presenter to populate
                //PopulateDocument();

                ////thie information panel loads when a document is in sharePoint that has metadata
                ////clients don't wish to see this so force the close of the panel once the wizard completes.
                //_presenter.CloseInformationPanel();

                //Close();

            }
            else
            {
                LoadForm(); //standard load either on file-> new in word or clicking the button on a generated form.
            }


        }

        private void RegenerateFormLoad()
        {
            // var uiScheduler = TaskScheduler.FromCurrentSynchronizationContext();
            // base.LoadGenericImageTabs(uiScheduler, tbcWizardScreens, lblCoverPageTitle.Text, lblLogoTitle.Text);

            //   var cacheName = _presenter.ReadDocumentProperty(Constants.CacheNames.RegenerateTemplate);
            if (Cache.Contains(Constants.CacheNames.RegenerateTemplate))
            {
                var template = (IInsuranceRenewalReport)Cache.Get(Constants.CacheNames.RegenerateTemplate);
                Cache.Remove(Constants.CacheNames.RegenerateTemplate);

                WizardBeingUpdated = true;
                //create document properties
                _presenter.CreateOrUpdateDocumentProperty(Constants.WordDocumentProperties.Segment, template.Segment.ToString());
                _presenter.CreateOrUpdateDocumentProperty(Constants.WordDocumentProperties.StatutoryInformation, template.Statutory.ToString());
                _presenter.CreateOrUpdateDocumentProperty(Constants.WordDocumentProperties.Remuneration, template.Remuneration.ToString());
                _presenter.CreateOrUpdateDocumentProperty(Constants.WordDocumentProperties.UFI, template.PopulateUFI.ToString());

                LoadDataSources(null);
                ReloadFields(template);
                ReloadPolicyClasses(template);
                ReloadSegments();
                ReloadStatutory();
                ReloadRemuneration();
                ReloadUFI();
                var uiScheduler = TaskScheduler.FromCurrentSynchronizationContext();
                base.LoadGenericImageTabs(uiScheduler, tbcWizardScreens, lblCoverPageTitle.Text, lblLogoTitle.Text);

                WizardBeingUpdated = false;
            }
        }

        private void LoadForm()
        {
            //var selectedCoverPage =  _presenter.ReadDocumentProperty(Constants.WordDocumentProperties.CoverPageTitle);
            //selectedCoverPage = (String.IsNullOrEmpty(selectedCoverPage) ? lblCoverPageTitle.Text : selectedCoverPage);

            //var selectedLogo = _presenter.ReadDocumentProperty(Constants.WordDocumentProperties.LogoTitle);
            //selectedLogo = (String.IsNullOrEmpty(selectedLogo) ? lblLogoTitle.Text : selectedLogo);

            var uiScheduler = TaskScheduler.FromCurrentSynchronizationContext();
            base.LoadGenericImageTabs(uiScheduler, tbcWizardScreens, lblCoverPageTitle.Text, lblLogoTitle.Text);

            if (Reload) // this happens if they click the button on the ribbon.
            {
                var template = new InsuranceRenewalReport();
                var values = _presenter.LoadData(template);
                var v = ((IInsuranceRenewalReport)values);

                LoadDataSources(null);
                ReloadFields(v);
                ReloadPolicyClasses(v);
                ReloadSegments();
                ReloadStatutory();
                ReloadRemuneration();
                ReloadUFI();
            }
            else  //new template
            {
                Task.Factory.StartNew(() => LoadDataSources(uiScheduler));
            }
        }

        private void LoadDataSources(TaskScheduler uiScheduler)
        {
            base.LoadTreeViewClasses(uiScheduler);
            //    _insurers = LoadInsurers();
        }

        //TODO MOVE TO PARENT
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

        private void ReloadFields(IInsuranceRenewalReport v)
        {
            DateTime outDate;
            txtClientName.Text = v.ClientName;
            txtClientCommonName.Text = v.ClientCommonName;
            txtExecutiveEmail.Text = v.ExecutiveEmail;
            txtExecutiveName.Text = v.ExecutiveName;
            txtExecutivePhone.Text = v.ExecutivePhone;
            txtExecutiveTitle.Text = v.ExecutiveTitle;
            txtExecutiveDepartment.Text = v.ExecutiveDepartment;

            //TODO: bug here as dates stored in document can be just year.  therefor they wont cast.
            dtpPeriodOfInsuranceFrom.Text = DateTime.TryParse(v.PeriodOfInsuranceFrom, out outDate)
                                                ? v.PeriodOfInsuranceFrom
                                                : String.Empty;
            dtpPeriodOfInsuranceTo.Text = DateTime.TryParse(v.PeriodOfInsuranceTo, out outDate)
                                              ? v.PeriodOfInsuranceTo
                                              : string.Empty;
            lblLogoTitle.Text = v.LogoTitle;
            lblCoverPageTitle.Text = v.CoverPageTitle;
            txtAssistantExecutiveName.Text = v.AssistantExecutiveName;
            txtAssistantExecutiveTitle.Text = v.AssistantExecutiveTitle;
            txtAssistantExecutivePhone.Text = v.AssistantExecutivePhone;
            txtAssistantExecutiveEmail.Text = v.AssistantExecutiveEmail;
            txtAssitantExecDepartment.Text = v.AssistantExecDepartment;

            txtClaimsExecutiveEmail.Text = v.ClaimsExecutiveEmail;
            txtClaimsExecutiveName.Text = v.ClaimsExecutiveName;
            txtClaimsExecutivePhone.Text = v.ClaimsExecutivePhone;
            txtClaimsExecutiveTitle.Text = v.ClaimsExecutiveTitle;
            txtClaimExecDepartment.Text = v.ClaimsExecDepartment;

            txtOtherContactEmail.Text = v.OtherContactEmail;
            txtOtherContactName.Text = v.OtherContactName;
            txtOtherContactPhone.Text = v.OtherContactPhone;
            txtOtherContactTitle.Text = v.OtherContactTitle;
            txtOtherExecDepartment.Text = v.OtherExecDepartment;

            lblCoverPageTitle.Text = v.CoverPageTitle;
            lblLogoTitle.Text = v.LogoTitle;
        }

        private void ReloadPolicyClasses(IInsuranceRenewalReport v)
        {
            if (v.SelectedDocumentFragments == null)
                v = _presenter.LoadIncludedPolicyClasses(v);

            //repopulate selected fields
            foreach (var f in v.SelectedDocumentFragments)
            {

                var found = MinorItems.FirstOrDefault(x => x.Id == f.Id);
                if (found != null)
                {
                    foreach (var no in tvaPolicies.AllNodes)
                    {

                        if (String.Equals(no.Tag.ToString(), found.MajorClass, StringComparison.OrdinalIgnoreCase))
                        {
                            foreach (var cno in no.Children)
                            {
                                if (String.Equals(cno.Tag.ToString(), found.Title, StringComparison.OrdinalIgnoreCase))
                                {

                                    var path = tvaPolicies.GetPath(cno);
                                    var node = ((AdvancedTreeNode)path.LastNode);
                                    node.CheckState = CheckState.Checked;
                                    node.Checked = true;
                                    no.Expand(false);

                                    //var item = MinorItems.FirstOrDefault(i => String.Equals(i.Title, found.Title, StringComparison.OrdinalIgnoreCase));
                                   
                                    node.Current = found.CurrentInsurer;
                                    node.Reccommended = found.RecommendedInsurer;

                                    _selectedDocumentFragments.Add(found);

                                }
                            }
                        }

                    }
                }

            }
        }

        private void ReloadRemuneration()
        {
            _storedRemuneration = _presenter.ReadDocumentProperty(Constants.WordDocumentProperties.Remuneration);
            if (string.Equals(_storedRemuneration, Enums.Remuneration.Combined.ToString(),
                              StringComparison.OrdinalIgnoreCase))
            {
                rdoCombination.Checked = true;
            }
            else if (string.Equals(_storedRemuneration, Enums.Remuneration.Commission.ToString(),
                                   StringComparison.OrdinalIgnoreCase))
            {
                rdoCommissionOnly.Checked = true;
            }
            else if (string.Equals(_storedRemuneration, Enums.Remuneration.Fee.ToString(),
                                   StringComparison.OrdinalIgnoreCase))
            {
                rdoFeeOnly.Checked = true;
            }
        }

        private void ReloadUFI()
        {
            _storedUFI = _presenter.ReadDocumentProperty(Constants.WordDocumentProperties.UFI);

            if (String.Equals(_storedUFI, "true", StringComparison.OrdinalIgnoreCase))
                rdoUFIYes.Checked = true;
            else
                rdoUFINo.Checked = true;
        }

        private void ReloadStatutory()
        {
            _storedStatutory = _presenter.ReadDocumentProperty(Constants.WordDocumentProperties.StatutoryInformation);
            if (string.Equals(_storedStatutory, Enums.Statutory.Wholesale.ToString(), StringComparison.OrdinalIgnoreCase))
            {
                rdoWholesale.Checked = true;
            }
            else if (string.Equals(_storedStatutory, Enums.Statutory.Retail.ToString(),
                                   StringComparison.OrdinalIgnoreCase))
            {
                rdoRetailFSG.Checked = true;
            }
            else if (string.Equals(_storedStatutory, Enums.Statutory.WholesaleWithResale.ToString(),
                                   StringComparison.OrdinalIgnoreCase))
            {
                rdoWholesaleWithRetail.Checked = true;
            }
        }

        private void ReloadSegments()
        {
            _storedSegment = _presenter.ReadDocumentProperty(Constants.WordDocumentProperties.Segment);
            if (string.Equals(_storedSegment, Enums.Segment.Two.ToString(), StringComparison.OrdinalIgnoreCase))
            {
                rdoSegment2.Checked = true;
            }
            else if (string.Equals(_storedSegment, Enums.Segment.Three.ToString(), StringComparison.OrdinalIgnoreCase))
            {
                rdoSegment3.Checked = true;
            }
            else if (string.Equals(_storedSegment, Enums.Segment.Four.ToString(), StringComparison.OrdinalIgnoreCase))
            {
                rdoSegment4.Checked = true;
            }
            else if (string.Equals(_storedSegment, Enums.Segment.Five.ToString(), StringComparison.OrdinalIgnoreCase))
            {
                rdoSegment5.Checked = true;
            }
        }

        private InsuranceRenewalReport GenerateTempalteObject()
        {
            //buid the marketing template
            var template = new InsuranceRenewalReport
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
                DatePrepared = DateTime.Now.ToString(CultureInfo.InvariantCulture),

                SelectedDocumentFragments = _selectedDocumentFragments
            };


            var baseTemplate = (BaseTemplate)template;
            var logoTab = tbcWizardScreens.TabPages[Constants.ControlNames.TabPageLogosName];
            //foreach (Control c in logoTab.Controls)
            //{
            //    if (c.GetType() == typeof(ValueRadioButton))
            //    {
            //        var v = ((ValueRadioButton)c);
            //        if (v.Checked)
            //        {
            //            template.LogoImageUrl = v.GetValue(Constants.RadioButtonValues.ImageUrl);
            //            template.LogoTitle = v.Text;
            //            template.OAMPSCompanyName = v.Text;
            //            template.OAMPSAfsl = v.GetValue(Constants.RadioButtonValues.AFSLKey);
            //            template.OAMPSAbnNumber = v.GetValue(Constants.RadioButtonValues.ABNKey);
            //            template.WebSite = v.GetValue(Constants.RadioButtonValues.WebsiteKey);
            //        }
            //    }
            //}

            PopulateLogosToTemplate(logoTab, ref baseTemplate);

            var covberTab =
                tbcWizardScreens.TabPages[Constants.ControlNames.TabPageCoverPagesName];

            PopulateCoversToTemplate(covberTab, ref baseTemplate);
            //foreach (Control c in covberTab.Controls)
            //{
            //    if (c.GetType() == typeof(ValueRadioButton))
            //    {
            //        var v = ((ValueRadioButton)c);
            //        if (v.Checked)
            //        {
            //            template.CoverPageImageUrl = v.GetValue(Constants.RadioButtonValues.ImageUrl);
            //            template.CoverPageTitle = v.Text;
            //        }
            //    }
            //}

            template.Segment = _selectedSegment;
            template.Remuneration = _selectedRemuneration;
            template.Statutory = _selectedStatutory;
            template.PopulateUFI = _populateUFI;

            return template;
        }

        private void PopulateDocument()
        {

            var template = GenerateTempalteObject();
            //change the graphics selected
            //if (Streams == null) return;
            _presenter.PopulateGraphics(template, lblCoverPageTitle.Text, lblLogoTitle.Text);

            if (_loadType == Helpers.Enums.FormLoadType.RibbonClick)
            {
                _presenter.PopulateData(template);
                return;
            }

            //popualte the basis of cover sections
            _presenter.PopulateBasisOfCover(_selectedDocumentFragments, Settings.Default.FragmentClassOfInsurance);



            var segmentDocuments = new Dictionary<Enums.Segment, string>
                {
                    {Enums.Segment.One, String.Empty},
                    {Enums.Segment.Two, Settings.Default.FragmentServicePlanSeg2},
                    {Enums.Segment.Three, Settings.Default.FragmentServicePlanSeg3},
                    {Enums.Segment.Four, Settings.Default.FragmentServicePlanSeg4},
                    {Enums.Segment.Five, Settings.Default.FragmentServicePlanSeg5},
                    {Enums.Segment.PersonalLines, String.Empty}
                };

            //remove basis of covers that have been unticked.
            // _presenter.RemoveBasisOfCover(_unslectedDocumentFragements);

            //populate service level segments
            _presenter.PopulateServiceLineAgrement(_selectedSegment, segmentDocuments, dtpPeriodOfInsuranceFrom.Value);

            _presenter.PopulatePurposeOfReport(_selectedSegment, Settings.Default.FragmentRRPurposeReport23, Settings.Default.FragmentRRPurposeReport45);

            //build remuneration documents 
            var remunerationDocuments = new List<DocumentFragment>
                {
                    new DocumentFragment
                        {
                            Title = Enums.Remuneration.Fee.ToString(),
                            Url = Settings.Default.FragementFeesRemuneration
                        },
                    new DocumentFragment
                        {
                            Title = Enums.Remuneration.Combined.ToString(),
                            Url = Settings.Default.FragementFeesCommission
                        },
                    new DocumentFragment
                        {
                            Title = Enums.Remuneration.Commission.ToString(),
                            Url = String.Empty
                        }
                };

            _presenter.PopulateRemuneration(_selectedRemuneration, remunerationDocuments);

            _presenter.PopulateExecutiveSummary(_selectedRemuneration, Settings.Default.FragmentRREexSumFeeCommission, Settings.Default.FragmentRREexSumFeeCombine);

            _presenter.PopulateImportantNotices(_selectedStatutory, Settings.Default.FragmentStatutory, Settings.Default.FragmentPrivacy, Settings.Default.FragmentFSG, Settings.Default.FragmentTermsOfEngagement);
            _presenter.PopulateUFI(_populateUFI, Settings.Default.FragmentUFI);

            _presenter.PopulatePremiumSummary(_selectedDocumentFragments);

            //try
            //{
            //    _presenter.SendUFIMessage(_populateUFI, Settings.Default.EmailAddressUFITo,Settings.Default.EmailAddressUFIFrom);
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);

            //  }

            //populate the content controls
            //populate data should be called last, as it ensures any inserted fragments get their controls populated.
            _presenter.PopulateData(template);

            _presenter.MoveToStartOfDocument();


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

                if (_generateNewTemplate)
                {
                    var template = GenerateTempalteObject();

                    Cache.Add(Constants.CacheNames.RegenerateTemplate, template, new CacheItemPolicy());

                    _presenter.GenerateNewTemplate(Constants.CacheNames.RegenerateTemplate, Settings.Default.TemplateInsuranceRenewalReport);
                }
                else
                {
                    //call presenter to populate
                    PopulateDocument();

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
            var peoplePicker = new Popups.PeoplePicker { txtFind = { Text = txtExecutiveName.Text } };
            this.TopMost = false;
            peoplePicker.ShowDialog();
            if (peoplePicker.SelectedUser == null) return;

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
            var peoplePicker = new Popups.PeoplePicker { txtFind = { Text = txtAssistantExecutiveName.Text } };
            this.TopMost = false;
            peoplePicker.ShowDialog();
            if (peoplePicker.SelectedUser == null) return;

            txtAssistantExecutiveName.Text = peoplePicker.SelectedUser.DisplayName;
            txtAssistantExecutiveTitle.Text = peoplePicker.SelectedUser.Title;
            txtAssistantExecutiveEmail.Text = peoplePicker.SelectedUser.EmailAddress;
            txtAssistantExecutivePhone.Text = peoplePicker.SelectedUser.VoiceTelephoneNumber;
            txtAssitantExecDepartment.Text = peoplePicker.SelectedUser.Branch;
        }

        private void btnClaimsExecutiveLookup_Click(object sender, EventArgs e)
        {
            var peoplePicker = new Popups.PeoplePicker { txtFind = { Text = txtClaimsExecutiveName.Text } };
            this.TopMost = false;
            peoplePicker.ShowDialog();
            if (peoplePicker.SelectedUser == null) return;

            txtClaimsExecutiveName.Text = peoplePicker.SelectedUser.DisplayName;
            txtClaimsExecutiveTitle.Text = peoplePicker.SelectedUser.Title;
            txtClaimsExecutiveEmail.Text = peoplePicker.SelectedUser.EmailAddress;
            txtClaimsExecutivePhone.Text = peoplePicker.SelectedUser.VoiceTelephoneNumber;
            txtClaimExecDepartment.Text = peoplePicker.SelectedUser.Branch;
        }

        private void btnOtherContactLookup_Click(object sender, EventArgs e)
        {
            var peoplePicker = new Popups.PeoplePicker { txtFind = { Text = txtOtherContactName.Text } };
            this.TopMost = false;
            peoplePicker.ShowDialog();
            if (peoplePicker.SelectedUser == null) return;

            txtOtherContactName.Text = peoplePicker.SelectedUser.DisplayName;
            txtOtherContactTitle.Text = peoplePicker.SelectedUser.Title;
            txtOtherContactEmail.Text = peoplePicker.SelectedUser.EmailAddress;
            txtOtherContactPhone.Text = peoplePicker.SelectedUser.VoiceTelephoneNumber;
            txtOtherExecDepartment.Text = peoplePicker.SelectedUser.Branch;
        }

        private void rdoSegment2_CheckedChanged(object sender, EventArgs e)
        {
            _selectedSegment = Enums.Segment.Two;

            if (Reload &&
                string.Equals(_storedSegment, Enums.Segment.Two.ToString(), StringComparison.OrdinalIgnoreCase))
            {
                if (rdoSegment2.Checked == false)
                {
                    if (ContinueWithSignificantChange(sender, true)) _generateNewTemplate = true;
                }
            }
        }

        //todo: check if the checked event always fires after the unchecked event
        private void rdoSegment3_CheckedChanged(object sender, EventArgs e)
        {
            _selectedSegment = Enums.Segment.Three;

            if (Reload &&
                string.Equals(_storedSegment, Enums.Segment.Three.ToString(), StringComparison.OrdinalIgnoreCase))
            {
                if (rdoSegment3.Checked == false)
                {
                    if (ContinueWithSignificantChange(sender, true)) _generateNewTemplate = true;
                }
            }
        }

        private void rdoSegment4_CheckedChanged(object sender, EventArgs e)
        {
            _selectedSegment = Enums.Segment.Four;

            if (Reload &&
                string.Equals(_storedSegment, Enums.Segment.Four.ToString(), StringComparison.OrdinalIgnoreCase))
            {
                if (rdoSegment4.Checked == false)
                {
                    if (ContinueWithSignificantChange(sender, true)) _generateNewTemplate = true;
                }
            }
        }

        private void rdoSegment5_CheckedChanged(object sender, EventArgs e)
        {
            _selectedSegment = Enums.Segment.Five;

            if (Reload &&
                string.Equals(_storedSegment, Enums.Segment.Five.ToString(), StringComparison.OrdinalIgnoreCase))
            {
                if (rdoSegment5.Checked == false)
                {
                    if (ContinueWithSignificantChange(sender, true)) _generateNewTemplate = true;
                }
            }
        }

        private void dtpPeriodOfInsuranceFrom_ValueChanged(object sender, EventArgs e)
        {
            dtpPeriodOfInsuranceTo.Value = dtpPeriodOfInsuranceFrom.Value.AddMonths(12);
        }

        private void tdoCommissionOnly_CheckedChanged(object sender, EventArgs e)
        {
            _selectedRemuneration = Enums.Remuneration.Commission;

            if (Reload &&
                string.Equals(_storedRemuneration, Enums.Remuneration.Commission.ToString(),
                              StringComparison.OrdinalIgnoreCase))
            {
                if (rdoCommissionOnly.Checked == false)
                {
                    if (ContinueWithSignificantChange(sender, true)) _generateNewTemplate = true;
                }
            }
        }

        private void rdoFeeOnly_CheckedChanged(object sender, EventArgs e)
        {
            _selectedRemuneration = Enums.Remuneration.Fee;

            if (Reload && string.Equals(_storedRemuneration, Enums.Remuneration.Fee.ToString(), StringComparison.OrdinalIgnoreCase))
            {
                if (rdoFeeOnly.Checked == false)
                {
                    if (ContinueWithSignificantChange(sender, true)) _generateNewTemplate = true;
                }
            }
        }

        private void rdoCombination_CheckedChanged(object sender, EventArgs e)
        {
            _selectedRemuneration = Enums.Remuneration.Combined;

            if (Reload && string.Equals(_storedRemuneration, Enums.Remuneration.Combined.ToString(), StringComparison.OrdinalIgnoreCase))
            {
                if (rdoCombination.Checked == false)
                {
                    if (ContinueWithSignificantChange(sender, true)) _generateNewTemplate = true;
                }
            }
        }

        private void rdoRetailFSG_CheckedChanged(object sender, EventArgs e)
        {
            _selectedStatutory = Enums.Statutory.Retail;

            if (Reload && string.Equals(_storedStatutory, Enums.Statutory.Retail.ToString(), StringComparison.OrdinalIgnoreCase))
            {
                if (rdoRetailFSG.Checked == false)
                {
                    if (ContinueWithSignificantChange(sender, true)) _generateNewTemplate = true;
                }
            }
        }

        private void rdoWholesaleWithRetail_CheckedChanged(object sender, EventArgs e)
        {
            _selectedStatutory = Enums.Statutory.WholesaleWithResale;

            if (Reload && string.Equals(_storedStatutory, Enums.Statutory.WholesaleWithResale.ToString(), StringComparison.OrdinalIgnoreCase))
            {
                if (rdoWholesaleWithRetail.Checked == false)
                {
                    if (ContinueWithSignificantChange(sender, true)) _generateNewTemplate = true;
                }
            }
        }

        private void rdoWholesale_CheckedChanged(object sender, EventArgs e)
        {
            _selectedStatutory = Enums.Statutory.Wholesale;

            if (Reload && string.Equals(_storedStatutory, Enums.Statutory.Wholesale.ToString(), StringComparison.OrdinalIgnoreCase))
            {
                if (rdoWholesale.Checked == false)
                {
                    if (ContinueWithSignificantChange(sender, true)) _generateNewTemplate = true;
                }
            }
        }

        private void rdoUFIYes_CheckedChanged(object sender, EventArgs e)
        {
            _populateUFI = true;
        }

        private void rdoUFINo_CheckedChanged(object sender, EventArgs e)
        {
            _populateUFI = false;
        }

        private void tvaPolicies_Click(object sender, EventArgs e)
        {

        }
    }
}
