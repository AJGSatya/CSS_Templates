using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Runtime.Caching;
using System.Threading.Tasks;
using System.Windows.Forms;
using Aga.Controls.Tree;
using Aga.Controls.Tree.NodeControls;
using OAMPS.Office.BusinessLogic.Helpers;
using OAMPS.Office.BusinessLogic.Interfaces.Template;
using OAMPS.Office.BusinessLogic.Interfaces.Wizards;
using OAMPS.Office.BusinessLogic.Models.Template;
using OAMPS.Office.BusinessLogic.Models.Wizards;
using OAMPS.Office.BusinessLogic.Presenters;
using OAMPS.Office.BusinessLogic.Presenters.SharePoint;
using OAMPS.Office.Word.Helpers.Controls;
using OAMPS.Office.Word.Models.SharePoint;
using OAMPS.Office.Word.Models.Word;
using OAMPS.Office.Word.Properties;
using OAMPS.Office.Word.Views.Wizards.Popups;

namespace OAMPS.Office.Word.Views.Wizards
{
    public partial class InsuranceRenewalReportWizard : BaseWizardForm
    {
        //private readonly List<IPolicyClass> _unslectedDocumentFragements = new List<IPolicyClass>();
        private readonly InsuranceRenewealReportWizardPresenter _wizardPresenter;
        private readonly List<IPolicyClass> _selectedDocumentFragments = new List<IPolicyClass>();
        private bool _populateClientProfile;
        private bool _populateUFI;
        private Enums.Remuneration _selectedRemuneration;
        public Enums.Segment SelectedSegment;
        public Enums.Statutory SelectedStatutory;
        private string _storedClientProfile;
        private string _storedRemuneration;
        private string _storedSegment;
        private string _storedStatutory;
        private string _storedUFI;

        public InsuranceRenewalReportWizard(OfficeDocument document, Helpers.Enums.FormLoadType loadType)
        {
            InitializeComponent();

            tbcWizardScreens.SelectedIndexChanged += tbcWizardScreens_SelectedIndexChanged;
            _checked.CheckStateChanged += _checked_CheckStateChanged;
            _orderPolicy.ChangesApplied += _textbox_ValueChanged;

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

            _name.DrawText += _name_DrawText;
            //send marketing template to the presenter

            _wizardPresenter = new InsuranceRenewealReportWizardPresenter(document, this);
            base.BaseWizardPresenter = _wizardPresenter;
            LoadType = loadType;

            tvaPolicies.Resize += tvaPolicies_Resize;
        }


        private void _name_DrawText(object sender, DrawEventArgs e)
        {
            if (String.IsNullOrEmpty(txtFindPolicyClass.Text)) return;
            if (e.Node.Parent == null) return;

            e.TextColor = e.Text.ToUpper().Contains(txtFindPolicyClass.Text.ToUpper()) ? Color.Black : Color.Gray;
        }

        private void _orderPolicy_ChangesApplied(object sender, EventArgs e)
        {
        }

        private void tvaPolicies_Resize(object sender, EventArgs e)
        {
            tvcCurrent.Width = (tvaPolicies.Width/5) + 150;
            tvcReccomended.Width = (tvaPolicies.Width/5) + 150;
        }

        private void rdoUFI_Validating(object sender, CancelEventArgs e)
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

        private void rdoFee_Validating(object sender, CancelEventArgs e)
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

        private void rdoRetailWholesale_Validating(object sender, CancelEventArgs e)
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

        private void rdoSegment_Validating(object sender, CancelEventArgs e)
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

        private void txtClientCommonName_Validating(object sender, CancelEventArgs e)
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

        private void txtClientName_Validating(object sender, CancelEventArgs e)
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

        private void _textbox_ValueChanged(object sender, EventArgs e)
        {
            if (Reload && !WizardBeingUpdated)
            {
                if (ContinueWithSignificantChange(sender, false))
                {
                    GenerateNewTemplate = true;
                }
            }
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

            if (isChecked)
            {
                var parent = ((AdvancedTreeNode) e.Path.FirstNode);
                if (parent != null)
                {
                    var insurers = new Insurers();
                    TopMost = false;
                    Cursor = Cursors.WaitCursor;
                    insurers.ShowDialog();
                    Cursor = Cursors.Default;

                    if (insurers.SelectedReccomendedInsurers.Count == 0 || insurers.SelectedCurrentInsurers.Count == 0)
                    {
                        node.Checked = false;
                    }
                    else
                    {
                        if (insurers.SelectedReccomendedInsurers.Count > 1)
                        {
                            foreach (IInsurer r in insurers.SelectedReccomendedInsurers)
                            {
                                node.Reccommended += r.Title + Constants.Seperators.Spaceseperator + r.Percent + "%" + Constants.Seperators.Spaceseperator + Constants.Seperators.Lineseperator;
                                node.ReccommendedId += r.Id + Constants.Seperators.Lineseperator;
                            }
                        }
                        else
                        {
                            node.Reccommended += insurers.SelectedReccomendedInsurers[0].Title;
                            node.ReccommendedId += insurers.SelectedReccomendedInsurers[0].Id;
                        }

                        if (insurers.SelectedCurrentInsurers.Count > 1)
                        {
                            foreach (IInsurer c in insurers.SelectedCurrentInsurers)
                            {
                                node.Current += c.Title + Constants.Seperators.Spaceseperator + c.Percent + "%" + Constants.Seperators.Spaceseperator + Constants.Seperators.Lineseperator;
                                node.CurrentId += c.Id + Constants.Seperators.Lineseperator;
                            }
                        }
                        else
                        {
                            node.Current += insurers.SelectedCurrentInsurers[0].Title;
                            node.CurrentId += insurers.SelectedCurrentInsurers[0].Id;
                        }

                        //TreeNodeAdv maxOrder = null;
                        int maxNumber = 0;
// ReSharper disable LoopCanBeConvertedToQuery
                        foreach (TreeNodeAdv x in tvaPolicies.AllNodes)
// ReSharper restore LoopCanBeConvertedToQuery
                        {
                            object currentNumber = _orderPolicy.GetValue(x);
                            if (currentNumber == null) continue;
                            var intCurrentNumber = 0;
                            if (!int.TryParse(currentNumber.ToString(), out intCurrentNumber)) continue;
                            if (intCurrentNumber > maxNumber)
                            {
                                maxNumber = intCurrentNumber;
                            }
                        }
                        node.OrderPolicy = (maxNumber + 1).ToString(CultureInfo.InvariantCulture);
                    }
                }
            }
            else
            {
                node.Reccommended = string.Empty;
                node.ReccommendedId = string.Empty;
                node.Current = string.Empty;
                node.CurrentId = string.Empty;
                node.OrderPolicy = string.Empty;

                //var item = MinorItems.Find(i => i.Title == node.Text);
                //// var item  = fragementPresenter.
                //if (item != null)
                //{
                //    //    _selectedDocumentFragments.Remove(item);
                //}
            }
        }

        private void InsuranceRenewalReport_Load(object sender, EventArgs e)
        {
            txtClientName.Focus();
            txtClientName.Select();

            if (LoadType == Helpers.Enums.FormLoadType.RegenerateTemplate)
            {
                RegenerateFormLoad();
                lblLogoTitle.Text = @"OAMPS Insurance Brokers Ltd";
                lblCoverPageTitle.Text = @"Boulder Opal";

                //MessageBox.Show(BusinessLogic.Helpers.Constants.Miscellaneous.RegenerateOnLoadMsg, String.Empty,
                //                MessageBoxButtons.OK, MessageBoxIcon.Information);

                base.StartTimer();
            }
            else
            {
                _loadComplete = false;
                LoadForm(); //standard load either on file-> new in word or clicking the button on a generated form.
                _loadComplete = true;
            }
        }

        private void RegenerateFormLoad()
        {
            // var uiScheduler = TaskScheduler.FromCurrentSynchronizationContext();
            // base.LoadGenericImageTabs(uiScheduler, tbcWizardScreens, lblCoverPageTitle.Text, lblLogoTitle.Text);

            //   var cacheName = _presenter.ReadDocumentProperty(Constants.CacheNames.RegenerateTemplate);
            if (Cache.Contains(Constants.CacheNames.RegenerateTemplate))
            {
                var template = (IInsuranceRenewalReport) Cache.Get(Constants.CacheNames.RegenerateTemplate);
                Cache.Remove(Constants.CacheNames.RegenerateTemplate);

                _populateClientProfile = template.PopulateClientProfile;
                WizardBeingUpdated = true;
                //create document properties
                _wizardPresenter.CreateOrUpdateDocumentProperty(Constants.WordDocumentProperties.Segment, template.Segment.ToString());
                _wizardPresenter.CreateOrUpdateDocumentProperty(Constants.WordDocumentProperties.StatutoryInformation, template.Statutory.ToString());
                _wizardPresenter.CreateOrUpdateDocumentProperty(Constants.WordDocumentProperties.Remuneration, template.Remuneration.ToString());
                _wizardPresenter.CreateOrUpdateDocumentProperty(Constants.WordDocumentProperties.UFI, template.PopulateUFI.ToString());
                _wizardPresenter.CreateOrUpdateDocumentProperty(Constants.WordDocumentProperties.ClientProfile, template.PopulateClientProfile.ToString());

                LoadDataSources(null);
                ReloadFields(template);
                ReloadPolicyClasses(template, true);
                ReloadSegments();
                ReloadStatutory();
                ReloadRemuneration();
                ReloadUFI();
                ReloadClientProfile();
                TaskScheduler uiScheduler = TaskScheduler.FromCurrentSynchronizationContext();
                base.LoadGenericImageTabs(null, tbcWizardScreens, lblCoverPageTitle.Text, lblLogoTitle.Text);

                WizardBeingUpdated = false;
            }
        }

        private void LoadForm()
        {
            //var selectedCoverPage =  _presenter.ReadDocumentProperty(Constants.WordDocumentProperties.CoverPageTitle);
            //selectedCoverPage = (String.IsNullOrEmpty(selectedCoverPage) ? lblCoverPageTitle.Text : selectedCoverPage);

            //var selectedLogo = _presenter.ReadDocumentProperty(Constants.WordDocumentProperties.LogoTitle);
            //selectedLogo = (String.IsNullOrEmpty(selectedLogo) ? lblLogoTitle.Text : selectedLogo);

            TaskScheduler uiScheduler = TaskScheduler.FromCurrentSynchronizationContext();
            base.LoadGenericImageTabs(uiScheduler, tbcWizardScreens, lblCoverPageTitle.Text, lblLogoTitle.Text);

            if (Reload) // this happens if they click the button on the ribbon.
            {
                var template = new InsuranceRenewalReport();
                object values = _wizardPresenter.LoadData(template);
                var v = ((IInsuranceRenewalReport) values);

                LoadDataSources(null);
                ReloadFields(v);
                ReloadPolicyClasses(v, false);
                ReloadSegments();
                ReloadStatutory();
                ReloadRemuneration();
                ReloadUFI();
                ReloadClientProfile();
            }
            else //new template
            {
                Task.Factory.StartNew(() => LoadDataSources(uiScheduler));
            }
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
            txtExecutiveMobile.Text = v.ExecutiveMobile;

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

            txtBranchAddress1.Text = v.OAMPSBranchAddress;
            txtBranchAddress2.Text = v.OAMPSBranchAddressLine2;
        }

        private void ReloadPolicyClasses(IInsuranceRenewalReport v, bool reload)
        {
            if (v.SelectedDocumentFragments == null)
                v = _wizardPresenter.LoadIncludedPolicyClasses(v);
            int reOrder = 0;
            //repopulate selected fields
            v.SelectedDocumentFragments.Sort((x, y) => y.Order.CompareTo(x.Order));

            for (int index = v.SelectedDocumentFragments.Count - 1; index >= 0; index--)
            {
                IPolicyClass f = v.SelectedDocumentFragments[index];

                IPolicyClass found = MinorItems.FirstOrDefault(x => x.Id == f.Id);
                if (found != null)
                {
                    foreach (TreeNodeAdv no in tvaPolicies.AllNodes)
                    {
                        if (String.Equals(no.Tag.ToString(), found.MajorClass, StringComparison.OrdinalIgnoreCase))
                        {
                            foreach (TreeNodeAdv cno in no.Children)
                            {
                                if (String.Equals(cno.Tag.ToString(), found.Title, StringComparison.OrdinalIgnoreCase))
                                {
                                    reOrder = reOrder + 1;
                                    TreePath path = tvaPolicies.GetPath(cno);
                                    var node = ((AdvancedTreeNode) path.LastNode);
                                    node.CheckState = CheckState.Checked;
                                    node.Checked = true;
                                    no.Expand(false);

                                    //if (reload) //on generate get them from cache as they're passed thru
                                    //{
                                    //    node.Current = found.CurrentInsurer;
                                    //    node.Reccommended = found.RecommendedInsurer;
                                    //    node.OrderPolicy = reOrder.ToString();
                                    //    node.ReccommendedId = found.RecommendedInsurerId;
                                    //    node.CurrentId = found.CurrentInsurerId;

                                    //}
                                    //else
                                    //{
                                    node.Current = f.CurrentInsurer;
                                    node.Reccommended = f.RecommendedInsurer;
                                    node.OrderPolicy = reOrder.ToString(CultureInfo.InvariantCulture);
                                    node.ReccommendedId = f.RecommendedInsurerId;
                                    node.CurrentId = f.CurrentInsurerId;

                                    //}
                                }
                            }


                            //var item =
                            //    MinorItems.FirstOrDefault(
                            //        i => String.Equals(i.Title, f.Title, StringComparison.OrdinalIgnoreCase));
                            //if (item != null)
                            //{


                            //_selectedDocumentFragments.Add(item);    
                            //}
                        }
                    }
                }
            }
        }

        private void ReloadRemuneration()
        {
            _storedRemuneration = _wizardPresenter.ReadDocumentProperty(Constants.WordDocumentProperties.Remuneration);
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

        private void ReloadClientProfile()
        {
            _storedClientProfile = _wizardPresenter.ReadDocumentProperty(Constants.WordDocumentProperties.ClientProfile);

            if (String.Equals(_storedClientProfile, "true", StringComparison.OrdinalIgnoreCase))
            {
                _populateClientProfile = true;
                rdoClitProfileYes.Checked = true;
            }

            else
            {
                rdoClitProfileNo.Checked = true;
                _populateClientProfile = false;
            }
        }

        private void ReloadUFI()
        {
            _storedUFI = _wizardPresenter.ReadDocumentProperty(Constants.WordDocumentProperties.UFI);

            if (String.Equals(_storedUFI, "true", StringComparison.OrdinalIgnoreCase))
                rdoUFIYes.Checked = true;
            else
                rdoUFINo.Checked = true;
        }

        private void ReloadStatutory()
        {
            _storedStatutory = _wizardPresenter.ReadDocumentProperty(Constants.WordDocumentProperties.StatutoryInformation);
            if (string.Equals(_storedStatutory, Enums.Statutory.Wholesale.ToString(), StringComparison.OrdinalIgnoreCase))
            {
                rdoWholesale.Checked = true;
            }
            else if (string.Equals(_storedStatutory, Enums.Statutory.Retail.ToString(),
                                   StringComparison.OrdinalIgnoreCase))
            {
                rdoRetailFSG.Checked = true;
            }
            else if (string.Equals(_storedStatutory, Enums.Statutory.WholesaleWithRetail.ToString(),
                                   StringComparison.OrdinalIgnoreCase))
            {
                rdoWholesaleWithRetail.Checked = true;
            }
        }

        private void ReloadSegments()
        {
            _storedSegment = _wizardPresenter.ReadDocumentProperty(Constants.WordDocumentProperties.Segment);
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
            bool outProfile = false;
            Boolean.TryParse(BaseWizardPresenter.ReadDocumentProperty(Constants.WordDocumentProperties.ClientProfile), out outProfile);
            //buid the marketing template
            var template = new InsuranceRenewalReport
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
                    PopulateClientProfile = outProfile,
                    OtherContactName = txtOtherContactName.Text,
                    OtherContactTitle = txtOtherContactTitle.Text,
                    OtherContactPhone = txtOtherContactPhone.Text,
                    OtherContactEmail = txtOtherContactEmail.Text,
                    OtherExecDepartment = txtOtherExecDepartment.Text,
                    OAMPSBranchAddress = txtBranchAddress1.Text,
                    OAMPSBranchAddressLine2 = txtBranchAddress2.Text,
                    DatePrepared = DateTime.Now.ToString(@"dd/MM/yyyy"),
                    SelectedDocumentFragments = _selectedDocumentFragments
                };


            var baseTemplate = (BaseTemplate) template;
            TabPage logoTab = tbcWizardScreens.TabPages[Constants.ControlNames.TabPageLogosName];

            PopulateLogosToTemplate(logoTab, ref baseTemplate);

            TabPage covberTab =
                tbcWizardScreens.TabPages[Constants.ControlNames.TabPageCoverPagesName];

            PopulateCoversToTemplate(covberTab, ref baseTemplate);

            template.Segment = SelectedSegment;
            template.Remuneration = _selectedRemuneration;
            template.Statutory = SelectedStatutory;
            template.PopulateUFI = _populateUFI;
            template.PopulateClientProfile = _populateClientProfile;

            return template;
        }

        private void StoreSelectedPolicies()
        {
            //tvaPolicies.AllNodes.ToList().ForEach(
            //    (x) =>
            foreach (TreeNodeAdv x in tvaPolicies.AllNodes)
            {
                IEnumerable<NodeControlInfo> nodeControl = tvaPolicies.GetNodeControls(x);
                NodeControlInfo checkbox = nodeControl.FirstOrDefault(y => (y.Control is NodeCheckBox));

                //checkbox found
                var dCheckBox = (NodeCheckBox) checkbox.Control;
                if (dCheckBox != null)
                {
                    object value = dCheckBox.GetValue(x);

                    if ((bool) value)
                    {
                        NodeControlInfo policyClass =
                            nodeControl.FirstOrDefault(
                                y =>
                                (y.Control is NodeTextBox && y.Control.ParentColumn.Header == @"Policy Class"));
                        object policyClassValue = ((NodeTextBox) policyClass.Control).GetValue(x);

                        NodeControlInfo currentInsurer =
                            nodeControl.FirstOrDefault(
                                y =>
                                (y.Control is NodeTextBox && y.Control.ParentColumn.Header == @"Current Insurer"));
                        object currentInsurerValue = ((NodeTextBox) currentInsurer.Control).GetValue(x);

                        NodeControlInfo reccommendedInsurer =
                            nodeControl.FirstOrDefault(
                                y =>
                                (y.Control is NodeTextBox &&
                                 y.Control.ParentColumn.Header == @"Recommended Insurer"));
                        object reccommendedInsurerValue = ((NodeTextBox) reccommendedInsurer.Control).GetValue(x);

                        NodeControlInfo reccommendedInsurerId =
                            nodeControl.FirstOrDefault(
                                y =>
                                (y.Control is NodeTextBox && y.Control.ParentColumn.Header == @"RecommendedId"));
                        object reccommendedInsurerIdValue =
                            ((NodeTextBox) reccommendedInsurerId.Control).GetValue(x);

                        NodeControlInfo currentId =
                            nodeControl.FirstOrDefault(
                                y => (y.Control is NodeTextBox && y.Control.ParentColumn.Header == @"CurrentId"));
                        object currentIdValue = ((NodeTextBox) currentId.Control).GetValue(x);

                        NodeControlInfo order =
                            nodeControl.FirstOrDefault(
                                y => (y.Control is NodeTextBox && y.Control.ParentColumn.Header == @"Order"));
                        object orderValue = ((NodeTextBox) order.Control).GetValue(x);


                        //var item = MinorItems.Find(i => i.Title == policyClassValue.ToString());
                        IPolicyClass item = MinorItems.Find(i => i.Title == policyClassValue.ToString() && i.MajorClass == x.Parent.Tag.ToString());

                        if (item != null)
                        {
                            item.RecommendedInsurer = reccommendedInsurerValue.ToString();
                            item.CurrentInsurer = currentInsurerValue.ToString();

                            int outOrder = 0;
                            int.TryParse(orderValue.ToString(), out outOrder);
                            //if (!String.IsNullOrEmpty(orderValue.ToString()) && int.TryParse(orderValue.ToString(), out outOrder))
                            //{

                            //}

                            item.Order = outOrder;

                            if (currentIdValue != null) item.CurrentInsurerId = currentIdValue.ToString();
                            if (reccommendedInsurerIdValue != null) item.RecommendedInsurerId = reccommendedInsurerIdValue.ToString();
                            _selectedDocumentFragments.Add(item);
                        }
                    }
                }
            }
            //  );
        }

        private void PopulateDocument()
        {
            InsuranceRenewalReport template = GenerateTempalteObject();
            //change the graphics selected
            //if (Streams == null) return;
            _wizardPresenter.PopulateGraphics(template, lblCoverPageTitle.Text, lblLogoTitle.Text);

            if (LoadType == Helpers.Enums.FormLoadType.RibbonClick)
            {
                LogUsage(template, Helpers.Enums.UsageTrackingType.UpdateData);
                _wizardPresenter.PopulateData(template);
                return;
            }

            //popualte the basis of cover sections
            _wizardPresenter.PopulateBasisOfCover(_selectedDocumentFragments, Settings.Default.FragmentClassOfInsurance);

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
            _wizardPresenter.PopulateServiceLineAgrement(SelectedSegment, segmentDocuments, dtpPeriodOfInsuranceFrom.Value);

            _wizardPresenter.PopulatePurposeOfReport(SelectedSegment, Settings.Default.FragmentRRPurposeReport23,
                                               Settings.Default.FragmentRRPurposeReport45, Settings.Default.FragmentRRPurposeReportGeneric);

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


            _wizardPresenter.PopulateRemuneration(_selectedRemuneration, remunerationDocuments);

            _wizardPresenter.PopulateExecutiveSummary(_selectedRemuneration, Settings.Default.FragmentRREexSumFeeCommission,
                                                Settings.Default.FragmentRREexSumFeeCombine);

            _wizardPresenter.PopulateImportantNotices(SelectedStatutory, Settings.Default.FragmentStatutory,
                                                Settings.Default.FragmentPrivacy, Settings.Default.FragmentFSG,
                                                Settings.Default.FragmentTermsOfEngagement);
            _wizardPresenter.PopulateUFI(_populateUFI, Settings.Default.FragmentUFI);

            _wizardPresenter.PopulatePremiumSummary(_selectedDocumentFragments);

            if (rdoClitProfileYes.Checked)
                _wizardPresenter.PopulateclientProfile(Settings.Default.FragmentClientProfile);

           


            //populate the content controls
            //populate data should be called last, as it ensures any inserted fragments get their controls populated.
            _wizardPresenter.PopulateData(template);

            _wizardPresenter.MoveToStartOfDocument();


            //tracking
            LogUsage(template,
                     LoadType == Helpers.Enums.FormLoadType.RegenerateTemplate
                         ? Helpers.Enums.UsageTrackingType.RegenerateDocument
                         : Helpers.Enums.UsageTrackingType.NewDocument);

            

            if (_populateUFI)
            {
                this.TopMost = false; //this is important, as there is a chance that outlook will prompt the users and if our wizard is forced to the top and outlook is minimised they will not see the outlook dialog and processing will hang.
                Task.Factory.StartNew(() => SendUFI(template)).ContinueWith((task=>
                {
                    if(task.IsFaulted) OnError(task.Exception);
                }));
            }
        }

        private void SendUFI(InsuranceRenewalReport template)
        {
            var list = new SharePointList(Settings.Default.SharePointContextUrl, "Configuration", Constants.SharePointQueries.GetItemByTitleQuery); //todo: make setting for new config list
            var presenter = new SharePointListPresenter(list, null);

            var autoSendItem = presenter.GetItemByTitle("UFI.AutoSendEmail");
            var shouldAutoSend = autoSendItem.GetFieldValue("Value");

            if (String.Equals(shouldAutoSend, "no", StringComparison.OrdinalIgnoreCase))
                return;

            var item = presenter.GetItemByTitle("UFI.Recipients");
            var address = item.GetFieldValue("Value");
            if (String.IsNullOrEmpty(address)) return;
            _wizardPresenter.SendUFIMessage(template, address);
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
                    //   BasePresenter.SwitchScreenUpdating(false);
                    StoreSelectedPolicies();

                    if (GenerateNewTemplate)
                    {
                        InsuranceRenewalReport template = GenerateTempalteObject();

                        Cache.Add(Constants.CacheNames.RegenerateTemplate, template, new CacheItemPolicy());

                        _wizardPresenter.GenerateNewTemplate(Constants.CacheNames.RegenerateTemplate, Settings.Default.TemplateInsuranceRenewalReport);
                    }
                    else
                    {
                        //call presenter to populate
                        PopulateDocument();

                        //thie information panel loads when a document is in sharePoint that has metadata
                        //clients don't wish to see this so force the close of the panel once the wizard completes.
                        _wizardPresenter.CloseInformationPanel();
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

        private void btnClaimsExecutiveLookup_Click(object sender, EventArgs e)
        {
            var peoplePicker = new PeoplePicker(txtClaimsExecutiveName.Text, this);
            if (peoplePicker.SelectedUser == null)
            {
                TopMost = false;

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
            var peoplePicker = new PeoplePicker(txtOtherContactName.Text, this);
            if (peoplePicker.SelectedUser == null)
            {
                TopMost = false;

                peoplePicker.ShowDialog();
                if (peoplePicker.SelectedUser == null) return;
            }

            txtOtherContactName.Text = peoplePicker.SelectedUser.DisplayName;
            txtOtherContactTitle.Text = peoplePicker.SelectedUser.Title;
            txtOtherContactEmail.Text = peoplePicker.SelectedUser.EmailAddress;
            txtOtherContactPhone.Text = peoplePicker.SelectedUser.VoiceTelephoneNumber;
            txtOtherExecDepartment.Text = peoplePicker.SelectedUser.Branch;
        }

        private void rdoSegment2_CheckedChanged(object sender, EventArgs e)
        {
            SelectedSegment = Enums.Segment.Two;

            if (Reload &&
                string.Equals(_storedSegment, Enums.Segment.Two.ToString(), StringComparison.OrdinalIgnoreCase))
            {
                if (rdoSegment2.Checked == false)
                {
                    if (ContinueWithSignificantChange(sender, true)) GenerateNewTemplate = true;
                }
            }
        }

        //todo: check if the checked event always fires after the unchecked event
        private void rdoSegment3_CheckedChanged(object sender, EventArgs e)
        {
            SelectedSegment = Enums.Segment.Three;

            if (Reload &&
                string.Equals(_storedSegment, Enums.Segment.Three.ToString(), StringComparison.OrdinalIgnoreCase))
            {
                if (rdoSegment3.Checked == false)
                {
                    if (ContinueWithSignificantChange(sender, true)) GenerateNewTemplate = true;
                }
            }
        }

        private void rdoSegment4_CheckedChanged(object sender, EventArgs e)
        {
            SelectedSegment = Enums.Segment.Four;

            if (Reload &&
                string.Equals(_storedSegment, Enums.Segment.Four.ToString(), StringComparison.OrdinalIgnoreCase))
            {
                if (rdoSegment4.Checked == false)
                {
                    if (ContinueWithSignificantChange(sender, true)) GenerateNewTemplate = true;
                }
            }
        }

        private void rdoSegment5_CheckedChanged(object sender, EventArgs e)
        {
            SelectedSegment = Enums.Segment.Five;

            if (Reload &&
                string.Equals(_storedSegment, Enums.Segment.Five.ToString(), StringComparison.OrdinalIgnoreCase))
            {
                if (rdoSegment5.Checked == false)
                {
                    if (ContinueWithSignificantChange(sender, true)) GenerateNewTemplate = true;
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
                    if (ContinueWithSignificantChange(sender, true)) GenerateNewTemplate = true;
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
                    if (ContinueWithSignificantChange(sender, true)) GenerateNewTemplate = true;
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
                    if (ContinueWithSignificantChange(sender, true)) GenerateNewTemplate = true;
                }
            }
        }

        private void rdoRetailFSG_CheckedChanged(object sender, EventArgs e)
        {
            SelectedStatutory = Enums.Statutory.Retail;

            if (Reload && string.Equals(_storedStatutory, Enums.Statutory.Retail.ToString(), StringComparison.OrdinalIgnoreCase))
            {
                if (rdoRetailFSG.Checked == false)
                {
                    if (ContinueWithSignificantChange(sender, true)) GenerateNewTemplate = true;
                }
            }
        }

        private void rdoWholesaleWithRetail_CheckedChanged(object sender, EventArgs e)
        {
            SelectedStatutory = Enums.Statutory.WholesaleWithRetail;

            if (Reload && string.Equals(_storedStatutory, Enums.Statutory.WholesaleWithRetail.ToString(), StringComparison.OrdinalIgnoreCase))
            {
                if (rdoWholesaleWithRetail.Checked == false)
                {
                    if (ContinueWithSignificantChange(sender, true)) GenerateNewTemplate = true;
                }
            }
        }

        private void rdoWholesale_CheckedChanged(object sender, EventArgs e)
        {
            SelectedStatutory = Enums.Statutory.Wholesale;

            if (Reload && string.Equals(_storedStatutory, Enums.Statutory.Wholesale.ToString(), StringComparison.OrdinalIgnoreCase))
            {
                if (rdoWholesale.Checked == false)
                {
                    if (ContinueWithSignificantChange(sender, true)) GenerateNewTemplate = true;
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

        private void rdoClitProfileNo_CheckedChanged(object sender, EventArgs e)
        {
            if (!LoadComplete)
                return;

            if (Reload && !WizardBeingUpdated)
            {
                if (ContinueWithSignificantChange(sender, true)) GenerateNewTemplate = true;
            }
            _populateClientProfile = false;
        }

        private void rdoClitProfileYes_CheckedChanged(object sender, EventArgs e)
        {
            if (!LoadComplete)
                return;

            if (Reload && !WizardBeingUpdated)
            {
                if (ContinueWithSignificantChange(sender, true)) GenerateNewTemplate = true;
            }
            _populateClientProfile = true;
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

        private void label19_Click(object sender, EventArgs e)
        {
        }
    }
}