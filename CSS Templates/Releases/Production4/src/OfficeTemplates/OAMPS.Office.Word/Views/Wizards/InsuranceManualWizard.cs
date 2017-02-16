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
using OAMPS.Office.BusinessLogic.Presenters;
using OAMPS.Office.BusinessLogic.Presenters.SharePoint;
using OAMPS.Office.Word.Helpers.Controls;
using OAMPS.Office.Word.Models.SharePoint;
using OAMPS.Office.Word.Models.Word;
using OAMPS.Office.Word.Properties;
using OAMPS.Office.Word.Views.Wizards.Popups;
using Enums = OAMPS.Office.Word.Helpers.Enums;

namespace OAMPS.Office.Word.Views.Wizards
{
    public partial class InsuranceManualWizard : BaseWizardForm
    {
        //private readonly List<IPolicyClass> _unslectedDocumentFragements = new List<IPolicyClass>();
        private readonly InsuranceManualWizardPresenter _wizardPresenter;
        private readonly List<IPolicyClass> _selectedPolicyClasses = new List<IPolicyClass>();

        private bool _populateClientProfile;

        private string _storedClientProfile;


        public InsuranceManualWizard(OfficeDocument document, Enums.FormLoadType loadType)
        {
            InitializeComponent();

            tbcWizardScreens.SelectedIndexChanged += tbcWizardScreens_SelectedIndexChanged;
            _checked.CheckStateChanged += _checked_CheckStateChanged;
            _orderPolicy.ChangesApplied += _textbox_ValueChanged;

            txtClientName.Validating += txtClientName_Validating;
            txtClientCommonName.Validating += txtClientCommonName_Validating;


            _name.DrawText += _name_DrawText;
            //send marketing template to the presenter

            _wizardPresenter = new InsuranceManualWizardPresenter(document, this);
            base.BaseWizardPresenter = _wizardPresenter;
            LoadType = loadType;

            //tvaPolicies.Resize += new EventHandler(tvaPolicies_Resize);
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

        //void tvaPolicies_Resize(object sender, EventArgs e)
        //{
        //    tvcInsurer.Width = (tvaPolicies.Width / 5) + 150;
        //    tvcReccomended.Width = (tvaPolicies.Width / 5) + 150;
        //}


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

                    Control grpRecommended = insurers.Controls["grpRecommended"];
                    if (grpRecommended != null)
                    {
                        grpRecommended.Visible = false;
                        insurers.Width = (insurers.Width - grpRecommended.Width);

                        Control grpCurrent = insurers.Controls["grpCurrent"];
                        if (grpCurrent != null)
                        {
                            grpCurrent.Text = @"Insurer";
                        }
                    }

                    insurers.ShowDialog();
                    Cursor = Cursors.Default;

                    if (insurers.SelectedCurrentInsurers.Count == 0)
                    {
                        node.Checked = false;
                    }
                    else
                    {
                        if (insurers.SelectedCurrentInsurers.Count > 1)
                        {
                            foreach (IInsurer c in insurers.SelectedCurrentInsurers)
                            {
                                node.Insurer += c.Title + Constants.Seperators.Spaceseperator + c.Percent + "%" + Constants.Seperators.Spaceseperator + Constants.Seperators.Lineseperator;
                                node.InsurerId += c.Id + Constants.Seperators.Lineseperator;
                            }
                        }
                        else
                        {
                            node.Insurer += insurers.SelectedCurrentInsurers[0].Title;
                            node.InsurerId += insurers.SelectedCurrentInsurers[0].Id;
                        }

                        int maxNumber = 0;
// ReSharper disable LoopCanBeConvertedToQuery
                        foreach (TreeNodeAdv x in tvaPolicies.AllNodes)
// ReSharper restore LoopCanBeConvertedToQuery
                        {
                            object currentNumber = _orderPolicy.GetValue(x);
                            if (currentNumber == null) continue;
                            int intCurrentNumber = 0;
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
                node.Insurer = string.Empty;
                node.InsurerId = string.Empty;
                node.PolicyNumber = string.Empty;
                node.OrderPolicy = string.Empty;
            }
        }

        private void InsuranceRenewalReport_Load(object sender, EventArgs e)
        {
            txtClientName.Focus();
            txtClientName.Select();

            if (LoadType == Enums.FormLoadType.RegenerateTemplate)
            {
                RegenerateFormLoad();
                lblLogoTitle.Text = @"OAMPS Insurance Brokers Ltd";
                lblCoverPageTitle.Text = @"Boulder Opal";

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
                var template = (IInsuranceManual) Cache.Get(Constants.CacheNames.RegenerateTemplate);
                Cache.Remove(Constants.CacheNames.RegenerateTemplate);

                _populateClientProfile = template.PopulateClientProfile;
                WizardBeingUpdated = true;
                //create document properties
                _wizardPresenter.CreateOrUpdateDocumentProperty(Constants.WordDocumentProperties.ClientProfile, template.PopulateClientProfile.ToString());

                LoadDataSources(null);
                ReloadFields(template);
                ReloadPolicyClasses(template, true);

                ReloadClientProfile();
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
                var template = new InsuranceManual();
                object values = _wizardPresenter.LoadData(template);
                var v = ((IInsuranceManual) values);

                LoadDataSources(null);
                ReloadFields(v);
                ReloadPolicyClasses(v, false);

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

        private void ReloadFields(IInsuranceManual v)
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

        private void ReloadPolicyClasses(IInsuranceManual v, bool reload)
        {
            if (v.SelectedPolicyClasses == null)
                v = _wizardPresenter.LoadIncludedPolicyClasses(v);
            int reOrder = 0;
            //repopulate selected fields
            v.SelectedPolicyClasses.Sort((x, y) => y.Order.CompareTo(x.Order));

            for (int index = v.SelectedPolicyClasses.Count - 1; index >= 0; index--)
            {
                IPolicyClass f = v.SelectedPolicyClasses[index];

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


                            //_selectedPolicyClasses.Add(item);    
                            //}
                        }
                    }
                }
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


        private IInsuranceManual GenerateTempalteObject()
        {
            bool outProfile = false;
            Boolean.TryParse(BaseWizardPresenter.ReadDocumentProperty(Constants.WordDocumentProperties.ClientProfile), out outProfile);
            //buid the marketing template
            var template = new InsuranceManual
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
                    SelectedPolicyClasses = _selectedPolicyClasses
                };


            var baseTemplate = (BaseTemplate) template;
            TabPage logoTab = tbcWizardScreens.TabPages[Constants.ControlNames.TabPageLogosName];

            PopulateLogosToTemplate(logoTab, ref baseTemplate);

            TabPage covberTab =
                tbcWizardScreens.TabPages[Constants.ControlNames.TabPageCoverPagesName];

            PopulateCoversToTemplate(covberTab, ref baseTemplate);


            template.PopulateClientProfile = _populateClientProfile;

            return template;
        }

        private void StoreSelectedPolicies()
        {
            //tvaPolicies.AllNodes.ToList().ForEach(
            //    (x) =>
            foreach (var x in tvaPolicies.AllNodes)
            {
                var nodeControl = tvaPolicies.GetNodeControls(x);
                var checkbox = nodeControl.FirstOrDefault(y => (y.Control is NodeCheckBox));

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
                                (y.Control is NodeTextBox && y.Control.ParentColumn.Header == @"Insurer/s"));
                        object currentInsurerValue = ((NodeTextBox) currentInsurer.Control).GetValue(x);


                        NodeControlInfo currentId =
                            nodeControl.FirstOrDefault(
                                y => (y.Control is NodeTextBox && y.Control.ParentColumn.Header == @"InsurerId"));
                        object currentIdValue = ((NodeTextBox) currentId.Control).GetValue(x);

                        NodeControlInfo order =
                            nodeControl.FirstOrDefault(
                                y => (y.Control is NodeTextBox && y.Control.ParentColumn.Header == @"Order"));
                        object orderValue = ((NodeTextBox) order.Control).GetValue(x);


                        NodeControlInfo policyNumber =
                            nodeControl.FirstOrDefault(
                                y => (y.Control is NodeTextBox && y.Control.ParentColumn.Header == @"Policy Number"));
                        object policyNumberValue = ((NodeTextBox) policyNumber.Control).GetValue(x);

                        //var item = MinorItems.Find(i => i.Title == policyClassValue.ToString());
                        var item = MinorItems.Find(i => i.Title == policyClassValue.ToString() && i.MajorClass == x.Parent.Tag.ToString());

                        if (item != null)

                            if (currentInsurerValue != null) item.CurrentInsurer = currentInsurerValue.ToString();

                        var outOrder = 0;
                        int.TryParse(orderValue.ToString(), out outOrder);
                        //if (!String.IsNullOrEmpty(orderValue.ToString()) && int.TryParse(orderValue.ToString(), out outOrder))
                        //{

                        //}

                        if (item != null)
                        {
                            item.Order = outOrder;

                            if (policyNumberValue != null) item.PolicyNumber = policyNumberValue.ToString();

                            if (currentIdValue != null) item.CurrentInsurerId = currentIdValue.ToString();


                            //todo update listname to settings
                            if (policyClassValue != null)
                            {
                                var ps = LoadClaimsProcedure(Settings.Default.SharePointContextUrl, "Insurance Manual Claims Procedures", policyClassValue.ToString());
                                if (ps != null && ps.Url != null) item.FragmentPolicyUrl = ps.Url;
                            }
                            _selectedPolicyClasses.Add(item);
                        }
                    }
                }
            }

            if(_selectedPolicyClasses != null)
                _selectedPolicyClasses.Sort((x,y) => x.Order.CompareTo(y.Order));
        }


        protected IManualClaimsProcedure LoadClaimsProcedure(string contextUrl, string listName, string title)
        {
            var list = new SharePointList(contextUrl, listName, String.Format(Constants.SharePointQueries.GetItemByPolicyTypeQuery, title));
            var presenter = new SharePointListPresenter(list, this);
            List<IManualClaimsProcedure> t = presenter.GetManualClaimsProcedure();
            return (t != null && t.Count > 0) ? t[0] : null;
        }

        private void PopulateDocument()
        {
            IInsuranceManual template = GenerateTempalteObject();
            //change the graphics selected
            //if (Streams == null) return;
            _wizardPresenter.PopulateGraphics(template, lblCoverPageTitle.Text, lblLogoTitle.Text);

            if (LoadType == Enums.FormLoadType.RibbonClick)
            {
                LogUsage(template, Enums.UsageTrackingType.UpdateData);
                _wizardPresenter.PopulateData(template);
                return;
            }

            _wizardPresenter.PopulateImportantNotices();

            if (rdoClitProfileYes.Checked)
                _wizardPresenter.PopulateclientProfile(Settings.Default.FragmentClientProfile);

            _wizardPresenter.PopulatePolicyTable(template.SelectedPolicyClasses);

            _wizardPresenter.PopulateProgramSummarys(template.SelectedPolicyClasses, Settings.Default.InsuranceManualProgramSummary);

            _wizardPresenter.PopulateClaimsProcedures(template.SelectedPolicyClasses);

            _wizardPresenter.PopulateData(template);
            _wizardPresenter.MoveToStartOfDocument();

            //tracking
            LogUsage(template,
                     LoadType == Enums.FormLoadType.RegenerateTemplate
                         ? Enums.UsageTrackingType.RegenerateDocument
                         : Enums.UsageTrackingType.NewDocument);
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
                        IInsuranceManual template = GenerateTempalteObject();

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

        private void dtpPeriodOfInsuranceFrom_ValueChanged(object sender, EventArgs e)
        {
            dtpPeriodOfInsuranceTo.Value = dtpPeriodOfInsuranceFrom.Value.AddMonths(12);
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