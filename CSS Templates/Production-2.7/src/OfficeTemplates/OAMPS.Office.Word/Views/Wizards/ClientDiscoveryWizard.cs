using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.Caching;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Aga.Controls.Tree;
using OAMPS.Office.BusinessLogic.Helpers;
using OAMPS.Office.BusinessLogic.Interfaces.SharePoint;
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
    public partial class ClientDiscoveryWizard : BaseWizardForm
    {
        private readonly ClientDiscoveryWizardPresenter _wizardPresenter;

        private readonly List<IQuestionClass> _selectedQuestions;
        private List<IQuestionClass> _questions;

        public ClientDiscoveryWizard(OfficeDocument document, Enums.FormLoadType loadType)
        {
            InitializeComponent();

            tbcWizardScreens.SelectedIndexChanged += tbcWizardScreens_SelectedIndexChanged;

            _wizardPresenter = new ClientDiscoveryWizardPresenter(document, this);
            base.BaseWizardPresenter = _wizardPresenter;
            _checked.CheckStateChanged += _checked_CheckStateChanged;

            _questions = new List<IQuestionClass>();
            _selectedQuestions = new List<IQuestionClass>();
            LoadType = loadType;
        }

        private void _checked_CheckStateChanged(object sender, TreePathEventArgs e)
        {
            var node = ((AdvancedTreeNode) e.Path.LastNode);
            bool isChecked = node.Checked;
            //var item = _questions.FirstOrDefault(i => i.Title == node.Text);

            if (Reload && !WizardBeingUpdated)
            {
                bool previous = !isChecked;

                if (ContinueWithSignificantChange(sender, previous)) GenerateNewTemplate = true;
            }


            AutoSelectChildren(node, isChecked);
            TreeNodeAdv tvaNode = tvaQuestions.AllNodes.FirstOrDefault(x => x.ToString() == node.Text);
            if (tvaNode != null)
                tvaNode.ExpandAll();
        }

        private void AutoSelectChildren(AdvancedTreeNode node, bool value)
        {
            if (node != null && node.Checked != value)
            {
                node.Checked = value;
            }

            if (node != null && (node.Nodes != null && node.Nodes.Count > 0))
            {
                node.Nodes.ToList().ForEach(x => AutoSelectChildren((AdvancedTreeNode) x, value));
            }

            AddOrRemoveSelectedQuestion(node, value);
        }

        private void AddOrRemoveSelectedQuestion(AdvancedTreeNode node, bool value)
        {
            IQuestionClass item = _questions.FirstOrDefault(i => i.Title == node.Text);

            if (item != null)
            {
                if (value)
                {
                    if (!_selectedQuestions.Contains(item))
                    {
                        _selectedQuestions.Add(item);
                    }
                }
                else
                {
                    _selectedQuestions.Remove(item);
                }
            }
        }

        private void tbcWizardScreens_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnNext.Text = tbcWizardScreens.SelectedIndex == tbcWizardScreens.TabCount - 1 ? "&Finish" : "&Next";
            btnBack.Enabled = tbcWizardScreens.SelectedIndex != 0;
            btnNext.Enabled = btnNext.Text != @"&Finish" || LoadComplete;
        }

        private void OnLoad_ClientDiscoveryWizard(object sender, EventArgs e)
        {
            bool doAuto = false;

            txtClientName.Focus();
            txtClientName.Select();

            if (Cache.Contains(Constants.CacheNames.RegenerateTemplate))
            {
                doAuto = true;
            }


            TaskScheduler uiScheduler = TaskScheduler.FromCurrentSynchronizationContext();
            if (Reload)
            {
                LoadTreeViewClasses(null);
                LoadAllFields();
            }
            else
            {
                Task.Factory.StartNew(() => LoadTreeViewClasses(uiScheduler));
            }


            base.LoadGenericImageTabs(uiScheduler, tbcWizardScreens, lblCoverPageTitle.Text, lblLogoTitle.Text);

            if (doAuto) base.StartTimer();
        }

        private void LoadAllFields()
        {
            var template = new ClientDiscovery();
            if (Cache.Contains(Constants.CacheNames.RegenerateTemplate))
            {
                template = GetCachedTempalteObject<ClientDiscovery>();

                string[] keys = template.Questions.Select(x => x.Id).ToArray();
                foreach (string key in keys)
                {
                    LoadChildNode(key);
                }
            }
            else
            {
                template = (ClientDiscovery) _wizardPresenter.LoadData(template);
                string questions = _wizardPresenter.ReadSelectedQuestionsFromDocument();

                if (questions != null)
                {
                    string[] keys = questions.Split(';');

                    foreach (string key in keys)
                    {
                        LoadChildNode(key);
                    }
                }
            }

            //load all fields

            txtClientName.Text = template.ClientName;
            txtClientContactName.Text = template.ClientContactName;
            txtCltiBAIS.Text = template.ClientiBais;


            dateDiscussion.Text = template.DiscussionDate;

            txtExecutiveName.Text = template.ExecutiveName;
            txtExecutiveEmail.Text = template.ExecutiveEmail;

            txtExecutivePhone.Text = template.ExecutivePhone;
            txtExecutiveTitle.Text = template.ExecutiveTitle;

            txtExecutiveDepartment.Text = template.ExecutiveDepartment;

            txtBranchAddress1.Text = template.OAMPSBranchAddress;
            txtBranchAddress2.Text = template.OAMPSBranchAddressLine2;

            lblLogoTitle.Text = template.LogoTitle;
            lblCoverPageTitle.Text = template.CoverPageTitle;

            DateTime outDate;
            datePrepared.Value = DateTime.TryParse(template.DatePrepared, out outDate)
                                     ? outDate
                                     : DateTime.Today;
        }

        private void LoadChildNode(string key)
        {
            IQuestionClass found = _questions.FirstOrDefault(x => x.Id == key);

            if (found != null)
            {
                foreach (TreeNodeAdv no in tvaQuestions.AllNodes)
                {
                    string title = no.Tag.ToString();

                    if (String.Equals(title, found.Title, StringComparison.OrdinalIgnoreCase))
                    {
                        LoadChildNodeImpl(no, title);
                    }
                }
            }
        }


        private void LoadChildNodeImpl(TreeNodeAdv child, string title)
        {
            TreePath path = tvaQuestions.GetPath(child);

            var node = ((AdvancedTreeNode) path.LastNode);
            node.CheckState = CheckState.Checked;
            node.Checked = true;

            _selectedQuestions.Add(_questions.FirstOrDefault(x => x.Title == title));

            var parent = child.Parent;

            while (parent != null)
            {
                parent.Expand();
                parent = parent.Parent;
            }
        }

        protected override void LoadTreeViewClasses(TaskScheduler uiScheduler)
        {
            if (Cache.Contains(Constants.CacheNames.MajorQuestionClassItems))
            {
                MajorItems =
                    ((List<ISharePointListItem>)
                     Cache.Get(Constants.CacheNames.MajorQuestionClassItems));
            }
            else
            {
                MajorItems = LoadMajorTreeNodeTypes(Settings.Default.SharePointContextUrl, Settings.Default.TopLevelQuestionClassesListName);
                Cache.Add(Constants.CacheNames.MajorQuestionClassItems, MajorItems,
                          new CacheItemPolicy());
            }


            if (Cache.Contains(Constants.CacheNames.MinorQuestionClassItems))
            {
                _questions =
                    ((List<IQuestionClass>) Cache.Get(Constants.CacheNames.MinorQuestionClassItems));
            }
            else
            {
                _questions = LoadMinorQuestionTypes();
                Cache.Add(Constants.CacheNames.MinorQuestionClassItems, _questions,
                          new CacheItemPolicy());
            }

            if (uiScheduler == null)
                WriteQuestionsToTree();
            else
                Task.Factory.StartNew(WriteQuestionsToTree, CancellationToken.None, TaskCreationOptions.None, uiScheduler);
        }

        private List<IQuestionClass> LoadMinorQuestionTypes()
        {
            var list = new SharePointList(Settings.Default.SharePointContextUrl, Settings.Default.MinorQuestionClassesListName);
            var presenter = new SharePointListPresenter(list, this);
            return presenter.GetMinorQuestionItems();
        }


        private void WriteQuestionsToTree()
        {
            Control[] tv = Controls.Find("tvaQuestions", true);
            if (tv.Length == 0) return;
            var tvaTreeView = (TreeViewAdv) tv[0];

            tvaTreeView.Model = TreeModel;

            tvaTreeView.BeginUpdate();
            foreach (ISharePointListItem majorItem in MajorItems)
            {
                Node rootNode = AddRoot(majorItem.Title);

                List<IQuestionClass> found =
                    _questions.FindAll(
                        i =>
                        String.Equals(i.TopCategory, majorItem.Title, StringComparison.OrdinalIgnoreCase));

                //group them

                Dictionary<string, List<IQuestionClass>> sDictionary = found.GroupBy(x => x.SubCategory).ToDictionary(x => x.Key, x => x.ToList());
                //and render the 2nd level 
                foreach (var dict in sDictionary)
                {
                    if (!string.IsNullOrEmpty(dict.Key))
                    {
                        Node subCatNode = AddChild(rootNode, dict.Key);
                        foreach (IQuestionClass question in dict.Value)
                        {
                            Node childNode = AddChild(subCatNode, question.Title);
                        }
                    }
                    else
                    {
                        //if has parent 2nd level then go straight to the minor items
                        foreach (IQuestionClass question in dict.Value)
                        {
                            Node childNode = AddChild(rootNode, question.Title);
                        }
                    }
                }
            }
            tvaTreeView.EndUpdate();
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
                    RunPopulate();
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
                SwitchTab(tbcWizardScreens.SelectedIndex + 1);
            }
        }

        private void SwitchTab(int index)
        {
            tbcWizardScreens.SelectedIndex = index;
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            SwitchTab(tbcWizardScreens.SelectedIndex - 1);
        }

        private void RunPopulate()
        {
            var clientDiscovery = new ClientDiscovery
                {
                    ClientContactName = txtClientContactName.Text,
                    ClientName = txtClientName.Text,
                    ClientiBais = txtCltiBAIS.Text,
                    ExecutiveName = txtExecutiveName.Text,
                    ExecutiveEmail = txtExecutiveEmail.Text,
                    ExecutivePhone = txtExecutivePhone.Text,
                    ExecutiveTitle = txtExecutiveTitle.Text,
                    ExecutiveDepartment = txtExecutiveDepartment.Text,
                    DiscussionDate = dateDiscussion.Text,
                    OAMPSBranchAddress = txtBranchAddress1.Text,
                    OAMPSBranchAddressLine2 = txtBranchAddress2.Text,
                    Questions = _selectedQuestions,
                    DatePrepared = datePrepared.Text
                };

            var basetemplate = (BaseTemplate) clientDiscovery;
            //assign selected logo to the document
            TabPage logoTab = tbcWizardScreens.TabPages[Constants.ControlNames.TabPageLogosName];
            PopulateLogosToTemplate(logoTab, ref basetemplate);
            //assign selected cover page to the document
            TabPage covberTab = tbcWizardScreens.TabPages[Constants.ControlNames.TabPageCoverPagesName];
            PopulateCoversToTemplate(covberTab, ref basetemplate);

            if (GenerateNewTemplate)
            {
                Cache.Add(Constants.CacheNames.RegenerateTemplate, clientDiscovery, new CacheItemPolicy());
                _wizardPresenter.GenerateNewTemplate(Constants.CacheNames.RegenerateTemplate, Settings.Default.TemplateDiscoveryGuide);
            }
            else
            {
                //call presenter to populate
                PopulateDocument(clientDiscovery, lblCoverPageTitle.Text, lblLogoTitle.Text);
                if (LoadType != Enums.FormLoadType.RibbonClick)
                    _wizardPresenter.PopulateClientQuestions(_selectedQuestions);

                _wizardPresenter.MoveToStartOfDocument();
            }

            //thie information panel loads when a document is in sharePoint that has metadata
            //clients don't wish to see this so force the close of the panel once the wizard completes.
            //Presenter.CloseInformationPanel();

            //tracking
            if (GenerateNewTemplate)
            {
                LogUsage(clientDiscovery, Enums.UsageTrackingType.RegenerateDocument);
            }
            else if (Reload)
            {
                LogUsage(clientDiscovery, Enums.UsageTrackingType.UpdateData);
            }
            else
            {
                LogUsage(clientDiscovery, Enums.UsageTrackingType.NewDocument);
            }
            Close();
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
            txtBranchAddress2.Text = peoplePicker.SelectedUser.BranchAddressLine2;

            txtExecutiveDepartment.Text = peoplePicker.SelectedUser.Branch;
        }

        private void chkSelectAll_CheckedChanged(object sender, EventArgs e)
        {
            foreach (AdvancedTreeNode node in tvaQuestions.AllNodes.Select(n => tvaQuestions.GetPath(n)).Select(path => ((AdvancedTreeNode) path.LastNode)))
            {
                node.Checked = chkSelectAll.Checked;
                AddOrRemoveSelectedQuestion(node, chkSelectAll.Checked);
            }

            if (chkSelectAll.Checked)
                tvaQuestions.ExpandAll();
            else
                tvaQuestions.CollapseAll();
        }
    }
}