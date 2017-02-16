using System;
using System.Threading.Tasks;
using OAMPS.Office.BusinessLogic.Helpers;
using OAMPS.Office.BusinessLogic.Interfaces.Template;
using OAMPS.Office.BusinessLogic.Models.Template;
using OAMPS.Office.BusinessLogic.Presenters.Wizards;
using OAMPS.Office.Word.Helpers;
using OAMPS.Office.Word.Models.Word;
using OAMPS.Office.Word.Properties;

namespace OAMPS.Office.Word.Views.Wizards
{
    public partial class ThemesOnlyWizard : BaseWizardForm
    {
        private readonly BaseWizardPresenter _wizardPresenter;

        public ThemesOnlyWizard()
        {
            InitializeComponent();
        }

        public ThemesOnlyWizard(OfficeDocument document)
        {
            InitializeComponent();

            tbcWizardScreens.SelectedIndexChanged += tbcWizardScreens_SelectedIndexChanged;

            _wizardPresenter = new BaseWizardPresenter(document, this);
            BaseWizardPresenter = _wizardPresenter;

            //ShouldUpdateTemplate(Settings.Default.TemplateLibraryName, "Multi Page Document.docx");
        }

        private void tbcWizardScreens_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnNext.Text = tbcWizardScreens.SelectedIndex == tbcWizardScreens.TabCount - 1 ? "&Finish" : "&Next";
            btnBack.Enabled = tbcWizardScreens.SelectedIndex != 0;
            btnNext.Enabled = btnNext.Text != @"&Finish" || LoadComplete;
        }

        private void ThemesOnlyWizard_Load(object sender, EventArgs e)
        {
            var uiScheduler = TaskScheduler.FromCurrentSynchronizationContext();
            //base.LoadCompanyLogoImagesTab(uiScheduler, tbcWizardScreens, lblLogoTitle.Text);
            LoadBrandingImagesTab(null, tbcWizardScreens, lblCoverPageTitle.Text, lblSpeciality.Text);
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            PopulateImages();
            Close();
        }


        private IBaseTemplate GenerateTempalteObject()
        {
            var baseTemplate = new BaseTemplate();
            var logoTab = tbcWizardScreens.TabPages[Constants.ControlNames.TabPageLogosName];

            PopulateLogosToTemplate(logoTab, ref baseTemplate);

            var covberTab =
                tbcWizardScreens.TabPages[Constants.ControlNames.TabPageCoverPagesName];

            PopulateCoversToTemplate(covberTab, ref baseTemplate);

            return baseTemplate;
        }

        private void PopulateImages()
        {
            var template = GenerateTempalteObject();
            //change the graphics selected
            //if (Streams == null) return;

            LogTemplateDetails(template);

            _wizardPresenter.PopulateGraphics(template, lblCoverPageTitle.Text, lblLogoTitle.Text);
            BaseWizardPresenter.CreateOrUpdateDocumentProperty(Constants.WordDocumentProperties.Speciality,
                lblSpeciality.Text);
        }

        private void SwitchTab(int index)
        {
            tbcWizardScreens.SelectedIndex = index;
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            SwitchTab(tbcWizardScreens.SelectedIndex - 1);
        }
    }
}