using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using OAMPS.Office.BusinessLogic.Helpers;
using OAMPS.Office.Word.Helpers;
using OAMPS.Office.Word.Models.Word;
using OAMPS.Office.Word.Properties;
using Task = System.Threading.Tasks.Task;

namespace OAMPS.Office.Word.Views.Word
{
    public partial class TemplateUpdateChecker : Form
    {
        private readonly string _shouldHide;
        private readonly string _fragmentsUsedPropertyValue;
        private readonly string _documentId;
        public DateTime LatestChange { get; set; }
        public bool hide { get; set; }
        public TemplateUpdateChecker(string shouldHide, string fragmentsUsedPropertyValue, string documentId)
        {
            _shouldHide = shouldHide;
            _fragmentsUsedPropertyValue = fragmentsUsedPropertyValue;
            _documentId = documentId;
            InitializeComponent();
        }

        private void TemplateUpdateChecker_Load(object sender, EventArgs e)
        {

        }

        public bool CheckDocument(string shouldHide, string fragmentsUsedPropertyValue, string themes, string logos, string mainTemplate)
        {
            var showMeFrag = false;
            var showMeTheme= false;
            var showMeLogo = false;
            var showMeMainTemplate = false;

            try
            {
                
                    showMeFrag = CheckFramements(fragmentsUsedPropertyValue);
                    showMeLogo = CheckLogo(logos);
                    showMeTheme = CheckTheme(themes);
                    showMeMainTemplate = CheckMainTemplate(mainTemplate);
                
            }
            catch
            {

            }
            return (showMeFrag || showMeLogo || showMeTheme || showMeMainTemplate);
        }

        private bool CheckMainTemplate(string mainTemplate)
        {
            if (!String.IsNullOrEmpty(mainTemplate))
            {
                var mainTemplateInfo = mainTemplate.Split(';');
                var name = mainTemplateInfo[0];
                var itemList = new Models.SharePoint.SharePointList(Settings.Default.SharePointContextUrl, Settings.Default.MainTemplatesLibraryName);
                var itemItem = itemList.GetListItemByTitle(name);
                var itemItemModifiedDate = itemItem.GetFieldValue("Modified");
                var releaseNotes = itemItem.GetFieldValue("ReleaseNotes");
                
                DateTime logoUsedDateAsDate;
                DateTime spModifiedDateAsDate;

                DateTime.TryParse(mainTemplateInfo[1], out logoUsedDateAsDate);
                DateTime.TryParse(itemItemModifiedDate, out spModifiedDateAsDate);

                var usedCompateDate = new DateTime(logoUsedDateAsDate.Year, logoUsedDateAsDate.Month, logoUsedDateAsDate.Day);
                var modifiedCompateDate = new DateTime(spModifiedDateAsDate.Year, spModifiedDateAsDate.Month, spModifiedDateAsDate.Day);

                if (usedCompateDate <= modifiedCompateDate)
                {
                    if (LatestChange < modifiedCompateDate)
                        LatestChange = modifiedCompateDate;

                    dgUpdatedItems.Rows.Add();
                    dgUpdatedItems[0, (dgUpdatedItems.Rows.Count - 1)].Value = name + " Main Template";
                    dgUpdatedItems[1, (dgUpdatedItems.Rows.Count - 1)].Value = releaseNotes;
                    dgUpdatedItems[2, (dgUpdatedItems.Rows.Count - 1)].Value = logoUsedDateAsDate.ToString("dd/MM/yyyy");
                    dgUpdatedItems[3, (dgUpdatedItems.Rows.Count - 1)].Value = spModifiedDateAsDate.ToString("dd/MM/yyyy");

                    return true;
                }

            }

            return false;
        }

        private bool CheckTheme(string themes)
        {

            if (!String.IsNullOrEmpty(themes))
            {
                var themeInfo = themes.Split(';');
                var name = themeInfo[0].Substring(themeInfo[0].LastIndexOf("/") + 1, themeInfo[0].Length - themeInfo[0].LastIndexOf("/") - 1);
                var themeList = new Models.SharePoint.SharePointList(Settings.Default.SharePointContextUrl, Settings.Default.GraphicsPictureLibraryTitle);
                var themeItem = themeList.GetListItemByName(name);
                var themeItemModifiedDate = themeItem.GetFieldValue("Modified");
                var title = themeItem.GetFieldValue("Title");
                var releaseNotes = themeItem.GetFieldValue("ReleaseNotes");

                DateTime logoUsedDateAsDate;
                DateTime spModifiedDateAsDate;

                DateTime.TryParse(themeInfo[1], out logoUsedDateAsDate);
                DateTime.TryParse(themeItemModifiedDate, out spModifiedDateAsDate);

                var usedCompateDate = new DateTime(logoUsedDateAsDate.Year, logoUsedDateAsDate.Month, logoUsedDateAsDate.Day);
                var modifiedCompateDate = new DateTime(spModifiedDateAsDate.Year, spModifiedDateAsDate.Month, spModifiedDateAsDate.Day);

                if (usedCompateDate <= modifiedCompateDate)
                {
                    if (LatestChange < modifiedCompateDate)
                        LatestChange = modifiedCompateDate;


                    dgUpdatedItems.Rows.Add();
                    dgUpdatedItems[0, (dgUpdatedItems.Rows.Count - 1)].Value = title + " Theme";
                    dgUpdatedItems[1, (dgUpdatedItems.Rows.Count - 1)].Value = releaseNotes;
                    dgUpdatedItems[2, (dgUpdatedItems.Rows.Count - 1)].Value = logoUsedDateAsDate.ToString("dd/MM/yyyy");
                    dgUpdatedItems[3, (dgUpdatedItems.Rows.Count - 1)].Value = spModifiedDateAsDate.ToString("dd/MM/yyyy");

                    return true;
                }
            }

            return false;
        }

        private bool CheckLogo(string logos)
        {
            if (logos != null)
            {
                if (!String.IsNullOrEmpty(logos))
                {
                    var logoInfo = logos.Split(';');
                    var name = logoInfo[0].Substring(logoInfo[0].LastIndexOf("/") + 1, logoInfo[0].Length - logoInfo[0].LastIndexOf("/") - 1);
                    var logosList = new Models.SharePoint.SharePointList(Settings.Default.SharePointContextUrl, Settings.Default.GraphicsPictureLibraryTitleLogos);
                    var logoItem = logosList.GetListItemByName(name);
                    var logoItemModifiedDate = logoItem.GetFieldValue("Modified");
                    var title = logoItem.GetFieldValue("Title");
                    var releaseNotes = logoItem.GetFieldValue("ReleaseNotes");

                    DateTime logoUsedDateAsDate;
                    DateTime spModifiedDateAsDate;

                    DateTime.TryParse(logoInfo[1], out logoUsedDateAsDate);
                    DateTime.TryParse(logoItemModifiedDate, out spModifiedDateAsDate);

                    var usedCompateDate = new DateTime(logoUsedDateAsDate.Year, logoUsedDateAsDate.Month, logoUsedDateAsDate.Day);
                    var modifiedCompateDate = new DateTime(spModifiedDateAsDate.Year, spModifiedDateAsDate.Month, spModifiedDateAsDate.Day);
                    
                    if (usedCompateDate <= modifiedCompateDate) //all done in a day buddy, i named it compate deal with it.
                    {
                        if (LatestChange < modifiedCompateDate)
                            LatestChange = modifiedCompateDate;


                        dgUpdatedItems.Rows.Add();
                        dgUpdatedItems[0, (dgUpdatedItems.Rows.Count - 1)].Value = title + " Company Logo";
                        dgUpdatedItems[1, (dgUpdatedItems.Rows.Count - 1)].Value = releaseNotes;
                        dgUpdatedItems[2, (dgUpdatedItems.Rows.Count - 1)].Value = logoUsedDateAsDate.ToString("dd/MM/yyyy");
                        dgUpdatedItems[3, (dgUpdatedItems.Rows.Count - 1)].Value = spModifiedDateAsDate.ToString("dd/MM/yyyy");

                        return true;
                    }
                }
            }
            return false;
        }

        private bool CheckFramements(string fragmentsUsedPropertyValue)
        {
            var fragDates = fragmentsUsedPropertyValue.Split(';');
            var showMe = false;
            foreach (var usedFragment in fragDates)
            {
                if (String.IsNullOrEmpty(usedFragment)) continue;

                var usedFragmentInfo = usedFragment.Split(':');
                var listAbbrivation = usedFragmentInfo[0].Substring(0, 2);
                var itemId = usedFragmentInfo[0].Substring(2, usedFragmentInfo[0].Length - 2);
                var libraryName = FrameworkExtensions.GetListTitleByAbbrivation(listAbbrivation);
                var templateUsedDate = usedFragmentInfo[1];
                var spList = new Models.SharePoint.SharePointList(Settings.Default.SharePointContextUrl, libraryName);
                var spItem = spList.GetListItemById(itemId);

                var spModifiedDate = spItem.GetFieldValue("Modified");
                var title = spItem.GetFieldValue(Constants.SharePointFields.Title);
                var notes = spItem.GetFieldValue("ReleaseNotes");

                DateTime templateUsedDateAsDate;
                DateTime spModifiedDateAsDate;

                DateTime.TryParse(templateUsedDate, out templateUsedDateAsDate);
                DateTime.TryParse(spModifiedDate, out spModifiedDateAsDate);

                var usedComponent = new DateTime(templateUsedDateAsDate.Year, templateUsedDateAsDate.Month, templateUsedDateAsDate.Day);
                var modifiedCompateDate = new DateTime(spModifiedDateAsDate.Year, spModifiedDateAsDate.Month, spModifiedDateAsDate.Day);

                if (usedComponent <= modifiedCompateDate)
                {
                    if (LatestChange < modifiedCompateDate)
                        LatestChange = modifiedCompateDate;


                    showMe = true;
                    dgUpdatedItems.Rows.Add();
                    dgUpdatedItems[0, (dgUpdatedItems.Rows.Count - 1)].Value = title;
                    dgUpdatedItems[1, (dgUpdatedItems.Rows.Count - 1)].Value = notes;
                    dgUpdatedItems[2, (dgUpdatedItems.Rows.Count - 1)].Value = templateUsedDateAsDate.ToString("dd/MM/yyyy");
                    dgUpdatedItems[3, (dgUpdatedItems.Rows.Count - 1)].Value = spModifiedDateAsDate.ToString("dd/MM/yyyy");
                    
                }

            }
            return showMe;
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            var spList = new Models.SharePoint.SharePointList(Settings.Default.SharePointContextUrl, "TemplateUpdateTracking");

            var item = spList.GetListItemByTitle(_documentId);
            if (item == null)
            {
                var values = new Dictionary<string, string>
                    {
                        {"Title", _documentId},
                        {"Hide", chkHide.Checked.ToString()}
                    };
                spList.AddItem(values);
            }
            else
            {
                var id = item.GetFieldValue("ID");
                spList.UpdateField(id, "Hide", chkHide.Checked.ToString());
            }

            Close();
        }

        private void chkHide_CheckedChanged(object sender, EventArgs e)
        {
            hide = chkHide.Checked;
        }
    }
}
