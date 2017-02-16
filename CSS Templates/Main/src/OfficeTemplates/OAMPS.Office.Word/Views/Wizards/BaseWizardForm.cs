using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.DirectoryServices.AccountManagement;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.Caching;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using System.Windows.Forms;
using Aga.Controls.Tree;
//using Microsoft.Office.Interop.Word;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json.Linq;
using OAMPS.Office.BusinessLogic.Helpers;
using OAMPS.Office.BusinessLogic.Interfaces;
using OAMPS.Office.BusinessLogic.Interfaces.SharePoint;
using OAMPS.Office.BusinessLogic.Interfaces.Template;
using OAMPS.Office.BusinessLogic.Interfaces.Wizards;
using OAMPS.Office.BusinessLogic.Loggers;
using OAMPS.Office.BusinessLogic.Models.Template;
using OAMPS.Office.BusinessLogic.Models.Wizards;
using OAMPS.Office.BusinessLogic.Presenters.SharePoint;
using OAMPS.Office.BusinessLogic.Presenters.Wizards;
using OAMPS.Office.Word.Helpers;
using OAMPS.Office.Word.Helpers.Controls;
using OAMPS.Office.Word.Helpers.LocalSharePoint;
using OAMPS.Office.Word.Models;
using OAMPS.Office.Word.Models.ActiveDirectory;
using OAMPS.Office.Word.Models.Local;
using OAMPS.Office.Word.Models.SharePoint;
using OAMPS.Office.Word.Properties;
using CheckBox = System.Windows.Forms.CheckBox;
using Enums = OAMPS.Office.Word.Helpers.Enums;
using Font = System.Drawing.Font;
using Form = System.Windows.Forms.Form;
using Task = System.Threading.Tasks.Task;
using Timer = System.Windows.Forms.Timer;

namespace OAMPS.Office.Word.Views.Wizards
{
    public class BaseWizardForm : Form, IBaseView, INotifyPropertyChanged
    {
        public static string HeaderType = string.Empty;
        protected static List<ISharePointListItem> MajorItems;
        protected static List<IPolicyClass> MinorItems;


        private static Timer _regenerationTimer;
        internal bool _loadComplete;
        public ObjectCache Cache = MemoryCache.Default;
        protected bool GenerateNewTemplate;
        protected Enums.FormLoadType LoadType;

        protected int ProtectionType;
        protected TreeModel TreeModel = new TreeModel();

        public BaseWizardForm()
        {
            AutoValidate = AutoValidate.Disable;
            ThisAddIn.IsWizzardRunning = true;
            PropertyChanged += BaseForm_PropertyChanged;
        }

        public override sealed AutoValidate AutoValidate
        {
            get { return base.AutoValidate; }
            set { base.AutoValidate = value; }
        }

        public bool Reload { get; set; }
        public bool AutoComplete { get; set; }

        public bool LoadComplete
        {
            get { return _loadComplete; }
            set
            {
                _loadComplete = value;
                OnPropertyChanged("LoadComplete");
            }
        }

        /// <summary>
        /// Method to check server for updated templates and download them id necessary 
        /// </summary>
        /// <param name="folder">The sharepoint source folder fragment, this will NOT be same as local folder.</param>
        /// <param name="localFolder">The local folder fragment.</param>
        /// <param name="file">The full filename of the template to check.</param>
        //public void ShouldUpdateTemplate(string folder, string file)
        //{
        //    string div = System.IO.Path.DirectorySeparatorChar.ToString();
        //    string remoteFile = string.Concat(
        //        Settings.Default.RelativeSitePath,
        //        folder, "/",
        //        file);

        //    string localFile = string.Concat(
        //        Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), div,
        //        Settings.Default.LocalPathRoot, div,
        //        file);

        //    bool localFileExists = System.IO.File.Exists(localFile);
        //    Debug.WriteLine(string.Format(localFileExists ? "\n  **  Local file '{0}' found" : "\n  **  Local file '{0}' not found", localFile));

        //    using (var ctx = new SharePointUserDirectClientContext(Settings.Default.SharePointContextUrl))
        //    {
        //        var rFile = ctx.Web.GetFileByServerRelativeUrl(remoteFile);
        //        ctx.Load(rFile, i => i.TimeLastModified);
        //        ctx.ExecuteQuery();

        //        if (!localFileExists || DateTime.Compare(rFile.TimeLastModified.ToUniversalTime(), System.IO.File.GetLastWriteTimeUtc(localFile)) > 0)
        //        {
        //            // Download and write new template to a temp file. 
        //            var temp = Path.GetTempFileName();
        //            var fileInformation = Microsoft.SharePoint.Client.File.OpenBinaryDirect(ctx, remoteFile); //needs to be template url.
        //            var stream = fileInformation.Stream;
        //            using (var output = System.IO.File.OpenWrite(temp))
        //            {
        //                stream.CopyTo(output);
        //            }

        //            if (System.IO.File.Exists(temp))
        //            {
        //                System.IO.File.Copy(temp, localFile, true);
        //                MessageBox.Show(Settings.Default.UpdateMessage, "Template updated", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        //                this.Close();
        //                Application.Exit();
        //            }
        //        }
        //    }
        //}

        //public void ShouldUpdateTemplate(string folder, string localFolder, string file)
        //{
        //    bool update = false;
        //    string div = System.IO.Path.DirectorySeparatorChar.ToString();
        //    string remoteFile = string.Concat(
        //        Settings.Default.RelativeSitePath,
        //        folder, "/",
        //        file);
        //    string localFile = string.Concat(
        //        Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), div,
        //        Settings.Default.LocalPathRoot, div,
        //        localFolder, div,
        //        file);


        //    using (var ctx = new SharePointUserDirectClientContext(Settings.Default.SharePointContextUrl))
        //    {
        //        var rFile = ctx.Web.GetFileByServerRelativeUrl(remoteFile);
        //        ctx.Load(rFile, i => i.TimeLastModified);
        //        ctx.ExecuteQuery();

        //        DateTime rDate, lDate;

        //        if (null != rFile.TimeLastModified)
        //        {
        //            rDate = rFile.TimeLastModified;
        //            lDate = System.IO.File.GetLastAccessTime(localFile);

        //            update = DateTime.Compare(rDate, lDate) > 0;
        //        }

        //        if (update)
        //        {
        //            // Download and write new template to a temp file. 
        //            var temp = Path.GetTempFileName();
        //            var fileInformation = Microsoft.SharePoint.Client.File.OpenBinaryDirect(ctx, remoteFile); //needs to be template url.
        //            var stream = fileInformation.Stream;
        //            using (var output = System.IO.File.OpenWrite(temp))
        //            {
        //                stream.CopyTo(output);
        //            }

        //            if (System.IO.File.Exists(temp))
        //            {
        //                System.IO.File.Copy(temp, localFile, true);
        //                MessageBox.Show(Settings.Default.UpdateMessage, "Template updated", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        //                this.Close();
        //                Application.Exit();
        //            }
        //        }
        //    }
        //}

        protected BaseWizardPresenter BaseWizardPresenter { get; set; }
        protected bool WizardBeingUpdated { get; set; }

        public virtual void DisplayMessage(string text, string caption)
        {
            MessageBox.Show(text, caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnTimedEvent(object source, ElapsedEventArgs e)
        {
            var btn = Controls.Find("btnNext", true);
            ((Button) btn[0]).PerformClick();

            if (_regenerationTimer == null) return;
            _regenerationTimer.Dispose();
        }

        public void StartTimer()
        {
            _regenerationTimer = new Timer();
            _regenerationTimer.Tick += _regenerationTimer_Tick;
            //_regenerationTimer.Elapsed += new ElapsedEventHandler(OnTimedEvent);
            _regenerationTimer.Interval = 1000;
            _regenerationTimer.Enabled = true;

            AutoComplete = true;
        }

        private void _regenerationTimer_Tick(object sender, EventArgs e)
        {
            using (_regenerationTimer)
            {
                var tabControl = Controls.Find(Constants.ControlNames.TabControl, true);

                if (tabControl.Length == 0)
                    return;

                var button = Controls.Find("btnNext", true);
                if (button.Length == 0)
                    return;
                var c = ((TabControl) tabControl[0]);

                c.SelectedIndex = c.TabCount - 1;
                ((Button) button[0]).PerformClick();
            }
        }

        private void BaseForm_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (LoadComplete)
            {
                var button = Controls.Find("btnNext", true);
                if (button.Length == 1)
                {
                    button[0].Enabled = true;
                }
            }
        }


        protected void OnPropertyChanged(PropertyChangedEventArgs e)
        {
            var handler = PropertyChanged;
            if (handler != null)
                handler(this, e);
        }

        protected void OnPropertyChanged(string propertyName)
        {
            OnPropertyChanged(new PropertyChangedEventArgs(propertyName));
        }

        internal virtual void LoadCompanyLogoImagesTab(TaskScheduler uiScheduler, TabControl tab,
            string logoTitleForDefault)
        {
            tab.TabPages.Add(Constants.ControlNames.TabPageLogosName, Constants.ControlNames.TabPageLogosTitle);
            var tbpOampsLogo = tab.TabPages[Constants.ControlNames.TabPageLogosName];

            if (uiScheduler == null)
            {
                LoadCompanyLogoGraphicsSync(null, tbpOampsLogo, logoTitleForDefault);
            }
            else
            {
                Task.Factory.StartNew(() => LoadCompanyLogoGraphicsSync(uiScheduler, tbpOampsLogo, logoTitleForDefault));
            }
        }


        internal virtual void LoadBrandingImagesTab(TaskScheduler uiScheduler, TabControl tab,
            string coverPageTitleForDefault, string defaultSpeciality)
        {
            HeaderType = BaseWizardPresenter.ReadDocumentProperty(Constants.SharePointFields.HeaderType);
            IEnumerable<string> specialities;

            var list = new Helpers.LocalSharePoint.ListLoader();
            var settings = list.GetListSettings(Settings.Default.GraphicsPictureLibraryTitle);
            specialities = ((JArray)settings["Speciality.Choices"]).ToObject<List<string>>();
            if (string.IsNullOrEmpty(defaultSpeciality))
            {
                defaultSpeciality = (string)settings["Speciality.DefaultValue"];
            }
          
            if (!string.Equals(HeaderType, "No Header Image", StringComparison.OrdinalIgnoreCase))
            {
                tab.TabPages.Add(Constants.ControlNames.TabPageCoverPagesName,
                    Constants.ControlNames.TabPageCoverPagesTitle);
                var tbpMainGraphic = tab.TabPages[Constants.ControlNames.TabPageCoverPagesName];
                tbpMainGraphic.AutoScroll = true;
                tbpMainGraphic.HorizontalScroll.Enabled = false;
                tbpMainGraphic.HorizontalScroll.Visible = false;
                tbpMainGraphic.VerticalScroll.Visible = true;
                tbpMainGraphic.VerticalScroll.Enabled = true;

                if (uiScheduler == null)
                {
                    LoadCoverPageGraphicsSync(null, tbpMainGraphic, coverPageTitleForDefault, specialities,
                        defaultSpeciality);
                }
                else
                {
                    Task.Factory.StartNew(
                        () =>
                            LoadCoverPageGraphicsSync(uiScheduler, tbpMainGraphic, coverPageTitleForDefault,
                                specialities, defaultSpeciality));
                }
            }
        }

        public void SwitchTab(TabControl tab, int index)
        {
            tab.SelectedIndex = index;
        }

        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            ThisAddIn.IsWizzardRunning = false;
            if (BaseWizardPresenter != null)
                BaseWizardPresenter.OnDocumentLoadCompleted(ProtectionType);
        }

        private string ConvertSegementToNumberical(BusinessLogic.Helpers.Enums.Segment segment)
        {
            switch (segment)
            {
                case BusinessLogic.Helpers.Enums.Segment.Five:
                    return "5";
                case BusinessLogic.Helpers.Enums.Segment.Four:
                    return "4";
                case BusinessLogic.Helpers.Enums.Segment.Three:
                    return "3";
                case BusinessLogic.Helpers.Enums.Segment.Two:
                    return "2";
                case BusinessLogic.Helpers.Enums.Segment.One:
                    return "1";
                case BusinessLogic.Helpers.Enums.Segment.PersonalLines:
                    return "Personal Lines";
                default:
                    return string.Empty;
            }
        }

        public void LogUsage(IBaseTemplate template, Enums.UsageTrackingType trackingType)
        {
            try
            {
                //var list = ListFactory.Create(Settings.Default.SharePointContextUrl,
                //    Settings.Default.UsageReportingListName);
                //var presenter = new SharePointListPresenter(list, this);

                var segment = string.Empty;
                var wholesaleOrRetail = string.Empty;
                var title = "Unknown";
                var userDep = "Unable to locate";
                var userOffice = "Unable to locate";

                var type = GetType().Name;

                switch (type)
                {
                    case "InsuranceRenewalReportWizard":
                    {
                        var form = ((InsuranceRenewalReportWizard) this);
                        segment = ConvertSegementToNumberical(form.SelectedSegment);
                        wholesaleOrRetail = form.SelectedStatutory.ToString();
                        title = Constants.TemplateNames.InsuranceRenewalReport;
                        break;
                    }
                    case "ClientDiscoveryWizard":
                    {
                        //var form = ((ClientDiscoveryWizard) this);
                        title = Constants.TemplateNames.ClientDiscovery;
                        break;
                    }

                    case "PreRenewalAgendaWizard":
                    {
                        //var form = ((PreRenewalAgendaWizard) this);
                        title = Constants.TemplateNames.PreRenewalAgenda;
                        break;
                    }

                    case "RenewalLetterWizard":
                    {
                        //var form = ((RenewalLetterWizard) this);
                        title = Constants.TemplateNames.RenewalLetter;
                        break;
                    }
                    case "SummaryOfDiscussionWizard":
                    {
                        //var form = ((SummaryOfDiscussionWizard) this);
                        title = Constants.TemplateNames.FileNote;
                        break;
                    }


                    case "GenericLetterWizard":
                    {
                        //var form = ((GenericLetterWizard)this);
                        title = Constants.TemplateNames.GenericLetter;
                        break;
                    }

                    case "PreRenewalQuestionareWizard":
                    {
                        title = "Fact Finder";
                            //todo: need to move this to constants once we can update value in ShaerPoint. (currently Pre Renewal Questionaire and cannot be changed until new wizard is released)
                        break;
                    }
                    case "FactFinderWizard":
                    {
                        title = "Fact Finder";
                            //todo: need to move this to constants once we can update value in ShaerPoint. (currently Pre Renewal Questionaire and cannot be changed until new wizard is released)
                        break;
                    }

                    case "QuoteSlipWizard":
                    {
                        title = Constants.TemplateNames.QuoteSlip;
                        break;
                    }

                    case "InsuranceManualWizard":
                    {
                        title = Constants.TemplateNames.InsuranceManual;
                        break;
                    }

                    case "ShortFormProposalWizard":
                    {
                        title = Constants.TemplateNames.ShortFormProposal + " - " + template.DocumentTitle;
                        break;
                    }

                    case "PlacementSlip":
                    {
                        title = "Placement Slip";
                        break;
                    }

                    default:
                    {
                        title = template.DocumentTitle;
                        break;
                    }
                }

                Task.Factory.StartNew(() =>
                {
                    var cacheKey = $"OAMPS.Word.CurrentUser:{Environment.UserName}";
                    var user = LocalCache.Get<UserPrincipalEx>(cacheKey, FindCurrentUserInAd);

                    if (user != null)
                    {
                        userDep = user.Branch;
                        userOffice = user.Suburb;
                    }

                    var reportDetails = string.Format("DocumentTitle:{0}, CoverPageTitle:{1}, LogoTitle:{2}", title, template.CoverPageTitle ?? "-", template.LogoTitle ?? "-");
                    var usageLog = new UsageLog
                    {
                        AccountExecutive = template.ExecutiveName ?? String.Empty,
                        CaptureDate = DateTime.Now.ToShortDateString(),
                        CaptureTime = DateTime.Now.TimeOfDay.ToString(),
                        ClientName = template.ClientName ?? String.Empty,
                        Report = reportDetails,
                        Segment = segment,
                        TrackingType = trackingType.ToString(),
                        UserDepartment = userDep,
                        UserName = Environment.UserName,
                        UserOffice = userOffice,
                        VersionNumber = Constants.Configuration.VersionNumber
                    };

                    var usageLogger = QueuedUsageLogger.Instance;
                    usageLogger.LogUsage(usageLog);

                }, CancellationToken.None).ContinueWith((task =>
                {
                    if (task.IsFaulted)
                    {
                        OnError(task.Exception);
                    }
                }));
            }
            catch (Exception ex)
            {
                OnError(ex);
            }
        }

        private UserPrincipalEx FindCurrentUserInAd()
        {
            return LocalCache.Get<UserPrincipalEx>("CurrentUserInAD:" + Environment.UserName, () => {
                try
                {
                    using (var context = new PrincipalContext(ContextType.Domain, Settings.Default.PeoplePickerSearchDomain, Settings.Default.PeoplePickerSearchOU))
                    {
                        using (var userPrincipal = new UserPrincipalEx(context) { Enabled = true })
                        {
                            userPrincipal.SamAccountName = Environment.UserName;
                            using (var searcher = new PrincipalSearcher(userPrincipal))
                            {
                                searcher.QueryFilter = userPrincipal;
                                var users = searcher.FindAll().ToList();

                                if (users.Count == 1)
                                    return (UserPrincipalEx)users.FirstOrDefault();
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    OnError(ex);
                }

                return null;
            });
        }

        protected override void OnLoad(EventArgs e)
        {
            var stopwatch = new Stopwatch();
            Debug.WriteLine("\n  *** STOPWATCH START - BaseWizardForm ***\n");
            stopwatch.Start();

            if (BaseWizardPresenter != null)
            {
                ProtectionType = BaseWizardPresenter.OnDocumentLoaded();

                var coverPageLabel = Controls.Find("lblCoverPageTitle", true);
                if (coverPageLabel.Length == 1)
                {
                    var selectedCoverPage =
                        BaseWizardPresenter.ReadDocumentProperty(Constants.WordDocumentProperties.CoverPageTitle);

                    //catch all document upgrades.  
                    //in october 2014 oamps renamed to author j gall
                    //to ensure no document was left upgraded the below business rules catches any document missed
                    if (selectedCoverPage != null &&
                        selectedCoverPage.Equals("boulder opal", StringComparison.OrdinalIgnoreCase))
                    {
                        selectedCoverPage = "Woodshed";
                    }

                    coverPageLabel[0].Text = (string.IsNullOrEmpty(selectedCoverPage)
                        ? coverPageLabel[0].Text
                        : selectedCoverPage);
                }

                var lblSpeciality = Controls.Find("lblSpeciality", true);
                if (lblSpeciality.Length == 1)
                {
                    var selectedSpeciality =
                        BaseWizardPresenter.ReadDocumentProperty(Constants.WordDocumentProperties.Speciality);
                    lblSpeciality[0].Text = (string.IsNullOrEmpty(selectedSpeciality)
                        ? lblSpeciality[0].Text
                        : selectedSpeciality);
                }

                var titleLabel = Controls.Find("lblLogoTitle", true);
                if (titleLabel.Length == 1)
                {
                    var selectedLogo =
                        BaseWizardPresenter.ReadDocumentProperty(Constants.WordDocumentProperties.LogoTitle);

                    //catch all document upgrades.  
                    //in october 2014 oamps renamed to author j gall
                    //to ensure no document was left upgraded the below business rules catches any document missed
                    if (selectedLogo != null &&
                        selectedLogo.Equals("oamps insurance brokers ltd", StringComparison.OrdinalIgnoreCase))
                    {
                        selectedLogo = "Arthur J. Gallagher & Co (Aus) Limited";
                    }

                    titleLabel[0].Text = (string.IsNullOrEmpty(selectedLogo) ? titleLabel[0].Text : selectedLogo);
                }

                BaseWizardPresenter.CloseInformationPanel(true);
            }

            CenterToScreen();
            CenterToParent();

            //Task.Factory.StartNew(() => Thread.Sleep(3000)).ContinueWith(task => this.Activate(), TaskScheduler.FromCurrentSynchronizationContext());
            base.OnLoad(e);

            stopwatch.Stop();
            Debug.WriteLine($"\n  *** STOPWATCH STOP - BaseWizardForm: {stopwatch.Elapsed} ***\n");
        }


        private void AddCompanyLogoImagesToTab(TaskScheduler uiScheduler, TabPage tabPage, string defaultControlName,
            List<IThumbnail> items)
        {
            var xpos = 15;
            var ypos = 15;
            var counter = 0;
            var ycounter = 0;

            if (uiScheduler == null)
            {
                ycounter = AddCompanyLogoImagesToTab(tabPage, defaultControlName, items, ycounter, ref xpos, ref ypos,
                    ref counter);
                LoadComplete = true;
            }
            else
            {
                Task.Factory.StartNew(() =>
                {
                    //requirements for this section has simplified.  we could now just use a table layout control and clone rows.
                    //implement if layout requiremnts change
                    ycounter = AddCompanyLogoImagesToTab(tabPage, defaultControlName, items, ycounter, ref xpos,
                        ref ypos, ref counter);
                    LoadComplete = true;
                }, CancellationToken.None, TaskCreationOptions.None, uiScheduler);
            }
        }


        private int AddCompanyLogoImagesToTab(TabPage tabPage, string defaultControlName, IEnumerable<IThumbnail> items,
            int ycounter, ref int xpos, ref int ypos, ref int counter)
        {
            foreach (var t in items)
            {
                using (var ms = new MemoryStream(t.ImageStream))
                {
                    var b = new Bitmap(ms);

                    //Cache.Add(t.ImageStream)

                    //calculate positioning of image.
                    if (ycounter < 3)
                    {
                        if (ycounter == 0)
                            xpos = 15;
                        else
                        {
                            xpos += b.Width + 35;
                        }


                        ycounter++;
                    }
                    else
                    {
                        xpos = 15;
                        ypos += b.Height + 195;

                        ycounter = 1; //reset back to 1 for second loop and onwards;
                    }
                    //display image
                    var pictureBox = new PictureBox
                    {
                        Name = Constants.ControlNames.TabPageCoverPagesPictureBoxName,
                        Visible = true,
                        Image = b,
                        Width = b.Width,
                        Height = b.Height,
                        Left = xpos,
                        Top = ypos
                    };

                    //pictureBox.Load(t.FullImageUrl);
                    tabPage.Controls.Add(pictureBox);

                    //display the image radio button
                    var valueRadioButton = new ValueRadioButton
                    {
                        Text = t.ImageTitle,
                        Font = new Font(pictureBox.Font.FontFamily, 7, pictureBox.Font.Style),
                        Left = xpos,
                        Top = ypos + b.Height + 5,
                        Width = b.Width + 15,
                        Height = 35,
                        AutoSize = false,
                        UseMnemonic = false,
                        Values = new Dictionary<string, string>
                        {
                            {Constants.RadioButtonValues.ImageUrl, t.FullImageUrl},
                            {Constants.RadioButtonValues.AbnKey, t.Abn},
                            {Constants.RadioButtonValues.AfslKey, t.Afsl},
                            {Constants.RadioButtonValues.WebsiteKey, t.WebSite},
                            {Constants.RadioButtonValues.OAMPSCompanyName, t.ImageTitle},
                            {Constants.RadioButtonValues.LongBrandingDesc, t.LongDescription}
                        }
                    };


                    if (!string.IsNullOrEmpty(defaultControlName))
                    {
                        if (string.Equals(t.ImageTitle, defaultControlName, StringComparison.OrdinalIgnoreCase))
                        {
                            valueRadioButton.Checked = true;
                        }
                    }
                    else if (counter == 0)
                        valueRadioButton.Checked = true;

                    if (t.ShortDescription != null)
                    {
                        var lblShortDescriptionLabel = new Label
                        {
                            Text = t.ShortDescription,
                            Left = xpos,
                            Top = ypos + b.Height + 45,
                            //AutoSize = true,
                            Width = 170,
                            Height = 105,
                            Font = new Font(Font.FontFamily, 7)
                        };
                        tabPage.Controls.Add(lblShortDescriptionLabel);
                    }


                    if (!string.IsNullOrEmpty(t.Abn))
                    {
                        var txtAbn = new Label
                        {
                            Text = @"ABN: " + t.Abn,
                            Left = xpos,
                            Top = ypos + b.Height + 50,
                            AutoSize = true
                        };
                        tabPage.Controls.Add(txtAbn);
                    }

                    if (!string.IsNullOrEmpty(t.Afsl))
                    {
                        var txtAfsl = new Label
                        {
                            Text = @"AFSL: " + t.Afsl,
                            Left = xpos,
                            Top = ypos + b.Height + 70,
                            AutoSize = true
                        };
                        tabPage.Controls.Add(txtAfsl);
                    }


                    if (!string.IsNullOrEmpty(t.WebSite))
                    {
                        var txtWebsite = new Label
                        {
                            Text = t.WebSite,
                            Left = xpos,
                            Top = ypos + b.Height + 90,
                            AutoSize = true
                        };
                        tabPage.Controls.Add(txtWebsite);
                    }


                    // tabPage.Click += new EventHandler(tbpMainGraphic_Click);
                    tabPage.Controls.Add(valueRadioButton);

                    counter++;
                }
            }
            return ycounter;
        }


        private void AddBrandingImagesToTab(TaskScheduler uiScheduler, TabPage tabPage, string defaultControlName,
            List<IThumbnail> items)
        {
            var xpos = 15;
            var ypos = 25;
            var counter = 0;
            var ycounter = 0;

            if (uiScheduler == null)
            {
                ycounter = AddBrandingImagesToTab(tabPage, defaultControlName, items, ycounter, ref xpos, ref ypos,
                    ref counter);
                LoadComplete = true;
            }
            else
            {
                Task.Factory.StartNew(() =>
                {
                    //requirements for this section has simplified.  we could now just use a table layout control and clone rows.
                    //implement if layout requiremnts change
                    ycounter = AddBrandingImagesToTab(tabPage, defaultControlName, items, ycounter, ref xpos, ref ypos,
                        ref counter);
                    LoadComplete = true;
                }, CancellationToken.None, TaskCreationOptions.None, uiScheduler);
            }
        }


        private int AddBrandingImagesToTab(TabPage tabPage, string defaultControlName, IEnumerable<IThumbnail> items,
            int ycounter, ref int xpos, ref int ypos, ref int counter)
        {
            foreach (var t in items)
            {
                using (var ms = new MemoryStream(t.ImageStream))
                {
                    var b = new Bitmap(ms);

                    //Cache.Add(t.ImageStream)

                    //calculate positioning of image.
                    if (ycounter < 3)
                    {
                        if (ycounter == 0)
                            xpos = 15;
                        else
                        {
                            xpos += b.Width + 25;
                        }


                        ycounter++;
                    }
                    else
                    {
                        xpos = 15;
                        ypos += b.Height + 130;

                        ycounter = 1; //reset back to 1 for second loop and onwards;
                    }
                    //display image
                    var pictureBox = new PictureBox
                    {
                        Name = Constants.ControlNames.TabPageCoverPagesPictureBoxName,
                        Visible = true,
                        Image = b,
                        Width = b.Width,
                        Height = b.Height,
                        Left = xpos,
                        Top = ypos
                    };

                    //pictureBox.Load(t.FullImageUrl);
                    tabPage.Controls.Add(pictureBox);

                    //display the image radio button
                    var valueRadioButton = new ValueRadioButton
                    {
                        Text = t.ImageTitle,
                        Font = new Font(pictureBox.Font.FontFamily, 7, pictureBox.Font.Style),
                        Left = xpos,
                        Top = ypos + b.Height + 5,
                        AutoSize = true,
                        UseMnemonic = false,
                        Values = new Dictionary<string, string>
                        {
                            {Constants.RadioButtonValues.ImageUrl, t.FullImageUrl},
                            {Constants.RadioButtonValues.AbnKey, t.Abn},
                            {Constants.RadioButtonValues.AfslKey, t.Afsl},
                            {Constants.RadioButtonValues.WebsiteKey, t.WebSite},
                            {Constants.RadioButtonValues.OAMPSCompanyName, t.ImageTitle},
                            {Constants.RadioButtonValues.LongBrandingDesc, t.LongDescription}
                        }
                    };


                    if (!string.IsNullOrEmpty(defaultControlName))
                    {
                        if (string.Equals(t.ImageTitle, defaultControlName, StringComparison.OrdinalIgnoreCase))
                        {
                            valueRadioButton.Checked = true;
                        }
                    }
                    else if (counter == 0)
                        valueRadioButton.Checked = true;

                    if (t.ShortDescription != null)
                    {
                        var lblShortDescriptionLabel = new Label
                        {
                            Text = t.ShortDescription,
                            Left = xpos,
                            Top = ypos + b.Height + 25,
                            //AutoSize = true,
                            Width = 170,
                            Height = 105,
                            Font = new Font(Font.FontFamily, 7),
                            UseMnemonic = false
                        };
                        tabPage.Controls.Add(lblShortDescriptionLabel);
                    }

                    // tabPage.Click += new EventHandler(tbpMainGraphic_Click);
                    valueRadioButton.UseMnemonic = false;
                    tabPage.Controls.Add(valueRadioButton);

                    counter++;
                }
            }
            return ycounter;
        }


        protected void PopulateLogosToTemplate(TabPage logoTab, ref BaseTemplate template)
        {
            if (logoTab == null)
                return;

            foreach (Control c in logoTab.Controls)
            {
                if (c.GetType() == typeof (ValueRadioButton))
                {
                    var v = ((ValueRadioButton) c);
                    if (v.Checked)
                    {
                        template.LogoImageUrl = v.GetValue(Constants.RadioButtonValues.ImageUrl);
                        template.LogoTitle = v.Text;
                        template.OAMPSAfsl = v.GetValue(Constants.RadioButtonValues.AfslKey);
                        template.OAMPSAbnNumber = v.GetValue(Constants.RadioButtonValues.AbnKey);
                        template.WebSite = v.GetValue(Constants.RadioButtonValues.WebsiteKey);
                        template.OAMPSCompanyName = v.GetValue(Constants.RadioButtonValues.OAMPSCompanyName);
                    }
                }
            }
        }

        protected void PopulateCoversToTemplate(TabPage covberTab, ref BaseTemplate template)
        {
            if (covberTab == null)
                return;

            foreach (Control c in covberTab.Controls)
            {
                if (c.GetType() == typeof (ValueRadioButton))
                {
                    var v = ((ValueRadioButton) c);
                    if (v.Checked)
                    {
                        template.CoverPageImageUrl = v.GetValue(Constants.RadioButtonValues.ImageUrl);
                        template.CoverPageTitle = v.Text;
                        template.LongBrandingDescription =
                            v.GetValue(Constants.RadioButtonValues.LongBrandingDesc);
                    }
                }
            }
        }

        internal virtual void LoadCoverPageGraphicsSync(TaskScheduler uiScheduler, TabPage tabPage,
            string coverPageTitleForDefault, IEnumerable<string> specialities, string defaultSpeciality)
        {
            var specialitySelected = AddFilterControlsToCoverPageImagesTab(tabPage, specialities, defaultSpeciality);
            var cacheName = HeaderType + Constants.ControlNames.TabPageCoverPagesName;
            var th = new List<IThumbnail>();

            if (Cache.Contains(cacheName + specialitySelected))
            {
                th = ((List<IThumbnail>)Cache.Get(cacheName + specialitySelected));
            }
            else
            {
                var listTitle = Settings.Default.GraphicsPictureLibraryTitle;
                var list = new Helpers.LocalSharePoint.ListLoader();
                var listFolder = list.GetLocalPath(listTitle);
                var listItems = list.GetItems(listTitle);

                IEnumerable<Helpers.LocalSharePoint.ListItem> items;
                if (string.Equals("full", HeaderType, StringComparison.OrdinalIgnoreCase))
                {
                    items = listItems.Where(x => x[Constants.SharePointFields.HeaderType].ToString().Equals(HeaderType, StringComparison.InvariantCultureIgnoreCase) &&
                        x[Constants.SharePointFields.Speciality].ToString().Equals(specialitySelected, StringComparison.InvariantCultureIgnoreCase)) 
                        .OrderBy(x => Convert.ToInt32(x["SortOrder"]));
                }
                else
                {
                    items = listItems.Where(x => x[Constants.SharePointFields.HeaderType].ToString().Equals(HeaderType, StringComparison.InvariantCultureIgnoreCase))
                        .OrderBy(x => Convert.ToInt32(x["SortOrder"]));
                }

                foreach (var item in items)
                {
                    var imagePath = Path.Combine(listFolder, (string) item["FileLeafRef"]);
                    var thumbcontent = System.IO.File.ReadAllBytes(FileHelpers.GetThumbnailPath(imagePath));
                    th.Add(new Thumbnail()
                    {
                        ImageStream = thumbcontent,
                        ImageTitle = item.GetFieldValue("Title"),
                        Order = item.GetFieldValueAsInteger("Order"),
                        FullImageUrl = "file://" + imagePath,
                        HeaderType = item.GetFieldValue(Constants.SharePointFields.HeaderType),
                        LongDescription = item.GetFieldValue(Constants.SharePointFields.LongDescription),
                        ShortDescription = item.GetFieldValue(Constants.SharePointFields.ShortDescription),
                        RelativeUrl = item.GetFieldValue(Constants.SharePointFields.FileRef)
                    });
                }
                th = th.OrderBy(x => x.Order).ToList();
                Cache.Add(cacheName + specialitySelected, th, new CacheItemPolicy());
            }

            AddBrandingImagesToTab(uiScheduler, tabPage, coverPageTitleForDefault, th);
        }

        private string AddFilterControlsToCoverPageImagesTab(TabPage tabPage, IEnumerable<string> specialities,
            string defaultSpeciality)
        {
            var lblSpecialityFilter = new Label
            {
                Text = "Filter on speciality:"
            };
            tabPage.Controls.Add(lblSpecialityFilter);
            var cboSpecialityFilter = new ComboBox
            {
                Left = lblSpecialityFilter.Width + 10,
                Width = 250,
                DropDownHeight = 600
            };

            cboSpecialityFilter.DataSource = specialities;
            tabPage.Controls.Add(cboSpecialityFilter);
            cboSpecialityFilter.SelectedText = defaultSpeciality;
            cboSpecialityFilter.Text = defaultSpeciality;

            cboSpecialityFilter.Refresh();
            cboSpecialityFilter.SelectedIndexChanged += cboSpecialityFilter_SelectedIndexChanged;


            return defaultSpeciality;
        }

        private void cboSpecialityFilter_SelectedIndexChanged(object sender, EventArgs e)
        {
            var comboBox = ((ComboBox) sender);
            var tabpage = comboBox.Parent as TabPage;
            var defaultText = comboBox.Text;
            var items = (from object i in comboBox.Items select i.ToString()).ToList();


            while (tabpage != null && tabpage.Controls.Count > 0)
            {
                tabpage.Controls.RemoveAt(0);
            }

            var lblSpeciality = Controls.Find("lblSpeciality", true);
            if (lblSpeciality.Length == 1)
            {
                lblSpeciality[0].Text = defaultText;
            }

            LoadCoverPageGraphicsSync(null, tabpage, string.Empty, items, defaultText);
        }

        internal virtual void LoadCompanyLogoGraphicsSync(TaskScheduler uiScheduler, TabPage tabPage,
            string logoTitleForDefault)
        {
            List<IThumbnail> logoStreams= new List<IThumbnail>();

            var cacheName = HeaderType + "LogoStreams";
            if (Cache.Contains(cacheName))
            {
                logoStreams = (List<IThumbnail>)Cache.Get(cacheName);
            }
            else
            {
                try
                {
                    var imagePath = Path.Combine(Environment.ExpandEnvironmentVariables(Settings.Default.LocalFullPathRoot),
                    Settings.Default.GraphicsPictureLibraryTitleLogos);
                    var dirInfo = new DirectoryInfo(imagePath);
                    foreach (var file in dirInfo.GetFiles("*.jpg"))
                    {
                        if (file.FullName.EndsWith("_jpg.jpg", StringComparison.InvariantCultureIgnoreCase))
                            continue;

                        var fieldValues = new Dictionary<string, object>();
                        var content = FileHelpers.GetFileStreamWithMetadata(FileHelpers.GetThumbnailPath(file.FullName), out fieldValues);

                        logoStreams.Add(new Thumbnail
                        {
                            ImageStream = content,
                            ImageTitle = fieldValues.GetFieldValue("Title"),
                            Order = fieldValues.GetFieldValueAsInteger("SortOrder"),
                            FullImageUrl = "file://" + file.FullName,
                            Abn = fieldValues.GetFieldValue(Constants.SharePointFields.Abn),
                            Afsl = fieldValues.GetFieldValue(Constants.SharePointFields.Afsl),
                            WebSite = fieldValues.GetFieldValue(Constants.SharePointFields.Website)
                        });
                    }
                    logoStreams = logoStreams.OrderBy(x => x.Order).ToList();
                }
                catch (Exception ex)
                {
                    OnError(ex);
                }

                Cache.Add(cacheName, logoStreams, new CacheItemPolicy());
            }            

            AddCompanyLogoImagesToTab(uiScheduler, tabPage, logoTitleForDefault, logoStreams);
        }

        protected T GetCachedTempalteObject<T>()
        {
            var template = (T) Cache.Get(Constants.CacheNames.RegenerateTemplate);
            Cache.Remove(Constants.CacheNames.RegenerateTemplate);
            return template;
        }

        protected void LogTemplateDetails(IBaseTemplate template)
        {
            ErrorLog.TraceLog("BaseWizardForm:LogTemplateDetails - DocumentTitle:{0}, CoverPageTitle:{1}, LogoTitle:{2}", template.DocumentTitle, template.CoverPageTitle, template.LogoTitle);
        }

        public virtual void PopulateDocument(IBaseTemplate template, string previousCoverPageTitle,
            string previousLogoTitle)
        {
            //populate the content controls
            BaseWizardPresenter.PopulateData(template);
            //change the graphics selected
            // if (Streams == null) return;

            LogTemplateDetails(template);

            BaseWizardPresenter.PopulateGraphics(template, previousCoverPageTitle, previousLogoTitle);
            BaseWizardPresenter.UpdateFields();
            var type = template.GetType();

            //var propertyName = Constants.WordDocumentProperties.UsedDateOfLogo;

            //if(key.Equals("Themes",StringComparison.OrdinalIgnoreCase))
            //    propertyName = Constants.WordDocumentProperties.UsedDateOfTheme;


            var date = DateTime.Now.ToString("dd/MM/yy");
            var logoVal = $"{template.LogoImageUrl};{date}";
            var themeVal = $"{template.CoverPageImageUrl};{date}";

            BaseWizardPresenter.CreateOrUpdateDocumentProperty(Constants.WordDocumentProperties.UsedDateOfLogo, logoVal);
            BaseWizardPresenter.CreateOrUpdateDocumentProperty(Constants.WordDocumentProperties.UsedDateOfTheme,
                themeVal);
            BaseWizardPresenter.CreateOrUpdateDocumentProperty(Constants.WordDocumentProperties.DocumentGeneratedDate,
                date);
        }

        public IEnumerable<Control> GetAll(Control control, Type type)
        {
            var controls = control.Controls.Cast<Control>();

            var enumerable = controls as IList<Control> ?? controls.ToList();
            foreach (var c in enumerable.SelectMany(ctrl => GetAll(ctrl, type)).Concat(enumerable))
            {
                if (c.GetType() == type) yield return c;
            }
        }


        protected List<ISharePointListItem> LoadMajorTreeNodeTypes(string contextUrl, string listName)
        {
            var list = ListFactory.Create(listName);
            var presenter = new SharePointListPresenter(list, this);
            return presenter.GetItems();
        }

        protected virtual List<IPolicyClass> LoadMinorPolicyTypes()
        {
            var list = ListFactory.Create(Settings.Default.MinorPolicyClassesListName);
            var presenter = new SharePointListPresenter(list, this);
            return presenter.GetMinorPolicyItems();
        }

        protected void LoadDataSources(TaskScheduler uiScheduler)
        {
            LoadTreeViewClasses(uiScheduler);
            //    _insurers = LoadInsurers();
        }


        protected virtual void LoadTreeViewClasses(TaskScheduler uiScheduler)
        {
            if (Cache.Contains(Constants.CacheNames.MajorPolicyClassItems))
            {
                MajorItems =
                    ((List<ISharePointListItem>)
                        Cache.Get(Constants.CacheNames.MajorPolicyClassItems));
            }
            else
            {
                MajorItems = LoadMajorTreeNodeTypes(Settings.Default.SharePointContextUrl,
                    Settings.Default.MajorPolicyClassesListName);
                Cache.Add(Constants.CacheNames.MajorPolicyClassItems, MajorItems,
                    new CacheItemPolicy());
            }

            if (Cache.Contains(Constants.CacheNames.MinorPolicyClassItems))
            {
                MinorItems =
                    ((List<IPolicyClass>) Cache.Get(Constants.CacheNames.MinorPolicyClassItems));
            }
            else
            {
                MinorItems = LoadMinorPolicyTypes();
                //Cache.Add(Constants.CacheNames.MinorPolicyClassItems, MinorItems,
                //          new CacheItemPolicy());
            }

            if (uiScheduler == null)
                WriteNodesToTree();
            else
                Task.Factory.StartNew(WriteNodesToTree, CancellationToken.None, TaskCreationOptions.None, uiScheduler);
        }

        private void WriteNodesToTree()
        {
            var tv = Controls.Find("tvaPolicies", true);
            if (tv.Length == 0) return;
            var tvaPolicies = (TreeViewAdv) tv[0];

            tvaPolicies.Model = TreeModel;

            tvaPolicies.BeginUpdate();
            foreach (var majorItem in MajorItems)
            {
                // var parentNode = new TreeNode();

                var rootNode = AddRoot(majorItem.Title);

                //var mItemTitle = majorItem.Title;
                //   parentNode = tvPolicies.Nodes.Add(mItemTitle.Replace(" ", string.Empty), mItemTitle);

                var found =
                    MinorItems.FindAll(
                        i =>
                            string.Equals(i.MajorClass, majorItem.Title, StringComparison.OrdinalIgnoreCase));
                foreach (var minorItem in found)
                {
                    var childNode = AddChild(rootNode, minorItem.Title);
                    var node = ((AdvancedTreeNode) childNode);

                    node.Id = minorItem.Id;
                }
            }
            tvaPolicies.EndUpdate();
        }

        protected Node AddRoot(string text)
        {
            Node node = new AdvancedTreeNode(text);
            TreeModel.Nodes.Add(node);
            return node;
        }

        protected virtual bool ContinueWithSignificantChange(object sender, bool previousValue)
        {
            if (GenerateNewTemplate)
                return true;

            var r = MessageBox.Show(
                Constants.Miscellaneous.ConfirmMsg,
                @"Please Confirm", MessageBoxButtons.YesNo);

            if (r == DialogResult.No)
            {
                WizardBeingUpdated = true;
                var type = sender.GetType();
                if (type == typeof (CheckBox))
                {
                    ((CheckBox) sender).Checked = previousValue;
                }
                else if (type == typeof (RadioButton))
                {
                    ((RadioButton) sender).Checked = previousValue;
                }
                WizardBeingUpdated = false;
                return false;
            }
            return true;
        }

        protected Node AddChild(Node parent, string text)
        {
            Node node = new AdvancedTreeNode(text);
            parent.Nodes.Add(node);
            return node;
        }

        protected virtual void OnError(Exception e)
        {
            var logger = new EventViewerLogger();
            logger.Log(e.ToString(), BusinessLogic.Interfaces.Logging.Type.Error);

            ErrorLog.Error(e, "BaseWizardForm.OnError");

#if DEBUG
            MessageBox.Show(e.ToString(), @"sorry");
#endif
        }

        private void InitializeComponent()
        {
            SuspendLayout();
            // 
            // BaseWizardForm
            // 
            ClientSize = new Size(284, 262);
            Name = "BaseWizardForm";
            Load += BaseWizardForm_Load;
            ResumeLayout(false);
        }

        private void BaseWizardForm_Load(object sender, EventArgs e)
        {
        }
    }
}