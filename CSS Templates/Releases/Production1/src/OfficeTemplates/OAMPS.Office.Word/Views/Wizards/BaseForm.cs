using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.Caching;
using System.Text;
using Aga.Controls.Tree;
using OAMPS.Office.BusinessLogic.Helpers;
using OAMPS.Office.BusinessLogic.Interfaces;
using System.Windows.Forms;
using System.Threading.Tasks;
using OAMPS.Office.BusinessLogic.Interfaces.SharePoint;
using OAMPS.Office.BusinessLogic.Models.Template;
using OAMPS.Office.Word.Helpers.Controls;
using OAMPS.Office.Word.Properties;
using OAMPS.Office.BusinessLogic.Interfaces.Wizards;
using OAMPS.Office.BusinessLogic.Presenters;
using OAMPS.Office.Word.Models.SharePoint;
using OAMPS.Office.BusinessLogic.Presenters.SharePoint;
using System.Threading;
using System.Drawing;
using OAMPS.Office.BusinessLogic.Interfaces.Template;
using OAMPS.Office.Word.Helpers;

namespace OAMPS.Office.Word.Views.Wizards
{
    public class BaseForm : Form, IBaseView, INotifyPropertyChanged
    {
        protected Helpers.Enums.FormLoadType _loadType;
        protected bool _generateNewTemplate;
        public static string HeaderType = String.Empty;
        protected static List<ISharePointListItem> MajorItems = null;
        protected static List<IPolicyClass> MinorItems = null;
        public bool Reload { get; set; }

        private bool _loadComplete;
        public bool LoadComplete
        {
            get { return _loadComplete; }
            set
            {
                _loadComplete = value;
                OnPropertyChanged("LoadComplete");
            }
        }

        protected int ProtectionType;
        protected TreeModel TreeModel = new TreeModel();
        protected BasePresenter BasePresenter { get; set; }
        protected bool WizardBeingUpdated { get; set; }
        public ObjectCache Cache = MemoryCache.Default;

        public event PropertyChangedEventHandler PropertyChanged;


        protected BaseForm()
        {
            this.AutoValidate = AutoValidate.Disable;

            this.PropertyChanged += new PropertyChangedEventHandler(BaseForm_PropertyChanged);
        }

        void BaseForm_PropertyChanged(object sender, PropertyChangedEventArgs e)
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
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
                handler(this, e);
        }

        protected void OnPropertyChanged(string propertyName)
        {
            OnPropertyChanged(new PropertyChangedEventArgs(propertyName));
        }



        internal virtual void LoadGenericImageTabs(TaskScheduler uiScheduler, TabControl tab, string coverPageTitleForDefault, string logoTitleForDefault)
        {
            HeaderType = BasePresenter.ReadDocumentProperty(Constants.SharePointFields.HeaderType);

            if (!String.Equals(HeaderType, "No Header Image", StringComparison.OrdinalIgnoreCase))
            {
                tab.TabPages.Add(Constants.ControlNames.TabPageCoverPagesName, Constants.ControlNames.TabPageCoverPagesTitle);
                var tbpMainGraphic = tab.TabPages[Constants.ControlNames.TabPageCoverPagesName];
                tbpMainGraphic.AutoScroll = true;
                tbpMainGraphic.HorizontalScroll.Enabled = false;
                tbpMainGraphic.HorizontalScroll.Visible = false;
                tbpMainGraphic.VerticalScroll.Visible = true;
                tbpMainGraphic.VerticalScroll.Enabled = true;
                Task.Factory.StartNew(() => LoadCoverPageGraphicsSync(uiScheduler, tbpMainGraphic, coverPageTitleForDefault));
            }

            tab.TabPages.Add(Constants.ControlNames.TabPageLogosName, Constants.ControlNames.TabPageLogosTitle);
            var tbpOampsLogo = tab.TabPages[Constants.ControlNames.TabPageLogosName];

            if (uiScheduler == null)
            {
                LoadLogoGraphicsSync(uiScheduler, tbpOampsLogo, logoTitleForDefault);
            }
            else
            {
                Task.Factory.StartNew(() => LoadLogoGraphicsSync(uiScheduler, tbpOampsLogo, logoTitleForDefault));    

                
                
            }
            
        }

        public void SwitchTab(TabControl tab, int index)
        {
            tab.SelectedIndex = index;
            
        }



        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            if (BasePresenter != null)
                BasePresenter.OnDocumentLoadCompleted(ProtectionType);
        }

        protected override void OnLoad(EventArgs e)
        {
            if (BasePresenter != null)
            {
                ProtectionType = BasePresenter.OnDocumentLoaded();

                var coverPageLabel = Controls.Find("lblCoverPageTitle", true);
                if (coverPageLabel.Length == 1)
                {
                    var selectedCoverPage = BasePresenter.ReadDocumentProperty(Constants.WordDocumentProperties.CoverPageTitle);
                    coverPageLabel[0].Text = (String.IsNullOrEmpty(selectedCoverPage) ? coverPageLabel[0].Text : selectedCoverPage);
                }

                var titleLabel = Controls.Find("lblLogoTitle", true);
                if (titleLabel.Length == 1)
                {
                    var selectedLogo = BasePresenter.ReadDocumentProperty(Constants.WordDocumentProperties.LogoTitle);
                    titleLabel[0].Text = (String.IsNullOrEmpty(selectedLogo) ? titleLabel[0].Text : selectedLogo);
                }

            }

            base.OnLoad(e);
        }

        private void DisplayGraphics(TaskScheduler uiScheduler, TabPage tabPage, string defaultControlName,  List<IThumbnail> items, bool isLogo)
        {
            var xpos = 15;
            var ypos = 15;
            var counter = 0;
            var ycounter = 0;

            Task.Factory.StartNew(() =>
            {
                //requirements for this section has simplified.  we could now just use a table layout control and clone rows.
                //implement if layout requiremnts change
                ycounter = DisplayGraphicsImpl(tabPage, defaultControlName, items, isLogo, ycounter, ref xpos, ref ypos, ref counter);
                LoadComplete = true;

            }, CancellationToken.None, TaskCreationOptions.None, uiScheduler);
        }

        protected void PopulateLogosToTemplate(TabPage logoTab, ref BaseTemplate template)
        {

            foreach (Control c in logoTab.Controls)
            {
                if (c.GetType() == typeof(ValueRadioButton))
                {
                    var v = ((ValueRadioButton)c);
                    if (v.Checked)
                    {
                        template.LogoImageUrl = v.GetValue(Constants.RadioButtonValues.ImageUrl);
                        template.LogoTitle = v.Text;
                        template.OAMPSAfsl = v.GetValue(Constants.RadioButtonValues.AFSLKey);
                        template.OAMPSAbnNumber = v.GetValue(Constants.RadioButtonValues.ABNKey);
                        template.WebSite = v.GetValue(Constants.RadioButtonValues.WebsiteKey);
                        template.OAMPSCompanyName = v.GetValue(Constants.RadioButtonValues.OAMPSCompanyName);
                    }
                }
            }
        }

        protected void PopulateCoversToTemplate(TabPage covberTab, ref BaseTemplate template)
        {
            foreach (Control c in covberTab.Controls)
            {
                if (c.GetType() == typeof(ValueRadioButton))
                {
                    var v = ((ValueRadioButton)c);
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
        
        private int DisplayGraphicsImpl(TabPage tabPage, string defaultControlName, List<IThumbnail> items, bool isLogo,
                                        int ycounter, ref int xpos, ref int ypos, ref int counter)
        {
            foreach (var t in items)
            {
                var b = new Bitmap(t.ImageStream);
                //calculate positioning of image.
                if (ycounter < 3)
                {
                    if (ycounter == 0)
                        xpos = 15;
                    else
                    {
                        if (isLogo)
                            xpos += b.Width + 25;
                        else
                            xpos += b.Width + 25;
                    }


                    ycounter++;
                }
                else
                {
                    if (isLogo)
                    {
                        xpos = 15;
                        ypos += b.Height + 45;
                    }
                    else
                    {
                        xpos = 15;
                        ypos += b.Height + 140;
                    }

                    ycounter = 1; //reset back to 1 for second loop and onwards;
                }
                //display image
                var pictureBox = new PictureBox
                    {
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
                        Left = xpos,
                        Top = ypos + b.Height + 5,
                        AutoSize = true,
                        Values = new Dictionary<string, string>
                            {
                                {Constants.RadioButtonValues.ImageUrl, t.FullImageUrl},
                                {Constants.RadioButtonValues.ABNKey, t.ABN},
                                {Constants.RadioButtonValues.AFSLKey, t.AFSL},
                                {Constants.RadioButtonValues.WebsiteKey, t.WebSite},
                                {Constants.RadioButtonValues.OAMPSCompanyName, t.ImageTitle},
                                {Constants.RadioButtonValues.LongBrandingDesc, t.LongDescription}
                            }
                    };


                if (!String.IsNullOrEmpty(defaultControlName))
                {
                    if (String.Equals(t.ImageTitle, defaultControlName, StringComparison.OrdinalIgnoreCase))
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
                            Font = new Font(Font.FontFamily, 7)
                        };
                    tabPage.Controls.Add(lblShortDescriptionLabel);
                }


                if (isLogo)
                {
                    if (!String.IsNullOrEmpty(t.ABN))
                    {
                        var txtAbn = new Label
                            {
                                Text = @"ABN: " + t.ABN,
                                Left = xpos,
                                Top = ypos + b.Height + 30,
                                AutoSize = true
                            };
                        tabPage.Controls.Add(txtAbn);
                    }

                    if (!String.IsNullOrEmpty(t.AFSL))
                    {
                        var txtAfsl = new Label
                            {
                                Text = @"AFSL: " + t.AFSL,
                                Left = xpos,
                                Top = ypos + b.Height + 50,
                                AutoSize = true
                            };
                        tabPage.Controls.Add(txtAfsl);
                    }


                    if (!String.IsNullOrEmpty(t.WebSite))
                    {
                        var txtWebsite = new Label
                            {
                                Text = t.WebSite,
                                Left = xpos,
                                Top = ypos + b.Height + 70,
                                AutoSize = true
                            };
                        tabPage.Controls.Add(txtWebsite);
                    }
                }


                // tabPage.Click += new EventHandler(tbpMainGraphic_Click);
                tabPage.Controls.Add(valueRadioButton);

                counter++;
            }
            return ycounter;
        }

        internal virtual void LoadCoverPageGraphicsSync(TaskScheduler uiScheduler, TabPage tabPage, string coverPageTitleForDefault)
        {
            string itemsByHeaderTypeSortBySortOrder =
           "<View><Query><Where><Eq><FieldRef Name='" + Constants.SharePointFields.HeaderType +
           "' /> <Value Type='Choice'>" + HeaderType + "</Value></Eq></Query><OrderBy><FieldRef Name='" +
           Constants.SharePointFields.SortOrder + "'/></OrderBy></View></Where>";

            var contextUrl = Settings.Default.SharePointContextUrl;
            var listTitle = Settings.Default.GraphicsPictureLibraryTitle;

            var sharePointPictureLibry = new SharePointPictureLibrary(contextUrl, listTitle, false, itemsByHeaderTypeSortBySortOrder);
            var presenter = new SharePointPictureLibraryPresenter(sharePointPictureLibry, this);
            var cacheName = HeaderType + BusinessLogic.Helpers.Constants.ControlNames.TabPageCoverPagesName;
          
             List<IThumbnail> th;
           
            if (Cache.Contains(cacheName))
            {
                th = presenter.GetPictureLibraryCoverPageThumnails();
                //th = ((List<IThumbnail>)_cache.Get(cacheName));
            }
            else
            {
                th = presenter.GetPictureLibraryCoverPageThumnails();
                Cache.Add(cacheName, th, new CacheItemPolicy());
            }

            DisplayGraphics(uiScheduler, tabPage, coverPageTitleForDefault, th, false);
        }

        internal virtual void LoadLogoGraphicsSync(TaskScheduler uiScheduler, TabPage tabPage, string logoTitleForDefault)
        {
            string contextUrl = Settings.Default.SharePointContextUrl;
            string listTitle = Settings.Default.GraphicsPictureLibraryTitleLogos;
            
            var sharePointPictureLibry = new SharePointPictureLibrary(contextUrl, listTitle, true, BusinessLogic.Helpers.Constants.SharePointQueries.AllItemsSortBySortOrder);
            var presenter = new SharePointPictureLibraryPresenter(sharePointPictureLibry, this);

            List<IThumbnail> logoStreams;

            var cacheName = HeaderType + "LogoStreams";

            if (Cache.Contains(cacheName))
            {
                logoStreams = presenter.GetPictureLibraryLogoThumbnails();
                //logoStreams = (List<IThumbnail>)_cache.Get(cacheName);
            }
            else
            {
                logoStreams = presenter.GetPictureLibraryLogoThumbnails();
                Cache.Add(cacheName, logoStreams, new CacheItemPolicy());
            }
           
            DisplayGraphics(uiScheduler, tabPage, logoTitleForDefault, logoStreams, true);

           
        }

        public virtual void DisplayMessage(string text, string caption)
        {
            MessageBox.Show(text, caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        protected T GetCachedTempalteObject<T>()
        {
            var template = (T)Cache.Get(Constants.CacheNames.RegenerateTemplate);
            Cache.Remove(Constants.CacheNames.RegenerateTemplate);
            return template;
        }
        public virtual void PopulateDocument(IBaseTemplate template)
        {
            //populate the content controls
            BasePresenter.PopulateData(template);

            //change the graphics selected
           // if (Streams == null) return;
            BasePresenter.PopulateGraphics(template, string.Empty, string.Empty);
        }

        public IEnumerable<Control> GetAll(Control control, Type type)
        {
            var controls = control.Controls.Cast<Control>();

            var enumerable = controls as IList<Control> ?? controls.ToList();
            foreach (Control c in enumerable.SelectMany(ctrl => GetAll(ctrl, type)).Concat(enumerable))
            {
                if (c.GetType() == type) yield return c;
            }
        }


        protected List<ISharePointListItem> LoadMajorTreeNodeTypes(string contextUrl, string listName)
        {            
            var list = new SharePointList(contextUrl, listName);
            var presenter = new SharePointListPresenter(list, this);
            return presenter.GetItems();
        }

        private List<IPolicyClass> LoadMinorPolicyTypes()
        {
            var list = new SharePointList(Settings.Default.SharePointContextUrl, Settings.Default.MinorPolicyClassesListName);
            var presenter = new SharePointListPresenter(list, this);
            return presenter.GetMinorPolicyItems();
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

                MajorItems = LoadMajorTreeNodeTypes(Settings.Default.SharePointContextUrl, Settings.Default.MajorPolicyClassesListName);
                Cache.Add(Constants.CacheNames.MajorPolicyClassItems, MajorItems,
                          new CacheItemPolicy());
            }

            if (Cache.Contains(Constants.CacheNames.MinorPolicyClassItems))
            {
                MinorItems =
                    ((List<IPolicyClass>)Cache.Get(Constants.CacheNames.MinorPolicyClassItems));
            }
            else
            {
                
                MinorItems = LoadMinorPolicyTypes();
                Cache.Add(Constants.CacheNames.MinorPolicyClassItems, MinorItems,
                          new CacheItemPolicy());
            }

            if (uiScheduler == null)
                WriteNodesToTree();
            else
                Task.Factory.StartNew(WriteNodesToTree, CancellationToken.None, TaskCreationOptions.None, uiScheduler);

        }

        private void WriteNodesToTree()
        {
            var tv = this.Controls.Find("tvaPolicies",true);
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
                        String.Equals(i.MajorClass, majorItem.Title, StringComparison.OrdinalIgnoreCase));
                foreach (var minorItem in found)
                {
                    Node childNode = AddChild(rootNode, minorItem.Title);                    
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
            if (_generateNewTemplate == true)
                return true;

                var r = MessageBox.Show(
                    Constants.Miscellaneous.ConfirmMsg,
                    @"Please Confirm", MessageBoxButtons.YesNo);

            if (r == DialogResult.No)
            {
                WizardBeingUpdated = true;
                var type = sender.GetType();
                if (type == typeof(CheckBox))
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


    }
}
