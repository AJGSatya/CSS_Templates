using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.Caching;
using System.Windows.Forms;
using System.Xml;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.SharePoint.Client;
using Microsoft.Win32;
using OAMPS.Office.BusinessLogic;
using OAMPS.Office.BusinessLogic.Helpers;
using OAMPS.Office.BusinessLogic.Loggers;
using OAMPS.Office.Word.Helpers;
using OAMPS.Office.Word.Helpers.LocalSharePoint;
using OAMPS.Office.Word.Models.SharePoint;
using OAMPS.Office.Word.Models.Word;
using OAMPS.Office.Word.Properties;
using OAMPS.Office.Word.Views.Word;
using Enums = OAMPS.Office.Word.Helpers.Enums;
using File = Microsoft.SharePoint.Client.File;
using Office = Microsoft.Office.Core;
using Task = System.Threading.Tasks.Task;
using Type = OAMPS.Office.BusinessLogic.Interfaces.Logging.Type;

namespace OAMPS.Office.Word
{
    public partial class ThisAddIn
    {
        //private SynchronizationContext _sync;
        public static bool IsWizzardRunning { get; set; }
        //Ribbon _ribbon;


        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon();
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            try
            {
                var logger = new EventViewerLogger();
                logger.Log(string.Format("Startup event"), Type.Information);

                ((ApplicationEvents4_Event) Application).NewDocument += ThisAddIn_NewDocument;
                (Application).DocumentOpen += ThisAddIn_DocumentOpen;
                if (Application.Documents.Count > 0)
                {
                    logger.Log(string.Format("Startup event - Application document count greater than 0"), Type.Information);
                    Globals.ThisAddIn.Application.ActiveDocument.ActiveWindow.View.ShadeEditableRanges = 0;
                    newTemplate(Globals.ThisAddIn.Application.ActiveDocument);
                        //required as they open from SharePoint links now too
                }
            }
            catch (Exception ex)
            {
                OnError(ex);
            }
        }

        private void OnError(Exception ex)
        {
            var logger = new EventViewerLogger();
            logger.Log(ex.ToString(), Type.Error);

#if DEBUG
            MessageBox.Show(ex.ToString(), @"sorry");
#endif
        }


        private void ThisAddIn_DocumentOpen(Document doc)
        {
            try
            {

                var logger = new EventViewerLogger();
                logger.Log(string.Format("DocumentOpen event"), Type.Information);

                if (doc == null)
                {
                    logger.Log(string.Format("DocumentOpen event - doc is null, returning"), Type.Information);
                    return;
                }

                doc.ActiveWindow.View.ShadeEditableRanges = 0;
                Ribbon.ribbon.Invalidate();


                var document = new OfficeDocument(doc);
                var documentId = document.GetPropertyValue(Constants.WordDocumentProperties.DocumentId);
                logger.Log(string.Format("DocumentOpen event - documentId = {0}", documentId), Type.Information);

                // disabled update check Nov 2017 DSZ
                //var fragmentsUsedPropertyValue = document.GetPropertyValue(Constants.WordDocumentProperties.UsedDateOfFragements);
                //var logoPropertyValue = document.GetPropertyValue(Constants.WordDocumentProperties.UsedDateOfLogo);
                //var themePropertyValue = document.GetPropertyValue(Constants.WordDocumentProperties.UsedDateOfTheme);
                //var mainTemplatePropertyValue =
                //    document.GetPropertyValue(Constants.WordDocumentProperties.BuiltInTitle) + ";" +
                //    document.GetPropertyValue(Constants.WordDocumentProperties.DocumentGeneratedDate);

                //var spList = ListFactory.Create("TemplateUpdateTracking");
                //var item = spList.GetListItemByTitle(documentId);
                //var shouldHide = "false";
                //var hideChosedDate = "";
                //if (item != null)
                //{
                //    shouldHide = item.GetFieldValue("Hide");
                //    hideChosedDate = item.GetFieldValue("Modified");
                //}

                //var task =
                //    Task.Factory.StartNew(
                //        () =>
                //            ExecuteUpdateChecker(shouldHide, fragmentsUsedPropertyValue, themePropertyValue,
                //                logoPropertyValue, mainTemplatePropertyValue, documentId, hideChosedDate));

                //System.Threading.Tasks.Task.Factory.StartNew(() => ThisAddIn.CheckStartupTasks(document), CancellationToken.None, TaskCreationOptions.None, uiScheduler);
            }
            catch (Exception ex)
            {
                OnError(ex);
            }
        }


        private void ExecuteUpdateChecker(string shouldHide, string fragmentsUsedPropertyValue, string theme,
            string logo, string mainTemplate, string documentId, string hideChosenDate)
        {
            var f = new TemplateUpdateChecker(shouldHide, fragmentsUsedPropertyValue, documentId);

            var showMe = f.CheckDocument(shouldHide, fragmentsUsedPropertyValue, theme, logo, mainTemplate);

            if (shouldHide.Equals("true", StringComparison.OrdinalIgnoreCase))
            {
                DateTime outChosenDate;
                if (DateTime.TryParse(hideChosenDate, out outChosenDate))
                {
                    if (outChosenDate > f.LatestChange)
                    {
                        showMe = false;
                    }
                }
            }

            if (showMe)
            {
                f.ShowDialog();
            }
            else
            {
                f.Dispose();
            }
        }

        private void ThisAddIn_NewDocument(Document doc)
        {
            try
            {
                //   var cacheName = doc.GetPropertyValue(BusinessLogic.Helpers.Constants.WordDocumentProperties.CacheName);
                newTemplate(doc);
            }
            catch (Exception ex)
            {
                OnError(ex);
            }
        }

        public static void SetProperty(string name, string value)
        {
            var doc = new OfficeDocument(Globals.ThisAddIn.Application.ActiveDocument);
            doc.UpdateOrCreatePropertyValue(name, value);
        }

        private void newTemplate(Document doc)
        {
            var loadType = Enums.FormLoadType.OpenDocument;
            ObjectCache cache = MemoryCache.Default;
            if (cache.Contains(Constants.CacheNames.RegenerateTemplate))
            {
                //we have a cache so this is a re gen of a template
                loadType = Enums.FormLoadType.RegenerateTemplate;
            }
            else if (cache.Contains(Constants.CacheNames.GenerateQuoteSlip))
            {
                loadType = Enums.FormLoadType.NoWizard;
                var quoteSlip = cache.GetCacheItem(Constants.CacheNames.GenerateQuoteSlip);
                cache.Remove(Constants.CacheNames.GenerateQuoteSlip);
                //var document = new Models.Word.OfficeDocument(doc);
                //var quoteSlipPresenter = new BusinessLogic.Presenters.QuoteSlipPresenter(document, null);

                //quoteSlipPresenter.InsertPolicySchedule();
            }
            else if (cache.Contains(Constants.CacheNames.NoWizard))
            {
                loadType = Enums.FormLoadType.NoWizard;
                cache.Remove(Constants.CacheNames.NoWizard);
            }
            else if (cache.Contains(Constants.CacheNames.ConvertWizard))
            {
                loadType = Enums.FormLoadType.ConvertWizard;
            }

            Globals.ThisAddIn.Application.ActiveDocument.UpdateOrCreatePropertyValue(
                Constants.WordDocumentProperties.DocumentId, Guid.NewGuid().ToString());
            Ribbon.ribbon.Invalidate();

            if (loadType != Enums.FormLoadType.NoWizard)
            {
                Globals.ThisAddIn.Application.ActiveDocument.LoadWizard(loadType);
                
            }
            else
            {
                doc.Application.DisplayDocumentInformationPanel = false;
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        ///     Required method for Designer support - do not modify
        ///     the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        #endregion

        #region Not Used Methods

        //void Application_DocumentBeforeClose(Document Doc, ref bool Cancel)
        //{
        //    try
        //    {
        //        StoreDurtion(Doc);
        //    }
        //    catch (Exception ex)
        //    {
        //    }
        //}

        //private static void StoreDurtion(Document Doc)
        //{
        //    var d = Doc.GetPropertyValue(BusinessLogic.Helpers.Constants.WordDocumentProperties.ActiveDocumentDuration, false);
        //    var documentId = Globals.ThisAddIn.Application.ActiveDocument.GetPropertyValue(BusinessLogic.Helpers.Constants.WordDocumentProperties.DocumentId, false);
        //    double outd = 0.0;
        //    double.TryParse(d, out outd);
        //    LogDurationToSharePoint(outd, documentId,Doc.Name,"");
        //}

        //private static void OnTimedEvent(object source, ElapsedEventArgs e)
        //{
        //    //return;
        //    LogDuration(_timer.Interval);
        //}

        //private static void LogDuration(double duration)
        //{
        //    var st = Globals.ThisAddIn.Application.ActiveDocument.GetPropertyValue(BusinessLogic.Helpers.Constants.WordDocumentProperties.ActiveDocumentDuration, false);
        //    var t = 0.0;
        //    if (!String.IsNullOrEmpty(st)) double.TryParse(st, out t);
        //    var s = t + duration;
        //    Globals.ThisAddIn.Application.ActiveDocument.UpdateOrCreatePropertyValue(BusinessLogic.Helpers.Constants.WordDocumentProperties.ActiveDocumentDuration, s.ToString(CultureInfo.InvariantCulture));
        //}

        //private static void LogDurationToSharePoint(double duration, string documentId, string documentName, string title)
        //{
        //    if (!String.IsNullOrEmpty(documentId))
        //    {
        //        System.Threading.Tasks.Task.Factory.StartNew(() =>
        //            {
        //                using (var clientContext = new ClientContext(Properties.Settings.Default.SharePointContextUrl))
        //                {
        //                    var list = clientContext.Web.Lists.GetByTitle("Duration Report");
        //                    clientContext.Load(list);
        //                    var query = new CamlQuery { ViewXml = String.Format(BusinessLogic.Helpers.Constants.SharePointQueries.LogDurationGetQuery, documentId) };
        //                    var sitems = list.GetItems(query);
        //                    list.Context.Load(sitems,
        //                                      myitems =>
        //                                      myitems.Include(item => item.Id, item => item["Author"], item => item["Title"], item => item["Duration"], item => item["DocumentID"], item => item["DocumentName"]));
        //                    clientContext.ExecuteQuery();

        //                    if (sitems.Count == 0)
        //                    {
        //                        var itemCreateInfo = new ListItemCreationInformation();
        //                        var listItem = list.AddItem(itemCreateInfo);
        //                        listItem["Duration"] = duration;
        //                        listItem["Title"] = title;
        //                        listItem["DocumentName"] = documentName;
        //                        listItem["DocumentID"] = documentId;
        //                        listItem.Update();
        //                    }
        //                    else
        //                    {
        //                        var d = sitems[0]["Duration"].ToString();
        //                        int n = 0;
        //                        int.TryParse(d, out n);
        //                        sitems[0]["Duration"] = duration + n;
        //                        sitems[0].Update();
        //                    }
        //                    clientContext.ExecuteQuery();
        //                }
        //            });
        //    }
        //}


        //private void Application_DocumentChange()
        //{
        //   //var d = DateTime.Now.TimeOfDay - _becameActiveDocumentOn;
        //   // _becameActiveDocumentOn = DateTime.Now.TimeOfDay;

        //   // var t = d.TotalSeconds / 60;
        //   // //log previous duration
        //   // LogDuration(t);
        //   // StartTimer();

        //}

        //private static void StartTimer()
        //{
        //    if (_timer != null)
        //    {
        //        _timer.Dispose();
        //    }

        //    var documentId = Globals.ThisAddIn.Application.ActiveDocument.GetPropertyValue(BusinessLogic.Helpers.Constants.WordDocumentProperties.DocumentId, false);
        //    if (!String.IsNullOrEmpty(documentId))
        //    {
        //        _timer = new System.Timers.Timer(60000);
        //        _timer.Elapsed += new ElapsedEventHandler(OnTimedEvent);

        //        _timer.Interval = 60000;
        //        _timer.Enabled = true;
        //    }
        //}


        // Removed by Friman - No longer needed since we use local templates
//        private void DeleteAllContentSubKeys()
//        {
//            var key = Registry.CurrentUser.OpenSubKey(Constants.Spotlight.ContentKeyUrl, true);
//            if (key != null)
//            {
//                var subKeyNames = key.GetSubKeyNames();
//                foreach (var subKeyName in subKeyNames)
//                {
//                    key.DeleteSubKeyTree(subKeyName);
//                }
//            }
//        }

//        private void FixBackstageView()
//        {
//            CreateSpotlightKeys();
//            DeleteAllContentSubKeys();
//            var contentProviders = GetRegistryKeys();

//            foreach (var regKey in contentProviders)
//            {
//                var newProviders = new List<RegistryProviderKey>();
//                if ((!string.IsNullOrEmpty(regKey.SiteUrl)))
//                {
//                    var serverRelativeUrl = regKey.ServiceUrl.Replace(regKey.SiteUrl, string.Empty);
//                    var file = GenerateTempFile(regKey.SiteUrl, serverRelativeUrl);
//                    using (var reader = XmlReader.Create(file))
//                    {
//                        while (reader.Read())
//                        {
//                            if ((reader.Name == "o:featuredtemplate"))
//                            {
//                                if ((reader.HasAttributes))
//                                {
//                                    var title = reader.GetAttribute("title");
//                                    var source = reader.GetAttribute("source");
//                                    newProviders.Add(new RegistryProviderKey(title, source, string.Empty));
//                                }
//                            }
//                        }
//                    }
//                    CreateContentRegistryKeys(regKey.KeyName, newProviders);
//                }
//            }
//        }

//        public string GenerateTempFile(string siteUrl, string documentUrl)
//        {
//            var tf = Path.GetTempFileName();
//            using (var clientContext = new ClientContext(siteUrl))
//            {
//                documentUrl = !documentUrl.StartsWith("/") ? "/" + documentUrl : documentUrl;

//                var tempContext = new SharePointUserDirectClientContext(clientContext.Url);
//                var fInfo = File.OpenBinaryDirect(tempContext, documentUrl);
//                var fs = new FileStream(tf, FileMode.Create);
//                var read = new byte[256];
//                var count = fInfo.Stream.Read(read, 0, read.Length);
//                while (count > 0)
//                {
//                    fs.Write(read, 0, count);
//                    count = fInfo.Stream.Read(read, 0, read.Length);
//                }
//                fs.Close();
//                fInfo.Stream.Close();
//            }
//            return tf;
//        }

//        private void CreateContentRegistryKeys(string providerName, IEnumerable<RegistryProviderKey> providerKeys)
//        {
//            var key = Registry.CurrentUser.OpenSubKey(Constants.Spotlight.ContentKeyUrl, true);
//            if (key != null)
//            {
//                var subKeyMain = key.CreateSubKey(providerName);
//                if (subKeyMain != null)
//                {
//                    var subKeyWord = subKeyMain.CreateSubKey("WD1033");
//                    var currentDate = DateTime.Now.ToString("yyyy-MM-dd");
//                    if (subKeyWord != null)
//                    {
//                        subKeyWord.SetValue("Update", currentDate);
//                        var subKeyFeaturedTemplates = subKeyWord.CreateSubKey("FeaturedTemplates");

//                        if (subKeyFeaturedTemplates != null)
//                        {
//                            var subKeyOne = subKeyFeaturedTemplates.CreateSubKey("1");
//                            if (subKeyOne != null)
//                            {
//                                subKeyOne.SetValue("enddate", DateTime.Now.AddYears(1).ToString("yyyy-MM-dd"));
//                                subKeyOne.SetValue("startdate", DateTime.Now.AddYears(-1).ToString("yyyy-MM-dd"));

//                                var counter = 0;
//                                foreach (var regKey in providerKeys)
//                                {
//                                    counter = counter + 1;
//                                    var contentKey =
//                                        subKeyOne.CreateSubKey(counter.ToString(CultureInfo.InvariantCulture));
//                                    if (contentKey != null)
//                                    {
//                                        contentKey.SetValue("source", regKey.ServiceUrl);
//                                        contentKey.SetValue("title", regKey.KeyName);
//                                    }
//                                }
//                            }
//                        }
//                    }
//                }
//            }
//        }

//        private IEnumerable<RegistryProviderKey> GetRegistryKeys()
//        {
//            var key = Registry.CurrentUser.OpenSubKey(Constants.Spotlight.ProviderKeyUrl, false);
//            if (key != null)
//            {
//                var subKeyNames = key.GetSubKeyNames();
//                var providerKeys = new List<RegistryProviderKey>();

//// ReSharper disable LoopCanBeConvertedToQuery
//                foreach (var subKeyName in subKeyNames)
//// ReSharper restore LoopCanBeConvertedToQuery
//                {
//                    var subKey =
//                        Registry.CurrentUser.OpenSubKey(
//                            string.Format("{0}\\{1}", Constants.Spotlight.ProviderKeyUrl, subKeyName), false);
//                    if (((subKey != null)))
//                    {
//                        var sk = subKey.GetValue("SiteUrl");
//                        if (sk == null) continue;

//                        providerKeys.Add(new RegistryProviderKey(subKeyName, subKey.GetValue("ServiceURL").ToString(),
//                            subKey.GetValue("SiteUrl").ToString()));
//                    }
//                }
//                return providerKeys;
//            }
//            return null;
//        }

//        private void CreateSpotlightKeys()
//        {
//            var spotlightkey = Registry.CurrentUser.OpenSubKey(Constants.Spotlight.SpotlightKeyUrl, false);
//            if ((spotlightkey == null))
//            {
//                var key = Registry.CurrentUser.OpenSubKey(Constants.Spotlight.SpotlightParentKeyUrl, true);
//                if (key != null)
//                {
//                    var subKeySpotlight = key.CreateSubKey("Spotlight");
//                    if (subKeySpotlight != null)
//                    {
//                        subKeySpotlight.CreateSubKey("Providers");
//                        subKeySpotlight.CreateSubKey("Content");
//                    }
//                }
//            }
//        }

        #endregion
    }
}