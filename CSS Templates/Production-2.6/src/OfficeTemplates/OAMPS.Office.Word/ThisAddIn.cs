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
using Enums = OAMPS.Office.Word.Helpers.Enums;
using File = Microsoft.SharePoint.Client.File;
using Office = Microsoft.Office.Core;
using Type = OAMPS.Office.BusinessLogic.Interfaces.Logging.Type;

namespace OAMPS.Office.Word
{
    public partial class ThisAddIn
    {
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
                ((ApplicationEvents4_Event) Application).NewDocument += ThisAddIn_NewDocument;
                (Application).DocumentOpen += ThisAddIn_DocumentOpen;
                if (Application.Documents.Count > 0)
                    Globals.ThisAddIn.Application.ActiveDocument.ActiveWindow.View.ShadeEditableRanges = 0;
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
                if (doc == null)
                    return;
   
                 doc.ActiveWindow.View.ShadeEditableRanges = 0;
                Ribbon.ribbon.Invalidate();
            }
            catch (Exception ex)
            {
                OnError(ex);
            }
        }

        private void ThisAddIn_NewDocument(Document doc)
        {
            try
            {
                //   var cacheName = doc.GetPropertyValue(BusinessLogic.Helpers.Constants.WordDocumentProperties.CacheName);
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
                    CacheItem quoteSlip = cache.GetCacheItem(Constants.CacheNames.GenerateQuoteSlip);
                    cache.Remove(Constants.CacheNames.GenerateQuoteSlip);
                    //var document = new Models.Word.OfficeDocument(doc);
                    //var quoteSlipPresenter = new BusinessLogic.Presenters.QuoteSlipPresenter(document, null);

                    //quoteSlipPresenter.InsertPolicySchedule();
                }

                Globals.ThisAddIn.Application.ActiveDocument.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.DocumentId, Guid.NewGuid().ToString());
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
            catch (Exception ex)
            {
                OnError(ex);
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
        }

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


        private void DeleteAllContentSubKeys()
        {
            RegistryKey key = Registry.CurrentUser.OpenSubKey(Constants.Spotlight.ContentKeyUrl, true);
            if (key != null)
            {
                string[] subKeyNames = key.GetSubKeyNames();
                foreach (string subKeyName in subKeyNames)
                {
                    key.DeleteSubKeyTree(subKeyName);
                }
            }
        }

        private void FixBackstageView()
        {
            CreateSpotlightKeys();
            DeleteAllContentSubKeys();
            IEnumerable<RegistryProviderKey> contentProviders = GetRegistryKeys();

            foreach (RegistryProviderKey regKey in contentProviders)
            {
                var newProviders = new List<RegistryProviderKey>();
                if ((!string.IsNullOrEmpty(regKey.SiteUrl)))
                {
                    string serverRelativeUrl = regKey.ServiceUrl.Replace(regKey.SiteUrl, string.Empty);
                    string file = GenerateTempFile(regKey.SiteUrl, serverRelativeUrl);
                    using (XmlReader reader = XmlReader.Create(file))
                    {
                        while (reader.Read())
                        {
                            if ((reader.Name == "o:featuredtemplate"))
                            {
                                if ((reader.HasAttributes))
                                {
                                    string title = reader.GetAttribute("title");
                                    string source = reader.GetAttribute("source");
                                    newProviders.Add(new RegistryProviderKey(title, source, string.Empty));
                                }
                            }
                        }
                    }
                    CreateContentRegistryKeys(regKey.KeyName, newProviders);
                }
            }
        }

        public string GenerateTempFile(string siteUrl, string documentUrl)
        {
            string tf = Path.GetTempFileName();
            using (var clientContext = new ClientContext(siteUrl))
            {
                documentUrl = !documentUrl.StartsWith("/") ? "/" + documentUrl : documentUrl;

                FileInformation fInfo = File.OpenBinaryDirect(clientContext, documentUrl);
                var fs = new FileStream(tf, FileMode.Create);
                var read = new byte[256];
                int count = fInfo.Stream.Read(read, 0, read.Length);
                while (count > 0)
                {
                    fs.Write(read, 0, count);
                    count = fInfo.Stream.Read(read, 0, read.Length);
                }
                fs.Close();
                fInfo.Stream.Close();
            }
            return tf;
        }

        private void CreateContentRegistryKeys(string providerName, IEnumerable<RegistryProviderKey> providerKeys)
        {
            RegistryKey key = Registry.CurrentUser.OpenSubKey(Constants.Spotlight.ContentKeyUrl, true);
            if (key != null)
            {
                RegistryKey subKeyMain = key.CreateSubKey(providerName);
                if (subKeyMain != null)
                {
                    RegistryKey subKeyWord = subKeyMain.CreateSubKey("WD1033");
                    string currentDate = DateTime.Now.ToString("yyyy-MM-dd");
                    if (subKeyWord != null)
                    {
                        subKeyWord.SetValue("Update", currentDate);
                        RegistryKey subKeyFeaturedTemplates = subKeyWord.CreateSubKey("FeaturedTemplates");

                        if (subKeyFeaturedTemplates != null)
                        {
                            RegistryKey subKeyOne = subKeyFeaturedTemplates.CreateSubKey("1");
                            if (subKeyOne != null)
                            {
                                subKeyOne.SetValue("enddate", DateTime.Now.AddYears(1).ToString("yyyy-MM-dd"));
                                subKeyOne.SetValue("startdate", DateTime.Now.AddYears(-1).ToString("yyyy-MM-dd"));

                                Int32 counter = 0;
                                foreach (RegistryProviderKey regKey in providerKeys)
                                {
                                    counter = counter + 1;
                                    RegistryKey contentKey =
                                        subKeyOne.CreateSubKey(counter.ToString(CultureInfo.InvariantCulture));
                                    if (contentKey != null)
                                    {
                                        contentKey.SetValue("source", regKey.ServiceUrl);
                                        contentKey.SetValue("title", regKey.KeyName);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        private IEnumerable<RegistryProviderKey> GetRegistryKeys()
        {
            RegistryKey key = Registry.CurrentUser.OpenSubKey(Constants.Spotlight.ProviderKeyUrl, false);
            if (key != null)
            {
                string[] subKeyNames = key.GetSubKeyNames();
                var providerKeys = new List<RegistryProviderKey>();

// ReSharper disable LoopCanBeConvertedToQuery
                foreach (string subKeyName in subKeyNames)
// ReSharper restore LoopCanBeConvertedToQuery
                {
                    RegistryKey subKey = Registry.CurrentUser.OpenSubKey(string.Format("{0}\\{1}", Constants.Spotlight.ProviderKeyUrl, subKeyName), false);
                    if (((subKey != null)))
                    {
                        object sk = subKey.GetValue("SiteUrl");
                        if (sk == null) continue;

                        providerKeys.Add(new RegistryProviderKey(subKeyName, subKey.GetValue("ServiceURL").ToString(), subKey.GetValue("SiteUrl").ToString()));
                    }
                }
                return providerKeys;
            }
            return null;
        }

        private void CreateSpotlightKeys()
        {
            RegistryKey spotlightkey = Registry.CurrentUser.OpenSubKey(Constants.Spotlight.SpotlightKeyUrl, false);
            if ((spotlightkey == null))
            {
                RegistryKey key = Registry.CurrentUser.OpenSubKey(Constants.Spotlight.SpotlightParentKeyUrl, true);
                if (key != null)
                {
                    RegistryKey subKeySpotlight = key.CreateSubKey("Spotlight");
                    if (subKeySpotlight != null)
                    {
                        subKeySpotlight.CreateSubKey("Providers");
                        subKeySpotlight.CreateSubKey("Content");
                    }
                }
            }
        }

        #endregion

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
    }
}