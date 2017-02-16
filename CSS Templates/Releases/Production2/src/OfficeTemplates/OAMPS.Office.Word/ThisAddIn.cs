using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.Caching;
using System.Timers;
using System.Windows.Forms;
using System.Xml;
using Microsoft.Office.Interop.Word;
using Microsoft.SharePoint.Client;
using Microsoft.Win32;
using OAMPS.Office.BusinessLogic;
using OAMPS.Office.Word.Views.Wizards;
using Office = Microsoft.Office.Core;
using OAMPS.Office.Word.Helpers;
using OAMPS.Office.Word.Views.Wizards.Popups;

namespace OAMPS.Office.Word
{
    public partial class ThisAddIn
    {
        public static bool IsWizzardRunning { get; set; }    
      //  Ribbon ribbon;
        Ribbon ribbon;
      //  private static System.Timers.Timer _timer;
      //  private static string _documentId;

        //protected override object RequestService(Guid serviceGuid)
        //{
        //    if (serviceGuid == typeof(Microsoft.Office.Core.IRibbonExtensibility).GUID)
        //    {
        //        if (ribbon == null)
        //            ribbon = new Ribbon1();
        //        return ribbon;
        //    }

        //    return base.RequestService(serviceGuid);
        //}

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon();
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Application.DocumentChange += Application_DocumentChange;
            ((ApplicationEvents4_Event)this.Application).NewDocument +=new ApplicationEvents4_NewDocumentEventHandler(ThisAddIn_NewDocument);
            ((ApplicationEvents4_Event)this.Application).DocumentOpen += new ApplicationEvents4_DocumentOpenEventHandler(ThisAddIn_DocumentOpen);
            //FixBackstageView();

         //   PeoplePicker.LoadUsersIntoCache();

            if (Application.Documents.Count > 0)
                Globals.ThisAddIn.Application.ActiveDocument.ActiveWindow.View.ShadeEditableRanges = 0;
        }

        void ThisAddIn_DocumentOpen(Document Doc)
        {            
            Doc.ActiveWindow.View.ShadeEditableRanges = 0;
        }

        //private static void OnTimedEvent(object source, ElapsedEventArgs e)
        //{
        //    return;
        //    LogDuration(_timer.Interval);
        //}

        //private static void LogDuration(double duration)
        //{
        //    if (!String.IsNullOrEmpty(_documentId))
        //    {
        //        System.Threading.Tasks.Task.Factory.StartNew(() =>
        //            {
        //                using (var clientContext = new ClientContext(Properties.Settings.Default.SharePointContextUrl))
        //                {
        //                    var list = clientContext.Web.Lists.GetByTitle("Duration Report");
        //                    clientContext.Load(list);
        //                    var query = new CamlQuery { ViewXml = String.Format(BusinessLogic.Helpers.Constants.SharePointQueries.LogDurationGetQuery,_documentId) };
        //                    var sitems = list.GetItems(query);
        //                    list.Context.Load(sitems,
        //                                      myitems =>
        //                                      myitems.Include(item => item.Id,item => item["Author"],item => item["Title"],item => item["Duration"]));
        //                    clientContext.ExecuteQuery();
                            
        //                    if (sitems.Count == 0)
        //                    {
        //                        var itemCreateInfo = new ListItemCreationInformation();
        //                        var listItem = list.AddItem(itemCreateInfo);
        //                        listItem["Duration"] = duration;
        //                        listItem["Title"] = _documentId;
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

        private void ThisAddIn_NewDocument(Document doc)
        {
            //   var cacheName = doc.GetPropertyValue(BusinessLogic.Helpers.Constants.WordDocumentProperties.CacheName);
            var loadType = Enums.FormLoadType.OpenDocument;
            ObjectCache cache = MemoryCache.Default;
            if (cache.Contains(BusinessLogic.Helpers.Constants.CacheNames.RegenerateTemplate))
            {
                //we have a cache so this is a re gen of a template
                loadType = Enums.FormLoadType.RegenerateTemplate;
            }
                
         //   _documentId = Guid.NewGuid().ToString();

        //    Globals.ThisAddIn.Application.ActiveDocument.UpdateOrCreatePropertyValue(BusinessLogic.Helpers.Constants.WordDocumentProperties.DocumentId, _documentId);
            
            Globals.ThisAddIn.Application.ActiveDocument.LoadWizard(loadType);


          //  StartTimer();

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        private void Application_DocumentChange()
        {
            ////log previous duration
          //  LogDuration(_timer == null ?0:_timer.Interval);

            //_documentId =
            //    Globals.ThisAddIn.Application.ActiveDocument.GetPropertyValue(
            //        BusinessLogic.Helpers.Constants.WordDocumentProperties.DocumentId,false );

            //StartTimer();

        }

        //private static void StartTimer()
        //{
        //    if (_timer != null)
        //    {
        //        _timer.Dispose();
        //    }

        //    if (!String.IsNullOrEmpty(_documentId))
        //    {
        //        _timer = new System.Timers.Timer(10000);
        //        _timer.Elapsed += new ElapsedEventHandler(OnTimedEvent);

        //        _timer.Interval = 10000;
        //        _timer.Enabled = true;
        //    }
        //}

        private void CreateSpotlightKeys()
        {
            RegistryKey spotlightkey = Registry.CurrentUser.OpenSubKey(BusinessLogic.Helpers.Constants.Spotlight.SpotlightKeyUrl, false);
            if ((spotlightkey == null))
            {
                RegistryKey key = Registry.CurrentUser.OpenSubKey(BusinessLogic.Helpers.Constants.Spotlight.SpotlightParentKeyUrl, true);
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

        private void DeleteAllContentSubKeys()
        {
            RegistryKey key = Registry.CurrentUser.OpenSubKey(BusinessLogic.Helpers.Constants.Spotlight.ContentKeyUrl, true);
            if (key != null)
            {
                var subKeyNames = key.GetSubKeyNames();
                foreach (var subKeyName in subKeyNames)
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

            foreach (var regKey in contentProviders)
            {
                var newProviders = new List<RegistryProviderKey>();
                if ((!string.IsNullOrEmpty(regKey.SiteUrl)))
                {
                    var serverRelativeUrl = regKey.ServiceUrl.Replace(regKey.SiteUrl, string.Empty);
                    var file = GenerateTempFile(regKey.SiteUrl, serverRelativeUrl);
                    using (XmlReader reader = XmlReader.Create(file))
                    {
                        while (reader.Read())
                        {
                            if ((reader.Name == "o:featuredtemplate"))
                            {
                                if ((reader.HasAttributes))
                                {
                                    var title = reader.GetAttribute("title");
                                    var source = reader.GetAttribute("source");
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

                var fInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(clientContext, documentUrl);
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
            RegistryKey key = Registry.CurrentUser.OpenSubKey(BusinessLogic.Helpers.Constants.Spotlight.ContentKeyUrl, true);
            if (key != null)
            {
                RegistryKey subKeyMain = key.CreateSubKey(providerName);
                if (subKeyMain != null)
                {
                    var subKeyWord = subKeyMain.CreateSubKey("WD1033");
                    string currentDate = DateTime.Now.ToString("yyyy-MM-dd");
                    if (subKeyWord != null)
                    {
                        subKeyWord.SetValue("Update", currentDate);
                        var subKeyFeaturedTemplates = subKeyWord.CreateSubKey("FeaturedTemplates");

                        if (subKeyFeaturedTemplates != null)
                        {
                            var subKeyOne = subKeyFeaturedTemplates.CreateSubKey("1");
                            if (subKeyOne != null)
                            {
                                subKeyOne.SetValue("enddate", DateTime.Now.AddYears(1).ToString("yyyy-MM-dd"));
                                subKeyOne.SetValue("startdate", DateTime.Now.AddYears(-1).ToString("yyyy-MM-dd"));

                                Int32 counter = 0;
                                foreach (var regKey in providerKeys)
                                {
                                    counter = counter + 1;
                                    var contentKey =
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
            var key = Registry.CurrentUser.OpenSubKey(BusinessLogic.Helpers.Constants.Spotlight.ProviderKeyUrl, false);
            if (key != null)
            {
                var subKeyNames = key.GetSubKeyNames();
                var providerKeys = new List<RegistryProviderKey>();

                foreach (string subKeyName in subKeyNames)
                {
                    var subKey = Registry.CurrentUser.OpenSubKey(string.Format("{0}\\{1}", BusinessLogic.Helpers.Constants.Spotlight.ProviderKeyUrl, subKeyName), false);
                    if (((subKey != null)))
                    {
                        var sk = subKey.GetValue("SiteUrl");
                        if (sk == null) continue;

                        providerKeys.Add(new RegistryProviderKey(subKeyName, subKey.GetValue("ServiceURL").ToString(), subKey.GetValue("SiteUrl").ToString()));
                    }
                }
                return providerKeys;
            }
            return null;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
