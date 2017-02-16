using System;
using System.Collections.Generic;
using Microsoft.SharePoint.Client;
using OAMPS.Office.BusinessLogic.Interfaces.SharePoint;

namespace OAMPS.Office.Word.Models.SharePoint
{
    public class SharePointList : ISharePointList
    {
        //private readonly List _list;
        private readonly string _contextUrl;
        private readonly string _listName;
        private string _camlQuery;


        public SharePointList(string contextUrl, string listTitle, string itemsCamlQuery = "<View/>")
        {
            _camlQuery = itemsCamlQuery;
            _contextUrl = contextUrl;
            _listName = listTitle;


            //using (var clientContext = new ClientContext(contextUrl))
            //{
            //    _camlQuery = itemsCamlQuery;
            //    _list = clientContext.Web.Lists.GetByTitle(listTitle);
            //    clientContext.Load(_list);
            //    //clientContext.ExecuteQuery();
            //}
        }

        public string Title
        {
            //get { return _list.Title; }
            get { return string.Empty; }
        }

        public string FullUrl
        {
            //get { return _list.DefaultViewUrl; }
            get { return string.Empty; }
        }

        public void DisplayMessage(string text, string caption)
        {
            throw new NotImplementedException();
        }

        public ISharePointListItem GetListItemByTitle(string title)
        {
            using (var clientContext = new SharePointAdalClientContext(_contextUrl))
            {
                var test = "<View><Query><Where><Eq><FieldRef Name='Title' /><Value Type='Text'>" + title +
                           "</Value></Eq></Where></Query></View>";
                    // string.Format(Constants.SharePointQueries.GetItemByTitleQuery, title);
                var query = new CamlQuery
                {
                    ViewXml = test
                };
                var list = clientContext.Web.Lists.GetByTitle(_listName);
                var sitems = list.GetItems(query);
                clientContext.Load(sitems);
                clientContext.ExecuteQuery();
                return sitems.Count == 1 ? new SharePointListItem(sitems[0]) : null;
            }
        }

        public ISharePointListItem GetListItemById(string id)
        {
            using (var clientContext = new SharePointAdalClientContext(_contextUrl))
            {
                var test = "<View><Query><Where><Eq><FieldRef Name='ID' /><Value Type='Text'>" + id +
                           "</Value></Eq></Where></Query></View>";
                var query = new CamlQuery
                {
                    ViewXml = test
                };
                var list = clientContext.Web.Lists.GetByTitle(_listName);
                var sitems = list.GetItems(query);
                clientContext.Load(sitems, i => i.IncludeWithDefaultProperties(item => item.Id, item => item["Title"]));
                clientContext.ExecuteQuery();

                return sitems.Count == 1 ? new SharePointListItem(sitems[0]) : null;
            }
        }

        public ISharePointListItem GetListItemByLookupField(string fieldName, string value)
        {
            using (var clientContext = new SharePointAdalClientContext(_contextUrl))
            {
                var test = "<View><Query><Where><Eq><FieldRef Name='" + fieldName + "' /><Value Type='Lookup'>" + value +
                           "</Value></Eq></Where></Query></View>";
                var query = new CamlQuery
                {
                    ViewXml = test
                };
                var list = clientContext.Web.Lists.GetByTitle(_listName);
                var sitems = list.GetItems(query);
                clientContext.Load(sitems, i => i.IncludeWithDefaultProperties(item => item.Id, item => item["Title"]));
                clientContext.ExecuteQuery();

                return sitems.Count == 1 ? new SharePointListItem(sitems[0]) : null;
            }
        }

        public ISharePointListItem GetListItemByTextField(string fieldName, string value)
        {
            using (var clientContext = new SharePointAdalClientContext(_contextUrl))
            {
                var test = "<View><Query><Where><Eq><FieldRef Name='" + fieldName + "' /><Value Type='Text'>" + value +
                           "</Value></Eq></Where></Query></View>";
                var query = new CamlQuery
                {
                    ViewXml = test
                };
                var list = clientContext.Web.Lists.GetByTitle(_listName);
                var sitems = list.GetItems(query);
                clientContext.Load(sitems, i => i.IncludeWithDefaultProperties(item => item.Id, item => item["Title"]));
                clientContext.ExecuteQuery();

                return sitems.Count == 1 ? new SharePointListItem(sitems[0]) : null;
            }
        }


        List<ISharePointListItem> ISharePointList.ListItems
        {
            get
            {
                using (var clientContext = new SharePointAdalClientContext(_contextUrl))
                {
                    var items = new List<ISharePointListItem>();
                    var query = new CamlQuery {ViewXml = _camlQuery};
                    var list = clientContext.Web.Lists.GetByTitle(_listName);
                    var sitems = list.GetItems(query);
                    clientContext.Load(sitems,
                        myitems => myitems.IncludeWithDefaultProperties(item => item.Id, item => item["Title"]));
                    clientContext.ExecuteQuery();

                    foreach (var item in sitems)
                    {
                        ISharePointListItem fItem = new SharePointListItem(item);
                        items.Add(fItem);
                    }

                    return items;
                }
            }
        }

        List<ISharePointListItem> ISharePointList.HelpListItems
        {
            get
            {
                using (var clientContext = new SharePointAdalClientContext(_contextUrl))
                {
                    var items = new List<ISharePointListItem>();
                    var query = new CamlQuery {ViewXml = _camlQuery};
                    var list = clientContext.Web.Lists.GetByTitle(_listName);
                    var sitems = list.GetItems(query);
                    clientContext.Load(sitems,
                        myitems =>
                            myitems.IncludeWithDefaultProperties(item => item.Id, item => item["Template"],
                                item => item["Title"], item => item["Content"]));


                    clientContext.ExecuteQuery();

                    foreach (var item in sitems)
                    {
                        ISharePointListItem fItem = new SharePointListItem(item);

                        items.Add(fItem);
                    }

                    return items;
                }
            }
        }


        public void UpdateCamlQuery(Func<ISharePointListItem, bool> predicate)
        {
            throw new NotImplementedException();
        }

        public void AddItem(Dictionary<string, string> fieldValues)
        {
            using (var clientContext = new SharePointAdalClientContext(_contextUrl))
            {
                var itemCreateInfo = new ListItemCreationInformation();
                var list = clientContext.Web.Lists.GetByTitle(_listName);

                var listItem = list.AddItem(itemCreateInfo);
                foreach (var v in fieldValues)
                {
                    listItem[v.Key] = v.Value;
                }
                listItem.Update();
                clientContext.ExecuteQuery();
            }
        }

        public void UpdateField(string itemId, string fieldName, string value)
        {
            using (var clientContext = new SharePointAdalClientContext(_contextUrl))
            {
                var itemCreateInfo = new ListItemCreationInformation();
                var list = clientContext.Web.Lists.GetByTitle(_listName);


                var listItem = list.GetItemById(itemId);

                listItem[fieldName] = value;
                listItem.Update();
                clientContext.ExecuteQuery();
            }
        }

        public ISharePointListItem GetListItemByName(string name)
        {
            using (var clientContext = new SharePointAdalClientContext(_contextUrl))
            {
                var test = "<View><Query><Where><Eq><FieldRef Name='FileLeafRef' /><Value Type='File'>" + name +
                           "</Value></Eq></Where></Query></View>";
                    // string.Format(Constants.SharePointQueries.GetItemByTitleQuery, title);
                var query = new CamlQuery
                {
                    ViewXml = test
                };
                var list = clientContext.Web.Lists.GetByTitle(_listName);
                var sitems = list.GetItems(query);
                clientContext.Load(sitems);
                clientContext.ExecuteQuery();
                return sitems.Count == 1 ? new SharePointListItem(sitems[0]) : null;
            }
        }


        public ISharePointListItem GetListItemByServerRelativeUrl(string value)
        {
            using (var clientContext = new SharePointAdalClientContext(_contextUrl))
            {
                var test = "<View><Query><Where><Eq><FieldRef Name='FileRef' /><Value Type='Url'>" + value +
                           "</Value></Eq></Where></Query></View>";
                var query = new CamlQuery
                {
                    ViewXml = test
                };
                var list = clientContext.Web.Lists.GetByTitle(_listName);
                var sitems = list.GetItems(query);
                clientContext.Load(sitems,
                    i => i.IncludeWithDefaultProperties(item => item.Id, item => item["Title"], item => item["FileRef"]));
                clientContext.ExecuteQuery();

                return sitems.Count == 1 ? new SharePointListItem(sitems[0]) : null;
            }
        }

        public void UpdateCamlQuery(string query)
        {
            _camlQuery = query;
        }

        //}
        //    }
        //        return items;

        //        }                
        //            ISharePointListItem fItem = new SharePointListItem(item);
        //        {

        //        foreach (ListItem item in sitems)
        //        _list.Context.ExecuteQuery();
        //        _list.Context.Load(sitems, myitems => myitems.IncludeWithDefaultProperties(item => item.Id, item => item["Template"], item => item["Title"], item => item["Content"]));
        //        var sitems = _list.GetItems(query);
        //        var query = new CamlQuery { ViewXml = _camlQuery };
        //        var items = new List<IPolicyFragment>();
        //    {
        //    get

        //public List<IPolicyFragment> PolicyFragmentItems { 
    }
}