using System;
using System.Collections.Generic;
using System.Linq;
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
            using (var clientContext = new ClientContext(_contextUrl))
            {
                var test = "<View><Query><Where><Eq><FieldRef Name='Title' /><Value Type='Text'>" + title + "</Value></Eq></Where></Query></View>"; // string.Format(Constants.SharePointQueries.GetItemByTitleQuery, title);
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

        public ISharePointListItem GetListItemByName(string name)
        {
            using (var clientContext = new ClientContext(_contextUrl))
            {
                var test = "<View><Query><Where><Eq><FieldRef Name='FileLeafRef' /><Value Type='File'>" + name + "</Value></Eq></Where></Query></View>"; // string.Format(Constants.SharePointQueries.GetItemByTitleQuery, title);
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
          

            using (var clientContext = new ClientContext(_contextUrl))
            {
                var test = "<View><Query><Where><Eq><FieldRef Name='ID' /><Value Type='Text'>" + id + "</Value></Eq></Where></Query></View>"; 
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
            using (var clientContext = new ClientContext(_contextUrl))
            {
                var test = "<View><Query><Where><Eq><FieldRef Name='" + fieldName + "' /><Value Type='Lookup'>" + value + "</Value></Eq></Where></Query></View>";
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
            using (var clientContext = new ClientContext(_contextUrl))
            {
                var test = "<View><Query><Where><Eq><FieldRef Name='" + fieldName + "' /><Value Type='Text'>" + value + "</Value></Eq></Where></Query></View>";
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


        public ISharePointListItem GetListItemByServerRelativeUrl(string value)
        {
            using (var clientContext = new ClientContext(_contextUrl))
            {
                var test = "<View><Query><Where><Eq><FieldRef Name='FileRef' /><Value Type='Url'>" + value + "</Value></Eq></Where></Query></View>";
                var query = new CamlQuery
                {
                    ViewXml = test
                };
                var list = clientContext.Web.Lists.GetByTitle(_listName);
                var sitems = list.GetItems(query);
                clientContext.Load(sitems, i => i.IncludeWithDefaultProperties(item => item.Id, item => item["Title"], item => item["FileRef"]));
                clientContext.ExecuteQuery();

                return sitems.Count == 1 ? new SharePointListItem(sitems[0]) : null;
            }
        }


        List<ISharePointListItem> ISharePointList.ListItems
        {
            get
            {
                using (var clientContext = new ClientContext(_contextUrl))
                {
                    var items = new List<ISharePointListItem>();
                    var query = new CamlQuery {ViewXml = _camlQuery};
                    List list = clientContext.Web.Lists.GetByTitle(_listName);
                    ListItemCollection sitems = list.GetItems(query);
                    clientContext.Load(sitems, myitems => myitems.IncludeWithDefaultProperties(item => item.Id, item => item["Title"]));
                    clientContext.ExecuteQuery();

// ReSharper disable LoopCanBeConvertedToQuery
                    foreach (ListItem item in sitems)
// ReSharper restore LoopCanBeConvertedToQuery
                    {
                        ISharePointListItem fItem = new SharePointListItem(item);
                        items.Add(fItem);
                    }

                    //  return Enumerable.Cast<ISharePointListItem>(sitems.Select(item => new SharePointListItem(item))).ToList();
                    return items;
                }
            }
        }

//        public List<ISharePointListItem> ListItems2
//        {
//            get
//            {

//                var test = @"<View>          
//       
//          
//          <Joins>
//            <Join Type='LEFT' ListAlias='CamlLookList'>
//            
//              <Eq>
//                <FieldRef Name='PolicyLookup_x003a_ID' RefType='ID' />
//                <FieldRef List='CamlLookList' Name='ID' />
//              </Eq>
//            </Join>
//          </Joins>
//        </View>";


//                 //var _list2 = clientContext.Web.Lists.GetByTitle("CamlTest");
//                 //   var field = _list2.Fields.GetByInternalNameOrTitle("Policy_x0020_Type_x003a_ID");
//                 //   clientContext.Load(field);
//                 //   clientContext.ExecuteQuery();                

//                var items = new List<ISharePointListItem>();
//                var query = new CamlQuery { ViewXml =  test};
//                var sitems = _list.GetItems(query);

//                _list.Context.Load(sitems);
//                _list.Context.ExecuteQuery();

//                foreach (ListItem item in sitems)
//                {
//                    ISharePointListItem fItem = new SharePointListItem(item);
//                    items.Add(fItem);
//                }

//                //  return Enumerable.Cast<ISharePointListItem>(sitems.Select(item => new SharePointListItem(item))).ToList();
//                return items;
//            }
//        }

        //TEST THIS 

        List<ISharePointListItem> ISharePointList.HelpListItems
        {
            get
            {
                using (var clientContext = new ClientContext(_contextUrl))
                {
                    var items = new List<ISharePointListItem>();
                    var query = new CamlQuery {ViewXml = _camlQuery};
                    List list = clientContext.Web.Lists.GetByTitle(_listName);
                    ListItemCollection sitems = list.GetItems(query);
                    clientContext.Load(sitems, myitems => myitems.IncludeWithDefaultProperties(item => item.Id, item => item["Template"], item => item["Title"], item => item["Content"]));
                    clientContext.ExecuteQuery();

// ReSharper disable LoopCanBeConvertedToQuery
                    foreach (ListItem item in sitems)
// ReSharper restore LoopCanBeConvertedToQuery
                    {
                        ISharePointListItem fItem = new SharePointListItem(item);

                        items.Add(fItem);
                    }

                    //  return Enumerable.Cast<ISharePointListItem>(sitems.Select(item => new SharePointListItem(item))).ToList();
                    return items;
                }
            }
        }

        public void AddItem(Dictionary<string, string> fieldValues)
        {
            using (var clientContext = new ClientContext(_contextUrl))
            {
                var itemCreateInfo = new ListItemCreationInformation();
                List list = clientContext.Web.Lists.GetByTitle(_listName);

                ListItem listItem = list.AddItem(itemCreateInfo);
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
            using (var clientContext = new ClientContext(_contextUrl))
            {
                var itemCreateInfo = new ListItemCreationInformation();
                List list = clientContext.Web.Lists.GetByTitle(_listName);


                ListItem listItem = list.GetItemById(itemId);

                listItem[fieldName] = value;
                listItem.Update();
                clientContext.ExecuteQuery();
            }
        }

        public void UpdateCamlQuery(string query)
        {
            _camlQuery = query;
        }

        //public List<IPolicyFragment> PolicyFragmentItems { 
        //    get
        //    {
        //        var items = new List<IPolicyFragment>();
        //        var query = new CamlQuery { ViewXml = _camlQuery };
        //        var sitems = _list.GetItems(query);
        //        _list.Context.Load(sitems, myitems => myitems.IncludeWithDefaultProperties(item => item.Id, item => item["Template"], item => item["Title"], item => item["Content"]));
        //        _list.Context.ExecuteQuery();

        //        foreach (ListItem item in sitems)
        //        {
        //            ISharePointListItem fItem = new SharePointListItem(item);

        //        }                
        //        return items;
        //    }
        //}
    }
}