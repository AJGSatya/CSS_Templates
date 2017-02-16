using System;
using System.Collections.Generic;

using OAMPS.Office.BusinessLogic.Interfaces.SharePoint;

using Microsoft.SharePoint.Client;

namespace OAMPS.Office.Word.Models.SharePoint
{
    public class SharePointList : ISharePointList
    {
        private readonly List _list;
        private string _camlQuery;

        public SharePointList(string contextUrl, string listTitle, string itemsCamlQuery = "<View/>")
        {
            using (var clientContext = new ClientContext(contextUrl))
            {
                _camlQuery = itemsCamlQuery;
                _list = clientContext.Web.Lists.GetByTitle(listTitle);
                clientContext.Load(_list);
                clientContext.ExecuteQuery();
            }
        }

        public string Title
        {
            get { return _list.Title; }
        }

        public string FullUrl
        {
            get { return _list.DefaultViewUrl; }
        }
        
        public void DisplayMessage(string text, string caption)
        {
            throw new NotImplementedException();
        }

        public void UpdateCamlQuery(string query)
        {
            _camlQuery = query;
        }

        List<ISharePointListItem> ISharePointList.ListItems
        {
            get 
            {
                var items = new List<ISharePointListItem>();
                var query = new CamlQuery { ViewXml = _camlQuery };
                var sitems = _list.GetItems(query);
                _list.Context.Load(sitems, myitems => myitems.IncludeWithDefaultProperties(item => item.Id, item => item["Title"]));
                _list.Context.ExecuteQuery();

                foreach (ListItem item in sitems)
                {                    
                    ISharePointListItem fItem = new SharePointListItem(item);
                    items.Add(fItem);
                }

                //  return Enumerable.Cast<ISharePointListItem>(sitems.Select(item => new SharePointListItem(item))).ToList();
                return items;
            }
        }

        //TEST THIS 

        List<ISharePointListItem> ISharePointList.HelpListItems
        {
            get
            {
                var items = new List<ISharePointListItem>();
                var query = new CamlQuery { ViewXml = _camlQuery };
                var sitems = _list.GetItems(query);
                _list.Context.Load(sitems, myitems => myitems.IncludeWithDefaultProperties(item => item.Id, item => item["Template"], item => item["Title"], item => item["Content"]));
                _list.Context.ExecuteQuery();

                foreach (ListItem item in sitems)
                {
                    ISharePointListItem fItem = new SharePointListItem(item);
                    
                    items.Add(fItem);
                }

                //  return Enumerable.Cast<ISharePointListItem>(sitems.Select(item => new SharePointListItem(item))).ToList();
                return items;
            }
        }
    }
}
