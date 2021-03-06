﻿using System;
using System.Collections.Generic;
using Microsoft.SharePoint.Client;
using OAMPS.Office.BusinessLogic.Helpers;
using OAMPS.Office.BusinessLogic.Interfaces.SharePoint;

namespace OAMPS.Office.Word.Models.SharePoint
{
    public class SharePointPictureLibrary : ISharePointPictureLibrary
    {
        private readonly string _camlQuery;
        private readonly List _list;

        private readonly bool _isLogo;

        public SharePointPictureLibrary(string contextUrl, string listTitle, bool isLogo, string itemsCamlQuery = "<View/>")
        {
            using (var clientContext = new ClientContext(contextUrl))
            {
                _camlQuery = itemsCamlQuery;
                _list = clientContext.Web.Lists.GetByTitle(listTitle);
                clientContext.Load(_list);
                _isLogo = isLogo;
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


        List<ISharePointPictureLibraryItem> ISharePointPictureLibrary.ListItems
        {
            get
            {
                var items = new List<ISharePointPictureLibraryItem>();

                var query = new CamlQuery {ViewXml = _camlQuery};

                ListItemCollection sitems = _list.GetItems(query);
                var context = (ClientContext) _list.Context;

                if (_isLogo)
                {
                    context.Load(sitems); //, i => i.Include(o => o["FileRef"]));    
                }
                else
                {
                    context.Load(sitems, x => x.IncludeWithDefaultProperties(item => item.Id, item => item[Constants.SharePointFields.ShortDescription], item => item[Constants.SharePointFields.LongDescription]));
                }

                context.ExecuteQuery();

// ReSharper disable LoopCanBeConvertedToQuery
                foreach (ListItem item in sitems)
// ReSharper restore LoopCanBeConvertedToQuery
                {
                    ISharePointPictureLibraryItem fItem = new SharePointPictureLibraryItem(item, context);
                    items.Add(fItem);
                }

                //  return Enumerable.Cast<ISharePointPictureLibraryItem>(sitems.Select(item => new SharePointPictureLibraryItem(item, context))).ToList();

                return items;
            }
        }
    }
}