using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Client;
using OAMPS.Office.BusinessLogic.Interfaces.SharePoint;

namespace OAMPS.Office.Word.Models.SharePoint
{
    public class MinorPolicyClassList //: IMinorPolicyClassList
    {
     //private readonly List _list;
     //   private readonly string _camlQuery;

     //   public MinorPolicyClassList(string contextUrl, string listTitle, string itemsCamlQuery = "<View/>")
     //   {
     //       using (var clientContext = new ClientContext(contextUrl))
     //       {
     //           _camlQuery = itemsCamlQuery;
     //           _list = clientContext.Web.Lists.GetByTitle(listTitle);
     //           clientContext.Load(_list);
     //       }
     //   }

     //   public string Title
     //   {
     //       get { return _list.Title; }
     //   }

     //   public string FullUrl
     //   {
     //       get { return _list.DefaultViewUrl; }
     //   }
        
     //   public void DisplayMessage(string text, string caption)
     //   {
     //       throw new NotImplementedException();
     //   }

     //   List<IMinorPolicyClassListItem> IMinorPolicyClassList.ListItems
     //   {
     //       get 
     //       {
     //           var items = new List<IMinorPolicyClassListItem>();
     //           var query = new CamlQuery { ViewXml = _camlQuery };
     //           var sitems = _list.GetItems(query);
     //           _list.Context.Load(sitems);
     //           _list.Context.ExecuteQuery();

     //           foreach (ListItem item in sitems)
     //           {
     //               IMinorPolicyClassListItem fItem = new MinorPolicyClassListItem(item);
     //               items.Add(fItem);
     //           }

     //           //  return Enumerable.Cast<ISharePointListItem>(sitems.Select(item => new SharePointListItem(item))).ToList();
     //           return items;
     //       }
     //   }

    }
}
