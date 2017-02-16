using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OAMPS.Office.BusinessLogic.Interfaces.SharePoint;
using OAMPS.Office.BusinessLogic.Loggers;
using OAMPS.Office.Word.Helpers.LocalSharePoint;
using Type = OAMPS.Office.BusinessLogic.Interfaces.Logging.Type;

namespace OAMPS.Office.Word.Models.Local
{
    public class LocalList: ISharePointList
    {
        private EventViewerLogger _logger = new EventViewerLogger();

        public string Title { get; }

        public List<ISharePointListItem> ListItems {
            get
            {
                if (_sortBySortOrder)
                {
                    return _spListItems
                        .Where(_predicate)
                        .OrderBy(x => Convert.ToInt32(x.GetFieldValue("SortOrder")))
                        .ToList();
                }
                else
                {
                    return _spListItems
                        .Where(_predicate)
                        .OrderBy(x => x.Title)
                        .ToList();
                }
            }
        }
        public List<ISharePointListItem> HelpListItems => ListItems;

        private Func<ISharePointListItem, bool> _predicate;
        private readonly IEnumerable<ISharePointListItem> _spListItems;
        private readonly bool _sortBySortOrder; // default by Title, set to true to use SortORder

        public LocalList(string listTitle, Func<ISharePointListItem, bool> predicate, bool sortBySortOrder = false)
        {
            _logger.Log("Initialise local list: " + listTitle, Type.Information);

            _predicate = predicate;
            _sortBySortOrder = sortBySortOrder;
            Title = listTitle;

            var listLoader = new ListLoader();
            var listItems = new ListItems();
            _spListItems = new List<ISharePointListItem>();

            try
            {
                listItems = listLoader.GetItems(listTitle);
                _logger.Log($"Loaded {listItems.Count} list items for list {listTitle}", Type.Information);

                _spListItems = listItems
                    .Select(x => new LocalListItem(x))
                    .Cast<ISharePointListItem>().ToList();
            }
            catch (Exception ex)
            {
                _logger.Log(string.Format("List not loaded {0}, does the list exists and has items? Message: {1}", listTitle, ex.Message), Type.Warning);
            }
            
        }

        public void DisplayMessage(string text, string caption)
        {
            // unused
            throw new NotImplementedException();
        }
        
        public ISharePointListItem GetListItemByTitle(string title)
        {
            return ListItems.FirstOrDefault(x => x.Title.Equals(title, StringComparison.InvariantCultureIgnoreCase));
        }

        public ISharePointListItem GetListItemById(string id)
        {
            return GetListItemByTextField("ID", id);
        }

        public ISharePointListItem GetListItemByLookupField(string fieldName, string value)
        {
            return ListItems.FirstOrDefault(x => x.GetLookupFieldValue(fieldName).Equals(value, StringComparison.InvariantCultureIgnoreCase));
        }

        public ISharePointListItem GetListItemByTextField(string fieldName, string value)
        {
            return ListItems.FirstOrDefault(x => x.GetFieldValue(fieldName).Equals(value, StringComparison.InvariantCultureIgnoreCase));
        }

        public ISharePointListItem GetListItemByName(string name)
        {
            return GetListItemByTextField("FileLeafRef", name);
        }

        public void UpdateCamlQuery(Func<ISharePointListItem, bool> predicate)
        {
            _predicate = predicate;
        }

        public void AddItem(Dictionary<string, string> fieldValues)
        {
            throw new NotImplementedException();
        }

        public void UpdateField(string itemId, string fieldName, string value)
        {
            throw new NotImplementedException();
        }
    }
}
