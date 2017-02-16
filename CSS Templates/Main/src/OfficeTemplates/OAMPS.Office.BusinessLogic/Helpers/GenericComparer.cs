using System;
using System.Collections.Generic;

namespace OAMPS.Office.BusinessLogic.Helpers
{
    public class GenericComparer<T> : IComparer<T>
    {
        // 0:Ascending, 1:Descending

        public GenericComparer(string sortExpression, int sortDirection)
        {
            SortExpression = sortExpression;
            SortDirection = sortDirection;
        }

        public GenericComparer()
        {
        }

        public string SortExpression { get; set; }
        public int SortDirection { get; set; }

        #region IComparer<T> Members

        public int Compare(T x, T y)
        {
            var propertyInfo = typeof (T).GetProperty(SortExpression);
            var obj1 = (IComparable) propertyInfo.GetValue(x, null);
            var obj2 = (IComparable) propertyInfo.GetValue(y, null);

            switch (SortDirection)
            {
                case 0:
                    return obj1.CompareTo(int.Parse(obj2.ToString()));
                default:
                    return obj2.CompareTo(int.Parse(obj1.ToString()));
            }
        }

        #endregion
    }
}