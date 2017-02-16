using System;
using System.Collections.Generic;

namespace OAMPS.Office.BusinessLogic.Helpers
{
    public class AnonymousComparer<T> : IComparer<T>
    {
        private readonly Comparison<T> _comparison;

        public AnonymousComparer(Comparison<T> comparison)
        {
            if (comparison == null)
                throw new ArgumentNullException(nameof(comparison));
            _comparison = comparison;
        }

        public int Compare(T x, T y)
        {
            return _comparison(x, y);
        }
    }
}