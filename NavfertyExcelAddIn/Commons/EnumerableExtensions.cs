using System;
using System.Collections.Generic;

namespace NavfertyExcelAddIn.Commons
{
    public static class EnumerableExtensions
    {
        public static void ForEach<T>(this IEnumerable<T> source, Action<T> action)
        {
            foreach (T element in source)
            {
                action(element);
            }
        }

    }
}
