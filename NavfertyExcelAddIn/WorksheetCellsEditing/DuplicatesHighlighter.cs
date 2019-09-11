using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using NLog;

namespace NavfertyExcelAddIn.WorksheetCellsEditing
{
    public class DuplicatesHighlighter : IDuplicatesHighlighter
    {
        private readonly ILogger logger = LogManager.GetCurrentClassLogger();

        public void HighlightDuplicates(Range range)
        {
            var values = range.Cast<Range>();

            // select duplicate values (where count > 1)
            var valueGroups = ExtractDuplicates(values);

            var colors = GetUnionRanges(valueGroups);

            // set colorIndex for each range
            foreach (var currentRange in colors.Keys)
            {
			    currentRange.Interior.ColorIndex = colors[currentRange];
            }

        }

        private Dictionary<string, Range[]> ExtractDuplicates(IEnumerable<Range> cells)
        {
            return cells
                .Where(x => !string.IsNullOrEmpty(x.Value.ToString()))
                .GroupBy(x => x.ToString().ToLowerInvariant(), new ValueEqualityComparer())
                .Where(x => x.Count() > 1)
                .ToDictionary(x => x.Key, x => x.ToArray());
        }

        private Dictionary<Range, int> GetUnionRanges(Dictionary<string, Range[]> valueGroups)
        {
            var colors = new Dictionary<Range, int>();
            var colorIndexCounter = 3;

            // for each of duplicates group make range
            foreach (var group in valueGroups)
            {
                var unionRange = GetUnion(group.Value);
                colors.Add(unionRange, colorIndexCounter++ % 57); // fill with colors from 3 to 57

                logger.Debug(() =>
                    $"Grouped cells with value {group.Key}, range is {unionRange.Address}, {group.Value.Length} items");
            }
            return colors;
        }

        private Range GetUnion(IReadOnlyCollection<Range> cells)
        {
            var cellsArr = cells.ToArray();

            var app = Globals.ThisAddIn.Application;

            var firstCell = cellsArr.First();

            if (cells.Count == 1)
                return firstCell;

            return cellsArr
                // .Skip(1)
                .Aggregate(firstCell, (current, cell) => app.Union(current, cell));
        }

        private class ValueEqualityComparer : IEqualityComparer<string>
        {
            public bool Equals(string x, string y)
            {
                if (string.IsNullOrWhiteSpace(x) || string.IsNullOrWhiteSpace(y))
                {
                    return false;
                }

                return x.Equals(y, StringComparison.InvariantCultureIgnoreCase);
            }
            public int GetHashCode(string obj)
            {
                return obj.GetHashCode();
            }
        }
    }
}
