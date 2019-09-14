using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using NavfertyExcelAddIn.Commons;
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
            foreach (var rangeByColor in colors)
            {
			    rangeByColor.Value.SetColor(rangeByColor.Key);
            }

        }

        private Dictionary<string, Range[]> ExtractDuplicates(IEnumerable<Range> cells)
        {
            return cells
                .GroupBy(x => ((object)x.Value).ToString().ToLowerInvariant(), new ValueEqualityComparer())
                .Where(x => x.Count() > 1 && !string.IsNullOrEmpty(x.Key))
                .ToDictionary(x => x.Key, x => x.ToArray());
        }

        private Dictionary<int, Range> GetUnionRanges(Dictionary<string, Range[]> valueGroups)
        {
            var rangesByColors = new Dictionary<int, Range>();
            var colorIndexCounter = 3;

            // for each of duplicates group make range
            foreach (var group in valueGroups)
            {
                var unionRange = group.Value.Aggregate((current, cell) => current.Union(cell));
                
                // fill with colors from 3 to 57
                int colorIndex = colorIndexCounter++ % 57;

                if (rangesByColors.TryGetValue(colorIndex, out var range))
                {
                    rangesByColors[colorIndex] = range.Union(unionRange);
                    logger.Debug(() =>
                        $"For value {group.Key} add range {unionRange.GetRelativeAddress()} to union with same " +
                        $"colorIndex {colorIndex}, with {group.Value.Length} items");
                }
                else
                {
                    rangesByColors.Add(colorIndex, unionRange);
                    logger.Debug(() =>
                        $"Grouped cells with value {group.Key}, range is {unionRange.GetRelativeAddress()}, {group.Value.Length} items");
                }
            }
            return rangesByColors;
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
