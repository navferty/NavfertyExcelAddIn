using Microsoft.Office.Interop.Excel;

namespace NavfertyExcelAddIn.Commons
{
    public static class RangeExtensions
    {
        private static IRangeExtensionsImplementation rangeExtensionsImplementation
            = new RangeExtensionsImplementation();

        // to mock in tests. some Range interface methods return dynamic and can't be mocked directly
        // otherwise, some props like Range.Address[false,false] are too long to call with indexes
        public static void ResetImplementation(IRangeExtensionsImplementation newRangeExtensionsImplementation)
        {
            rangeExtensionsImplementation = newRangeExtensionsImplementation;
        }

        public static string GetFormula(this Range range)
            => rangeExtensionsImplementation.GetFormula(range);

        public static string GetRelativeAddress(this Range range)
            => rangeExtensionsImplementation.GetRelativeAddress(range);

        public static Range Union(this Range current, Range range)
            => rangeExtensionsImplementation.Union(current, range);

        public static void SetColor(this Range range, int ColorIndex)
            => rangeExtensionsImplementation.SetColor(range, ColorIndex);

        private class RangeExtensionsImplementation : IRangeExtensionsImplementation
        {
            private Application App => Globals.ThisAddIn.Application;
            public string GetFormula(Range range) => (string)range.Formula;

            public string GetRelativeAddress(Range range) => range.Address[false, false];

            public void SetColor(Range range, int colorIndex) => range.Interior.ColorIndex = colorIndex;

            public Range Union(Range r1, Range r2) => App.Union(r1, r2);
        }
    }

    public interface IRangeExtensionsImplementation
    {
        string GetFormula(Range range);
        string GetRelativeAddress(Range range);
        void SetColor(Range range, int colorIndex);
        Range Union(Range r1, Range r2);
    }
}
