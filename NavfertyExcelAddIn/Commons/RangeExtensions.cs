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

        private class RangeExtensionsImplementation : IRangeExtensionsImplementation
        {
            public string GetFormula(Range range) => (string)range.Formula;

            public string GetRelativeAddress(Range range) => range.Address[false, false];
        }
    }

    public interface IRangeExtensionsImplementation
    {
        string GetFormula(Range range);
        string GetRelativeAddress(Range range);
    }
}
