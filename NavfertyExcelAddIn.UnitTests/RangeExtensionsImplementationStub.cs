using System.Collections.Generic;
using NavfertyExcelAddIn.Commons;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace NavfertyExcelAddIn.UnitTests
{
    public class RangeExtensionsImplementationStub : IRangeExtensionsImplementation
    {
        public IReadOnlyCollection<Range> GetFormulaInvocations => getFormulaInvocations;
        public IReadOnlyCollection<Range> GetRelativeAddressInvocations => getRelativeAddressInvocations;
        public IReadOnlyCollection<(Range, int)> SetColorInvocations => setColorInvocations;
        public IReadOnlyCollection<(Range, Range)> UnionInvocations => unionInvocations;

        private readonly List<Range> getFormulaInvocations = new List<Range>();
        private readonly List<Range> getRelativeAddressInvocations = new List<Range>();
        private readonly List<(Range, int)> setColorInvocations = new List<(Range, int)>();
        private readonly List<(Range, Range)> unionInvocations = new List<(Range, Range)>();

        public string GetFormula(Range range)
        {
            getFormulaInvocations.Add(range);
            return "=1+1";
        }

        public string GetRelativeAddress(Range range)
        {
            getRelativeAddressInvocations.Add(range);
            return "A1";
        }

        public void SetColor(Range range, int colorIndex)
        {
            setColorInvocations.Add((range, colorIndex));
        }

        public Range Union(Range r1, Range r2)
        {
            unionInvocations.Add((r1, r2));
            return r1;
        }
    }
}
