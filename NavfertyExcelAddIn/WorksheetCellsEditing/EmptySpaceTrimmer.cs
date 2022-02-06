using System.Linq;

using Microsoft.Office.Interop.Excel;

using Navferty.Common;

using NavfertyExcelAddIn.Commons;

using NLog;

namespace NavfertyExcelAddIn.WorksheetCellsEditing
{
    public class EmptySpaceTrimmer : IEmptySpaceTrimmer
    {
        private static readonly ILogger logger = LogManager.GetCurrentClassLogger();

        public void TrimExtraSpaces(Range range)
        {
            logger.Info($"Trim spaces for range {range.GetRelativeAddress()}");

            range.ApplyForEachCellOfType<string, string>(value => value.TrimSpaces());
        }

        public void RemoveAllSpaces(Range range)
        {
            logger.Info($"Trim spaces for range {range.GetRelativeAddress()}");

            range.ApplyForEachCellOfType<string, string>(value =>
            {
                if (string.IsNullOrWhiteSpace(value))
                    return null;

                return new string(value.Where(c => !char.IsWhiteSpace(c)).ToArray());
            });
        }
    }
}
