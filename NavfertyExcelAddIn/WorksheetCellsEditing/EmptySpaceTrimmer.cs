using Microsoft.Office.Interop.Excel;
using NavfertyExcelAddIn.Commons;
using NLog;

namespace NavfertyExcelAddIn.WorksheetCellsEditing
{
    public class EmptySpaceTrimmer : IEmptySpaceTrimmer
    {
        private static readonly ILogger logger = LogManager.GetCurrentClassLogger();

        public void TrimSpaces(Range range)
        {
            logger.Info($"Trim spaces for range {range.Address}");

            range.ApplyForEachCellOfType<string, string>(value => value.TrimSpaces());
        }
    }
}
