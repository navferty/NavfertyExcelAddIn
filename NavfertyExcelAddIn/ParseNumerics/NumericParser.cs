using Microsoft.Office.Interop.Excel;
using NLog;
using Autofac.Extras.DynamicProxy;

namespace NavfertyExcelAddIn.ParseNumerics
{
    [Intercept(typeof(ExceptionLogger))]
    public class NumericParser : INumericParser
    {
        private readonly ILogger _logger = LogManager.GetCurrentClassLogger();

        public void Parse(Range selection)
        {
            // _logger.Info($"Start executing Parse for range {selection.Address}");

            var values = (object[,])selection.Value;

            int upperI = values.GetUpperBound(0); // Columns
            int upperJ = values.GetUpperBound(1); // Rows

            var isChanged = false;

            for (int i = values.GetLowerBound(0); i <= upperI; i++)
            {
                for (int j = values.GetLowerBound(1); j <= upperJ; j++)
                {
                    var value = values[i, j];
                    if (value is string s)
                    {
                        var parsedValue = s.ParseDecimal();
                        if (parsedValue != null)
                        {
                            isChanged = true;
                            values[i, j] = parsedValue;
                        }
                    }
                }
            }

            if (isChanged)
            {
                selection.Value = values;
            }
        }
    }
}
