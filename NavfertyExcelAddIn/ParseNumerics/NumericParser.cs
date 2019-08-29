using System;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Linq;
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
            _logger.Info($"Start executing Parse for range {selection.Address}");

            foreach (Range cell in selection.Cells)
            {
                var value = cell.Value;
            }

            throw new NotImplementedException();
        }
    }
}
