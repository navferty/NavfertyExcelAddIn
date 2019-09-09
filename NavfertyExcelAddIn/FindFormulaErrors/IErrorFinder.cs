﻿using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace NavfertyExcelAddIn.FindFormulaErrors
{
    public interface IErrorFinder
    {
        IEnumerable<ErroredRange> GetAllErrorCells(Range range);
    }
}
