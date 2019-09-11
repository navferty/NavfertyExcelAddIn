using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;

namespace NavfertyExcelAddIn.DataValidation
{
    public interface ICellsValueValidator
    {
        IReadOnlyCollection<ValidationError> Validate(Range range, ValidationType validationType);
    }
}