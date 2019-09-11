using System;
using System.Collections.Generic;
using NLog;
using Microsoft.Office.Interop.Excel;
using NavfertyExcelAddIn.Commons;
using NavfertyExcelAddIn.InteractiveRangeReport;

namespace NavfertyExcelAddIn.DataValidation
{
    public class CellsValueValidator : ICellsValueValidator
    {
        private readonly IValidatorFactory validatorFactory;
        private readonly Logger logger = LogManager.GetCurrentClassLogger();

        public CellsValueValidator(IValidatorFactory validatorFactory)
        {
            this.validatorFactory = validatorFactory;
        }

        public IReadOnlyCollection<InteractiveErrorItem> Validate(Range range, ValidationType validationType)
        {
			var validator = validatorFactory.CreateValidator(validationType);

            var errors = new List<InteractiveErrorItem>();

            range.ForEachCell(c => CheckCell(c, validator, errors, range.Worksheet.Name));

            throw new NotImplementedException();
        }

        private void CheckCell(Range cell, IValidator validator, ICollection<InteractiveErrorItem> errors, string wsName)
        {
            if (string.IsNullOrEmpty(cell.Value?.ToString()))
            {
                return;
            }

            // Value instead of Value2 can return datetime
            var value = (object)cell.Value;

            ValidationResult result = validator.CheckValue(value);

            if (result.IsSuccess)
            {
                return;
            }

            var error = new InteractiveErrorItem
            {
                Range = cell,
                Value = value.ToString(),
                ErrorMessage = result.Message,
                //Info = $"Invalid value '{cell.Value}', {result.Message}",
                Address = cell.Address[false, false],
                WorksheetName = wsName
            };

            errors.Add(error);
        }
    }
}
