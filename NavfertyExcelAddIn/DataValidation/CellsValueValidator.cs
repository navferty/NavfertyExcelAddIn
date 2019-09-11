using System;
using System.Collections.Generic;
using NLog;
using Microsoft.Office.Interop.Excel;
using NavfertyExcelAddIn.Commons;

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

        public IReadOnlyCollection<ValidationError> Validate(Range range, ValidationType validationType)
        {
			var validator = validatorFactory.CreateValidator(validationType);

            var errors = new List<ValidationError>();

            range.ForEachCell(c => CheckCell(c, validator, errors, range.Worksheet.Name));

            throw new NotImplementedException();
        }

        private void CheckCell(Range cell, IValidator validator, ICollection<ValidationError> errors, string wsName)
        {
            if (string.IsNullOrEmpty(cell.Value?.ToString()))
            {
                return;
            }

            // Value instead of Value2 can return datetime
            ValidationResult result = validator.CheckValue(cell.Value);

            if (result.IsSuccess)
            {
                return;
            }

            var error = new ValidationError
            {
                Range = cell,
                ErrorMessage = result.Message,
                Info = $"Invalid value '{cell.Value}', {result.Message}",
                WorksheetName = wsName
            };

            errors.Add(error);
        }
    }

    public enum ValidationType
    {
        Numeric,
        Xml,
        Date,
        TinPersonal,
        TinOrganization
    }
}