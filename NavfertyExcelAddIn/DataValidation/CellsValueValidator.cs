using System.Collections.Generic;

using Microsoft.Office.Interop.Excel;

using NavfertyExcelAddIn.Commons;
using NavfertyExcelAddIn.InteractiveRangeReport;

namespace NavfertyExcelAddIn.DataValidation
{
	public class CellsValueValidator : ICellsValueValidator
	{
		private readonly IValidatorFactory validatorFactory;

		public CellsValueValidator(IValidatorFactory validatorFactory)
		{
			this.validatorFactory = validatorFactory;
		}

		public IReadOnlyCollection<InteractiveErrorItem> Validate(Range range, ValidationType validationType)
		{
			var validator = validatorFactory.CreateValidator(validationType);

			var errors = new List<InteractiveErrorItem>();

			range.ForEachCell(c => CheckCell(c, validator, errors, range.Worksheet.Name));

			return errors.ToArray();
		}

		private void CheckCell(Range cell, IValidator validator, ICollection<InteractiveErrorItem> errors, string wsName)
		{
			// Value instead of Value2 can return datetime
			var value = (object?)cell.Value;

			if (value is null || string.IsNullOrEmpty(value.ToString()))
			{
				return;
			}

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
				Address = cell.GetRelativeAddress(),
				WorksheetName = wsName
			};

			errors.Add(error);
		}
	}
}
