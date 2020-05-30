using System.Globalization;

using NavfertyExcelAddIn.Localization;

namespace NavfertyExcelAddIn.DataValidation.Validators
{
	public class NumericValidator : IValidator
	{
		public ValidationResult CheckValue(object cellValue)
		{
			if (cellValue is double || cellValue is decimal || cellValue is int) // TODO int stands for CVErr...
			{
				return ValidationResult.Success;
			}
			string value = cellValue.ToString();

			var isMatch = decimal.TryParse(value.Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out _);

			return isMatch
				? ValidationResult.Success
				: ValidationResult.Fail(ValidationMessages.ValueNotNumeric);
		}
	}
}
