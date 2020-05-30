using System;
using System.Globalization;

using NavfertyExcelAddIn.Localization;

namespace NavfertyExcelAddIn.DataValidation.Validators
{
	public class DateValidator : IValidator
	{
		public ValidationResult CheckValue(dynamic cellValue)
		{
			if (cellValue is DateTime)
			{
				return ValidationResult.Success;
			}
			DateTime value;
			var s = cellValue.ToString();
			bool isMatch = DateTime.TryParse(s, DateTimeFormatInfo.InvariantInfo, DateTimeStyles.None, out value);
			if (!isMatch)
			{
				isMatch = DateTime.TryParseExact(s, "dd.MM.yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None, out value);
			}
			return isMatch ? ValidationResult.Success : ValidationResult.Fail(ValidationMessages.ValueCantBeConvertedToDate); // TODO localize
		}
	}
}
