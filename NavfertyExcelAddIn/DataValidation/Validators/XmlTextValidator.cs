using System.Text.RegularExpressions;

using NavfertyExcelAddIn.Localization;

namespace NavfertyExcelAddIn.DataValidation.Validators
{
	public class XmlTextValidator : IValidator
	{
		private const string Pattern = "^[^\u0000-\u001F\u0022\u0026\u0027\u003C\u003E\u007E-\u040F\u00B4\u0060]+$";

		public ValidationResult CheckValue(object cellValue)
		{
			string stringValue = cellValue.ToString();
			var isMatch = Regex.IsMatch(stringValue, Pattern);

			return isMatch
				? ValidationResult.Success
				: ValidationResult.Fail(ValidationMessages.ValueContainsInvalidChars);
		}
	}
}
