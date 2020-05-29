using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

using NavfertyExcelAddIn.Localization;

namespace NavfertyExcelAddIn.DataValidation.Validators
{
	public class TinPersonalValidator : IValidator
	{
		public ValidationResult CheckValue(object cellValue)
		{
			string value = cellValue.ToString();

			if (!Regex.IsMatch(value.Trim(), @"^\d{12}$"))
			{
				return ValidationResult.Fail(string.Format(ValidationMessages.NotMathesPatternNDigits1, 12));
			}
			// check last 2 digits as checksum
			var multipliers = new[] { 7, 2, 4, 10, 3, 5, 9, 4, 6, 8 };
			var sum = multipliers
				.Select((t, i) => ParseChar(value[i]) * t)
				.Sum();
			if (sum % 11 % 10 != ParseChar(value[10]))
			{
				return ValidationResult.Fail(string.Format(ValidationMessages.NthDigitShouldBeAs4, 11, value[10], sum % 11 % 10, sum));
			}
			multipliers = new[] { 3, 7, 2, 4, 10, 3, 5, 9, 4, 6, 8 };
			sum = multipliers
				.Select((t, i) => ParseChar(value[i]) * t)
				.Sum();

			var isTin = sum % 11 % 10 == ParseChar(value[11]);

			return isTin
				? ValidationResult.Success
				: ValidationResult.Fail(string.Format(ValidationMessages.NthDigitShouldBeAs4, 12, value[11], sum % 11 % 10, sum));
		}

		private static int ParseChar(char character)
		{
			return int.Parse(character.ToString(CultureInfo.InvariantCulture), NumberStyles.Integer, CultureInfo.InvariantCulture);
		}
	}
}
