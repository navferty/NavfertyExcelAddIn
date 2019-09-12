using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using NavfertyExcelAddIn.Localization;

namespace NavfertyExcelAddIn.DataValidation.Validators
{
    public class TinOrganizationValidator : IValidator
    {
        public ValidationResult CheckValue(object cellValue)
        {
            // Алгоритм вычисления контрольных цифр для ИНН ФЛ и ЮЛ, см.:
            // https://ru.wikipedia.org/wiki/Идентификационный_номер_налогоплательщика

            string value = cellValue.ToString();

            if (!Regex.IsMatch(value.Trim(), @"^\d{10}$"))
            {
                return ValidationResult.Fail(string.Format(ValidationMessages.NotMathesPatternNDigits1, 10));
            }
            // check last digit as checksum
            var multipliers = new[] { 2, 4, 10, 3, 5, 9, 4, 6, 8 };
            var sum = multipliers
                .Select((t, i) => ParseChar(value[i]) * t)
                .Sum();
            // последняя цифра - остаток от последовательного деления на 11 и на 10 суммы остальных цифр,
            // взятых с соответствующими коэффициентами
            var isTin = sum % 11 % 10 == ParseChar(value[9]);
            return isTin
                ? ValidationResult.Success
                : ValidationResult.Fail(string.Format(ValidationMessages.NthDigitShouldBeAs4, 10, value[9], sum % 11 % 10, sum));
        }

        private static int ParseChar(char character)
        {
            return int.Parse(character.ToString(CultureInfo.InvariantCulture), NumberStyles.Integer, CultureInfo.InvariantCulture);
        }
    }
}
