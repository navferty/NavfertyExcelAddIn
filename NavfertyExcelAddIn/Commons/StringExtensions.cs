using System.Text;
using System.Text.RegularExpressions;
using NLog;

namespace NavfertyExcelAddIn.Commons
{
    public static class StringExtensions
    {
        private static readonly ILogger logger = LogManager.GetCurrentClassLogger();
        private static readonly Regex spacesRegex = new Regex("\\s+", RegexOptions.None);

        public static string TrimSpaces(this string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                logger.Trace("Value is Null Or WhiteSpace => '<null>'");
                return null;
            }

            // replace any single or multiple space chars with single space
            var newValue = spacesRegex.Replace(value, " ");

            newValue = string.IsNullOrEmpty(newValue)
                ? null
                : newValue.Trim();

            logger.Debug($"'{value}' => '{newValue}'");

            return newValue;
        }
    }
}
