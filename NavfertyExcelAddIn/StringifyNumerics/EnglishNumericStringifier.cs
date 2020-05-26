using System;
using System.Text;

namespace NavfertyExcelAddIn.StringifyNumerics
{
	public class EnglishNumericStringifier : INumericStringifier
	{
		public string StringifyNumber(double number)
		{
			var mainPart = Convert.ToInt32(number);

			var main = NumberToWords(mainPart);
			var fractional = GeconverttFractionalPart(number, mainPart);
			var result = main + fractional;
			return result.Trim();
		}

		private static string GeconverttFractionalPart(double number, int mainPart)
		{
			int fractionalPart = (int)Math.Round((number - mainPart) * 1000);
			if (fractionalPart == 0)
				return string.Empty;

			(int Multiplyer, string Word) power;

			if (fractionalPart % 100 == 0)
				power = (100, "tenths");
			else if (fractionalPart % 10 == 0)
				power = (10, "hundredths");
			else
				power = (1, "thousandths");

			return " and " + NumberToWords(fractionalPart / power.Multiplyer) + " " + power.Word;
		}

		private static string NumberToWords(int input)
		{
			if (input == 0)
				return "zero";

			var number = Math.Abs(input);

			var sb = new StringBuilder();
			if (input < 0)
				return "minus " + NumberToWords(number);

			if ((number / 1000000) > 0)
			{
				sb.Append(NumberToWords(number / 1000000) + " million ");
				number %= 1000000;
			}

			if ((number / 1000) > 0)
			{
				sb.Append(NumberToWords(number / 1000) + " thousand ");
				number %= 1000;
			}

			if ((number / 100) > 0)
			{
				sb.Append(NumberToWords(number / 100) + " hundred ");
				number %= 100;
			}

			if (number > 0)
			{
				if (sb.Length > 0)
					sb.Append("and ");

				var unitsMap = new[]
				{
					"zero", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten", "eleven",
					"twelve", "thirteen", "fourteen", "fifteen", "sixteen", "seventeen", "eighteen", "nineteen"
				};
				var tensMap = new[]
				{
					"zero", "ten", "twenty", "thirty", "forty",
					"fifty", "sixty", "seventy", "eighty", "ninety"
				};

				if (number < 20)
				{
					sb.Append(unitsMap[number]);
				}
				else
				{
					sb.Append(tensMap[number / 10]);
					if ((number % 10) > 0)
						sb.Append("-" + unitsMap[number % 10]);
				}
			}
			return sb.ToString();
		}
	}
}
