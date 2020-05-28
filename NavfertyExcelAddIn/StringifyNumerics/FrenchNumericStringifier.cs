using System;
using System.Collections.Generic;
using System.Linq;

namespace NavfertyExcelAddIn.StringifyNumerics
{
	public class FrenchNumericStringifier : INumericStringifier
	{
		private Dictionary<int, string> allNumbers;

		public FrenchNumericStringifier()
		{
			InitializeFirstThousand();
		}

		// thanks to pikabu.ru/@iakki for idea of algorythm
		public string StringifyNumber(double number)
		{
			if (number == 0)
				return "zéro";

			if (number < 0)
				return "moins " + StringifyNumber(Math.Abs(number));

			var billionsNumber = (int)(number / 1_000_000_000);
			var billionsName = GetMultiplierName(billionsNumber, "milliard");
			var billions = billionsNumber == 0
				? string.Empty
				: allNumbers[billionsNumber] + " " + billionsName + " ";

			var millionsNumber = (int)((number % 1_000_000_000) / 1_000_000);
			var millionsName = GetMultiplierName(millionsNumber, "million");
			var millions = millionsNumber == 0
				? string.Empty
				: allNumbers[millionsNumber] + " " + millionsName + " ";

			var thousandsNumber = (int)((number % 1_000_000) / 1_000);
			var thousandsName = "mille";
			var thousands = thousandsNumber == 0
				? string.Empty
				: allNumbers[thousandsNumber] + " " + thousandsName + " ";
			thousands = FixThousands(thousands);

			int fractionalPart = (int)Math.Round((number - (long)number) * 1000);

			int zeroPaddingLeft;
			int zeroPaddingRight;

			if (fractionalPart > 100)
				zeroPaddingLeft = 0;
			else if (fractionalPart > 10)
				zeroPaddingLeft = 1;
			else
				zeroPaddingLeft = 2;

			if (fractionalPart % 100 == 0)
				zeroPaddingRight = 2;
			else if (fractionalPart % 10 == 0)
				zeroPaddingRight =1;
			else
				zeroPaddingRight = 0;

			fractionalPart /= (int)Math.Pow(10, zeroPaddingRight);

			var fractional = fractionalPart == 0
				? string.Empty
				: string.Concat(Enumerable.Repeat(" zéro", zeroPaddingLeft))
					+ " " + allNumbers[fractionalPart]; // 0,003 - zero virgule zero zero trois

			var numbers = (int)(number % 1000);

			var ending = number != 0
				?  allNumbers[numbers] + (fractionalPart != 0 ? " virgule" : string.Empty)
				: string.Empty;
			if (numbers == 0 && fractionalPart == 0)
				ending = string.Empty;


			return (billions + millions + thousands + ending + fractional).Trim();
		}

		private string FixThousands(string thousands)
		{
			var result = thousands.Replace("un mille", "mille");
			return result;
		}

		private static string GetMultiplierName(int thousandsNumber, string name)
		{
			if (thousandsNumber == 1)
				return name;
			return name + "s";
		}

		private void InitializeFirstThousand()
		{
			allNumbers = new Dictionary<int, string>
			{
				{ 0, "zéro" },
				{ 1, "un" },
				{ 2, "deux" },
				{ 3, "trois" },
				{ 4, "quatre" },
				{ 5, "cinq" },
				{ 6, "six" },
				{ 7, "sept" },
				{ 8, "huit" },
				{ 9, "neuf" },
				{ 10, "dix" },
				{ 11, "onze" },
				{ 12, "douze" },
				{ 13, "treize" },
				{ 14, "quatorze" },
				{ 15, "quinze" },
				{ 16, "seize" },
				{ 17, "dix-sept" },
				{ 18, "dix-huit" },
				{ 19, "dix-neuf" }
			};

			var dozens = new Dictionary<int, string>
			{
				{ 2, "vingt" },
				{ 3, "trente" },
				{ 4, "quarante" },
				{ 5, "cinquante" },
				{ 6, "soixante" }
			};

			var firstHundred = Enumerable.Range(20, 50) // 20-69
				.ToDictionary(i => i, i =>
				{
					if (i % 10 == 0)
						return dozens[(int)(i / 10)]; // trente (30)
					else if (i % 10 == 1)
						return dozens[(int)(i / 10)] + " et " + allNumbers[i % 10]; // trente et un (31)
					else
						return dozens[(int)(i / 10)] + "-" + allNumbers[i % 10]; // quarante-cinq (45)
				});
			MergeDicts(allNumbers, firstHundred);

			var soixante = dozens[6];
			var seventies = Enumerable.Range(70, 10) // 70-79
				.ToDictionary(i => i, i =>
				{
					if (i % 10 == 1)
						return soixante + " et " + allNumbers[10 + (i % 10)]; // soixante et onze (71)
					else
						return soixante + "-" + allNumbers[10 + (i % 10)]; // soixante-douze (72)
				});
			MergeDicts(allNumbers, seventies);

			allNumbers.Add(80, "quatre-vingts");

			const string quatreVingtSuffix = "quatre-vingt-";

			// quatre-vingt-un (81), quatre-vingt-deux (82), quatre-vingt-dix (90), quatre-vingt-onze (91)
			var eightyToHundred = Enumerable.Range(81, 19) // 80-99
				.ToDictionary(i => i, i => quatreVingtSuffix + allNumbers[i - 80]);

			MergeDicts(allNumbers, eightyToHundred);

			allNumbers.Add(100, "cent");

			var firstThousand = Enumerable.Range(101, 899) // 101-999
				.ToDictionary(i => i, i =>
				{
					if (i < 200)
						return allNumbers[100] + " " + allNumbers[i % 100]; // cent un (101)
					else if (i % 100 == 0)
						return allNumbers[(int)(i / 100)] + " " + allNumbers[100]; // deux cent (200)
					else
						return allNumbers[(int)(i / 100)]
							+ " " + allNumbers[100]
							+ " " + allNumbers[i % 100]; // deux cent un (201)
				});

			MergeDicts(allNumbers, firstThousand);
		}

		private void MergeDicts<TKey, TValue>(IDictionary<TKey, TValue> first, IDictionary<TKey, TValue> second)
		{
			foreach (var item in second)
			{
				first[item.Key] = item.Value;
			}
		}
	}
}
