using System;
using System.Collections.Generic;
using System.Linq;

namespace NavfertyExcelAddIn.StringifyNumerics
{
	public class RussianNumericStringifier : INumericStringifier
	{
		private Dictionary<int, string> allNumbers;

		public RussianNumericStringifier()
		{
			InitializeFirstThousand();
		}

		// thanks to pikabu.ru/@iakki for idea of algorythm
		public string StringifyNumber(double number)
		{
			if (number >= 1_000_000_000_000)
				return null;

			if (Math.Abs(number) < 0.001)
				return "ноль";

			if (number < 0)
				return "минус " + StringifyNumber(Math.Abs(number));

			var billionsNumber = (int)(number / 1_000_000_000);
			var billionsName = GetMultiplierName(billionsNumber, "миллиарда", "миллиардов", "миллиард");
			var billions = billionsNumber == 0
				? string.Empty
				: allNumbers[billionsNumber] + " " + billionsName + " ";

			var millionsNumber = (int)((number % 1_000_000_000) / 1_000_000);
			var millionsName = GetMultiplierName(millionsNumber, "миллиона", "миллионов", "миллион");
			var millions = millionsNumber == 0
				? string.Empty
				: allNumbers[millionsNumber] + " " + millionsName + " ";

			var thousandsNumber = (int)((number % 1_000_000) / 1_000);
			var thousandsName = GetMultiplierName(thousandsNumber, "тысячи", "тысяч", "тысяча");
			var thousands = thousandsNumber == 0
				? string.Empty
				: allNumbers[thousandsNumber] + " " + thousandsName + " ";
			thousands = FixThousands(thousands);

			int fractionalPart = (int)Math.Round((number - (long)number) * 1000);

			(int Multiplyer, string Word1, string Word2, string Word3) power;

			if (fractionalPart % 100 == 0)
				power = (100, "десятых", "десятых", "десятая");
			else if (fractionalPart % 10 == 0)
				power = (10, "сотых", "сотых", "сотая");
			else
				power = (1, "тысячных", "тысячных", "тысячная");

			fractionalPart /= power.Multiplyer;

			var fractionalName = GetMultiplierName(fractionalPart, power.Word1, power.Word2, power.Word3);
			var fractional = fractionalPart == 0
				? string.Empty
				: " и " + allNumbers[fractionalPart] + " " + fractionalName;

			var numbers = (int)(number % 1000);
			// TODO what about currency... can user select currency in addition?
			// var endingName = GetMultiplierName(numbers, "российских рубля", "российских рублей", "российский рубль");
			var endingName = GetMultiplierName(numbers, "целых", "целых", "целая");
			var ending = number != 0
				? allNumbers[numbers] + (fractionalPart != 0 ? " " + endingName : string.Empty)
				: string.Empty;
			ending = FixIfHasFraction(ending);
			if (numbers == 0 && fractionalPart == 0)
				ending = string.Empty;


			return (billions + millions + thousands + ending + fractional).Trim();
		}

		private string FixThousands(string thousands)
		{
			var result = thousands.Replace("один тысяча", "одна тысяча");
			result = result.Replace("два тысячи", "две тысячи");
			return result;
		}
		private string FixIfHasFraction(string number)
		{
			var result = number.Replace("один целая", "одна целая");
			result = result.Replace("два целых", "две целых");
			result = result.Replace("два тысячных", "две тысячных");
			result = result.Replace("один тысячная", "одна тысячная");
			return result;
		}

		private static string GetMultiplierName(int thousandsNumber, string t1, string t2, string t3)
		{
			string thousandsName;
			if (thousandsNumber % 10 == 2 || thousandsNumber % 10 == 3 || thousandsNumber % 10 == 4)
				thousandsName = t1;
			else
				thousandsName = t2;

			if (thousandsNumber % 10 == 1)
				thousandsName = t3;

			if (thousandsNumber % 100 > 4 && thousandsNumber % 100 < 21)
				thousandsName = t2;
			return thousandsName;
		}

		private void InitializeFirstThousand()
		{
			allNumbers = new Dictionary<int, string>
			{
				{ 0, "ноль" },
				{ 1, "один" },
				{ 2, "два" },
				{ 3, "три" },
				{ 4, "четыре" },
				{ 5, "пять" },
				{ 6, "шесть" },
				{ 7, "семь" },
				{ 8, "восемь" },
				{ 9, "девять" },
				{ 10, "десять" },
				{ 11, "одиннадцать" },
				{ 12, "двенадцать" },
				{ 13, "тринадцать" },
				{ 14, "четырнадцать" },
				{ 15, "пятнадцать" },
				{ 16, "шестнадцать" },
				{ 17, "семнадцать" },
				{ 18, "восемнадцать" },
				{ 19, "девятнадцать" },
			};

			var dozens = new Dictionary<int, string>
			{
				{ 2, "двадцать" },
				{ 3, "тридцать" },
				{ 4, "сорок" },
				{ 5, "пятьдесят" },
				{ 6, "шестьдесят" },
				{ 7, "семьдесят" },
				{ 8, "восемьдесят" },
				{ 9, "девяносто" }
			};

			var hundreds = new Dictionary<int, string>
			{
				{ 1, "сто" },
				{ 2, "двести" },
				{ 3, "триста" },
				{ 4, "четыреста" },
				{ 5, "пятьсот" },
				{ 6, "шестьсот" },
				{ 7, "семьсот" },
				{ 8, "восемьсот" },
				{ 9, "девятьсот" }
			};

			var firstHundred = Enumerable.Range(20, 80)
				.ToDictionary(i => i, i =>
				{
					//if (i < 20)
					//    return "";

					if (i % 10 == 0)
						return dozens[(int)(i / 10)]; // Если число делится на 10 без остатка - то не соединяем его с наименованием единицы
					else
						return dozens[(int)(i / 10)] + " " + allNumbers[i % 10]; // в противном случае соединяем название десятки и единицы
				});

			MergeDicts(allNumbers, firstHundred);

			var firstThousand = Enumerable.Range(100, 900)
				.ToDictionary(i => i, i =>
				{
					if (i % 100 == 0)
						return hundreds[(int)(i / 100)]; // Если число делится на 100 без остатка - то не соединяем его с наименованием единицы
					else
						return hundreds[(int)(i / 100)] + " " + allNumbers[i % 100]; // в противном случае соединяем название сотни и единицы
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
