using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using NavfertyExcelAddIn.Localization;

namespace NavfertyExcelAddIn.Web.CurrencyExchangeRates
{

	[TypeConverter(typeof(ExpandableObjectConverter))]
	internal class WebResultRow
	{
		public readonly string Name;
		public readonly DateTime ValidFrom;

		/// <summary>Number of units</summary>
		public readonly double Units;

		public readonly double Curs;
		public readonly string CursAsString;

		/// <summary>Internal bank code</summary>
		public readonly int Code;

		public readonly string ISOCode;

		public uint PriorityInGrid = uint.MaxValue;

		/// <summary>Contructor for CBRF record</summary>
		public WebResultRow(DataRow row, DateTime dt)
		{
			//Vname — Название валюты
			//Vnom — Номинал
			//Vcurs — Курс
			//Vcode — ISO Цифровой код валюты
			//VchCode — ISO Символьный код валюты

			Name = row.Field<string>("Vname").Trim();
			ISOCode = row.Field<string>("VchCode").Trim().ToUpper();
			Code = row.Field<int>("Vcode");

			Units = Convert.ToDouble(row.Field<decimal>("Vnom"));
			Curs = Convert.ToDouble(row.Field<decimal>("Vcurs"));
			CursAsString = row[2].ToString().Trim();

			ValidFrom = dt;
			PriorityInGrid = uint.MaxValue;
		}

		public WebResultRow(NBU.ExchangeRatesForDateRecord nbu)
		{
			Name = nbu.Name;

			ISOCode = nbu.ISOCode;
			Code = nbu.r030;

			CursAsString = nbu.RateString;
			var fi = (NumberFormatInfo)NumberFormatInfo.InvariantInfo.Clone();
			//fi.CurrencyDecimalSeparator = ".";
			fi.NumberDecimalSeparator = ".";
			fi.NumberGroupSeparator = "";
			fi.CurrencyGroupSeparator = "";
			var bParsed = double.TryParse(CursAsString, NumberStyles.Float, fi, out double parsedCurs);
			if (bParsed) Curs = parsedCurs;

			Units = 1;

			ValidFrom = DateTime.Now;
			PriorityInGrid = uint.MaxValue;
		}


		public string FullNameWithUnits =>
			(Units == 1.0)
			? Name
			: (Name + " (" + string.Format(UIStrings.CurrencyExchangeRates_UnitsFormat, Units.ToString("N0")) + ")");


		public static int GetMaxDecimalDigitsCount(WebResultRow[] wrr)
		{
			return 4;// some UA NBU records has 8 decimal numbers, and number formatted with this value looks bad!

			var exchangeRatesDecimalDigitsCount = wrr
												.Select(wrr => wrr.CursAsString)
												.Select(s =>
												{
													var last = s.LastIndexOfAny(new[] { ',', '.' });
													if (last < 0) return 0;

													var cDecimalSeparator = s[last];
													var numberParts = s.Split(new[] { cDecimalSeparator });
													var sDecimalPart = numberParts.Last();
													return sDecimalPart.Length;
												}).Max();


			return exchangeRatesDecimalDigitsCount;
		}
	}
}
