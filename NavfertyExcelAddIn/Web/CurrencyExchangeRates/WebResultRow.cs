using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using NavfertyExcelAddIn.Localization;

namespace NavfertyExcelAddIn.Web.CurrencyExchangeRates
{
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
		public WebResultRow(
			DateTime date,
			string Vname,
			double Vnom,
			string sCurs,
			Int32 Vcode,
			string VchCode)
		{
			Name = Vname;
			ValidFrom = date;
			Units = Vnom;
			CursAsString = sCurs;
			Curs = Convert.ToDouble(CursAsString);
			Code = Vcode;
			ISOCode = VchCode;
			PriorityInGrid = uint.MaxValue;
		}


		public string FullNameWithUnits =>
			(Units == 1.0)
			? Name
			: (Name + " (" + string.Format(UIStrings.CurrencyExchangeRates_UnitsFormat, Units.ToString("N0")) + ")");


		public static int GetMaxDecimalDigitsCount(WebResultRow[] wrr)
		{
			var exchangeRatesDecimalDigitsCount = wrr
												.Select(wrr => wrr.CursAsString)
												.Select(s =>
												{
													var last = s.LastIndexOfAny(new[] { ',', '.' });
													if (last < 0) return 0;

													var cDecimalSeparator = s[last];
													var sDecimalPart = s.Split(new[] { cDecimalSeparator }).Last();
													return sDecimalPart.Length;
												}).Max();
			return exchangeRatesDecimalDigitsCount;
		}

	}
}
