﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.Linq;

using Navferty.ExcelAddIn.Web.Localization;

#nullable enable

namespace Navferty.ExcelAddIn.Web.CurrencyExchangeRates
{

	[TypeConverter(typeof(ExpandableObjectConverter))]
	public class ExchangeRateRecord
	{
		public readonly string Name = string.Empty;
		public readonly string ISOCode = string.Empty;
		public readonly double Units = 1.0;
		public readonly double Curs = 0.0;

		internal readonly string CursAsString = "";
		/// <summary>Internal bank code</summary>
		internal readonly int Code = 0;

		internal uint PriorityInGrid = uint.MaxValue;
		internal readonly DateTime ValidFrom;

		private ExchangeRateRecord() : base()
		{
			ValidFrom = DateTime.Now;
			PriorityInGrid = uint.MaxValue;
			Units = 1;
		}
		/// <summary>Contructor for CBRF record</summary>
		internal ExchangeRateRecord(DataRow row, DateTime dt)
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
		}

		/// <summary>Constructor for NBU record</summary>
		internal ExchangeRateRecord(Providers.NBU.JsonExchangeRatesForDateRecord nbu) : this()
		{
			Name = nbu.Name;

			ISOCode = nbu.ISOCode;
			Code = nbu.R030;

			CursAsString = nbu.RateString;
			var fi = (NumberFormatInfo)NumberFormatInfo.InvariantInfo.Clone();
			fi.NumberDecimalSeparator = ".";
			fi.NumberGroupSeparator = "";
			fi.CurrencyGroupSeparator = "";
			var bParsed = double.TryParse(CursAsString, NumberStyles.Float, fi, out double parsedCurs);
			if (bParsed) Curs = parsedCurs;
		}

		/// <summary>Constructor for NBU record</summary>
		internal ExchangeRateRecord(Providers.ECB.ECBExchangeRatesRecord ecb) : this()
		{
			Name = ecb.Title;

			ISOCode = ecb.ISO;
			Code = 1;
			CursAsString = ecb.ObsValue;

			var fi = (NumberFormatInfo)NumberFormatInfo.InvariantInfo.Clone();
			fi.NumberDecimalSeparator = ".";
			fi.NumberGroupSeparator = "";
			fi.CurrencyGroupSeparator = "";
			var bParsed = double.TryParse(CursAsString, NumberStyles.Float, fi, out double parsedCurs);
			if (bParsed) Curs = (1 / parsedCurs);

			ValidFrom = DateTime.Now;
		}



		public string DisplayName =>
			(Units == 1.0)
			? Name
			: (Name + " (" + string.Format(UIStrings.CurrencyExchangeRates_UnitsFormat, Units.ToString("N0")) + ")");

		public double CursFor1Unit
		 =>
			(Units == 1.0 || Units == 0.0)
			? Curs
			: (Curs / Units);

		public static int GetMaxDecimalDigitsCount(ExchangeRateRecord[] wrr)
		{
			return 4;// some UA NBU records has 8 decimal numbers, and number formatted with this value looks bad!
			/*			 
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
			*/
		}

		internal static CurrencyExchangeRatesDataset.ExchangeRatesDataTable ToDataTable(IEnumerable<ExchangeRateRecord> webRows)
		{
			//Sort rows
			webRows = (from r in webRows
					   orderby r.PriorityInGrid ascending, r.Name ascending
					   select r);

			CurrencyExchangeRatesDataset.ExchangeRatesDataTable dtResult = new();
			webRows.ToList().ForEach(wrr =>
			{
				var newRow = dtResult.NewExchangeRatesRow();
				{
					newRow.Raw = wrr;
					newRow.Name = wrr.DisplayName;
					newRow.ISO = wrr.ISOCode;
					newRow.Rate = wrr.Curs;
				}
				dtResult.Rows.Add(newRow);
			});

			return dtResult;
		}
	}
}
