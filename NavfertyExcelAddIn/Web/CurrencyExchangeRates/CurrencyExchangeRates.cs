using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

using Microsoft.Office.Interop.Excel;

using NavfertyExcelAddIn.Commons;
using NavfertyExcelAddIn.Localization;

using Newtonsoft.Json;

using DataTable = System.Data.DataTable;

namespace NavfertyExcelAddIn.Web.CurrencyExchangeRates
{
	public class CurrencyExchangeRates : ICurrencyExchangeRates
	{
		internal readonly IDialogService dialogService;
		private Microsoft.Office.Interop.Excel.Application App => Globals.ThisAddIn.Application;

		public CurrencyExchangeRates(IDialogService dialogService)
			=> this.dialogService = dialogService;


		public void ShowCurrencyExchangeRates(Workbook wb)
		{
			if (App.Selection == null
				|| ((Range)App.Selection).Cells == null
				|| ((Range)App.Selection).Cells.Count != 1)
			{
				dialogService.ShowError(UIStrings.CurrencyExchangeRates_NedAnyCellSelection);
				return;
			}

			using (var f = new frmExchangeRates(this, wb))
			{
				if (f.ShowDialog() != DialogResult.OK) return;
			};
		}



		internal static async Task<WebResultRow[]> GetCurrencyExchabgeRates_CBRF(DateTime dt)
		{
			using (var cbr = new Web.CBR.DailyInfoSoapClient())
			{
				var dtsResult = await cbr.GetCursOnDateAsync(dt);
				if (dtsResult == null) throw new Exception("Failed to get remote data with no errors!");

				var dtFirst = dtsResult.Tables.Cast<DataTable>().FirstOrDefault();
				if (dtFirst == default) throw new Exception("Remote dstaset does not containt Tables!");
				var rows = dtFirst.RowsAsEnumerable().Select(row => new WebResultRow(row, dt)).ToArray();
				return rows;
			};
		}




		/// <summary>
		/// https://bank.gov.ua/ua/open-data/api-dev
		/// </summary>
		internal static async Task<WebResultRow[]> GetCurrencyExchabgeRates_NBU(DateTime dt)
		{
			//https://bank.gov.ua/NBUStatService/v1/statdirectory/exchange
			//https://bank.gov.ua/NBUStatService/v1/statdirectory/exchange?date=20200302&json

			string sDateForNBU = dt.ToString("yyyyMMdd");
			var urlNBUExchangeForDate = @$"https://bank.gov.ua/NBUStatService/v1/statdirectory/exchange?date={sDateForNBU}&json";
			//Debug.WriteLine(urlNBUExchangeForDate);
			/*
			using (WebClient wc = new WebClient())
			{
				string sJson = string.Empty;
				wc.Encoding = Encoding.UTF8;
				sJson = await wc.DownloadStringTaskAsync(urlNBUExchangeForDate);

			}			
			*/
			using (var htc = new HttpClient())
			{
				var sJson = await (await htc.GetAsync(urlNBUExchangeForDate)).
					EnsureSuccessStatusCode().
					Content.ReadAsStringAsync();

				var nbuResultRows = JsonConvert.DeserializeObject<NBU.ExchangeRatesForDateRecord[]>(sJson);
				var rows = nbuResultRows.Select(row => new WebResultRow(row)).ToArray();
				return rows;
			}
		}
	}
}
