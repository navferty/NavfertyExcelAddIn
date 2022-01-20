using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

using Microsoft.Office.Interop.Excel;

using NavfertyExcelAddIn.Commons;
using NavfertyExcelAddIn.Localization;

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
			var nbu = GetCurrencyExchabgeRates_NBU(dt);

			using (var cbr = new Web.CBR.DailyInfoSoapClient())
			{
				var dtsResult = await cbr.GetCursOnDateAsync(dt);
				if (dtsResult == null) throw new Exception("Failed to get remote data with no errors!");

				var dtFirst = dtsResult.Tables.Cast<DataTable>().FirstOrDefault();
				if (dtFirst == default) throw new Exception("Remote dstaset does not containt Tables!");

				//Vname — Название валюты
				//Vnom — Номинал
				//Vcurs — Курс
				//Vcode — ISO Цифровой код валюты
				//VchCode — ISO Символьный код валюты

				var cbrfRows = dtFirst.RowsAsEnumerable();
				var aData = (from oldRow in cbrfRows
							 let oldValues = oldRow.ItemArray
							 let Vname = oldValues[0].ToString().Trim()
							 let Vnom = Convert.ToDouble(oldValues[1])
							 let sVcurs = oldValues[2].ToString().Trim()
							 let Vcurs = Convert.ToDouble(sVcurs)
							 let Vcode = Convert.ToInt32(oldValues[3])
							 let VchCode = oldValues[4].ToString().Trim().ToUpper()
							 let result = new WebResultRow(dt, Vname, Vnom, sVcurs, Vcode, VchCode)
							 select result).ToArray();

				return aData;
			};
		}




		/// <summary>
		/// https://bank.gov.ua/ua/open-data/api-dev
		/// </summary>
		internal static WebResultRow[] GetCurrencyExchabgeRates_NBU(DateTime dt)
		{
			//https://bank.gov.ua/NBUStatService/v1/statdirectory/exchange
			//https://bank.gov.ua/NBUStatService/v1/statdirectory/exchange?date=20200302&json

			string sDateForNBU = dt.ToString("yyyymmdd");
			var urlNBUExchangeForDate = @$"https://bank.gov.ua/NBUStatService/v1/statdirectory/exchange?date={sDateForNBU}&json";
			Debug.WriteLine(urlNBUExchangeForDate);


			//System.Text.Json ddd;


			return new WebResultRow[] { };
		}
	}
}
