using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using NavfertyExcelAddIn.Commons;
using NavfertyExcelAddIn.Localization;

namespace NavfertyExcelAddIn.Web.CurrencyExchangeRates
{
	public partial class frmExchangeRates : Form
	{
		private static readonly CultureInfo ciRU = CultureInfo.GetCultureInfo("ru-RU");
		private static readonly CultureInfo ciUA = CultureInfo.GetCultureInfo("uk-UA");
		//private static readonly CultureInfo ciUS = CultureInfo.GetCultureInfo("en-US");

		private static readonly Dictionary<string, uint> vipCurrencies = new()
		{
			{ "USD", 1u },
			{ "EUR", 2u },
			{ "GBP", 3u },
			{ "CNY", 4u }
		};
		private static readonly DataGridViewCellStyle cellStyle_ExchangeRate = new() { Alignment = DataGridViewContentAlignment.MiddleRight };
		private string columnCurrencyTitle = "Валюта";
		private string columnCurrencyCode = "Код";
		private string columnCurrencyRate = "Курс";
		private DataTable dtResult = null;
		private static int exchangeRatesDecimalDigitsCount = 2;
		private CultureInfo ciResult = ciRU;

		public frmExchangeRates()
		{
			InitializeComponent();
		}

		private void Form_Load(object sender, EventArgs e)
		{
			//Text = $"Курсы валют по данным ЦБРФ, по отношению к {ciResult.NumberFormat.CurrencySymbol}";
			Text = string.Format(UIStrings.CurrencyExchangeRates_FormTitle, UIStrings.CurrencyExchangeRates_Sources_CBRF, ciResult.NumberFormat.CurrencySymbol);
			var dtNow = DateTime.Now;
			dtpDate.Value = dtNow;
			dtpDate.MaxDate = dtNow;
			dtpDate.MinDate = dtNow.AddYears(-10);
		}

		private async void Form_Displayed(object sender, EventArgs e)
		{
			dtpDate.ValueChanged += DtpDate_ValueChanged;
			gridResult.CellFormatting += FormatCell_Rates;

			//txtFilter.Setcuebanner ("фильтр строк")
			await UpdateExchangeRates();
			txtFilter.AttachDelayedFilter(() => FilterResultInView());
		}

		private async void DtpDate_ValueChanged(object sender, EventArgs e)
		{
			await UpdateExchangeRates();
		}

		private async Task UpdateExchangeRates()
		{
			this.UseWaitCursor = true;
			try
			{
				dtResult = null;
				var dtDate = dtpDate.Value;

				using (var cbr = new Web.CBR.DailyInfoSoapClient())
				{
					var dtsResult = await cbr.GetCursOnDateAsync(dtDate);
					if (dtsResult == null) throw new Exception("Failed to get remote data with no errors!");

					var dtFirst = dtsResult.Tables.Cast<DataTable>().FirstOrDefault();
					if (dtFirst == default) throw new Exception("Dstaset does not containt Tables!");

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
								 let iPriority = vipCurrencies.TryGetValue(VchCode, out uint iPriorityFound) ? iPriorityFound : uint.MaxValue
								 let result = new { Vname, Vnom, sVcurs, Vcurs, Vcode, VchCode, iPriority }
								 orderby result.iPriority, result.Vname
								 select result).ToArray();

					var dtView = new DataTable();
					{
						DataColumn colName = new(columnCurrencyTitle, typeof(string));
						DataColumn colISO3 = new(columnCurrencyCode, typeof(string));
						DataColumn colExchangeRate = new(columnCurrencyRate, typeof(double));
						dtView.Columns.AddRange(new[] { colName, colISO3, colExchangeRate });
					}

					//Count max decimal digits length from all rows
					exchangeRatesDecimalDigitsCount = aData
									.Select(old => old.sVcurs)
									.Select(s =>
									{
										var last = s.LastIndexOfAny(new[] { ',', '.' });
										if (last < 0) return 0;

										var cDecimalSeparator = s[last];
										var sDecimalPart = s.Split(new[] { cDecimalSeparator }).Last();
										return sDecimalPart.Length;
									}).Max();

					foreach (var old in aData)
					{
						string sCurrency = old.Vname;
						if (old.Vnom != 1.0)
						{
							sCurrency += $" (за {old.Vnom})";
						}

						var newRow = dtView.NewRow();
						newRow.ItemArray = new object[] { sCurrency, old.VchCode, old.Vcurs };
						dtView.Rows.Add(newRow);
					}

					dtResult = dtView;
				};

				gridResult.DataSource = dtResult;
				if (dtResult == null) return;

				if (gridResult.Columns.Count > 0)
				{
					var colRate = gridResult.ColumnsAsEnumerable().Last();
					if (colRate.DefaultCellStyle != cellStyle_ExchangeRate) colRate.DefaultCellStyle = cellStyle_ExchangeRate;
				}
				//gridResult.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				gridResult.DataSource = null;
				//Commons.DialogService.ReferenceEquals
				MessageBox.Show(ex.Message);
			}
			finally
			{
				this.UseWaitCursor = false;
				this.Refresh();
			}
		}

		private void FilterResultInView()
		{
			var sFilter = txtFilter.Text.Trim();
			try
			{
				if (string.IsNullOrWhiteSpace(sFilter))
				{
					sFilter = string.Empty;
				}
				else
				{
					var columnNames = dtResult.ColumnsAsEnumerable()
						.Select(col =>
						{
							if (col.DataType == typeof(string))
								return $"[{col.ColumnName}] LIKE '%{sFilter}%'";

							return "";// $"[{col.ColumnName}] = '{sFilter}'"; ;
						})
						.Where(s => !string.IsNullOrWhiteSpace(s)).ToArray();

					sFilter = string.Join(" OR ", columnNames);
					//Debug.WriteLine("Row filter = " + sFilter);
					//= string.Format($"[{columnCurrencyTitle}] LIKE '%{0}%' OR [{columnCurrencyCode}] LIKE '%{0}%' OR [{columnCurrencyRate}] LIKE '%{0}%' ", sFilter);
				}
				dtResult.DefaultView.RowFilter = sFilter;
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Apply Row filter ('{sFilter}') ERROR!\n" + ex.Message);
			}
		}

		/// <summary>
		/// Format Rate column cells like number with thouthand separator
		/// </summary>
		private void FormatCell_Rates(object sender, DataGridViewCellFormattingEventArgs e)
		{
			if (e.Value == null || e.RowIndex == gridResult.NewRowIndex) return;

			if (e.ColumnIndex < (gridResult.ColumnCount - 1)) return;
			if (e.Value is not double dRate) return;

			e.Value = dRate.ToString($"C{exchangeRatesDecimalDigitsCount}", ciResult);
		}

		private void btnApplyResult_Click(object sender, EventArgs e)
		{
			//
		}
	}
}
