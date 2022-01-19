using Microsoft.Office.Interop.Excel;

namespace NavfertyExcelAddIn.Web.CurrencyExchangeRates
{
	public interface ICurrencyExchangeRates
	{
		void ShowCurrencyExchangeRates(Workbook wb);
	}
}
