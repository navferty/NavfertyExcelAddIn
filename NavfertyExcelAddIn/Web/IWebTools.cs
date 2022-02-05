using Microsoft.Office.Interop.Excel;

namespace NavfertyExcelAddIn.Web
{
	public interface IWebTools
	{
		void CurrencyExchangeRates_Show(Workbook wb);
	}
}
