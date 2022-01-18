using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NavfertyExcelAddIn.CurrencyExchangeRates
{
	internal static class CBRWebServiceTest
	{
		static public void Test1()
		{
			using (var f = new frmExchangeRates())
			{
				f.ShowDialog();
			};
		}
	}
}
