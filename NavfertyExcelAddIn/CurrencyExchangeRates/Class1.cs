using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NavfertyExcelAddIn.CurrencyExchangeRates
{
	internal class GetCBR2
	{
		void DDD()
		{
			using (var cbr = new Web.CBR.DailyInfoSoapClient())
			{
				var curses = cbr.GetCursOnDate(DateTime.Now);
			};
		}

	}
}
