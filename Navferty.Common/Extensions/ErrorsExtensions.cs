using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Navferty.Common.Controls;

using NLog;

#nullable enable

namespace Navferty.Common
{
	[DebuggerStepThrough]
	public static class ErrorsExtensions
	{
		public static void ShowErrorUI(
			this Exception ex,
			string? errorTitle = null,
			ILogger? logger = null,
			string? loggerTitle = null,
			bool allowErrorReporting = true)
		{

			//errorTitle ??= Application.ProductName;
			loggerTitle ??= errorTitle ?? "[Error description empty]";
			logger ??= NLog.LogManager.GetCurrentClassLogger();
			logger?.Error(ex, loggerTitle);

			using var ef = new ErrorForm(ex, errorTitle, allowErrorReporting);
			ef.ShowDialog();
		}
	}
}
