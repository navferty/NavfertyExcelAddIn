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
		/// <summary>Displays error message box to user, with ability to send feedback, and wites error to the Log</summary>
		/// <param name="errorTitle">Some description about error to display to the user</param>
		/// <param name="logger"></param>
		/// <param name="loggerTitle">Some description about error to write to Log. If null - use errorTitle</param>
		/// <param name="allowErrorReporting">Allow user to send bug report</param>
		public static void ShowErrorUI(
			this Exception ex,
			string? errorTitle = null,
			ILogger? logger = null,
			string? loggerTitle = null,
			bool allowErrorReporting = true)
		{
			loggerTitle ??= errorTitle ?? "[Error description empty]";
			logger ??= NLog.LogManager.GetCurrentClassLogger();
			logger?.Error(ex, loggerTitle);

			using var ef = new ErrorDialog(ex, errorTitle, allowErrorReporting);
			ef.ShowDialog();
		}
	}
}
