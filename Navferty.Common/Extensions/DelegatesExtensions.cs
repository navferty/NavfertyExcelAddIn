using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using NLog;

#nullable enable

namespace Navferty.Common
{
	[DebuggerStepThrough]
	public static class DelegatesExtensions
	{
		public static bool TryCatch(
			this Action a,
			bool displayErrorMessage = true,
			string? errorTitle = null,
			ILogger? logger = null,
			string? loggerTitle = null,
			bool allowErrorReporting = true
			)
		{
			try
			{
				a.Invoke();
				return true;
			}
			catch (Exception ex)
			{
				errorTitle ??= Application.ProductName;
				loggerTitle ??= errorTitle;

				if (displayErrorMessage)
				{
					ex.ShowErrorUI(errorTitle, logger, loggerTitle, allowErrorReporting);
					//MessageBox.Show(ex.Message,						errorTitle!,						MessageBoxButtons.OK,						MessageBoxIcon.Error);
				}
				else
				{
					logger?.Error(ex, loggerTitle);
				}
			}
			return false;
		}
	}
}
