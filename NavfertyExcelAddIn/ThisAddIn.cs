using System;
using System.Threading;

using Microsoft.Office.Core;

using NavfertyExcelAddIn.Commons;

using NLog;

namespace NavfertyExcelAddIn
{
	public partial class ThisAddIn
	{

		protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
		{
			return new NavfertyRibbon();
		}

		private void ThisAddInStartup(object sender, EventArgs e)
		{
			LogManager.GetCurrentClassLogger().Info(
				$"Add-in {GetType().Assembly.GetName().Version} started. Excel {Application.Version}, " +
				$"OS {Environment.OSVersion}, UI culture {Application.LanguageSettings.get_LanguageID(MsoAppLanguageID.msoLanguageIDUI)}");

			Thread.CurrentThread.CurrentUICulture =
				new System.Globalization.CultureInfo(
					Application.LanguageSettings.get_LanguageID(
						MsoAppLanguageID.msoLanguageIDUI));
			Thread.CurrentThread.CurrentCulture =
				new System.Globalization.CultureInfo(
					Application.LanguageSettings.get_LanguageID(
						MsoAppLanguageID.msoLanguageIDUI));
		}

		private void ThisAddInShutdown(object sender, EventArgs e)
		{
			LogSetup.Shutdown();
		}

		private void InternalStartup()
		{
			LogSetup.Initialize();
			Startup += new EventHandler(ThisAddInStartup);
			Shutdown += new EventHandler(ThisAddInShutdown);
		}
	}
}
