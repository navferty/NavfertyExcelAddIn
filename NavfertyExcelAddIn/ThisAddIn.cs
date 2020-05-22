using System;
using Microsoft.Office.Core;
using System.Threading;

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
            // TODO
        }

        private void InternalStartup()
        {
            Startup += new EventHandler(ThisAddInStartup);
            Shutdown += new EventHandler(ThisAddInShutdown);
        }
    }
}
