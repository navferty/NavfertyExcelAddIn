using System;
using Microsoft.Office.Core;

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
            // TODO
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
