using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Diagnostics.CodeAnalysis;

using Autofac;
using NavfertyExcelAddIn.ParseNumerics;
using Microsoft.Office.Core;

using NLog;
using Microsoft.Office.Interop.Excel;

using Application = Microsoft.Office.Interop.Excel.Application;

namespace NavfertyExcelAddIn
{
    [ComVisible(true)]
    [SuppressMessage("Style", "IDE0060:Remove unused parameter", Justification = "Ribbon callbacks must have certain signature")]
    public class NavfertyRibbon : IRibbonExtensibility
    {
        private static readonly IContainer container = Registry.CreateContainer();
        private readonly ILogger logger = LogManager.GetCurrentClassLogger();
        private Application App => Globals.ThisAddIn.Application;

        public NavfertyRibbon()
        {
            // TODO
        }

        #region IRibbonExtensibility
        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText();
        }
        #endregion

        #region Ribbon callbacks
        public void RibbonLoad(IRibbonUI ribbonUI)
        {
            // TODO
        }

        public void ParseNumerics(IRibbonControl ribbonControl)
        {
            var selection = (Range)App.Selection;

            if (selection == null)
                return;

            logger.Debug($"Parse numerics for range {selection.Address}");

            using (var scope = container.BeginLifetimeScope())
            {
                var parser = scope.Resolve<INumericParser>();
                parser.Parse(selection);
            }
        }

        public void UnprotectWorkbook(IRibbonControl ribbonControl)
        {

        }

        public Bitmap GetImage(string imageName)
        {
            return (Bitmap)RibbonImages.ResourceManager.GetObject(imageName);
        }
        #endregion

        #region Utils
        private static string GetResourceText()
        {
            var asm = typeof(NavfertyRibbon).Assembly;

            using (var stream = asm.GetManifestResourceStream("NavfertyExcelAddIn.NavfertyRibbon.xml"))
            using (var resourceReader = new StreamReader(stream))
            {
                return resourceReader.ReadToEnd();
            }
        }
        #endregion
    }
}
